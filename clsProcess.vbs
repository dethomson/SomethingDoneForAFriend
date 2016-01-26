Dim strFileScript : strFileScript = "clsProcess.vbs"
'***********************************************************************
'File:     clsProcess.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 5/3/2014
'
'          This script file is based on the Configuration Manager Health Check Tool
'          (http://configmgrclienthtc.codeplex.com/)
'
'Notes:    
'
'Requires: 
'
'Solution version: 4.01.03
'
'Disclaimer:   As with many freeware tools, use of this script and its associated scripts is at
'          YOUR OWN RISK.  There is NO WARRANTY implied or otherwise to the form, fitness,
'          or function of these scripts.  Prior to using these scripts, they should be thoroughly
'          evaluated and tested in a lab environment.
'
' Unless otherwise noted at the procedure level, all code is copyright
' Dan Thomson.  You may modify, use, and share this code for non-profit
' activities only.
'***********************************************************************


'***********************************************************************
Class clsProcess
  
  Private m_strGeneralOutput
  Private m_strErrorOutput
  Private m_dblExitCode
  Private m_blnSilent
  Private m_blnOutputBlankLines
  
  '********************************************************************************
  Private Sub Class_Initialize()
    Silent = False
    OutputBlankLines = False
  End Sub
  
  '********************************************************************************
  Private Sub Class_Terminate()
  End Sub
   
  '***********************************************************************
  Public Function Output()
    Dim strOutput : strOutput = Empty
    Dim arrOutput
    
    If Not IsNullOrEmpty(m_strGeneralOutput) Then strOutput = m_strGeneralOutput
    If Not IsNullOrEmpty(m_strErrorOutput) Then
      If IsNullOrEmpty(strOutput) Then
        strOutput = m_strErrorOutput
      Else
        strOutput = strOutput & VbCrLf & m_strErrorOutput
      End If
    End If
    If StrIn(1, strOutput, VbCrLf) > 0 Then
      arrOutput = StrSplit(strOutput, VbCrLf, -1)
    Else
      arrOutput = Array(strOutput)
    End If
    
    Output = arrOutput
  End Function
  
  '***********************************************************************
  Public Function ExitCode()
    ExitCode = m_dblExitCode
  End Function
  
  '********************************************************************************
  'Should blank lines be included in the captured output?
  Public Property Let OutputBlankLines(ByVal blnValue)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    m_blnOutputBlankLines = blnValue
  End Property
  
  Public Property Get OutputBlankLines()
    OutputBlankLines = m_blnOutputBlankLines
  End Property
  
  '********************************************************************************
  'Sets logging off or on
  Public Property Let Silent(ByVal blnValue)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    m_blnSilent = blnValue
  End Property
  
  Public Property Get Silent()
    Silent = m_blnSilent
  End Property
  
  '***********************************************************************
  Public Function Exec(ByVal strCommand, ByVal strSuccessCodes)

    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Exec = False
    
    Dim objExec, strLine
    Dim dicSuccessCodes
    Dim arrSuccessCodes
    Dim strSuccessCode
    Dim blnOutputIt
    
    m_strGeneralOutput = Empty
    m_strErrorOutput = Empty
    m_dblExitCode = Empty
    
    Set dicSuccessCodes = New clsDictionary
    
    If StrIn(1, strSuccessCodes, ",") > 0 Then
      arrSuccessCodes = StrSplit(strSuccessCodes, ",", -1)
      For Each strSuccessCode in arrSuccessCodes
        Call dicSuccessCodes.Add(Trim(strSuccessCode), Empty)
      Next
    Else
      Call dicSuccessCodes.Add(Trim(strSuccessCodes), Empty)
    End If
    
    If Not Silent() Then Call LogIt("Preparing to exec '" & strCommand & "'", "Exec", LogTypeInfo)
    
    Do
      g_objError.Clear
      Set objExec = g_objWshShell.Exec(strCommand)
      If (Not IsObject(objExec)) Or (g_objError.Check()) Then
        Call LogIt("  Could not run command (" & strCommand & "). Error message = " & g_objError.Message, "Exec", LogTypeError)
        Exit Do
      End If
  
      Do While objExec.Status = 0
        strLine = Empty
        If Not objExec.StdOut.AtEndOfStream Then
          Do While Not objExec.StdOut.AtEndOfStream
            strLine = objExec.StdOut.ReadLine
            If IsNullOrEmpty(strLine) And (OutputBlankLines = False) Then
              blnOutputIt = False
            Else
              blnOutputIt = True
            End If
            If blnOutputIt Then
              If Not Silent() Then Call LogIt("  Output: " & strLine, "Exec", LogTypeInfo)  ' + LogTypeVerbose)
              If IsNullOrEmpty(m_strGeneralOutput) Then
                m_strGeneralOutput = strLine
              Else
                m_strGeneralOutput = m_strGeneralOutput & VbCrLf & strLine
              End If
            End If
          Loop
        End If
        If Not objExec.StdErr.AtEndOfStream Then
          Do While Not objExec.StdErr.AtEndOfStream
            strLine = objExec.StdErr.ReadLine
            If IsNullOrEmpty(strLine) And (OutputBlankLines = False) Then
              blnOutputIt = False
            Else
              blnOutputIt = True
            End If
            If blnOutputIt Then
              If Not Silent() Then Call LogIt("    Error: " & strLine, "Exec", LogTypeWarning)
              If IsNullOrEmpty(m_strErrorOutput) Then
                m_strErrorOutput = strLine
              Else
                m_strErrorOutput = m_strErrorOutput & VbCrLf & strLine
              End If
            End If
          Loop
        End If
        
        Wscript.Sleep 100
      Loop
      If Not Silent() Then Call LogIt("  Exec completed.", "Exec", LogTypeInfo)
      
      If dicSuccessCodes.Exists(CStr(objExec.ExitCode)) Then
        Exec = True
      Else
        If Not Silent() Then Call LogIt("    Unexpected exit code (" & objExec.ExitCode & ")", "Exec", LogTypeError)
      End If
      
      m_dblExitCode = objExec.ExitCode
      Exit Do
    Loop
    
    If IsObject(dicSuccessCodes) Then Set dicSuccessCodes = Nothing
    
  End Function

  '***********************************************************************
  '
  ' Function: IsRunning
  '
  ' Purpose:  Determine if a process is running
  '
  ' Input:    Name or PID of process
  '
  ' Output:   True or False depending on if the process is running
  '
  '***********************************************************************
  Public Function IsRunning(ByVal strProcess)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    IsRunning = False
    
    Dim strCommand
    Dim arrOutput, strOutput
    
    If IsNumeric(strProcess) Then
      strCommand = "TaskList /FO TABLE /NH /FI " & Chr(34) & "PID eq " & strProcess & Chr(34)
    Else
      strCommand = "TaskList /FO TABLE /NH /FI " & Chr(34) & "ImageName eq " & strProcess & Chr(34)
    End If
    
    If Exec(strCommand, 0) Then
      arrOutput = Output()
      For Each strOutput In arrOutput
        If StrIn(1, strOutput, strProcess) > 0 Then
          IsRunning = True
        End If
      Next
    Else
      Call LogIt("  Could not determine if the process is running.", "IsRunning", LogTypeError)
    End If
    
  End Function
  
  '***********************************************************************
  Public Function Run(ByVal strCommand, ByVal blnWait, ByVal varSuccessCodes)

    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Run = False
    
    Dim dicSuccessCodes
    Dim arrSuccessCodes
    Dim strSuccessCodes, strSuccessCode
    
    blnWait = CBool(blnWait)
    
    m_strGeneralOutput = Empty
    m_strErrorOutput = Empty
    m_dblExitCode = Empty
    
    Set dicSuccessCodes = New clsDictionary
    
    strSuccessCodes = CStr(varSuccessCodes)
    
    If StrIn(1, strSuccessCodes, ",") > 0 Then
      arrSuccessCodes = StrSplit(strSuccessCodes, ",", -1)
      For Each strSuccessCode in arrSuccessCodes
        Call dicSuccessCodes.Add(Trim(strSuccessCode), Empty)
      Next
    Else
      Call dicSuccessCodes.Add(Trim(strSuccessCodes), Empty)
    End If
    
    If Not Silent() Then Call LogIt("Preparing to run '" & strCommand & "'", "Run", LogTypeInfo)
    
    Do
      g_objError.Clear
      m_dblExitCode = g_objWshShell.Run(strCommand, 0, blnWait)
      If g_objError.Check() Then
        If Not Silent() Then Call LogIt("  Could not run command (" & strCommand & "). Exit code = " & m_dblExitCode & " Error message = " & g_objError.Message, "Run", LogTypeError)
        Exit Do
      End If
      If blnWait Then
        If dicSuccessCodes.Exists(CStr(m_dblExitCode)) Then
          If Not Silent() Then Call LogIt("  The command completed. Exit code = " & m_dblExitCode, "", LogTypeInfo)
          Run = True
        Else
          If Not Silent() Then Call LogIt("  The command completed with an unexpected exit code. Exit code = " & m_dblExitCode, "Run", LogTypeError)
        End If
      Else
        Run = True
      End If
      Exit Do
    Loop
    
    If IsObject(dicSuccessCodes) Then Set dicSuccessCodes = Nothing
    
  End Function

  '***********************************************************************
  '
  ' Function: Terminate
  '
  ' Purpose:  Terminates a process
  '
  ' Input:    Name or PID of process
  '
  ' Output:   True or False depending on if the process was terminated
  '
  '***********************************************************************
  Public Function Terminate(ByVal strProcess)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Terminate = False
    
    Dim strCommand
    Dim strOutput
    
    If Not IsRunning(strProcess) Then
      Terminate = True
      Exit Function
    End If
    
    If IsNumeric(strProcess) Then
      strCommand = "TaskKill /F /PID " & strProcess
    Else
      strCommand = "TaskKill /F /IM " & strProcess
    End If
    
    If Exec(strCommand, 0) Then
      strOutput = Output()
      If StrIn(1, strOutput, "Success") = 0 Then
        Call LogIt("  Terminated the " & strProcess & " process.", "Terminate", LogTypeInfo)
        Terminate = True
      Else
        Call LogIt("  Could not terminate the " & strProcess & " process.", "Terminate", LogTypeError)
      End If
    Else
      Call LogIt("  Could not terminate the " & strProcess & " process.", "Terminate", LogTypeError)
    End If
    
  End Function

  '***********************************************************************
  '
  ' Sub: Wait
  '
  ' Purpose:  Waits for a process
  '
  ' Input:    Name or ID of process of process
  '           Wait time in seconds before termination.
  '             -1 will cause the script to wait indefinitely
  '             0 terminates the process imediately
  '             Any value > 0 will cause the script to wait the specified amount
  '                 of time in seconds berfore terminating the process
  '
  ' Output:   None
  '
  '***********************************************************************
  Public Sub Wait(ByVal strProcess, ByVal intWaitTime, ByVal strMode)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    If IsRunning(strProcess) Then
      
      Dim intWait : intWait = 0
      Dim intPause : intPause = 5   'Wait 5 seconds between checks
      
      'Loop while the process is running
      Do While IsRunning(strProcess)
        'Check to see if specified number of seconds have passed before terminating
        'the process. If yes, then terminate the process
        If (w >= intWaitTime) AND (intWaitTime >= 0) Then
          Call Terminate(strProcess)
          Exit Do
        End If
        
        'Increment the seconds counter
        intWait = intWait + intPause
        
        'Pause
        Wscript.Sleep(intPause * 1000)
      Loop
    End If
  End Sub
End Class
