Dim strFileScript : strFileScript = "libWMI.vbs"
'***********************************************************************
'File:     libWMI.vbs
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


'********************************************************************************
Private Function ExecQuery(ByRef colCol, ByVal objSource, ByVal strQuery)

  On Error Resume Next
  
  ExecQuery = False
  
  'Const wbemFlagReturnImmediately = &h10
  'Const wbemFlagForwardOnly = &h20

  'I had poor results with the following syntax getting all queries to return collection sets
  'Set colCol = objSource.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
  If Not IsObject(objSource) Then
    Call LogIt("  Could not run ExecQuery.  The source is not an object.", "ExecQuery", LogTypeError)
    Exit Function
  End If
  
  g_objError.Clear
  Set colCol = objSource.ExecQuery(strQuery)
  If (colCol Is Nothing) Then
    Call LogIt("  Could not run ExecQuery '" & strQuery & "'. " & g_objError.Message, "ExecQuery", LogTypeError)
  Else
    If (Not IsObject(colCol)) Or (g_objError.Check()) Then
      Call LogIt("  Could not run ExecQuery '" & strQuery & "'. " & g_objError.Message, "ExecQuery", LogTypeError)
    Else
      ExecQuery = True
    End If
  End If
End Function

'***********************************************************************
Private Function MOFComp(ByVal strFolder)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  MOFComp = False
  
  Dim objFolder
  Dim objFile
  Dim dicMofsFailedToCompile
  Dim strKey
  Dim strFileType
  
  If IsNullOrEmpty(strFolder) Then
    Exit Function
  End If
  
  If Not g_objFSO.FolderExists(strFolder) Then
    Exit Function
  End If
  
  Set objFolder = g_objFSO.GetFolder(strFolder)
  If Not IsObject(objFolder) Then
    
    Exit Function
  End If
  
  Set dicMofsFailedToCompile = New clsDictionary
  
  Call LogIt("  Attempting to compile .mof and .mfl files in the " & strFolder & " folder.", "MOFComp", LogTypeInfo + LogTypeVerbose)
  
  For Each strFileType In Array("mof", "mfl")
    For Each objFile In objFolder.Files
      If LCase(g_objFSO.GetExtensionName(objFile.Name)) = strFileType Then
        Call g_objProcess.Exec("mofcomp.exe " & Chr(34) & objFile.Path & Chr(34), 0)
        If StrIn(1, Join(g_objProcess.Output, "#!#"), "Done!") = 0 Then
          Call LogIt("  Failed to compile the " & objFile.Path & " file.", "MOFComp", LogTypeError)
          dicMofsFailedToCompile.Add objFile.Path, Empty
        Else
          Call LogIt("  Compiled the " & objFile.Path & " file.", "MOFComp", LogTypeInfo + LogTypeVerbose)
        End If
      End If
    Next
  Next
  
  Call LogIt("  Attempting to compile .mof and .mfl files that did not compile.", "MOFComp", LogTypeInfo + LogTypeVerbose)
  For Each strKey In dicMofsFailedToCompile.Keys
    If IsNullOrEmpty(Trim(strKey)) Then
      dicMofsFailedToCompile.Remove strKey
    Else
      Call g_objProcess.Exec("mofcomp.exe " & Chr(34) & strKey & Chr(34), 0)
      If StrIn(1, Join(g_objProcess.Output, "#!#"), "Done!") = 0 Then
        Call LogIt("  Failed to compile the " & strKey & " file.", "MOFComp", LogTypeError)
      Else
        Call LogIt("  Compiled the " & strKey & " file.", "MOFComp", LogTypeInfo + LogTypeVerbose)
        dicMofsFailedToCompile.Remove strKey
      End If
    End If
  Next
  
  If dicMofsFailedToCompile.Count > 0 Then
    Call LogIt("  Could not compile these mof or mfl files.", "MOFComp", LogTypeError)
    For Each strKey In dicMofsFailedToCompile.Keys
      Call LogIt("    " & strKey, "MOFComp", LogTypeError)
    Next
  Else
    MOFComp = True
  End If
  
  If IsObject(dicMofsFailedToCompile) Then Set dicMofsFailedToCompile = Nothing
  If IsObject(objFolder) Then Set objFolder = Nothing
End Function

'********************************************************************************
Public Function ProcessTime(ByVal strProcess, ByVal strDateInterval)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  ProcessTime = Null
  
  Dim colProcesses, objProcess, strQuery
  Dim strCreationDate
  
  If IsNumeric(strProcess) Then
    strQuery = "Select CreationDate from Win32_Process Where ID = '" & strProcess & "'"
  Else
    strQuery = "Select CreationDate from Win32_Process Where Name = '" & strProcess & "'"
  End If
  
  g_objError.Clear
  Set colProcesses = GetObject("Winmgmts:").ExecQuery(strQuery)
  If (Not IsObject(colProcesses)) Or g_objError.Check() Then
    Call LogIt("  Failed to query Win32_Process. " & g_objError.Message, "ProcessTime", LogTypeError)
  Else
    For Each objProcess in colProcesses
      g_objError.Clear
      strCreationDate = WMIDateStringToDate(objProcess.CreationDate)
      If g_objError.Check() Then
        Call LogIt("  Could not get the " & strProcess & " process CreationDate. " & g_objError.Message, "ProcessTime", LogTypeError)
        Exit For
      End If
    Next
  End If
  
  If IsObject(colProcesses) Then Set colProcesses = Nothing
  
  If Not IsNullOrEmpty(strCreationDate) Then ProcessTime = DateDiff(strDateInterval, strCreationDate, Now())
End Function

'***********************************************************************
Private Function WMIDateStringToDate(ByVal strWMIDateTime)
  
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  WMIDateStringToDate = Empty
  
  'This is the 'uncool' oldschool way
  WMIDateStringToDate = CDate(Mid(strWMIDateTime, 5, 2) & "/" & _
                              Mid(strWMIDateTime, 7, 2) & "/" & _
                              Left(strWMIDateTime, 4) & " " & _
                              Mid(strWMIDateTime, 9, 2) & ":" & _
                              Mid(strWMIDateTime, 11, 2) & ":" & _
                              Mid(strWMIDateTime, 13, 2))
End Function
