Dim strFileScript : strFileScript = "clsLogging.vbs"
'***********************************************************************
'File:     clsLogging.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 5/3/2014
'
'          This script file is based on the Configuration Manager Health Check Tool
'          (http://configmgrclienthtc.codeplex.com/)
'
'Notes:    
'
'Requires: g_objFSO
'          g_objError
'          g_objLog
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


'Format accepts values of FLAT or SMS

Const LogTypeInfo             = 1
Const LogTypeWarning          = 2
Const LogTypeError            = 3
Const LogTypeVerbose          = 4

'*******************************************************************
Class clsLogging

  Private m_strFile
  Private m_strComponent
  Private m_strContext
  Private m_objFile
  Private m_blnIsLogOpen
  Private m_intWriteCounter
  Private m_intSize_Max
  Private m_strFile_Path
  Private m_strFile_Name
  Private m_strFile_Extension
  Private m_strFormat
  Private m_blnVerbose
  Private m_blnQuiet
  Private m_arrCachedWrites()
  Private m_blnLoggingIncludesComponentName
    
  '*******************************************************************
  Private Sub Class_Initialize()
    m_strComponent           = Empty
    m_strContext             = Empty
    m_blnIsLogOpen           = False
    m_blnVerbose             = True
    m_blnQuiet               = False
    m_intWriteCounter        = 0
    m_strFormat              = "FLAT"   'Can be FLAT,SMS,SCCM
    m_blnLoggingIncludesComponentName = True
    
    ReDim m_arrCachedWrites(0)
    
    Call Size_Max_Convert(5)     '5 MB to start as default
  End Sub
  
  '*******************************************************************
  Private Sub Class_Terminate()
    On Error Resume Next
    
    If m_blnIsLogOpen Then m_objFile.Close
    If IsObject(m_objFile) Then Set m_objFile = Nothing
  End Sub
  
  '*******************************************************************
  Public Property Let Component(strValue)
    m_strComponent = Trim(strValue)
  End Property
  
  '*******************************************************************
  Public Property Get File()
    File = m_strFile
  End Property
  
  '*******************************************************************
  Public Property Let File(strValue)
    m_strFile = Trim(strValue)
    
    m_strFile_Path = g_objFSO.GetParentFolderName(m_strFile)
    m_strFile_Name = g_objFSO.GetBaseName(m_strFile)
    m_strFile_Extension = g_objFSO.GetExtensionName(m_strFile)
    
    If IsNullOrEmpty(m_strFile_Extension) Then
      m_strFile_Extension = "log"
      m_strFile = m_strFile & ".log"
    End If
  End Property
  
  '*******************************************************************
  Public Property Get Format()
    Format = m_strFormat
  End Property
  
  '*******************************************************************
  Public Property Let Format(strValue)
    If StrIn(1, "FLAT,SMS,SCCM", UCase(strValue)) > 0 Then
      m_strFormat = UCase(strValue)
    Else
      m_strFormat = "FLAT"
    End If
  End Property
  
  '*******************************************************************
  Public Property Get Verbose()
    Verbose = m_blnVerbose
  End Property
  
  '*******************************************************************
  Public Property Let Verbose(blnValue)
    m_blnVerbose = blnValue
  End Property
  
  '*******************************************************************
  Public Property Get Quiet()
    Quiet = m_blnQuiet
  End Property
  
  '*******************************************************************
  Public Property Let Quiet(blnValue)
    m_blnQuiet = blnValue
  End Property
  
  '*******************************************************************
  Public Property Get LoggingIncludesComponentName()
    LoggingIncludesComponentName = m_blnLoggingIncludesComponentName
  End Property
  
  '*******************************************************************
  Public Property Let LoggingIncludesComponentName(blnValue)
    m_blnLoggingIncludesComponentName = blnValue
  End Property
  
  '*******************************************************************
  Public Property Let Size_Max(strValue)
    Call Size_Max_Convert(strValue)
  End Property
  
  '*******************************************************************
  Private Sub Size_Max_Convert(strValue)
    'Let's limit the max log size to 50MB
    If CInt(strValue) > 50 Then Exit Sub
    
    'Convert specified max log size (MB) to bytes
    m_intSize_Max = CInt(strValue) * 1048576
  End Sub
  
  '*******************************************************************
  ' Sub Write
  '
  ' Purpose:  Preps for writing output to a file.
  '
  ' Input:    strLogMsg - The text to write out.
  '           intType - The type of message to write.
  '                     Valid types are:
  '                       LogTypeInfo             = 1
  '                       LogTypeWarning          = 2
  '                       LogTypeError            = 3
  '                       LogTypeVerbose          = 4
  '
  '*******************************************************************
  Public Sub Write(ByVal strLogMsg, ByVal intType, ByVal strComponent)
    
    If m_blnQuiet Then Exit Sub
    
    'Make sure the log file is open
    Call Create()
    
    If m_blnIsLogOpen = False Then
      Dim intUBound
      intUBound = UBound(m_arrCachedWrites)
      m_arrCachedWrites(intUBound) = strLogMsg & "#!#" & intType & "#!#" & strComponent
      ReDim Preserve m_arrCachedWrites(intUBound + 1)
      Exit Sub
    End If
    
    If intType > 3 Then
      If Not Verbose Then Exit Sub
      intType = intType - LogTypeVerbose
    End If
    
    If Not IsNullOrEmpty(m_arrCachedWrites) Then
      ReDim Preserve m_arrCachedWrites(UBound(m_arrCachedWrites) - 1)
      Dim strEntry, strLM, intT, strC
      For Each strEntry In m_arrCachedWrites
        Dim arrEntry
        arrEntry = StrSplit(strEntry, "#!#", -1)
        strLM = arrEntry(0)
        intT = CInt(arrEntry(1))
        strC = arrEntry(2)
        If m_strFormat = "FLAT" Then
          Call WriteEntryFlat(strLM, intT, strC)
        Else
          Call WriteEntryFormatted(strLM, intT, strC)
        End If
      Next
      Erase m_arrCachedWrites
    End If
    
    If m_strFormat = "FLAT" Then
      Call WriteEntryFlat(strLogMsg, intType, strComponent)
    Else
      Call WriteEntryFormatted(strLogMsg, intType, strComponent)
    End If
    
    m_intWriteCounter = m_intWriteCounter + 1
    
    If m_intWriteCounter > 49 Then
      m_intWriteCounter = 0
      Call Backup()
    End If
  End Sub
  
  '*******************************************************************
  Private Sub WriteEntryFlat(ByVal strLogMsg, ByVal intType, ByVal strComponent)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim strTime, strDate, strTempMsg, strType
    
    'Echo the message if running via Cscript
    If g_dicSettings.Key("IsCscript") Then Wscript.Echo strLogMsg
    
    If intType > 3 Then intType = intType - LogTypeVerbose
    
    Select Case intType
      Case LogTypeInfo    : strType = "Info"
      Case LogTypeWarning : strType = "Warn"
      Case LogTypeError   : strType = "Error"
    End Select
    
    If m_blnLoggingIncludesComponentName Then
      strComponent = "(" & strComponent & ")"
    Else
      strComponent = Empty
    End If
    
    ' The WriteLine has the potential to cause a runtime error.
    ' However, we must not stop operation if there is a failure, so always continue.
    On Error Resume Next
    
    'Create the log entry
    If m_blnIsLogOpen Then
      g_objError.Clear
      m_objFile.WriteLine Now() & "(" & strType & ")" & strComponent & " " & strLogMsg
      If g_objError.Check() Then
        If dicSettings("IsCscript") Then Wscript.Echo "Could not write to the '" & m_strFile & "' log file. " & g_objError.Message
      End If
    End If
  End Sub
  
  '*******************************************************************
  Private Sub WriteEntryFormatted(ByVal strLogMsg, ByVal intType, ByVal strComponent)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim strTime, strDate, strTempMsg
    
    'Echo the message if running via Cscript
    If g_dicSettings.Key("IsCscript") Then Wscript.Echo strLogMsg
    
    ' Each of the operations below has the potential to cause a runtime error.
    ' However, we must not stop operation if there is a failure, so always continue.
    On Error Resume Next
    
    ' Populate the variables to log
    strTime = Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & ".000+000"
    strDate = Right("0"& Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "-" & Year(Now)
    strTempMsg = "<![LOG[" & strLogMsg & "]LOG]!><time=""" & strTime & """ date=""" & strDate & """ component=""" & strComponent & """ context=""" & m_strContext & """ type=""" & intType & """ thread="""" file=""" & Wscript.ScriptName & """>"
    
    'Create the log entry
    If m_blnIsLogOpen Then
      g_objError.Clear
      m_objFile.WriteLine strTempMsg
      If g_objError.Check() Then
        If dicSettings("IsCscript") Then Wscript.Echo "Could not write to the '" & m_strFile & "' log file. " & g_objError.Message
      End If
    End If
  End Sub
  
  '***********************************************************************
  Private Function Create()
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Create = False
    
    If IsNullOrEmpty(m_strFile) Then Exit Function
    
    Call Backup()
    
    If m_blnIsLogOpen = False Then
      g_objError.Clear
      Set m_objFile = g_objFSO.OpenTextFile(m_strFile, ForAppending, True)
      If (Not IsObject(m_objFile)) Or g_objError.Check() Then
        If g_dicSettings.Key("IsCscript") Then Wscript.Echo g_objError.Message
        'Call LogIt("  Could not instantiate an object reference to new log file " & m_strFile & ". " & g_objError.Message, "Create", LogTypeError)
      Else
        m_blnIsLogOpen = True
        Create = True
      End If
    End If
  End Function
  
  '***********************************************************************
  Private Sub Backup()
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim strFile
    Dim lngSize
    Dim objFile
    
    'Check to see if file is there. If it is oversized, back it up to .??_
    If g_objFSO.FileExists(m_strFile) Then
      g_objError.Clear
      Set objFile = g_objFSO.GetFile(m_strFile)
      If (Not IsObject(objFile)) Or g_objError.Check() Then
        Call WriteEntryFormatted("  Could not instantiate an object reference to " & m_strFile & ". " & g_objError.Message, "Backup", LogTypeError)
        Exit Sub
      End If
      lngSize = objFile.Size
      Set objFile = Nothing
      
      If lngSize > m_intSize_Max Then
        strFile = m_strFile_Path & "\" & m_strFile_Name & "." & Left(m_strFile_Extension, Len(m_strFile_Extension) - 1) & "_"
        
        If m_blnIsLogOpen Then
          Call WriteEntryFormatted(">>>>>>>> Max log size reached <<<<<<<<", "Backup", LogTypeInfo)
        End If
        
        If g_objFSO.FileExists(strFile) Then
          Call File_Delete(strFile)
        End If
        
        If m_blnIsLogOpen Then
          If IsObject(m_objFile) Then
            m_objFile.Close
            Set m_objFile = Nothing
          End If
          m_blnIsLogOpen = False
        End If
        
        g_objError.Clear
        g_objFSO.MoveFile m_strFile, strFile
        If g_objError.Check() And (g_dicSettings.Key("IsCscript")) Then
          Wscript.Echo "Could not backup the log file '" & m_strFile & "' to '" & strFile & "'. " & g_objError.Message
        End If
      End If
    End If
  End Sub
End Class
