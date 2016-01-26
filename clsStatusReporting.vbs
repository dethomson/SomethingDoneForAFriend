Dim strFileScript : strFileScript = "clsStatusReporting.vbs"
'***********************************************************************
'File:     clsStatusReporting.vbs
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
Class clsStatusReporting
  
  Private m_dicStatus
  Private m_strFile
  Private m_blnReportStatus
  Private m_blnReportStatusError
  Private m_blnReportStatusInfo
  Private m_blnReportStatusHealthy
  
  '***********************************************************************
  Private Sub Class_Initialize()
    Set m_dicStatus = New clsDictionary
  End Sub
  
  '***********************************************************************
  Private Sub Class_Terminate()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    If IsObject(m_dicStatus) Then Set m_dicStatus = Nothing
  End Sub

  '***********************************************************************
  Public Property Get StatusCodes()
    StatusCodes = m_dicStatus.Keys
  End Property

  '***********************************************************************
  Public Property Get File()
    File = m_strFile
  End Property
  
  '***********************************************************************
  Public Property Let File(ByVal strValue)
    m_strFile = Trim(strValue)
  End Property

  '***********************************************************************
  Public Property Get ReportStatus()
    ReportStatus = m_blnReportStatus
  End Property
  
  '***********************************************************************
  Public Property Let ReportStatus(ByVal strValue)
    m_blnReportStatus = CBool(Trim(strValue))
  End Property

  '***********************************************************************
  Public Property Get ReportStatusError()
    ReportStatusError = m_blnReportStatusError
  End Property
  
  '***********************************************************************
  Public Property Let ReportStatusError(ByVal strValue)
    m_blnReportStatusError = CBool(Trim(strValue))
  End Property

  '***********************************************************************
  Public Property Get ReportStatusInfo()
    ReportStatusInfo = m_blnReportStatusInfo
  End Property
  
  '***********************************************************************
  Public Property Let ReportStatusInfo(ByVal strValue)
    m_blnReportStatusInfo = CBool(Trim(strValue))
  End Property

  '***********************************************************************
  Public Property Get ReportStatusHealthy()
    ReportStatusHealthy = m_blnReportStatusHealthy
  End Property
  
  '***********************************************************************
  Public Property Let ReportStatusHealthy(ByVal strValue)
    m_blnReportStatusHealthy = CBool(Trim(strValue))
  End Property
  
  '***********************************************************************
  Public Sub Add(ByVal strType, ByVal intCode)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    If Not IsNumeric(intCode) Then Exit Sub
    
    If Not IsObject(m_dicStatus) Then Exit Sub
  
    If (intCode > errSMS_StartStatusDefinitions) And (intCode < infSuccessfulCompletion) Then
      If HasSystemStatus Then
        Call LogIt("  One or more local administrator status codes exist in the status code list. Skip adding this status code (" & intCode & ").", "Add", LogTypeWarning)
        Exit Sub
      End If
    End If
    
    If m_dicStatus.Exists(intCode) Then Exit Sub
    
    Call LogIt("  Adding " & strType & " status code " & intCode & " to the status code list.", "Add", LogTypeWarning)
    
    m_dicStatus.Add intCode, strType
  End Sub
  
  '***********************************************************************
  Public Sub Add_Forced(ByVal strType, ByVal intCode)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    If Not IsNumeric(intCode) Then Exit Sub
    
    If Not IsObject(m_dicStatus) Then Exit Sub
  
    If m_dicStatus.Exists(intCode) Then Exit Sub
    
    Call LogIt("  Adding " & strType & " status code " & intCode & " to the status code list.", "Add_Forced", LogTypeWarning)
    
    m_dicStatus.Add intCode, strType
  End Sub
  
  '***********************************************************************
  Public Function HasSystemStatus()
    Dim strKey
    
    HasSystemStatus = False
    
    If m_dicStatus.Count > 0 Then
      For Each strKey In m_dicStatus.Keys
        If CLng(strKey) < errSMS_StartStatusDefinitions Then
          HasSystemStatus = True
          Exit For
        End If
      Next
    End If
  End Function
  
  '***********************************************************************
  Public Function HasInfoStatus()
    Dim strKey
    
    HasInfoStatus = False
    
    If m_dicStatus.Count > 0 Then
      For Each strKey In m_dicStatus.Keys
        If CLng(strKey) > infSuccessfulCompletion Then
          HasInfoStatus = True
          Exit For
        End If
      Next
    End If
  End Function
  
  '***********************************************************************
  Public Function HasStatusCodes(ByVal varCodes)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim arrCodes, strCodes, strCode, intCode
    
    HasStatusCodes = False
    
    If Not IsObject(m_dicStatus) Then Exit Function
    
    strCodes = CStr(varCodes)
    
    If StrIn(1, strCodes, ",") = 0 Then
      arrCodes = Array(strCodes)
    Else
      arrCodes = StrSplit(strCodes, ",", -1)
    End If
    For Each strCode In arrCodes
      intCode = CInt(Trim(strCode))
      If m_dicStatus.Exists(intCode) Then
        HasStatusCodes = True
        Exit For
      End If
    Next
  End Function
  
  '***********************************************************************
  Public Function HasStatus()
    If m_dicStatus.Count > 0 Then HasStatus = True Else HasStatus = False
  End Function
  
  '***********************************************************************
  Public Function HasStatusType(ByVal strTypes)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim arrTypes, strType
    
    HasStatusType = False
    
    If Not IsObject(m_dicStatus) Then Exit Function
    
    If StrIn(1, strTypes, ",") = 0 Then
      arrTypes = Array(strTypes)
    Else
      arrTypes = StrSplit(strTypes, ",", -1)
    End If
    For Each strType In arrTypes
      If m_dicStatus.Item_Exists(Trim(strType)) Then
        HasStatusType = True
        Exit For
      End If
    Next
  End Function
  
  '***********************************************************************
  Public Function IsHealthy()
    Dim strKey
    
    IsHealthy = False
    
    If m_dicStatus.Count = 1 Then
      For Each strKey In m_dicStatus.Keys
        If CLng(strKey) = infSuccessfulCompletion Then IsHealthy = True
      Next
    End If
  End Function
  
  '***********************************************************************
  Public Sub RemoveAll()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    If Not IsObject(m_dicStatus) Then Exit Sub
    
    m_dicStatus.RemoveAll
    
    Call LogIt("  Removed all status codes from the status code list.", "RemoveAll", LogTypeInfo + LogTypeVerbose)
  End Sub
  
  '***********************************************************************
  Public Sub RemoveStatusCodes(ByVal varCodes)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim arrCodes, strCodes, strCode, intCode
    
    If Not IsObject(m_dicStatus) Then Exit Sub
    
    Call LogIt("Removing " & varCodes & " from the status code list.", "RemoveStatusCodes", LogTypeInfo + LogTypeVerbose)
    
    strCodes = CStr(varCodes)
    
    If StrIn(1, strCodes, ",") = 0 Then
      arrCodes = Array(strCodes)
    Else
      arrCodes = StrSplit(strCodes, ",", -1)
    End If
    For Each strCode In arrCodes
      On Error Resume Next
      intCode = CInt(Trim(strCode))
      If TypeName(intCode) = "Numeric" Then
        If m_dicStatus.Exists(intCode) Then
          If m_dicStatus.Remove(intCode) Then
            Call LogIt("  Removed status code " & intCode & " from the status code list.", "RemoveStatusCode", LogTypeInfo + LogTypeVerbose)
          End If
        End If
      Else
        Call LogIt("  Procedure was provided a non-numeric status code to delete: " & strCode, "RemoveStatusCode", LogTypeError)
      End If
    Next
  End Sub
  
  '***********************************************************************
  Public Sub RemoveStatusTypes(ByVal strTypes)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim arrTypes, strType
    Dim strKey, strItem
    
    If Not IsObject(m_dicStatus) Then Exit Sub
    
    Call LogIt("Removing " & strTypes & " from the status code list.", "RemoveStatusTypes", LogTypeInfo + LogTypeVerbose)
    
    If StrIn(1, strTypes, ",") = 0 Then
      arrTypes = Array(strTypes)
    Else
      arrTypes = StrSplit(strTypes, ",", -1)
    End If
    For Each strType In arrTypes
      strType = Trim(strType)
      For Each strKey In m_dicStatus.Keys
        strItem = Trim(m_dicStatus.Item(strKey))
        If StrCompare(strType, strItem) = 0  Then
          If m_dicStatus.Item_Remove(strItem) Then
            Call LogIt("  Removed " & strItem & " status code " & strKey & " from the status code list.", "RemoveStatusTypes", LogTypeInfo + LogTypeVerbose)
          End If
        End If
      Next
    Next
  End Sub
  
  '***********************************************************************
  Public Sub WriteToFile()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim objFile
    Dim blnSuccess
    Dim strKey
    
    blnSuccess = False
  
    'Create the error file
    g_objError.Clear
    Set objFile = g_objFSO.CreateTextFile(m_strFile, True)
    If (Not IsObject(objFile)) Or (g_objError.Check()) Then
      Call LogIt("  Could not instantiate an object reference to " & strFile & ". " & g_objError.Message, "WriteToFile", LogTypeError)
    Else
      blnSuccess = True
    End If
    
    If blnSuccess Then
      objFile.WriteLine "name=" & g_objOS.Name
      objFile.WriteLine "domain=" & g_objOS.Domain
      objFile.WriteLine "ou=" & g_objOS.OU
      objFile.WriteLine "version=" & g_dicSettings.Key("ScriptVersion")
      For Each strKey In m_dicStatus.Keys
Call LogIt(">>>>>> " & strKey, "WriteToFile", LogTypeWarning)
        objFile.WriteLine(strKey & "|" & Now)
      Next
      Call LogIt("  Wrote status to " & m_strFile & ".", "WriteToFile", LogTypeInfo + LogTypeVerbose)
      objFile.Close
      Set objFile = Nothing
    Else
      Call LogIt("  Could not write status to " & m_strFile & ".", "WriteToFile", LogTypeError)
    End If
  End Sub
End Class
