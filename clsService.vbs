Dim strFileScript : strFileScript = "clsService.vbs"
'***********************************************************************
'File:     clsRegistry.vbs
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


'Constants for ADSI services
  Const ADS_SVC_START_BOOT       = 0
  Const ADS_SVC_START_SYSTEM     = 1
  Const ADS_SVC_START_AUTO       = 2
  Const ADS_SVC_START_DEMAND     = 3
  Const ADS_SVC_DISABLED         = 4

  Const ADS_SVC_STOPPED          = 1
  Const ADS_SVC_START_PENDING    = 2
  Const ADS_SVC_STOP_PENDING     = 3
  Const ADS_SVC_RUNNING          = 4
  Const ADS_SVC_CONTINUE_PENDING = 5
  Const ADS_SVC_PAUSE_PENDING    = 6
  Const ADS_SVC_PAUSED           = 7
  Const ADS_SVC_ERROR            = 8


'********************************************************************************
Class clsService
  
  Private m_objSvc
  Private m_strName
  Private m_strDisplayName
  Private m_strConnectTo
  Private m_blnIsConnected
  Private m_intState
  Private m_intType
  Private m_intStartType
  Private m_intWin32ExitCode
  Private m_intServiceExitCode
  Private m_strBinary
  'Private m_strLoadOrderGroup
  Private m_dicDependedOnBy
  'Private m_strServiceStartName
  Private m_strSecurityDescriptor
  Private m_strComputer
  Private m_strCommandsAvailable
  Private m_blnServiceRqdRestart
  Private m_blnHasStateBeenChecked
  Private m_blnHasDependedOnByBeenChecked
  Private m_blnIsConfigurationUpdated
  Private m_strEnumBytesRequired
  
  '********************************************************************************
  Private Sub Class_Initialize()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    m_strComputer = CreateObject("Wscript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%")
    m_blnIsConnected                = False
    m_blnHasStateBeenChecked        = False
    m_blnHasDependedOnByBeenChecked = False
    m_blnIsConfigurationUpdated     = False
    m_blnServiceRqdRestart          = False
    
    Set m_dicDependedOnBy           = New clsDictionary
  End Sub
  
  '********************************************************************************
  Private Sub Class_Terminate()
    If IsObject(m_objSvc) Then Set m_objSvc = Nothing
  End Sub
  
  '********************************************************************************
  Public Property Get AccountName()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    AccountName = m_objSvc.ServiceAccountName
  End Property
  
  '********************************************************************************
  Public Property Let AccountName(ByVal strLogon)
    Call SetAccountName(strLogon)
  End Property
    
  '********************************************************************************
  Public Property Let AccountPath(ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    g_objError.Clear
    m_objSvc.Put "ServiceAccountPath", Trim(strValue)
    m_objSvc.SetInfo
    
    If Not g_objError.Check() Then
      m_blnIsConfigurationUpdated = True
    End If
  End Property
  
  '********************************************************************************
  Public Property Get AccountPath()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    AccountPath = m_objSvc.ServiceAccountPath
  End Property
  
  '********************************************************************************
  Public Property Get Binary()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    If IsObject(m_objSvc) Then
      Binary = m_objSvc.Path
    Else
      Binary = Empty
    End If
  End Property
  
  '********************************************************************************
  Public Property Let Binary(strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    g_objError.Clear
    m_objSvc.Put "Path", Trim(strValue)
    m_objSvc.SetInfo
    
    If Not g_objError.Check() Then
      m_blnIsConfigurationUpdated = True
    End If
  End Property
  
  '***********************************************************************
  Public Function Config_CheckAndFix(ByVal strStartMode, _
                                     ByVal strState, _
                                     ByVal strLogon, _
                                     ByVal strSecurityDescriptor, _
                                     ByVal blnAutoRepair)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Config_CheckAndFix = False
    
    Dim strProcedure : strProcedure = "Config_CheckAndFix"
    Dim i, blnSuccess
    Dim objService
    Dim intErrorCount
    Dim intStartMode_Required
    Dim intStartMode_Current
    Dim intState_Required
    Dim intState_Current
    Dim strState_Allowed
    
    Dim strLogon_Current
    Dim strLogon_Required
    Dim arrLogon_Required, strLogon_Domain_Required, strLogon_Name_Required
    Dim arrLogon_Current, strLogon_Domain_Current, strLogon_Name_Current
    Dim blnFixIt
    
    blnSuccess = False
    m_blnServiceRqdRestart = False
    
    intErrorCount = 0
    
    strState_Allowed = strState
    
    'Does the global auto repair setting override the service specific auto repair setting?
    If Not g_dicSettings.Key("ServiceRepair") Then
      blnAutoRepair = False
      Call LogIt("Setting AutoRepair to False because ServiceRepair = False.", "" & strProcedure, LogTypeWarning)
    End If
    
    If Not HWProfile_Check() Then
      If blnAutoRepair Then Call HWProfile_Fix()
    End If
    
    blnFixIt = False
    
    strLogon_Required = Trim(strLogon)
    If IsNullOrEmpty(strLogon_Required) Then strLogon_Required = ".\LocalSystem"
    
    Call LogIt("Checking the " & m_strName & " service to verify the logon account is " & strLogon_Required & ".", "" & strProcedure, LogTypeInfo)
    
    strLogon_Current = Trim(m_objSvc.ServiceAccountName)
    Call LogIt("  The " & m_strName & " service logon account is currently set to " & strLogon_Current & ".", "" & strProcedure, LogTypeInfo)
    
    If StrIn(1, strLogon_Required, "\") > 0 Then
      arrLogon_Required = StrSplit(strLogon_Required, "\", -1)
      strLogon_Domain_Required = Trim(arrLogon_Required(0))
      strLogon_Name_Required = Trim(arrLogon_Required(1))
    Else
      strLogon_Domain_Required = "."
      strLogon_Name_Required = strLogon_Required
    End If
    If StrIn(1, strLogon_Current, "\") > 0 Then
      arrLogon_Current = StrSplit(strLogon_Current, "\", -1)
      strLogon_Domain_Current = Trim(arrLogon_Current(0))
      strLogon_Name_Current = Trim(arrLogon_Current(1))
    Else
      strLogon_Domain_Current = "."
      strLogon_Name_Current = strLogon_Current
    End If
    
    If strLogon_Domain_Required = "." Then strLogon_Domain_Required = m_strComputer
    If strLogon_Domain_Current = "." Then strLogon_Domain_Current = m_strComputer
    
    If StrCompare(strLogon_Domain_Required, strLogon_Domain_Current) <> 0 Then
      blnFixIt = True
    End If
    If StrCompare(strLogon_Name_Required, strLogon_Name_Current) <> 0 Then
      blnFixIt = True
    End If
    
    If blnFixIt Then
      Call LogIt("  The " & m_strName & " service logon account is not correctly set.", "" & strProcedure, LogTypeError)
      If blnAutoRepair Then
        If StrCompare(strLogon_Domain_Required, m_strComputer) = 0 Then
          strLogon_Required = strLogon_Name_Required
        Else
          strLogon_Required = strLogon_Domain_Required & "\" & strLogon_Name_Required
        End If
        
        If Not SetAccountName(strLogon_Required) Then intErrorCount = intErrorCount + 1
      Else
        Call LogIt("    Cannot fix the logon account for the " & m_strName & " service (AutoRepair = FALSE).", "" & strProcedure, LogTypeWarning)
        intErrorCount = intErrorCount + 1
      End If
    Else
      Call LogIt("  The " & m_strName & " service is configured with the correct logon account.", "" & strProcedure, LogTypeInfo)
    End If
    
    'Validate the security descriptor
    If IsNullOrEmpty(strSecurityDescriptor) Then
      Call LogIt("Security descriptor setting is blank. Skipping validation.", "" & strProcedure, LogTypeWarning)
    Else
      If Not ValidateSecurityDescriptor(strSecurityDescriptor) Then
        If blnAutoRepair Then
          If Not SetSecurityDescriptor(strSecurityDescriptor) Then
            intErrorCount = intErrorCount + 1
          End If
        Else
          Call LogIt("    Cannot fix the security descriptor for the " & m_strName & " service (AutoRepair = FALSE).", "" & strProcedure, LogTypeWarning)
          intErrorCount = intErrorCount + 1
        End If
      End If
    End If
    
    intStartMode_Required = Eval(strStartMode)
    
    Call LogIt("Checking the " & m_strName & " service to see if it is set to " & StartType_Enum(intStartMode_Required) & " and is " & strState, "" & strProcedure, LogTypeInfo)
    'check start mode and see if it matches what we want
    intStartMode_Current = m_objSvc.StartType
    
    Call LogIt("  The " & m_strName & " service StartMode is set to " & StartType_Enum(intStartMode_Current), "" & strProcedure, LogTypeInfo)
    
    If intStartMode_Current <> intStartMode_Required Then
      'It is set incorrectly ... attempt to fix
      Call LogIt("  The " & m_strName & " service StartMode is incorrectly set.", "" & strProcedure, LogTypeWarning)
      
      If blnAutoRepair Then
        If SetStartType(intStartMode_Required) Then
          Call LogIt("  The " & m_strName & " service StartMode is now set correctly.", "" & strProcedure, LogTypeInfo)
        Else
          Call LogIt("  Failed to set the " & m_strName & " service StartMode.", "" & strProcedure, LogTypeError)
          intErrorCount = intErrorCount + 1
        End If
      Else
        Call LogIt("    Cannot fix the StartMode for the " & m_strName & " service (AutoRepair = FALSE).", "" & strProcedure, LogTypeWarning)
        intErrorCount = intErrorCount + 1
      End If
    End If
    
    If StrIn(1, strState, ",") > 0 Then
      strState = Left(strState, StrIn(1, strState, ",") - 1)
    End If
    intState_Required = Eval(strState)
    
    If m_blnIsConfigurationUpdated Then
      Call LogIt("    The " & m_strName & " service configuration was updated", "" & strProcedure, LogTypeInfo)
      Call EventLog_Write(EVENT_INFO, "The " & m_strName & " service configuration was updated.")
      
      'If the configuration was updated, we restart the service
      'Get the current state of the service so we can start it later.
      If m_objSvc.Status <> ADS_SVC_STOPPED Then
        Call LogIt("  The " & m_strName & " service will be restarted", "" & strProcedure, LogTypeInfo)
        If IsNullOrEmpty(intState_Required) Then
          Select Case m_objSvc.Status
            Case ADS_SVC_START_PENDING
              intState_Required = ADS_SVC_RUNNING
            Case ADS_SVC_STOP_PENDING
            Case ADS_SVC_RUNNING
              intState_Required = ADS_SVC_RUNNING
            Case ADS_SVC_CONTINUE_PENDING
              intState_Required = ADS_SVC_RUNNING
            Case ADS_SVC_PAUSE_PENDING
            Case ADS_SVC_PAUSED
          End Select
        End If
        
        'Stop the service
        Call SetState(ADS_SVC_STOPPED)
      End If
    End If
    
    'Make sure the service is set to the desired state
    If Not IsNullOrEmpty(intState_Required) Then
      'check state and see if it matches what we want
      On Error Resume Next
      g_objError.Clear
      intState_Current = m_objSvc.Status
      If g_objError.Check() Then
        Call LogIt("  Could not get the " & m_strName & " service current state. " & g_objError.Message, "" & strProcedure, LogTypeError)
        intErrorCount = intErrorCount + 1
      Else
        If StrIn(1, strState_Allowed, SvcStateIntToString(intState_Current)) > 0 Then
          Call LogIt("  The " & m_strName & " state is correctly set to " & State_Enum(intState_Current), "" & strProcedure, LogTypeInfo)
          'Commented this out since it starts EVERYTHING
          'If intState_Required = ADS_SVC_RUNNING Then
          '  Call SetDependentSvcs(ADS_SVC_RUNNING)
          'End If
        Else
          Call LogIt("  The " & m_strName & " state is incorrectly set to " & State_Enum(intState_Current), "" & strProcedure, LogTypeWarning)
          'Are we allowed to make system changes?
          If blnAutoRepair Then
            'Set the service state
            If Not SetState(Eval(strState)) Then
              intErrorCount = intErrorCount + 1
            Else
              intState_Current = m_objSvc.Status
            
              Call LogIt("    The " & m_strName & " service is " & State_Enum(intState_Current), "" & strProcedure, LogTypeInfo)
              
              If StrIn(1, strState_Allowed, SvcStateIntToString(intState_Current)) > 0 Then
                'if it matches, return True
                Call LogIt("    The " & m_strName & " service operation succeeded.", "" & strProcedure, LogTypeInfo)
                m_blnServiceRqdRestart = True
              Else
                Call LogIt("    The " & m_strName & " service operation failed.", "" & strProcedure, LogTypeError)
                intErrorCount = intErrorCount + 1
              End If
            End If
          Else
            Call LogIt("    Cannot set the " & m_strName & " service to " & strState & " (AutoRepair = FALSE).", "" & strProcedure, LogTypeWarning)
            intErrorCount = intErrorCount + 1
          End If
        End If
      End If
    End If
    
    If intErrorCount = 0 Then
      Call LogIt("    The " & m_strName & " service configuration was verified.", "" & strProcedure, LogTypeInfo)
      Config_CheckAndFix = True
    Else
      Call LogIt("    Failed to verify or update the " & m_strName & " service configuration.", "" & strProcedure, LogTypeError)
    End If
  End Function
  
  '********************************************************************************
  Public Function Connect(ByVal strService)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Connect = False
    
    Dim strProcedure : strProcedure = "Connect"
    
    m_strName = strService
    
    If Not Service_Check_Exists(m_strComputer, m_strName) Then
      Call LogIt("The " & m_strName & " service does not exist.", "" & strProcedure, LogTypeWarning)
      Exit Function
    End If
    
    Call LogIt("Connecting to the " & m_strName & " service.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
    If Not ObjectRef_Get(m_objSvc, "WinNT://" & m_strComputer & "/" & m_strName & ",Service") Then
      Call LogIt("  Could not connect to the service. " & g_objError.Message, "" & strProcedure, LogTypeError)
      m_blnIsConnected = False
    Else
      Call LogIt("  Connected to the service.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
      Connect = True
      m_blnIsConnected = True
      m_strDisplayName = m_objSvc.DisplayName
    End If
  End Function
  
  '********************************************************************************
  Public Function Delete()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
 ' This procedure requires SC.exe to be located in the current folder or a folder
 '  defined by the Path environment variable
    
    Delete = False
    
    Dim strProcedure : strProcedure = "Delete"
    Dim strReturn
    Dim arrReturn
    Dim blnSuccess
    
    blnSuccess = False
    
    If IsObject(m_objSvc) Then Set m_objSvc = Nothing
    
    Call LogIt("Deleting the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
    
    If g_objProcess.Exec("sc.exe delete " & m_strName, "0") Then
      arrReturn = g_objProcess.Output
      For Each strReturn In arrReturn
        If Not IsNullOrEmpty(strReturn) Then
          If StrIn(1, strReturn, "[SC] DeleteService SUCCESS") > 0 Then
            blnSuccess = True
            Exit For
          End If
        End If
      Next
    End If
    
    If blnSuccess Then
      Call LogIt("  Deleted the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
    Else
       Call LogIt("  Could not delete the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
    End If
    
    Call g_objProcess.Exec("sc.exe interrogate " & m_strName, "0")
    Call g_objProcess.Exec("sc.exe query " & m_strName, "0")
    
    Delete = blnSuccess
  End Function
  
  '********************************************************************************
  Public Property Get DependsOn()
 '********************************************************************************
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Dim strDependsOn : strDependsOn = Empty
    Dim strDependency : strDependency = Empty
    
    If IsArray(m_objSvc.Dependencies) Then
      For Each strDependency In m_objSvc.Dependencies
        If IsNullOrEmpty(strDependsOn) Then
          strDependsOn = strDependency
        Else
          strDependsOn = strDependsOn & "," & strDependency
        End If
      Next
    Else
      If Len(m_objSvc.Dependencies) > 0 Then
        strDependsOn = m_objSvc.Dependencies
      End If
    End If
    DependsOn = strDependsOn
  End Property
  
  '********************************************************************************
  Public Function DependedOnBy(ByVal blnRescan)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
 ' This procedure requires SC.exe to be located in the current folder or a folder
 '  defined by the Path environment variable
 '  We can, sometime in the future, remove the need for SC.exe if we enumerate
 '  HKLM\System\CurrentControlSet\Services and look for DependOnGroup and
 '  DependOnService entries.
    
    Dim strProcedure : strProcedure = "DependedOnBy"
    Dim strResults
    Dim i
    Dim arrResults
    Dim strLine
    Dim strServiceName
    Dim blnDoReRun
    Dim intDependantCount
    
    blnDoReRun = False
    
    If Not m_blnHasDependedOnByBeenChecked Then blnRescan = True
    
    If blnRescan Then
      Call LogIt("Getting all services that depend on the " & m_strName & " service.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
      
      m_dicDependedOnBy.RemoveAll
      
      strResults = Empty
      Call g_objProcess.Exec("sc.exe EnumDepend " & m_strName & m_strEnumBytesRequired, "0, 234")
      arrResults = g_objProcess.Output
      If Not IsNullOrEmpty(arrResults) Then
          For i = 0 To UBound(arrResults)
            strLine = Trim(arrResults(i))
            If Not IsNullOrEmpty(strLine) Then
              If StrIn(1, strLine, ": more data, need ") > 0 Then
                m_strEnumBytesRequired = Mid(strLine, StrIn(1, strLine, ", need ") + 7)
                m_strEnumBytesRequired = " " & Left(m_strEnumBytesRequired, StrIn(1, m_strEnumBytesRequired, " ") - 1)
                
                Call LogIt("  The buffer is set too small to get all dependant services. Re-running with a buffer of '" & m_strEnumBytesRequired & "' bytes.", "" & strProcedure, LogTypeWarning + LogTypeVerbose)
                blnDoReRun = True
                Exit For
              End If
              If StrIn(1, strLine, "entriesread") > 0 Then
                '[SC] EnumDependentServices: entriesread = 0
                intDependantCount = CInt(Trim(Mid(strLine, StrIn(1, strLine, "=") + 1)))
                If intDependantCount = 0 Then
                  Call LogIt("  There are no services dependant on " & m_strName, "" & strProcedure, LogTypeInfo + LogTypeVerbose)
                  Exit For
                End If
              End If
              If UCase(Left(strLine, 12)) = "SERVICE_NAME" Then
                strServiceName = Mid(strLine, StrIn(1, strLine, ":") + 1)
                strServiceName = Trim(strServiceName)
                
                Call m_dicDependedOnBy.Add(strServiceName, Empty)
              End If
            End If
          Next
      End If
    End If
    
    If blnDoReRun Then Call DependedOnBy(True)
    
    m_blnHasDependedOnByBeenChecked = True
    
    DependedOnBy = m_dicDependedOnBy.Keys
  End Function
  
  '********************************************************************************
  Public Property Get DisplayName()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    DisplayName = m_objSvc.DisplayName
  End Property
  
  '********************************************************************************
  Public Property Let DisplayName(ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    m_strDisplayName = Trim(strValue)
    If IsObject(m_objSvc) Then
      g_objError.Clear
      m_objSvc.Put "DisplayName", m_strDisplayName
      m_objSvc.SetInfo
        
      If Not g_objError.Check() Then
        m_blnIsConfigurationUpdated = True
      End If
    End If
  End Property
  
  '********************************************************************************
'  Public Function Exists()
    
'    Exists = False
    
    'Wscript.echo "Service state for " & m_objSvc.Name & " is " & strState
    
'    Exists = True
'  End Function
  
  '***********************************************************************
  Public Function HWProfile_Check()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    HWProfile_Check = False
    
    Dim strProcedure : strProcedure = "HWProfile_Check"
    Dim strKeyPath, strValueName, varValue
    
    Call LogIt("Checking HW profile key for the " & m_strName & " service to make sure it is 0 or doesn't exist.", "" & strProcedure, LogTypeInfo)
  
    strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current\System\CurrentControlSet\Enum\ROOT\LEGACY_" & m_strName & "\0000"
    strValueName = "CSConfigFlags"
    
    If Not g_objReg.Exists(strKeyPath & "\" & strValueName) Then
      HWProfile_Check = True
      Exit Function
    End If
    
    If g_objReg.Read(strKeyPath & "\" & strValueName, varValue) Then
      'If the value is set to anything but 0 then the hardware profile is not enabled
      If Not IsNullOrEmpty(varValue) Then
        If CInt(varValue) <> 0 Then
          Call LogIt(" The HW profile value for the " & m_strName & " service is not set to 0.", "" & strProcedure, LogTypeError)
        Else
          Call LogIt(" The HW profile key for the " & m_strName & " is set to 0.", "" & strProcedure, LogTypeInfo)
          HWProfile_Check = True
        End If
      End If
    End If
  End Function
  
  '***********************************************************************
  Public Function HWProfile_Fix()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    'This currently defaults to enabling the service for ALL hardware profiles
  
    HWProfile_Fix = False
    
    Dim strProcedure : strProcedure = "HWProfile_Fix"
    Dim strHive, strKey, strValueName, dwValue
    Dim intReturn
    
    Call LogIt("Fixing HW profile registry key for the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
  
    strKey = "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current\System\CurrentControlSet\Enum\ROOT\LEGACY_" & m_strName & "\0000"
  
    If g_objReg.Delete(strKey & "\") Then
      Call EventLog_Write(EVENT_INFO, "Deleted the " & strKey & " registry key from the registry.")
      Call LogIt(" The HW profile registry key for the " & m_strName & " was deleted.", "" & strProcedure, LogTypeInfo)
      
      m_blnIsConfigurationUpdated = True
      HWProfile_Fix = True
    Else
      Call LogIt("  The HW profile registry key for the " & m_strName & " could not be deleted. " & g_objError.Message, "" & strProcedure, LogTypeError)
      Call g_objHealthStatus.Add("GEN", errService_HWProfileNotEnabled)
    End If
  End Function
  
  '********************************************************************************
  Private Function GetSecurityDescriptor()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
   ' This procedure requires SC.exe to be located in the current folder or a folder
   '  defined by the Path environment variable
    
    Dim strProcedure : strProcedure = "GetSecurityDescriptor"
    Dim strReturn, arrReturn
    Dim strSD
  
    GetSecurityDescriptor = Empty
    m_strSecurityDescriptor = Empty
    
    Call LogIt("Getting the security descriptor for the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
    
    If g_objProcess.Exec("sc.exe sdshow " & m_strName, "0") Then
      arrReturn = g_objProcess.Output
        For Each strReturn In arrReturn
          If Not IsNullOrEmpty(strReturn) Then
            strSD = strSD & " " & strReturn
          End If
        Next
    End If
    
    If IsNullOrEmpty(strSD) Then
      Call LogIt("  Could not get the security descriptor for the " & m_strName & " service.", "" & strProcedure, LogTypeError)
    Else
      GetSecurityDescriptor = strSD
      m_strSecurityDescriptor = strSD
    End If
  End Function
   
  Private Function SetSecurityDescriptor(ByVal strSecurityDescriptor)
  '***********************************************************************
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Dim strProcedure : strProcedure = "SetSecurityDescriptor"
    Dim strReturn, arrReturn
    Dim blnSuccess
    
    blnSuccess = False
    
    If IsNullOrEmpty(strSecurityDescriptor) Then Exit Function
    
    Call LogIt("Setting the security descriptor for the " & m_strName & " service to '" & strSecurityDescriptor & "'.", "" & strProcedure, LogTypeInfo)
    
    If g_objProcess.Exec("sc.exe sdset " & m_strName & " " & Chr(34) & strSecurityDescriptor & Chr(34), "0") Then
      arrReturn = g_objProcess.Output
      For Each strReturn In arrReturn
        If StrIn(1, strReturn, "[sc] setserviceobjectsecurity success") > 0 Then
          blnSuccess = True
          Exit For
        End If
      Next
    End If
    
    If blnSuccess Then
      Call EventLog_Write(EVENT_INFO, "Set the " & m_strName & " service security descriptor to " & strSecurityDescriptor & ".")
      Call LogIt("  Set the security descriptor.", "" & strProcedure, LogTypeInfo)
      
      m_blnIsConfigurationUpdated = True
      m_strSecurityDescriptor = strSecurityDescriptor
    Else
      Call LogIt("  Could not set security descriptor. Error = " & strReturn, "" & strProcedure, LogTypeError)
    End If
  End Function

 '********************************************************************************
  Public Property Get Host()
    Host = m_strComputer
  End Property
  
  '********************************************************************************
  Public Property Let Host(ByVal strValue)
    m_strComputer = Trim(strValue)
  End Property
  
  '********************************************************************************
  Private Function IsCommandAvailable(ByVal strCommand)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    IsCommandAvailable = False
    
    Dim strProcedure : strProcedure = "IsCommandAvailable"
    Dim strCommandsAvailable
    Dim arrResults
    Dim strResults
    Dim strLine
    Dim strStartType
    
    If StartType() = ADS_SVC_DISABLED Then
      Exit Function
    End If
    
    'Sample output from running sc query command on a service
    ' c:\>sc query tbs
    ' 
    ' SERVICE_NAME: tbs
    '         TYPE               : 20  WIN32_SHARE_PROCESS
    '         STATE              : 1  STOPPED
    '         WIN32_EXIT_CODE    : 1077  (0x435)
    ' ...
    '
    ' c:\>sc queryex eventlog
    ' 
    ' SERVICE_NAME: eventlog
    '         TYPE               : 20  WIN32_SHARE_PROCESS
    '         STATE              : 4  RUNNING
    '                                 (STOPPABLE, NOT_PAUSABLE, ACCEPTS_SHUTDOWN)
    '         WIN32_EXIT_CODE    : 0  (0x0)
    '...
    
    strResults = Empty
    Call g_objProcess.Exec("sc.exe query " & m_strName, "0, 234")
    arrResults = g_objProcess.Output
    If Not IsNullOrEmpty(arrResults) Then
      For i = 0 To UBound(arrResults)
        strLine = Trim(arrResults(i))
        If Not IsNullOrEmpty(strLine) Then
          If StrIn(1, strLine, "STATE") > 0 Then
            strLine = Trim(arrResults(i) + 1)
            If Left(strLine, 15) <> "WIN32_EXIT_CODE" Then
              strCommandsAvailable = Mid(strLine, 2, Len(strLine) - 1)
              Exit For
            End If
            Exit For
          End If
        End If
      Next
    End If
    
    strCommandsAvailable = StrReplace(strCommandsAvailable, "NOT_STOPPABLE", "", 1, -1)
    strCommandsAvailable = StrReplace(strCommandsAvailable, "NOT_PAUSABLE", "", 1, -1)
    strCommandsAvailable = StrReplace(strCommandsAvailable, "IGNORES_SHUTDOWN", "", 1, -1)
    strCommandsAvailable = StrReplace(strCommandsAvailable, "ACCEPTS_SHUTDOWN", "", 1, -1)
    
    strCommandsAvailable = StrReplace(strCommandsAvailable, "STOPPABLE", "STOP", 1, -1)
    strCommandsAvailable = StrReplace(strCommandsAvailable, " PAUSABLE", "PAUSE", 1, -1)
    
    If StrIn(1, strCommandsAvailable, strCommand) > 0 Then
      IsCommandAvailable = True
    Else
      IsCommandAvailable = False
    End If
  End Function
  
  '********************************************************************************
  Public Property Get IsConfigurationUpdated()
    IsConfigurationUpdated = m_blnIsConfigurationUpdated
  End Property
  
  '********************************************************************************
  Public Property Get IsConnected()
    IsConnected = m_blnIsConnected
  End Property
  
  '********************************************************************************
  Public Property Get LoadOrderGroup()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    LoadOrderGroup = m_objSvc.LoadOrderGroup
  End Property
  
  '********************************************************************************
  Public Property Let LoadOrderGroup(ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    g_objError.Clear
    m_objSvc.Put "LoadOrderGroup", Trim(strValue)
    m_objSvc.SetInfo
    
    If Not g_objError.Check() Then
      m_blnIsConfigurationUpdated = True
    End If
  End Property

  '********************************************************************************
  Public Property Get Name()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    If IsObject(m_objSvc) Then
      Name = m_objSvc.Name
    Else
      Name = m_strName
    End If
  End Property
  
  '********************************************************************************
  Public Property Let Name(ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    m_strName = Trim(strValue)
    If IsObject(m_objSvc) Then
      g_objError.Clear
      m_objSvc.Put "Name", m_strName
      m_objSvc.SetInfo
      
'>>>>>> Error handling?
      If Not g_objError.Check() Then
        m_blnIsConfigurationUpdated = True
      End If
    End If
  End Property
    
  '********************************************************************************
  Public Property Let Password(ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Dim strProcedure : strProcedure = "Password"
    Dim blnSuccess
    
    blnSuccess = False
    
    g_objError.Clear
    Do
      m_objSvc.SetPassword Trim(strValue)
      If g_objError.Check() Then Exit Do
      
      m_objSvc.SetInfo
      If g_objError.Check() Then Exit Do
      
      blnSuccess = True
    Loop
    
    If blnSuccess Then
      Call LogIt("  Set the service logon account password.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
      m_blnIsConfigurationUpdated = True
    Else
      Call LogIt("  Failed to set the service logon account password.", "" & strProcedure, LogTypeError)
    End If
  End Property
  
  '***********************************************************************
  Public Function Restart()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Restart = False
    
    Dim strProcedure : strProcedure = "Restart"
    Dim intErrorCount
    Dim intStatus
    
    If Not IsConnected() Then
      Call Connect(m_objSvc.Name)
    End If
    
    If Not m_blnIsConnected Then
      'Call LogIt("  failed to connect to the " & m_objSvc.Name & " service.", "" & strProcedure, LogTypeError)
      Exit Function
    End If
    
    On Error Resume Next
    intStatus = m_objSvc.Status
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    'check state and see if it is running
    Call LogIt("  The " & m_objSvc.Name & " service is " & State_Enum(intStatus), "" & strProcedure, LogTypeInfo)
    
    intErrorCount = 0
    
    If m_objSvc.Status <> ADS_SVC_STOPPED Then
'>>>>>>>>>
      Call SetDependentSvcs(ADS_SVC_STOPPED)
'>>>>>>>>>
      
      Call LogIt("  Attempting to set the " & m_objSvc.Name & " service to Stopped.", "" & strProcedure, LogTypeInfo)
       
      If Not SetState(ADS_SVC_STOPPED) Then
        intErrorCount = intErrorCount + 1
      Else
        intStatus = m_objSvc.Status
      
        Call LogIt("    The " & m_strName & " service is " & State_Enum(intStatus), "" & strProcedure, LogTypeInfo)
        
        If intStatus = ADS_SVC_STOPPED Then
          Call LogIt("    The operation succeeded.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
        Else
          Call LogIt("    The operation failed.", "" & strProcedure, LogTypeError)
'>>>>>>>>>
          Call SetDependentSvcs(ADS_SVC_RUNNING)
'>>>>>>>>>
          
          Exit Function
        End If
      End If
    End If
    
    Call LogIt("  Attempting to set the " & m_objSvc.Name & " service to Running.", "" & strProcedure, LogTypeInfo)
    
    If Not SetState(ADS_SVC_RUNNING) Then
      intErrorCount = intErrorCount + 1
    Else
      intStatus = m_objSvc.Status
      
      Call LogIt("    The " & m_strName & " service is " & State_Enum(intStatus), "" & strProcedure, LogTypeInfo)
      
      If intStatus = ADS_SVC_RUNNING Then
'>>>>>>>>>
        Call SetDependentSvcs(ADS_SVC_RUNNING)
'>>>>>>>>>
      Else
        Call LogIt("    The operation failed.", "" & strProcedure, LogTypeError)
        intErrorCount = intErrorCount + 1
      End If
    End If
    
    If intErrorCount = 0 Then Restart = True
'Call LogIt(">>>>>> end restart", "" & strProcedure, LogTypeError)
  End Function
  
  '********************************************************************************
  Public Function RqdStart()
    RqdStart = m_blnServiceRqdRestart
  End Function
  
  '********************************************************************************
  Private Function SetAccountName(ByVal strLogon)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    SetAccountName = False
    
    Dim strProcedure : strProcedure = "SetAccountName"
    Dim arrReturn, strReturn, blnSuccess
    
    Call LogIt("    Setting the service logon account to " & strLogon & ".", "" & strProcedure, LogTypeInfo)
    
    If g_objProcess.Exec("sc.exe config " & m_strName & " obj= " & Chr(34) & strLogon & Chr(34), "0") Then
      arrReturn = g_objProcess.Output
      For Each strReturn In arrReturn
        If StrIn(1, strReturn, "[sc] changeserviceconfig success") > 0 Then
          blnSuccess = True
          Exit For
        End If
      Next
    End If
    
    If blnSuccess Then
      Call LogIt("  Set the service logon account using SC.", "" & strProcedure, LogTypeInfo)
      Call EventLog_Write(EVENT_INFO, "Set the " & m_strName & " service logon account to " & strLogon & ".")
      
      m_blnIsConfigurationUpdated = True
      SetAccountName = True
    Else
      Call LogIt("  Could not set service logon account.", "" & strProcedure, LogTypeError)
    End If
  End Function
  
  '********************************************************************************
  Private Function SetAccountName_SC(ByVal strStartName, ByVal strStartPassword)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    SetAccountName_SC = False
    
    Dim strProcedure : strProcedure = "SetAccountName_SC"
    Dim arrReturn, strReturn, blnSuccess
    
    If g_objProcess.Exec("sc.exe config " & m_strName & " obj= " & Chr(34) & strStartName & Chr(34), "0") Then
      arrReturn = g_objProcess.Output
      For Each strReturn In arrReturn
        If StrIn(1, strReturn, "[sc] changeserviceconfig success") > 0 Then
          blnSuccess = True
          Exit For
        End If
      Next
    End If
    
    If blnSuccess Then
      Call LogIt("  Set the service logon account using SC.", "" & strProcedure, LogTypeInfo)
      SetAccountName_SC = True
    Else
      Call LogIt("  Failed to set the service logon account using SC. " & g_objError.Message, "" & strProcedure, LogTypeError)
    End If
  End Function
  
  '***********************************************************************
  Public Property Get SecurityDescriptor()
    SecurityDescriptor = GetSecurityDescriptor()
  End Property
  
  '***********************************************************************
  Public Property Let SecurityDescriptor(ByVal strSecurityDescriptor)
    Call SetSecurityDescriptor(strSecurityDescriptor)
  End Property
  
  '********************************************************************************
  Public Function SetDependentSvcs(ByVal intState_Required)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    SetDependentSvcs = False
    
    Dim strProcedure : strProcedure = "SetDependentSvcs"
    Dim blnErrorOccurred
    Dim objDependentSvc
    Dim arrDependentSvcs
    Dim strDependentSvc
    
    blnErrorOccurred = False
    
    arrDependentSvcs = DependedOnBy(False)
    If UBound(arrDependentSvcs) >= 0 Then
      Call LogIt("Setting all services that depend on the " & m_strName & " service to " & State_Enum(intState_Required), "" & strProcedure, LogTypeWarning)
      For Each strDependentSvc In arrDependentSvcs
        Set objDependentSvc = New clsService
        Do
          If Not objDependentSvc.Connect(strDependentSvc) Then Exit Do
          If objDependentSvc.StartType = ADS_SVC_DISABLED Then
            Call LogIt("    The " & strDependentSvc & " service is set to disabled. Will not attempt to set it to " & State_Enum(intState_Required) & ".", "" & strProcedure, LogTypeInfo)
            Exit Do
          End If
          objDependentSvc.State = intState_Required
          If objDependentSvc.State <> intState_Required Then blnErrorOccurred = True
          Exit Do
        Loop
        If IsObject(objDependentSvc) Then Set objDependentSvc = Nothing
      Next
    Else
      Call LogIt("The " & m_strName & " service does not have any dependant services.", "" & strProcedure, LogTypeInfo)
    End If
    
    If Not blnErrorOccurred Then SetDependentSvcs = True
  End Function
  
  '********************************************************************************
  Private Function SetStartType(ByVal intValue)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    SetStartType = False
    
    Dim strProcedure : strProcedure = "SetStartType"
    
    If Not IsNumeric(intValue) Then
      Call LogIt ("The StartType procedure was passed a non-numeric value.", "" & strProcedure, LogTypeInfo)
      Exit Function
    End If
    
    If IsObject(m_objSvc) Then
      g_objError.Clear
      m_objSvc.Put "StartType", intValue
      m_objSvc.SetInfo
      
      If Not g_objError.Check() Then
        m_blnIsConfigurationUpdated = True
        SetStartType = True
      End If
    End If
  End Function
  
  '********************************************************************************
  Private Function SetState(ByVal intState_Required)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Dim strProcedure : strProcedure = "SetState"
    Dim strState_Required, strState_Current
    Dim blnStateChanged
    Dim intErrorCount
    Dim intStartType
    
    SetState = False
    
    Call LogIt("Attempting to set the " & m_strName & " service to " & State_Enum(intState_Required) & ".", "" & strProcedure, LogTypeInfo)
    
    If Not IsNumeric(intState_Required) Then
      Call LogIt("The State procedure was passed a non-numeric desired state value.", "" & strProcedure, LogTypeError)
      Exit Function
    End If
    If Not IsObject(m_objSvc) Then
      Call LogIt("The " & m_strName & " service is not yet connected.", "" & strProcedure, LogTypeWarning)
      Exit Function
    End If
    If Not m_blnIsConnected Then
      Call LogIt("The " & m_strName & " service is not yet connected.", "" & strProcedure, LogTypeWarning)
      Exit Function
    End If
    
    strState_Required = State_Enum(intState_Required)
    
    If m_blnHasStateBeenChecked Then
      Call Connect(m_objSvc.Name)
    End If
    
    If Not m_blnIsConnected Then
      'Call LogIt("  failed to connect to the " & m_objSvc.Name & " service.", "" & strProcedure, LogTypeError)
      Exit Function
    End If
    
    m_blnHasStateBeenChecked = True
    
    If m_objSvc.Status = intState_Required Then
      Call LogIt("The " & m_strName & " service is already in a state of " & strState_Required, "" & strProcedure, LogTypeInfo)
      
      If intState_Required = ADS_SVC_RUNNING Then
        Call SetDependentSvcs(ADS_SVC_RUNNING)
      End If
      Exit Function
    End If
    
    strState_Current = State_Enum(m_objSvc.Status)
    
    Call LogIt("The " & m_strName & " service is in a state of " & strState_Current, "" & strProcedure, LogTypeInfo)
    
    If intState_Required = ADS_SVC_STOPPED Then
      If m_objSvc.Status <> ADS_SVC_STOPPED Then
        If Not SetDependentSvcs(ADS_SVC_STOPPED) Then
          Exit Function
        End If
      End If
    End If
    
    Select Case m_objSvc.Status
      Case ADS_SVC_START_PENDING
        If Not StateWait(ADS_SVC_RUNNING) Then Exit Function
      Case ADS_SVC_STOP_PENDING
        If Not StateWait(ADS_SVC_STOPPED) Then Exit Function
      Case ADS_SVC_CONTINUE_PENDING
        If Not StateWait(ADS_SVC_RUNNING) Then Exit Function
      Case ADS_SVC_PAUSE_PENDING
        If Not StateWait(ADS_SVC_PAUSED) Then Exit Function
      Case ADS_SVC_ERROR
        'handle service being in an error state???
    End Select
    
    blnStateChanged = False
    
    On Error Resume Next
    Select Case intState_Required
      Case ADS_SVC_STOPPED
        Select Case m_objSvc.Status
          Case ADS_SVC_STOP_PENDING
            Call LogIt("Cannot stop the " & m_strName & " service because it is currently pending a stop.", "" & strProcedure, LogTypeError)
          Case Else
            If IsCommandAvailable("Stop") Then
              Call LogIt("Sending a Stop command to the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
              g_objError.Clear
              m_objSvc.Stop
              If g_objError.Check()  Then
                Call LogIt("  Could not stop the " & m_strName & " service. " & g_objError.Message, "" & strProcedure, LogTypeError)
                intErrorCount = intErrorCount + 1
              Else
                blnStateChanged = True
              End If
            Else
              '
            End If
        End Select
      Case ADS_SVC_RUNNING
        Select Case m_objSvc.Status
          Case ADS_SVC_CONTINUE_PENDING
            Call LogIt("Cannot start the " & m_strName & " service because it is currently pending a continue.", "" & strProcedure, LogTypeError)
          Case ADS_SVC_PAUSED
            If IsCommandAvailable("Continue") Then
              Call LogIt("Sending a Continue command to the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
              g_objError.Clear
              m_objSvc.Continue
              If g_objError.Check()  Then
                Call LogIt("  Could not start the " & m_strName & " service. " & g_objError.Message, "" & strProcedure, LogTypeError)
                intErrorCount = intErrorCount + 1
              Else
                blnStateChanged = True
              End If
            Else
              '
            End If
          Case Else
            intStartType = StartType()
            If intStartType <> ADS_SVC_DISABLED Then
              If IsCommandAvailable("Start") Then
                Call LogIt("Sending a Start command to the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
                g_objError.Clear
                m_objSvc.Start
                If g_objError.Check() Then
                  Call LogIt("  Could not start the " & m_strName & " service. " & g_objError.Message, "" & strProcedure, LogTypeError)
                  intErrorCount = intErrorCount + 1
                Else
                  blnStateChanged = True
                End If
              Else
                '
              End If
            Else
              Call LogIt("Cannot start the " & m_strName & " service because it is " & StartType_Enum(intStartType) & " .", "" & strProcedure, LogTypeError)
              intErrorCount = intErrorCount + 1
            End If
        End Select
      Case ADS_SVC_PAUSED
        Select Case m_objSvc.Status
          Case ADS_SVC_STOP_PENDING
            Call LogIt("Cannot pause the " & m_strName & " service because it is currently pending a stop.", "" & strProcedure, LogTypeError)
            intErrorCount = intErrorCount + 1
          Case ADS_SVC_PAUSE_PENDING
            Call LogIt("Cannot pause the " & m_strName & " service because it is currently pending a pause.", "" & strProcedure, LogTypeError)
            intErrorCount = intErrorCount + 1
          Case ADS_SVC_STOPPED
            Call LogIt("Cannot pause the " & m_strName & " service because it is currently stopped.", "" & strProcedure, LogTypeError)
            intErrorCount = intErrorCount + 1
          Case Else
            If IsCommandAvailable("Pause") Then
              Call LogIt("Sending a Pause command to the " & m_strName & " service.", "" & strProcedure, LogTypeInfo)
              g_objError.Clear
              m_objSvc.Pause
              If g_objError.Check() Then
                Call LogIt("  Could not pause the " & m_strName & " service. " & g_objError.Message, "" & strProcedure, LogTypeError)
                intErrorCount = intErrorCount + 1
              Else
                blnStateChanged = True
              End If
            Else
              '
            End If
        End Select
    End Select
    
    If blnStateChanged Then
'>>>>>>
      m_blnIsConfigurationUpdated = False
'>>>>>>
      If StateWait(intState_Required) Then
      
      Else
        intErrorCount = intErrorCount + 1
      End If
    End If
    
    If intState_Required = ADS_SVC_RUNNING Then
      If SetDependentSvcs(ADS_SVC_RUNNING) = False Then
        intErrorCount = intErrorCount + 1
      End If
    End If
    
    If intErrorCount = 0 Then SetState = True
  End Function
  
  '********************************************************************************
  Public Property Get State()
    If IsObject(m_objSvc) Then
      On Error Resume Next
      g_objError.Clear
      State = m_objSvc.Status
      If g_objError.Check() Then State = ADS_Svc_Error
    Else
      State = m_intState
    End If
  End Property
  
  '********************************************************************************
  Public Property Let State(ByVal intState_Required)
    Call SetState(intState_Required)
  End Property
  
  '********************************************************************************
  Public Function State_Enum(ByVal intState)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    State_Enum = Empty
    
    Dim strProcedure : strProcedure = "State_Enum"
    
    If Not IsNumeric(intState) Then
      Call LogIt("The State_Enum procedure was passed a non-numeric value.", "" & strProcedure, LogTypeInfo)
      Exit Function
    End If
    
    Select Case CInt(intState)
      Case 1    : State_Enum = "Stopped"
      Case 2    : State_Enum = "Start Pending"
      Case 3    : State_Enum = "Stop Pending"
      Case 4    : State_Enum = "Running"
      Case 5    : State_Enum = "Continue Pending"
      Case 6    : State_Enum = "Pause Pending"
      Case 7    : State_Enum = "Paused"
      Case Else : State_Enum = "Error"
    End Select
  End Function
  
  '********************************************************************************
  Public Function StateWait(ByVal intState_Required)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Dim strProcedure : strProcedure = "StateWait"
    Dim intLoopCounter : intLoopCounter = 0
    Dim strResults
    Dim intState_Current
    
    If Not IsObject(m_objSvc) Then
      Call LogIt("The " & m_strName & " service is not yet connected.", "" & strProcedure, LogTypeWarning)
      Exit Function
    End If
    
    StateWait = False
    
    If SvcStateIntToString(intState_Required) = Empty Then
      Call LogIt("  The StateWait procedure was passed an incorrect state value.", "" & strProcedure, LogTypeError)
      Exit Function
    End If
    
    Call LogIt("  Waiting for the " & m_strName & " service to enter a " & State_Enum(intState_Required) & " state.", "" & strProcedure, LogTypeInfo)
    
    'Wait for the service to be in the desired state
    Do
      Call Connect(m_strName)
      
      If Not m_blnIsConnected Then Exit Do
      
      On Error Resume Next
      g_objError.Clear
      intState_Current = CInt(m_objSvc.Status)
      If g_objError.Check() Then
        Call LogIt("  Could not get the current service state for " & m_strName & ". " & g_objError.Message, "" & strProcedure, LogTypeError)
      Else
        If intState_Current <> CInt(intState_Required) Then
          Call LogIt("The " & m_strName & " service is in a state of " & State_Enum(m_objSvc.Status), "" & strProcedure, LogTypeInfo)
          Call LogIt("  Waiting for the " & m_strName & " service to enter a state of " & State_Enum(intState_Required), "" & strProcedure, LogTypeInfo)
        Else
          Call LogIt("The " & m_strName & " service is in a state of " & State_Enum(m_objSvc.Status), "" & strProcedure, LogTypeInfo)
          StateWait = True
          
          Exit Do
        End If
      End If
      
      intLoopCounter = intLoopCounter + 1
      If intLoopCounter > 36 Then 'don't wait too long before bailing
        Call EventLog_Write(EVENT_ERROR, "The " & m_strName & " service appears to be hung with a state of " & State_Enum(m_objSvc.Status) & ".")
        Call LogIt("   The " & m_strName & " service appears to be hung with a state of " & State_Enum(m_objSvc.Status), "" & strProcedure, LogTypeError)
        
        Exit Do
      End If
      Wscript.Sleep(2500)
    Loop
  End Function
  
  '********************************************************************************
  Public Function StartType_Enum(ByVal intStartType)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    StartType_Enum = Empty
    
    Dim strProcedure : strProcedure = "StartType_Enum"
    
    If Not IsNumeric(intStartType) Then
      Call LogIt("The StartType_Enum procedure was passed a non-numeric value.", "" & strProcedure, LogTypeInfo)
      Exit Function
    End If
    
    Select Case CInt(intStartType)
      Case 0    : StartType_Enum = "Boot"
      Case 1    : StartType_Enum = "System"
      Case 2    : StartType_Enum = "Auto"
      Case 3    : StartType_Enum = "Manual"
      Case 4    : StartType_Enum = "Disabled"
      Case Else : StartType_Enum = "Unknown"
    End Select
  End Function
  
  '********************************************************************************
  Public Property Get StartType()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    StartType = m_objSvc.StartType
  End Property
  
  '********************************************************************************
  Public Property Let StartType(ByVal intValue)
    Call SetStartType(intValue)
  End Property
  
  '********************************************************************************
  Public Function ValidateSecurityDescriptor(ByVal strDesiredSecurityDescriptor)
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    ValidateSecurityDescriptor = False
    
    Dim strProcedure : strProcedure = "ValidateSecurityDescriptor"
    Dim arrSecurityDescriptor
    Dim strSecurityDescriptor
    Dim dicSecurity
    
    'a sample security descriptor generated by running sc.exe sdshow <service>
    ' D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)
    
    'Refresh the security descriptor variable
    Call GetSecurityDescriptor()
    
    Call LogIt("Validate the " & m_strName & " service security descriptor.", "" & strProcedure, LogTypeInfo)
    Call LogIt("  Current security descriptor is: " & m_strSecurityDescriptor, "" & strProcedure, LogTypeInfo)
    Call LogIt("  Required security descriptor  : " & strDesiredSecurityDescriptor, "" & strProcedure, LogTypeInfo)
    
    Set dicSecurity = New clsDictionary
    '>>>>>> Add error handling?
    
    'Split current descriptor and put into an array
    If StrIn(1, m_strSecurityDescriptor, "(") > 0 Then
      arrSecurityDescriptor = StrSplit(m_strSecurityDescriptor, "(", -1)
      For Each strSecurityDescriptor In arrSecurityDescriptor
        dicSecurity.Key(Trim(strSecurityDescriptor)) = Empty
      Next
    End If
    
    'Split required descriptor and check for entries in the array
    If StrIn(1, strDesiredSecurityDescriptor, "(") > 0 Then
      arrSecurityDescriptor = StrSplit(strDesiredSecurityDescriptor, "(", -1)
      For Each strSecurityDescriptor In arrSecurityDescriptor
        If dicSecurity.Exists(strSecurityDescriptor) Then
          Call dicSecurity.Remove(strSecurityDescriptor)
        End If
      Next
    End If
    
    'Remove blank entries from the array
    For Each strSecurityDescriptor In dicSecurity.Keys
      strSecurityDescriptor = Trim(strSecurityDescriptor)
      If IsNullOrEmpty(strSecurityDescriptor) Then
        Call dicSecurity.Remove(strSecurityDescriptor)
      End If
    Next
    
    If dicSecurity.Count = 0 Then
      Call LogIt("The " & m_strName & " service security descriptor passed validation.", "" & strProcedure, LogTypeInfo)
      ValidateSecurityDescriptor = True
    Else
      Call LogIt("The " & m_strName & " service security descriptor failed validation.", "" & strProcedure, LogTypeError)
    End If
    
    If IsObject(dicSecurity) Then Set dicSecurity = Nothing
  End Function
End Class

'********************************************************************************
Private Sub Service_Delete(ByVal strComputer, ByVal strService)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  If Not Service_Check_Exists(strComputer, strService) Then Exit Sub
  
  Dim strProcedure : strProcedure = "Service_Delete"
  Dim objSvc
  
  'Connect to the service
  Set objSvc = New clsService
  If objSvc.Connect(strService) Then
    'Stop the service
    objSvc.State = ADS_SVC_STOPPED
  End If
  If IsObject(objSvc) Then Set objSvc = Nothing

  Dim strReturn
  Dim arrReturn
  Dim blnSuccess
  
  blnSuccess = False
  
  Call LogIt("Deleting the " & strService & " service.", "" & strProcedure, LogTypeInfo)
  
  If g_objProcess.Exec("sc.exe delete " & strService, "0") Then
    arrReturn = g_objProcess.Output
    For Each strReturn In arrReturn
      If Not IsNullOrEmpty(strReturn) Then
        If StrIn(1, strReturn, "[SC] DeleteService SUCCESS") > 0 Then
          blnSuccess = True
          Exit For
        End If
      End If
    Next
  End If
  
  Dim intCountLoop : intCountLoop = 0
  Call LogIt("Pausing 2 minutes or until the " & strService & " service no longer exists.", "" & strProcedure, LogTypeInfo)
  Do
    If Not Service_Check_Exists(g_objOS.Name, strService) Then
      Exit Do
    End If
    WScript.Sleep 2000
    If intCountLoop > 60 Then
      Exit Do
    End If
    intCountLoop = intCountLoop + 1
  Loop
  
  If blnSuccess Then
    Call LogIt("  Deleted the " & strService & " service.", "" & strProcedure, LogTypeInfo)
  Else
     Call LogIt("  Could not delete the " & strService & " service.", "" & strProcedure, LogTypeInfo)
  End If
End Sub

'********************************************************************************
Private Function Service_Check_Exists(ByVal strComputer, ByVal strService)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  Dim strProcedure : strProcedure = "Service_Check_Exists"
  Dim objComputer, objService
  'Dim blnFound
  
  Service_Check_Exists = False
  'blnFound = False
  
  Call LogIt("Checking if the " & strService & " service exists.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
  
  If Not ObjectRef_Get(objComputer, "WinNT://" & strComputer & ", computer") Then
    Call LogIt("  Failed to connect to the ADSI computer object.", "" & strProcedure, LogTypeError)
    Exit Function
  End If
  
  If IsObject(objComputer) Then
    Do
      objComputer.Filter = Array("Service")
      For Each objService In objComputer
        If LCase(objService.Name) = LCase(strService) Then
          Call LogIt("  Found service: " & objService.Name, "" & strProcedure, LogTypeInfo + LogTypeVerbose)
          Service_Check_Exists = True
          Exit Do
        End If
      Next
      Call LogIt("  The service does not appear to exist.", "" & strProcedure, LogTypeInfo + LogTypeVerbose)
      Exit Do
    Loop
  End If
  
  If IsObject(objService) Then Set objService = Nothing
  If IsObject(objComputer) Then Set objComputer = Nothing
End Function

'********************************************************************************
Private Function SvcStartTypeIntToString(ByVal intSvcStartType)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  SvcStartTypeIntToString = Empty
  
  Select Case UCase(intSvcStartType)
    Case 0 : SvcStartTypeIntToString = "ADS_SVC_START_BOOT"
    Case 1 : SvcStartTypeIntToString = "ADS_SVC_START_SYSTEM"
    Case 2 : SvcStartTypeIntToString = "ADS_SVC_START_AUTO"
    Case 3 : SvcStartTypeIntToString = "ADS_SVC_START_DEMAND"
    Case 4 : SvcStartTypeIntToString = "ADS_SVC_DISABLED"
  End Select
End Function

'********************************************************************************
Private Function SvcStartType_ReverseEnum(ByVal strSvcStartMode)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  SvcStartType_ReverseEnum = Empty
  
  Select Case UCase(strSvcStartMode)
    Case "ADS_SVC_START_BOOT"       : SvcStartType_ReverseEnum = 0
    Case "ADS_SVC_START_SYSTEM"     : SvcStartType_ReverseEnum = 1
    Case "ADS_SVC_START_AUTO"       : SvcStartType_ReverseEnum = 2
    Case "ADS_SVC_START_DEMAND"     : SvcStartType_ReverseEnum = 3
    Case "ADS_SVC_DISABLED"         : SvcStartType_ReverseEnum = 4
  End Select
End Function

'********************************************************************************
Private Function SvcState_ReverseEnum(ByVal strSvcState)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  SvcState_ReverseEnum = Empty
  
  Select Case UCase(strSvcState)
    Case "ADS_SVC_STOPPED"          : SvcState_ReverseEnum = 1
    Case "ADS_SVC_START_PENDING"    : SvcState_ReverseEnum = 2
    Case "ADS_SVC_STOP_PENDING"     : SvcState_ReverseEnum = 3
    Case "ADS_SVC_RUNNING"          : SvcState_ReverseEnum = 4
    Case "ADS_SVC_CONTINUE_PENDING" : SvcState_ReverseEnum = 5
    Case "ADS_SVC_PAUSE_PENDING"    : SvcState_ReverseEnum = 6
    Case "ADS_SVC_PAUSED"           : SvcState_ReverseEnum = 7
    Case "ADS_SVC_ERROR"            : SvcState_ReverseEnum = 8
  End Select
End Function

'********************************************************************************
Private Function SvcStateIntToString(ByVal intSvcState)
  If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
  
  SvcStateIntToString = Empty
  
  Select Case UCase(intSvcState)
    Case 1 : SvcStateIntToString = "ADS_SVC_STOPPED"
    Case 2 : SvcStateIntToString = "ADS_SVC_START_PENDING"
    Case 3 : SvcStateIntToString = "ADS_SVC_STOP_PENDING"
    Case 4 : SvcStateIntToString = "ADS_SVC_RUNNING"
    Case 5 : SvcStateIntToString = "ADS_SVC_CONTINUE_PENDING"
    Case 6 : SvcStateIntToString = "ADS_SVC_PAUSE_PENDING"
    Case 7 : SvcStateIntToString = "ADS_SVC_PAUSED"
    Case 8 : SvcStateIntToString = "ADS_SVC_ERROR"
    'Case Else  : SvcStateIntToString = "Error"
  End Select
End Function
