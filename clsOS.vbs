Dim strFileScript : strFileScript = "clsOS.vbs"
'***********************************************************************
'File:     clsOS.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 5/3/2014
'
'          This script file is based on the Configuration Manager Health Check Tool
'          (http://configmgrclienthtc.codeplex.com/)
'
'Notes:    
'
'Requires: g_objWshShell
'          g_objError
'          g_objWMIService
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


'*******************************************************************
Class clsOS
  Private m_arrProductSuite
  Private m_blnHasTimeZoneBiasBeenChecked
  Private m_blnIs64Bit
  Private m_blnIsServer
  Private m_intTimeZoneBias
  Private m_strADSite
  Private m_dblBuildNumber
  Private m_strCaption
  Private m_strConfigurationNamingContext
  Private m_strDir_WBEM_Repository
  Private m_strDir_Win
  Private m_strDistinguishedName
  Private m_strDrv_System
  Private m_strDNSDomain
  Private m_strDomain
  Private m_strDomainRole
  Private m_strLocale
  Private m_strName
  Private m_strOSType
  Private m_strOU
  Private m_strProductType
  Private m_intServicePack
  Private m_dblVersion
  
  '***********************************************************************
  Private Sub Class_Initialize()
    m_blnHasTimeZoneBiasBeenChecked = False
    m_blnIs64Bit = Empty
    
    Call Name()
    Call Dir_Win()
    Call Drv_System()
    Call Version()
    Call ServicePack()
    Call ProductType()
    Call Caption()
    Call BuildNumber()
  End Sub
  
  '***********************************************************************
  Private Sub Class_Terminate()
  End Sub
  
  '***********************************************************************
  Public Function ADSite()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    m_strADSite = Empty
    
    Call LogIt("Reading the system's AD site name.", "ADSite", LogTypeInfo + LogTypeVerbose)
    
    'Get the name of the AD site that this computer is in
    On Error Resume Next
    Call g_objReg.Read("HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters\DynamicSiteName", m_strADSite)
    If IsNullOrEmpty(m_strADSite) Then
      Call LogIt("  Could not read DynamicSiteName from the registry.", "ADSite", LogTypeError)
      
      Call g_objReg.Read("HKLM\Software\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Site-Name", m_strADSite)
      If IsNullOrEmpty(m_strADSite) Then
        Call LogIt("  Could not read Site-Name from the registry.", "ADSite", LogTypeError)
      End If
    End If
    
    Call LogIt("  The system's  AD site name is " & m_strADSite, "ADSite", LogTypeInfo + LogTypeVerbose)
    
    ADSite = m_strADSite
  End Function
  
  '***********************************************************************
  Public Function BuildNumber()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_dblBuildNumber) Then
      Dim strValue
      Call g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber", strValue)
      If Not IsNullOrEmpty(strValue) Then m_dblBuildNumber = CDbl(strValue)
    End If
    BuildNumber = m_dblBuildNumber
  End Function
 
  '***********************************************************************
  Public Function Caption()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strCaption) Then
      Call g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProductName", m_strCaption)
    End If
    Caption = m_strCaption
  End Function
  
  '***********************************************************************
'   Public Function ConfigurationNamingContext()
'     If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
'     
'     Dim objRootDSE
'     
'     If IsNullOrEmpty(m_strConfigurationNamingContext) Then
'       If ObjectRef_Get(objRootDSE, "LDAP://RootDSE") Then
'         m_strConfigurationNamingContext = objRootDSE.Get("configurationNamingContext")
'         Set objRootDSE = Nothing
'       End If
'     End If
'     
'     ConfigurationNamingContext = m_strConfigurationNamingContext
'   End Function

  '***********************************************************************
  Public Function Dir_WBEM_Repository()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strDir_WBEM_Repository) Then
      'Get the directory where the WMI repository exists
      g_objError.Clear
      If Not g_objReg.Read("HKLM\Software\Microsoft\WBEM\CIMOM\Repository Directory", m_strDir_WBEM_Repository) Then
        Call g_dicSystemHealth.Add("GEN", errSYS_ProblemWithScriptObject)
      End If
      m_strDir_WBEM_Repository = g_objWshShell.ExpandEnvironmentStrings(m_strDir_WBEM_Repository)
    End If
    Dir_WBEM_Repository = m_strDir_WBEM_Repository
  End Function
  
  '***********************************************************************
  Public Function Dir_Win()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    'Get the location of the WINNT or Windows folder
    If IsNullOrEmpty(m_strDir_Win) Then
      g_objError.Clear
      m_strDir_Win = g_objWshShell.ExpandEnvironmentStrings("%WinDir%")
      If IsNullOrEmpty(m_strDir_Win) Or g_objError.Check() Then
        Call LogIt("  Could not read WinDir from the System environment. " & g_objError.Message, "Dir_Win", LogTypeError)
      End If
    End If
    Dir_Win = m_strDir_Win
  End Function
  
  '***********************************************************************
  Public Function DistinguishedName()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    'Get the distinguished name of the computer.
    'This is formatted as CN=Computer,CN=Container,DC=Domain,DC=Suffix
    If IsNullOrEmpty(m_strDistinguishedName) Then
      Call g_objReg.Read("HKLM\Software\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Distinguished-Name", m_strDistinguishedName)
    End If
    DistinguishedName = m_strDistinguishedName
  End Function
  
  '***********************************************************************
  Public Function Domain()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strDomain) Then
      If g_objReg.Read("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Domain", m_strDomain) Then
        If StrIn(1, m_strDomain, ".") > 0 Then
          m_strDomain = Left(m_strDomain, StrIn(1, m_strDomain, ".") - 1)
        End If
      End If
    End If
    
    Domain = m_strDomain
  End Function
  
  '***********************************************************************
  Public Function DNSDomain()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strDNSDomain) Then
      If g_objReg.Read("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Domain", m_strDNSDomain) Then
        If StrIn(1, m_strDNSDomain, ".") > 0 Then
          m_strDNSDomain = "DC=" & StrReplace(m_strDNSDomain, ".", ",DC=", 1, -1)
        End If
      End If
    End If
    DNSDomain = m_strDNSDomain
  End Function
   
  '***********************************************************************
'   Public Property Get DomainRole()
'     DomainRole = m_strDomainRole
'   End Property
' 
'   Public Property Let DomainRole(strValue)
'     m_strDomainRole = Trim(strValue)
'   End Property
 
  '***********************************************************************
  Public Function Drv_System()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strDrv_System) Then
      'Get the drive where the operating system is installed
      g_objError.Clear
      m_strDrv_System = g_objWshShell.ExpandEnvironmentStrings("%SystemDrive%")
      If IsNullOrEmpty(m_strDrv_System) Or g_objError.Check() Then
        Call LogIt("  Could not read SystemDrive from the System environment. " & g_objError.Message, "Drv_System", LogTypeError)
      End If
    End If
    Drv_System = m_strDrv_System
  End Function
  
  '***********************************************************************
  Public Property Get EnvItem(ByVal strEnvironment, ByVal strItem)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    EnvItem = Empty
    
    If StrIn(1, "System, Process", strEnvironment) = 0 Then
      Call LogIt("Bad environment specified: " & strEnvironment & g_objError.Message, "EnvItem", LogTypeError)
      Exit Property
    End If
    
    Dim objEnv
    Dim strValue
    
    g_objError.Clear
    Set objEnv = g_objWshShell.Environment(strEnvironment)
    If (Not IsObject(objEnv)) Or g_objError.Check() Then
      Call LogIt("Could not create an object reference to the " & strEnvironment & " environment. " & g_objError.Message, "EnvItem", LogTypeError)
      Call EventLog_Write(EVENT_ERROR, "Could not create an object reference to the " & strEnvironment & " environment. " & g_objError.Message)
      Exit Property
    End If
    
    g_objError.Clear
    strValue = objEnv(strItem)
    If IsNullOrEmpty(strValue) Or g_objError.Check() Then
      Call LogIt("Could not read the " & strItem & " from the " & strEnvironment & " environment. " & g_objError.Message, "EnvItem", LogTypeError)
      Call EventLog_Write(EVENT_ERROR, "Could not read the " & strItem & " from the " & strEnvironment & " environment. " & g_objError.Message)
    Else
    '  Call LogIt("Updated the " & strItem & " in the " & strEnvironment & " environment.", "EnvItem", LogTypeInfo)
    '  Call EventLog_Write(EVENT_INFO, "Updated the " & strItem & " in the " & strEnvironment & " environment.")
      
      EnvItem = strValue
    End If
    If IsObject(objEnv) Then Set objEnv = Nothing
  End Property
  
  '***********************************************************************
  Public Property Let EnvItem(ByVal strEnvironment, ByVal strItem, ByVal strValue)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    If StrIn(1, "System, Process", strEnvironment) = 0 Then
      Call LogIt("Bad environment specified: " & strEnvironment & g_objError.Message, "EnvItem", LogTypeError)
      Exit Property
    End If
    
    Dim objEnv
    
    g_objError.Clear
    Set objEnv = g_objWshShell.Environment(strEnvironment)
    If (Not IsObject(objEnv)) Or g_objError.Check() Then
      Call LogIt("Could not create an object reference to the " & strEnvironment & " environment. " & g_objError.Message, "EnvItem", LogTypeError)
      Call EventLog_Write(EVENT_ERROR, "Could not create an object reference to the " & strEnvironment & " environment. " & g_objError.Message)
      Exit Property
    End If
  
    If IsNullOrEmpty(strValue) Then
      g_objError.Clear
      objEnv.Remove(strItem)
      If g_objError.Check() Then
        Call LogIt("Could not remove the " & strItem & " from the " & strEnvironment & " environment. " & g_objError.Message, "EnvItem", LogTypeError)
        Call EventLog_Write(EVENT_ERROR, "Could not remove the " & strItem & " from the " & strEnvironment & " environment. " & g_objError.Message)
      Else
        Call LogIt("Removed the " & strItem & " from the " & strEnvironment & " environment.", "EnvItem", LogTypeInfo + LogTypeVerbose)
        Call EventLog_Write(EVENT_INFO, "Removed the " & strItem & " from the " & strEnvironment & " environment.")
      End If
    Else
      g_objError.Clear
      objEnv(strItem) = strValue
      If g_objError.Check() Then
        Call LogIt("Could not update the " & strItem & " in the " & strEnvironment & " environment. " & g_objError.Message, "EnvItem", LogTypeError)
        Call EventLog_Write(EVENT_ERROR, "Could not update the " & strItem & " in the " & strEnvironment & " environment. " & g_objError.Message)
      Else
        Call LogIt("Updated the " & strItem & " in the " & strEnvironment & " environment.", "EnvItem", LogTypeInfo + LogTypeVerbose)
        Call EventLog_Write(EVENT_INFO, "Updated the " & strItem & " in the " & strEnvironment & " environment.")
      End If
    End If
    If IsObject(objEnv) Then Set objEnv = Nothing
    
  End Property
  
  '***********************************************************************
'   Public Function GetAnyDCName()
'     GetAnyDCName = CreateObject("ADSystemInfo").GetAnyDCName
'   End Function
  
  
  '***********************************************************************
  Public Function Is64Bit()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_blnIs64Bit) Then
      If EnvItem("System", "Processor_Architecture") = "x86" Then
      'If g_objWshShell.ExpandEnvironmentStrings("%Processor_Architecture%") = "x86" Then
        m_blnIs64Bit = False
      Else
        m_blnIs64Bit = True
      End If
    End If
    Is64Bit = m_blnIs64Bit
  End Function

  '***********************************************************************
  Public Function IsDC()
    IsDC = g_objFSO.FolderExists(m_strDir_Win & "\SYSVOL")
  End Function
  
  '***********************************************************************
  Public Function IsDomainMember()
    IsDomainMember = CreateObject("Shell.Application").GetSystemInformation("IsOS_DomainMember")
  End Function
  
  '***********************************************************************
  Public Function IsRebootPending()
  '
  ' Checks if a reboot is needed.
  '***********************************************************************
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    IsRebootPending = False
    
    Dim objSystemInfo
    Dim intCount
    Dim strKey
    Dim strValueName
    Dim arrValueNames, arrValues
    Dim arrSubKeys
    Dim i, intReturn
    Dim blnFound
  
    intCount = 0
    
    Call LogIt("Checking if this computer is pending a reboot.", "IsRebootPending", LogTypeInfo)
    
    Call LogIt("  Checking the Windows Update Agent.", "IsRebootPending", LogTypeInfo + LogTypeVerbose)
    
    If ObjectRef_Create(objSystemInfo, "Microsoft.Update.SystemInfo") Then
      If objSystemInfo.RebootRequired Then
        intCount = intCount + 1
        Call LogIt("    The Windows Update Agent reports a reboot is required.", "IsRebootPending", LogTypeWarning + LogTypeVerbose)
      Else
        Call LogIt("    The Windows Update Agent reports no reboot is required.", "IsRebootPending", LogTypeInfo + LogTypeVerbose)
      End If
      Set objSystemInfo = Nothing
    Else
      Call LogIt("    Can't use the Windows Update Agent to determine if a reboot is required.", "IsRebootPending", LogTypeWarning)
    End If
    
    Call LogIt("  Checking the PendingFileRenameOperations entry in the registry.", "IsRebootPending", LogTypeInfo + LogTypeVerbose)
    
    strKey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager"
    blnFound = False
    If g_objReg.ValueNames(strKey, arrValueNames) Then
      For Each strValueName In arrValueNames
        If StrCompare(strValueName, "PendingFileRenameOperations") = 0 Then
          blnFound = True
          Exit For
        End If
      Next
      
      If blnFound Then
        If g_objReg.Read(strKey & "\PendingFileRenameOperations", arrValues) Then
          If IsArray(arrValues) Then
            intCount = intCount + 1
            Call LogIt("    The existance of the PendingFileRenameOperations registry entry indicates a reboot is required.", "IsRebootPending", LogTypeWarning + LogTypeVerbose)
            For i = 0 To UBound(arrValues)
              Call LogIt("    Entry: '" & arrValues(i) & "'", "IsRebootPending", LogTypeInfo + LogTypeVerbose)
            Next
          End If
        End If
      Else
        Call LogIt("    The PendingFileRenameOperations registry entry is missing.", "IsRebootPending", LogTypeInfo + LogTypeVerbose)
      End If
    End If
    
    If intCount > 0 Then
      IsRebootPending = True
      Call LogIt("  This computer needs to be rebooted to complete one or more installations.", "IsRebootPending", LogTypeWarning)
    Else
      Call LogIt("  This computer is not pending a reboot.", "IsRebootPending", LogTypeInfo)
    End If
  End Function
  
  '***********************************************************************
  Public Function IsServer()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_blnIsServer) Then
      If IsNullOrEmpty(m_strProductType) Then Call ProductType()
      If IsNullOrEmpty(m_strProductType) Then Exit Function
    
      'Check OS for 'Server'.
      If StrIn(1, m_strProductType, "WinNT") > 0 Then
        m_blnIsServer = False
      Else
        m_blnIsServer = True
      End If
    End If
    IsServer = m_blnIsServer
  End Function

  '***********************************************************************
  Public Function Locale()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strLocale) Then
      'Get the locale of the computer
      g_objError.Clear
      m_strLocale = GetLocale()
      If IsNullOrEmpty(m_strLocale) Or g_objError.Check() Then
        Call LogIt("  Could not read the Locale. " & g_objError.Message, "Locale", LogTypeError)
      End If
    End If
    Locale = m_strLocale
  End Function

  '***********************************************************************
  Public Function Name()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strName) Then
      'Get the name of the computer
      g_objError.Clear
      m_strName = g_objWshShell.ExpandEnvironmentStrings("%ComputerName%")
      If IsNullOrEmpty(m_strName) Or g_objError.Check() Then
        Call LogIt("  Could not read ComputerName from the System environment. " & g_objError.Message, "Name", LogTypeError)
      End If
    End If
    Name = m_strName
  End Function
  
  '***********************************************************************
  Public Function OSType()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strOSType) Then
      'http://technet.microsoft.com/en-us/library/cc782360(WS.10).aspx
      '##Confirm others
      'Confirmed...
      ' 2003 server      ServerNT
      ' Win7             WinNT
      ' XP               WinNT
      ' 2008 R2 server   Domain Controller=LanmanNT
      ' 2008 R2 server   Member Server=ServerNT
      ' 2012 R2 server   Member Server=ServerNT
      Select Case UCase(ProductType)
        Case "WINNT"
          OSType = "WORKSTATION"
        Case "SERVERNT", "LANMANNT"
          OSType = "SERVER"
        Case Else
          OSType = "UNKNOWN"
      End Select
    End If
    OSType = m_strOSType
  End Function
  
  '***********************************************************************
  Public Function OU()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    Dim strTemp
    
    strTemp = "CN=" & Name() & ","
    m_strOU = DistinguishedName()
    If StrIn(1, m_strOU, strTemp) > 0 Then
      m_strOU = Mid(m_strOU, Len(strTemp) + 1)
    End If
    
    OU = m_strOU
  End Function
  
  '***********************************************************************
  Public Function ProcessorArchitecture()
    ProcessorArchitecture = CreateObject("Shell.Application").GetSystemInformation("ProcessorArchitecture")
  End Function
  
  '***********************************************************************
  Public Function ProductSuite()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_arrProductSuite) Then
      'need to adjust to handle array values (REG_MULTI_SZ)
      Call g_objReg.Read("HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductSuite", m_arrProductSuite)
    End If
    ProductSuite = m_arrProductSuite
  End Function
  
  '***********************************************************************
  Public Function ProductType()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_strProductType) Then
      Call g_objReg.Read("HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType", m_strProductType)
    End If
    ProductType = m_strProductType
  End Function
  
  '***********************************************************************
  Public Function ServicePack()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_intServicePack) Then
      Dim strValue
      If g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CSDVersion", strValue) Then
        If StrIn(1, strValue, "Service Pack ") > 0 Then
          strValue = StrReplace(strValue, "Service Pack ", "", 1, -1)
        End If
      End If
      If Not IsNullOrEmpty(strValue) Then m_intServicePack = CInt(strValue)
    End If
    ServicePack = m_intServicePack
  End Function
  
  '***********************************************************************
  Public Function TimeZoneBias()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If Not m_blnHasTimeZoneBiasBeenChecked Then
      Dim colOS, objOS, colTZ, objTZ
    
      Call LogIt("Querying for current time zone offset.", "TimeZoneBias", LogTypeInfo)
      
      If Not ExecQuery(colOS, g_objWMIService, "SELECT DayLightinEffect FROM Win32_ComputerSystem") Then Exit Function
      If Not ExecQuery(colTZ, g_objWMIService, "SELECT Bias, DaylightBias, StandardBias FROM Win32_TimeZone") Then Exit Function
    
      If (colOS.Count < 1) Or (colTZ.Count < 1) Then
        Call LogIt("  No instances were returned when querying Win32_ComputerSystem and/or Win32_TimeZone.", "TimeZoneBias", LogTypeError)
      Else
        For Each objOS In colOS
          For Each objTZ In colTZ
            m_intTimeZoneBias = objTZ.Bias
            If objOS.DayLightinEffect Then
              m_intTimeZoneBias = m_intTimeZoneBias - objTZ.DaylightBias
            Else
              m_intTimeZoneBias = m_intTimeZoneBias - objTZ.StandardBias
            End If
          Next
        Next
        Call LogIt("  Current timezone offset is " & m_intTimeZoneBias & ".", "TimeZoneBias", LogTypeInfo + LogTypeVerbose)
      End If
      m_blnHasTimeZoneBiasBeenChecked = True
    End If
    TimeZoneBias = m_intTimeZoneBias
  End Function
  
  '***********************************************************************
  Public Function Version()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
    If IsNullOrEmpty(m_dblVersion) Then
      If StrIn(1, Caption(), "Windows 10") > 0 Then
        Dim intMajor
        Dim intMinor
        Call g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentMajorVersionNumber", intMajor)
        Call g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentMinorVersionNumber", intMinor)
        m_dblVersion = CDbl(intMajor & "." & intMinor)
      Else
        Dim strValue
        Call g_objReg.Read("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion", strValue)
        If Not IsNullOrEmpty(strValue) Then m_dblVersion = CDbl(strValue)
      End If
    End If
    Version = m_dblVersion
  End Function
End Class
