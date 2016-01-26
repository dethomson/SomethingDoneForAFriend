Dim strFileScript : strFileScript = "clsRegistry.vbs"
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
'Requires: g_objWshShell
'          g_objLog
'          g_objError
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


'It's a bummer to have to code using legacy methods, but we can't
'expect WMI to be available at all times.  Not using WMI methods
'has us missing out on easier methods and better error reporting.

'Constants used to specify registry entry types
  Const REG_SZ                  = 1
  Const REG_EXPAND_SZ           = 2
  Const REG_BINARY              = 3
  Const REG_DWORD               = 4
  Const REG_MULTI_SZ            = 7
    
'Constants used to specify which registry hive to write to
  Const HKEY_CLASSES_ROOT       = &H80000000
  Const HKEY_CURRENT_USER       = &H80000001
  Const HKEY_LOCAL_MACHINE      = &H80000002
  Const HKEY_USERS              = &H80000003
  Const HKEY_CURRENT_CONFIG     = &H80000005
  
  Const KEY_QUERY_VALUE         = &H0001  'Required to query the values of a registry key.
  Const KEY_SET_VALUE           = &H0002  'Required to create, delete, or set a registry value.
  Const KEY_CREATE_SUB_KEY      = &H0004  'Required to create a subkey of a registry key.
  Const KEY_ENUMERATE_SUB_KEYS  = &H0008  'Required to enumerate the subkeys of a registry key.
  Const KEY_NOTIFY              = &H0016  'Required to request change notifications for a registry key or for subkeys of a registry key.
  Const KEY_CREATE              = &H0032  'Required to create a registry key.
  Const DELETE                  = &H10000 'Required to delete a registry key.
  Const READ_CONTROL            = &H20000 'Combines the STANDARD_RIGHTS_READ, KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY values.
  Const WRITE_DAC               = &H40000 'Required to modify the DACL in the object's security descriptor.
  Const WRITE_OWNER             = &H80000 'Required to change the owner in the object's security descriptor.
  
'*************************************************************************
Class clsRegistry
  Private m_objReg
  Private m_blnIsConnected
  Private m_objProcess
  'Private m_dblLastError
  
  '***********************************************************************
  Private Sub Class_Initialize()
    m_blnIsConnected = Connect()
    Set m_objProcess = New clsProcess
  End Sub
  
  '***********************************************************************
  Private Sub Class_Terminate()
    If IsObject(m_objReg) Then Set m_objReg = Nothing
    If IsObject(m_objProcess) Then Set m_objProcess = Nothing
  End Sub
  
  '***********************************************************************
  Public Function Access_Check(ByVal strKeyPath, ByVal lngAccessRequired, ByRef blnHasAccess)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Access_Check = False
    
    Dim intReturn
    Dim dblHive
    
    If Not Connection_Check Then Exit Function
    
    dblHive = HiveReverseEnum(Left(strKeyPath, StrIn(1, strKeyPath, "\") - 1))
    strKeyPath = Mid(strKeyPath, StrIn(1, strKeyPath, "\") + 1)
    
    g_objError.Clear
    intReturn = m_objReg.CheckAccess(dblHive, strKeyPath, lngAccessRequired, blnHasAccess)
    If (intReturn <> 0) Or g_objError.Check() Then
      Call LogIt("  Could not check access for " & strKeyPath & ". Return code: " & intReturn & ". " & g_objError.Message, "Access_Check", LogTypeError)
      If intReturn = 2 Then
        Call LogIt("  The registry key does not exist." , "Access_Check", LogTypeError)
      End If
    Else
      Access_Check = True
    End If
    
  End Function
  
  '***********************************************************************
  Public Function Connect()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    'Create an object reference to WMI Reg
    Connect = ObjectRef_Get(m_objReg, "winmgmts:{impersonationLevel=impersonate}!\root\default:StdRegProv")
    If IsObject(m_objReg) Then m_blnIsConnected = True
  End Function
  
  '***********************************************************************
  Private Function Connection_Check()
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Connection_Check = False
    
    Dim intReturn
    Dim strValue
    
    If Not IsObject(m_objReg) Then
      If Not Connect() Then Exit Function
    End If
    
    g_objError.Clear()
    intReturn = m_objReg.GetStringValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "CurrentVersion", strValue)
    If (intReturn <> 0) Or g_objError.Check() Then
      m_blnIsConnected = False
      If g_objError.ErrDec = -2147417848 Then
        m_blnIsConnected = Connect()
      End If
    Else
      m_blnIsConnected = True
    End If
    
    Connection_Check = m_blnIsConnected
  End Function
  
  '***********************************************************************
  Public Function IsConnected()
    IsConnected = Connection_Check
  End Function
  
  '***********************************************************************
  Public Function Delete(ByVal strPath)
    'Uses legacy methods to delete the specified key or value
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Delete = False
    
    If Right(strPath, 1) = "\" Then
      Delete = Key_Delete(strPath)
      Exit Function
    End If
    
    Delete = Value_Delete(strPath)
    
  End Function
  
  '***********************************************************************
  Public Function Exists(ByVal strPath)
    
    Exists = False
    
    Dim strKeyName, strValueName
    Dim strCommand
    
    On Error Resume Next
    
    If StrIn(1, "5.1, 6.0, 6.1", CStr(g_objOS.Version)) > 0 Then
      'If the key or value does not exist, the RegRead method will error
      'with one of the listed values
      Dim varReturn
      
      g_objError.Clear
      varReturn = g_objWshShell.RegRead(strPath)
      If g_objError.Check() Then
        Select Case g_objError.ErrDec
          Case -2147024893, -2147024894  ' Invalid root in registry
            'Call LogIt("Could not read the registry. " & g_objError.Message, "Exists", LogTypeInfo + LogTypeVerbose)
          Case Else
            Call LogIt("Could not read the registry. " & g_objError.Message, "Exists", LogTypeInfo + LogTypeVerbose)
        End Select
      Else
        Exists = True
      End If
    
      Exit Function
    End If
    
    'The RegRead method above doesn't seem to work for Server 2003, so we have to resort to
    'other methods to see if the key or value exists.
    If Right(strPath, 1) = "\" Then
      strKeyName = Left(strPath, Len(strPath) - 1)
      strCommand = "reg.exe QUERY " & Chr(34) & strKeyName & Chr(34) & " /VE"
    Else
      strKeyName = Left(strPath, InstrRev(strPath, "\", -1, vbTextCompare) - 1)
      strValueName = Mid(strPath, InstrRev(strPath, "\", -1, vbTextCompare) + 1)
      strCommand = "reg.exe QUERY " & Chr(34) & strKeyName & Chr(34) & " /V " & Chr(34) & strValueName & Chr(34)
    End If
    
    If m_objProcess.Exec(strCommand, "0, 1") Then
      If m_objProcess.ExitCode = 0 Then
        Exists = True
      End If
    End If
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
  End Function

  '***********************************************************************
  Public Function Export(ByVal strKeyPath, ByVal strRegFile)
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Export = False
    
    Dim strCommand
    Dim intReturn
    
    Call File_Delete(strRegFile)
    
    strKeyPath = HiveExpand(strKeyPath)
    
    Call LogIt("Exporting registry data '" & strKeyPath & "' to file.", "Export", LogTypeInfo + LogTypeVerbose)
    
    strCommand = "regedit.exe /s /e " & Chr(34) & strRegFile & Chr(34) & " " & Chr(34) & strKeyPath & Chr(34)
    If m_objProcess.Run(strCommand, True, 0) Then
      Call LogIt("  Exported the '" & strKeyPath & "' registry data.", "Export", LogTypeInfo + LogTypeVerbose)
      Export = True
    Else
      Call LogIt("  Failed to export the '" & strKeyPath & "' from the registry. The exit code was " & intReturn, "Export", LogTypeError)
    End If
    
  End Function
    
  '***********************************************************************
  Private Function Key_Delete(ByVal strKeyPath)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Key_Delete = False
    
    'Using legacy methods to delete a registry key can be done by running
    '"reg.exe delete HKLM\Software\AAA /F".
    
    'This will delete the AAA registry key whether or not any subkeys exist.
    
    'This assumes we're running XP or greater and reg.exe can be found somewhere on
    'the system in a folder listed in the %PATH% environment variable.
    
    Dim strCommand
    Dim intReturn
    
    If Right(strKeyPath, 1) <> "\" Then
      strKeyPath = strKeyPath & "\"
    End If
    
    If Not Exists(strKeyPath) Then
      Key_Delete = True
      Exit Function
    End If
    
    If Right(strKeyPath, 1) = "\" Then
      strKeyPath = Left(strKeyPath, Len(strKeyPath) - 1)
    End If
    
    Call LogIt("Delete registry key (" & strKeyPath & ").", "Key_Delete", LogTypeInfo + LogTypeVerbose)
    
    strCommand = "reg.exe DELETE " & Chr(34) & strKeyPath & Chr(34) & " /F"
    If m_objProcess.Run(strCommand, True, 0) Then
      Call LogIt("    Deleted the registry key (" & strKeyPath & ").", "Key_Delete", LogTypeInfo + LogTypeVerbose)
      Key_Delete = True
    Else
      Call LogIt("    Could not delete the registry key (" & strKeyPath & ").", "Key_Delete", LogTypeError)
    End If
    
  End Function
  
  '***********************************************************************
'   Public Property Get LastError()
'     m_dblLastError = m_dblLastError
'   End Property
'
'   Public Property Let LastError(ByVal varValue)
'     m_dblLastError = varValue
'   End Property
  
  '***********************************************************************
  Public Function Read(ByVal strPath, ByRef varValue)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Read = False
    
    Call LogIt("Reading the " & strPath & " value from the registry.", "Read", LogTypeInfo + LogTypeVerbose)
    
    On Error Resume Next
    g_objError.Clear
    varValue = g_objWshShell.RegRead(strPath)
    If g_objError.Check() Then
      Call LogIt("  Failed to read the " & strPath & " value from the registry. " & g_objError.Message, "Read", LogTypeError)
      varValue = Empty
    Else
      Read = True
    End If
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  End Function
  
  '***********************************************************************
  Public Function SubKeys(ByVal strKeyPath, ByRef arrSubKeys)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    SubKeys = False
    
    Dim dblHive
    
    dblHive = HiveReverseEnum(Left(strKeyPath, StrIn(1, strKeyPath, "\") - 1))
    strKeyPath = Mid(strKeyPath, StrIn(1, strKeyPath, "\") + 1)
    
    If SubKeys_GetWMI(dblHive, strKeyPath, arrSubKeys) Then
      SubKeys = True
      Exit Function
    End If
    If SubKeys_GetLegacy(HiveEnum(dblHive) & "\" & strKeyPath, arrSubKeys) Then
      SubKeys = True
    End If
    
  End Function
  
  '***********************************************************************
  Private Function SubKeys_GetLegacy(ByVal strKeyPath, ByRef arrSubKeys)
    'Uses legacy methods to return the immediate subkeys of the specified key
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    SubKeys_GetLegacy = False
    
    Dim strCommand
    Dim arrOutput
    Dim strLine
    Dim i
    Dim arrKeys()
    
    If IsArray(arrSubKeys) Then Erase arrSubKeys
    
    If Not Exists(strKeyPath & "\") Then Exit Function
    
    Call LogIt("Querying registry key (" & strKeyPath & ").", "SubKeys_GetLegacy", LogTypeInfo + LogTypeVerbose)
    
    strCommand = "reg.exe QUERY " & Chr(34) & strKeyPath & Chr(34)
    If Not m_objProcess.Exec(strCommand, 0) Then
      Exit Function
    End If
    
    arrOutput = m_objProcess.Output
    
    ReDim arrKeys(0)
    
    For i = 0 To UBound(arrOutput)
      strLine = Trim(arrOutput(i))
      If Not IsNullOrEmpty(strLine) Then
        If StrIn(1, strLine, strKeyPath & "\") > 0 Then
          strLine = Mid(strLine, InStrRev(strLine, "\", -1) + 1)
          arrKeys(UBound(arrKeys)) = strLine
          ReDim Preserve arrKeys(UBound(arrKeys) + 1)
        End If
      End If
    Next
    ReDim Preserve arrKeys(UBound(arrKeys) - 1)
    
    arrSubKeys = arrKeys
    
    SubKeys_GetLegacy = True
    
  End Function
  
  '***********************************************************************
  Private Function SubKeys_GetWMI(ByVal dblHive, ByVal strKeyPath, ByRef arrSubKeys)
    'Uses WMI methods to return the subkeys of the specified key
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    SubKeys_GetWMI = False
    
    Dim intReturn
    
    If IsArray(arrSubKeys) Then Erase arrSubKeys
    
    If Not Exists(HiveEnum(dblHive) & "\" & strKeyPath & "\") Then Exit Function
    
    If Not Connection_Check Then Exit Function
    
    g_objError.Clear
    intReturn = m_objReg.EnumKey(dblHive, strKeyPath, arrSubKeys)
    If (intReturn <> 0) Or g_objError.Check() Then
      Call LogIt("  Could not enumerate registry sub keys for (" & HiveEnum(dblHive) & "\" & strKeyPath & ". Return code: " & intReturn & ". " & g_objError.Message, "SubKeys_GetWMI", LogTypeError)
    Else
      Call LogIt("  Found " & UBound(arrSubKeys) + 1 & " registry sub keys for " & HiveEnum(dblHive) & "\" & strKeyPath, "SubKeys_GetWMI", LogTypeInfo + LogTypeVerbose)
      SubKeys_GetWMI = True
    End If
    
  End Function
  
  '***********************************************************************
  Public Function TypeCheck(ByVal strKeyPath, ByVal strValueName, ByVal varValueType)
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    TypeCheck = False
    
    Dim intReturn
    Dim intValueType_Current
    
    If Not IsNumeric(varValueType) Then varValueType = TypeReverseEnum(varValueType)
    
    If ValueType(strKeyPath, strValueName, intValueType_Current) Then
      If varValueType = intValueType_Current Then
        TypeCheck = True
        Call LogIt("The registry entry is the correct value type.", "TypeCheck", LogTypeInfo + LogTypeVerbose)
      Else
        Call LogIt("The registry entry is a " & TypeEnum(intValueType_Current) & " type, which is NOT correct.", "TypeCheck", LogTypeError)
      End If
    End If
    
  End Function
  
  '***********************************************************************
  Private Function Value_Delete(ByVal strValuePath)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Value_Delete = False
    
    Call LogIt("Deleting the " & strValuePath & " registry entry.", "Value_Delete", LogTypeInfo + LogTypeVerbose)
    
    On Error Resume Next
    g_objError.Clear
    g_objWshShell.RegDelete strValuePath
    If g_objError.Check() Then
      Select Case g_objError.ErrDec
        Case -2147024893, -2147024894
          Call LogIt("  The registry key or value (" & strValuePath & ") does not exist.", "Value_Delete", LogTypeError + LogTypeVerbose)
        Case Else
          Call LogIt("  Could not delete the value from the registry (" & strValuePath & "). " & g_objError.Message, "Value_Delete", LogTypeError)
      End Select
    Else
      Call LogIt("  Deleted the value from the registry (" & strValuePath & "). " & g_objError.Message, "Value_Delete", LogTypeInfo + LogTypeVerbose)
      Value_Delete = True
    End If
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
  End Function
  
  '***********************************************************************
  Public Function ValueNames(ByVal strKeyPath, ByRef arrValueNames)
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    ValueNames = False
    
    Dim dblHive
    
    If IsArray(arrValueNames) Then Erase arrValueNames
    
    dblHive = HiveReverseEnum(Left(strKeyPath, StrIn(1, strKeyPath, "\") - 1))
    strKeyPath = Mid(strKeyPath, StrIn(1, strKeyPath, "\") + 1)
    
    If ValueNames_GetWMI(dblHive,  strKeyPath, arrValueNames) Then
      ValueNames = True
      Exit Function
    End If
    If ValueNames_GetLegacy(HiveEnum(dblHive) & "\" & strKeyPath, arrValueNames) Then
      ValueNames = True
    End If
    
  End Function
  
  '***********************************************************************
  Private Function ValueNames_GetLegacy(ByVal strKeyPath, ByRef arrValueNames)
    'Uses legacy methods to return the immediate valuenames of the specified key
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    ValueNames_GetLegacy = False
    
    Dim strCommand
    Dim arrOutput
    Dim strLine
    Dim blnInValueList
    
    If IsArray(arrValueNames) Then Erase arrValueNames
    
    If Not Exists(strKeyPath & "\") Then Exit Function
    
    Call LogIt("Querying registry key (" & strKeyPath & ").", "ValueNames_GetLegacy", LogTypeInfo + LogTypeVerbose)
    
    strCommand = "reg.exe QUERY " & Chr(34) & strKeyPath & Chr(34)
    If Not m_objProcess.Exec(strCommand, 0) Then
      Exit Function
    End If
    
    arrOutput = m_objProcess.Output
    
    ReDim arrValueNames(0)
    
    For Each strLine In arrOutput
      strLine = Trim(strLine)
      If Not IsNullOrEmpty(strLine) Then
        If StrIn(1, strLine, strKeyPath) > 0 Then
          If StrCompare(strLine, strKeyPath) = 0 Then
            blnInValueList = True
          Else
            blnInValueList = False
          End If
        End If
        If blnInValueList Then
          strLine = Left(strLine, StrIn(strLine, " REG_", -1))
          strLine = Trim(strLine)
          arrValueNames(UBound(arrValueNames)) = strLine
          ReDim Preserve arrValueNames(UBound(arrValueNames) + 1)
        End If
      End If
    Next
    ReDim Preserve arrValueNames(UBound(arrValueNames) - 1)
    
    ValueNames_GetLegacy = True
    
  End Function
  
  '***********************************************************************
  Private Function ValueNames_GetWMI(ByVal dblHive, ByVal strKeyPath, ByRef arrValueNames)
    'Uses legacy methods to return the immediate subkeys of the specified key
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    ValueNames_GetWMI = False
    
    Dim intReturn
    Dim arrValueTypes
    
    If Not Connection_Check Then Exit Function
    
    If IsArray(arrValueNames) Then Erase arrValueNames
    
    g_objError.Clear
    intReturn = m_objReg.EnumValues(dblHive, strKeyPath, arrValueNames, arrValueTypes)
    If (intReturn <> 0) Or g_objError.Check() Then
      
      Exit Function
    Else
      ValueNames_GetWMI = True
    End If
    
  End Function
  
  '***********************************************************************
  Public Function ValueType(ByVal strKeyPath, ByVal strValueName, ByRef intValueType)
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    ValueType = False
    
    Dim strValueType
    Dim dblHive
    Dim strKeyPathForWMI
    
    intValueType = Empty
    
    If Not Exists(strKeyPath & "\" & strValueName) Then
      Exit Function
    End If
    
    dblHive = HiveReverseEnum(Left(strKeyPath, StrIn(1, strKeyPath, "\") - 1))
    strKeyPathForWMI = Mid(strKeyPath, StrIn(1, strKeyPath, "\") + 1)
    
    If Value_TypeWMI(dblHive, strKeyPathForWMI, strValueName, intValueType) Then
      ValueType = True
      Exit Function
    End If
    
    If Value_TypeLegacy(strKeyPath, strValueName, strValueType) Then
      intValueType = TypeReverseEnum(strValueType)
      ValueType = True
      Exit Function
    End If
    
    Call LogIt("There was an error trying to get the registry entry value type.", "ValueType", LogTypeError)
  End Function
  
  '***********************************************************************
  Private Function Value_TypeLegacy(ByVal strKeyPath, ByVal strValueName, ByRef strValueType)
  
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Value_TypeLegacy = False
    
    Dim strCommand
    Dim arrOutput
    Dim strLine
    Dim arrLine
    Dim strValueName_
    
    strValueType = Empty
    
    Call LogIt("Querying registry key (" & strKeyPath & "\" & strValueName & ").", "Value_TypeLegacy", LogTypeInfo + LogTypeVerbose)
    
    strCommand = "reg.exe QUERY " & Chr(34) & strKeyPath & Chr(34) & " /v " & Chr(34) & strValueName & Chr(34)
    If Not m_objProcess.Exec(strCommand, 0) Then
      Exit Function
    End If
    
    arrOutput = m_objProcess.Output
    
    For Each strLine In arrOutput
      strLine = Trim(strLine)
      If Not IsNullOrEmpty(strLine) Then
        If StrIn(1, strLine, vbTab) > 0 Then
          strLine = StrReplace(strLine, vbTab, "  ", 1, -1)
        End If
        strLine = Trim(strLine)
        
        strValueName_ = Left(strLine, Len(strValueName))
        If StrCompare(strValueName_, strValueName) = 0 Then
          strValueType = Mid(strLine, Len(strValueName) + 1)
          strValueType = Trim(strValueType)
          strValueType = Left(strValueType, StrIn(1, strValueType, " ") - 1)
    
          Value_TypeLegacy = True
          Exit For
        End If
      End If
    Next
    
  End Function
  
  '***********************************************************************
  Private Function Value_TypeWMI(ByVal dblHive, ByVal strKeyPath, ByVal strValueName, ByRef intValueType)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Value_TypeWMI = False
    
    Dim arrValueNames, arrValueTypes
    Dim intReturn
    Dim i
    
    intValueType = Empty
    
    If Not Connection_Check Then Exit Function
    
    g_objError.Clear
    intReturn = m_objReg.EnumValues(dblHive, strKeyPath, arrValueNames, arrValueTypes)
    If (intReturn <> 0) Or g_objError.Check() Then
      Select Case intReturn
        Case 2
          'Call LogIt("  Could not read the registry (" & dblHive & "\" & strKeyPath & "\" & strValueName & "). Return code: " & intReturn & ". The registry key does not exist.", "Value_TypeWMI", LogTypeError)
        Case Else
          Call LogIt("  Could not read the registry (" & dblHive & "\" & strKeyPath & "\" & strValueName & "). " & g_objError.Message, "Value_TypeWMI", LogTypeError)
      End Select
      Exit Function
    End If
    
    Value_TypeWMI = True
    
    If TypeName(arrValueNames) = "Variant()" Then
      'Loop through the array of value names looking for the requested value name and return its type
      For i = 0 To UBound(arrValueNames)
        If LCase(arrValueNames(i)) = LCase(strValueName) Then
          intValueType = arrValueTypes(i)
          Exit For
        End If
      Next
    End If
    
  End Function
  
  '***********************************************************************
  Public Function Write(ByVal strPath, ByVal varValue, ByVal varValueType)
    
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Write = False
    
    Dim strCommand
    Dim strKeyPath
    Dim strValueName
    
    If IsNumeric(varValueType) Then varValueType = TypeEnum(varValueType)
    
    If StrCompare(varValueType, "REG_MULTI_SZ") <> 0 Then 'This type is not supported for the RegWrite method
      On Error Resume Next
      g_objError.Clear
      If Right(strPath, 1) = "\" Then
        Call LogIt("Write registry key: " & strPath, "Write", LogTypeInfo + LogTypeVerbose)
        g_objWshShell.RegWrite strPath
      Else
        Call LogIt("Write the " & strPath & " registry value = " & varValue, "Write", LogTypeInfo + LogTypeVerbose)
        g_objWshShell.RegWrite strPath, varValue, varValueType
      End If
      If g_objError.Check() Then
        Call LogIt("  Could not write the registry (" & strPath & "). " & g_objError.Message, "Write", LogTypeError)
      Else
        Call LogIt("  Wrote the (" & strPath & ") registry key or value.", "Write", LogTypeInfo + LogTypeVerbose)
        Write = True
      End If
      If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
      Exit Function
    End If
    
    'Use legacy methods to create a REG_MULTI_SZ entry
    strKeyPath = Left(strPath, InStrRev(strPath, "\", -1) - 1)
    strValueName = Mid(strPath, InStrRev(strPath, "\", -1) + 1)
    
    strCommand = "reg.exe ADD " & Chr(34) & strKeyPath & Chr(34) & _
                        " /v " & Chr(34) & strValueName & Chr(34) & _
                        " /t " & varValueType & _
                        " /d " & Chr(34) & varValue & Chr(34) & _
                        " /f"
    
    If m_objProcess.Exec(strCommand, 0) Then
      Write = True
    Else
      Call LogIt("  Could not write the (" & strPath & ") value to the registry.", "Write", LogTypeError)
    End If
  End Function
  
  '***********************************************************************
  Public Function Path_Verify(ByVal strKeyPath)
  '***********************************************************************
  ' Purpose:  Verifies that a registry key path exists. Any missing
  '           keys are created.
  '
  ' Input:    strKeyPath - The registry key path to verify.
  '
  ' Returns:  True - The procedure succeeded.
  '           False - The procedure failed.
  '
  '***********************************************************************
    If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    
    Path_Verify = False
    
    Dim arrKeys
    Dim strKeyToVerify
    Dim strKey
    Dim i
    Dim blnErrorOccurred
  
    blnErrorOccurred = False
    
    'Strip double backslashes
    If StrIn(1, strKeyPath, "\\") > 0 Then
      strKeyPath = StrReplace(strKeyPath, "\\", "\", 1, -1)
    End If
    
    If StrIn(1, strKeyPath, "\") = 0 Then
      Call LogIt("  An invalid registry key path (" & strKeyPath & ") was passed to the Path_Verify procedure.", "Path_Verify", LogTypeError)
      Exit Function
    End If
    
    'Remove trailing backslash
    If Right(strKeyPath, 1) = "\" Then
      strKeyPath = Left(strKeyPath, Len(strKeyPath) - 1)
    End If
    
    'Split the path into an array
    arrKeys = StrSplit(strKeyPath, "\", -1)
    
    'Loop through the array verifying the path fully exists.  Missing subkeys will be created
    For i = 1 To UBound(arrKeys)
      'Append path elements
      If IsNullOrEmpty(strKeyToVerify) Then
        strKeyToVerify = arrKeys(0) & "\" & arrKeys(1)
      Else
        strKeyToVerify = strKeyToVerify & "\" & arrKeys(i)
      End If
      
      If Not Exists(strKeyToVerify & "\") Then
        'Create the missing subkey
        If Not Write(strKeyToVerify & "\", "Nothing", REG_SZ) Then
          blnErrorOccurred = True
          Exit For
        End If
      End If
    Next
    If Not blnErrorOccurred Then Path_Verify = True
  End Function
  
  '********************************************************************************
  Public Function HiveEnum(ByVal dblHive)
    Select Case dblHive
      Case HKEY_CLASSES_ROOT   : HiveEnum = "HKEY_CLASSES_ROOT"
      Case HKEY_CURRENT_USER   : HiveEnum = "HKEY_CURRENT_USER"
      Case HKEY_LOCAL_MACHINE  : HiveEnum = "HKEY_LOCAL_MACHINE"
      Case HKEY_USERS          : HiveEnum = "HKEY_USERS"
      Case HKEY_CURRENT_CONFIG : HiveEnum = "HKEY_CURRENT_CONFIG"
      Case Else                : HiveEnum = "unknown"
    End Select
  End Function
  
  '********************************************************************************
  Public Function HiveReverseEnum(ByVal strHive)
    Select Case strHive
      Case "HKEY_CLASSES_ROOT", "HKCR"   : HiveReverseEnum = HKEY_CLASSES_ROOT
      Case "HKEY_CURRENT_USER", "HKCU"   : HiveReverseEnum = HKEY_CURRENT_USER
      Case "HKEY_LOCAL_MACHINE", "HKLM"  : HiveReverseEnum = HKEY_LOCAL_MACHINE
      Case "HKEY_USERS", "HKU"           : HiveReverseEnum = HKEY_USERS
      Case "HKEY_CURRENT_CONFIG", "HKCC" : HiveReverseEnum = HKEY_CURRENT_CONFIG
      Case Else                          : HiveReverseEnum = Empty
    End Select
  End Function
  
  '********************************************************************************
  Public Function HiveExpand(ByVal strPath)
    
    HiveExpand = Empty
    
    Dim strHive
    If StrIn(1, strPath, "\") > 0 Then
      strHive = Left(strPath, StrIn(1, strPath, "\") - 1)
    Else
      strHive = strPath
    End If
    
    Select Case True
      Case StrIn(1, strHive, "HKLM") > 0
        HiveExpand = StrReplace(strPath, "HKLM", "HKEY_LOCAL_MACHINE", 1, -1)
      Case StrIn(1, strHive, "HKCU") > 0
        HiveExpand = StrReplace(strPath, "HKCU", "HKEY_CURRENT_USER", 1, -1)
      Case StrIn(1, strHive, "HKU") > 0
        HiveExpand = StrReplace(strPath, "HKU", "HKEY_USERS", 1, -1)
      Case StrIn(1, strHive, "HKCR") > 0
        HiveExpand = StrReplace(strPath, "HKCR", "HKEY_CLASSES_ROOT", 1, -1)
      Case StrIn(1, strHive, "HKCC") > 0
        HiveExpand = StrReplace(strPath, "HKCC", "HKEY_CURRENT_CONFIG", 1, -1)
    End Select
  End Function
  
  '********************************************************************************
  Public Function TypeEnum(ByVal intType)
    Select Case intType
      Case REG_SZ        : TypeEnum = "REG_SZ"
      Case REG_EXPAND_SZ : TypeEnum = "REG_EXPAND_SZ"
      Case REG_BINARY    : TypeEnum = "REG_BINARY"
      Case REG_DWORD     : TypeEnum = "REG_DWORD"
      Case REG_MULTI_SZ  : TypeEnum = "REG_MULTI_SZ"
      Case Else          : TypeEnum = "unknown"
    End Select
  End Function
  
  '********************************************************************************
  Public Function TypeReverseEnum(ByVal strType)
    Select Case strType
      Case "REG_SZ"        : TypeReverseEnum = REG_SZ
      Case "REG_EXPAND_SZ" : TypeReverseEnum = REG_EXPAND_SZ
      Case "REG_BINARY"    : TypeReverseEnum = REG_BINARY
      Case "REG_DWORD"     : TypeReverseEnum = REG_DWORD
      Case "REG_MULTI_SZ"  : TypeReverseEnum = REG_MULTI_SZ
      Case Else            : TypeReverseEnum = Empty
    End Select
  End Function
  
End Class
