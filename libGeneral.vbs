Dim strFileScript : strFileScript = "libGeneral.vbs"
'***********************************************************************
'File:     libGeneral.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 5/3/2014
'
'          This script file is based on the Configuration Manager Health Check Tool
'          (http://configmgrclienthtc.codeplex.com/)
'
'Notes:    
'
'Requires: g_objError
'          g_dicSettings
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

'EventLog constants
  Const EVENT_SUCCESS           = 0
  Const EVENT_ERROR             = 1
  Const EVENT_WARNING           = 2
  Const EVENT_INFO              = 4
  Const EVENT_AUDIT_SUCCESS     = 8
  Const EVENT_AUDIT_FAILURE     = 16

'Constants used to specify database table field types
  'May be used by ADOR.Recordset
  Const adVarChar                           = 200
  Const adInteger                           = 3
  Const adDouble                            = 5
  Const adLength                            = 50
  Const adMaxCharacters                     = 255
  Const adFldIsNullable                     = 32
  Const adFldKeyColumn                      = 32768
  Const adBoolean                           = 11
  
  Const adStateClosed                       = 0     'The object is closed.
  Const adStateOpen                         = 1     'The object is open.
  Const adStateConnecting                   = 2     'The object is connecting.
  Const adStateExecuting                    = 4     'The object is executing a command.
  Const adStateFetching                     = 8     'The rows of the object are being retrieved.
  
'***********************************************************************
Private Function DoDebug()
  Dim strArg, blnFound
  blnFound = False
  For Each strArg In Wscript.Arguments
    If StrCompare(strArg, "Debug") = 0 Then blnFound = True
  Next
  DoDebug = blnFound
End Function

'***********************************************************************
Private Function Drive_Space_Free(ByVal strDrive, ByRef lngFreeSpace)

  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Drive_Space_Free = False
  
  Dim objDisk
  
  lngFreeSpace = Empty
  
  Call LogIt("Checking free space on drive " & strDrive, "Drive_Space_Free", LogTypeInfo)
  
  If IsNullOrEmpty(strDrive) Then
    Call LogIt("An invalid drive spec was passed.", "Drive_Space_Free", LogTypeError)
    Exit Function
  End If
  
  On Error Resume Next
  Set objDisk = g_objFSO.GetDrive(strDrive)
  If g_objError.Check Then
    Call LogIt("An error occurred while getting an object reference to the drive." & g_objError.Message, "Drive_Space_Free", LogTypeError)
  Else
    lngFreeSpace = objDisk.FreeSpace / 1048576
    If Not IsNullOrEmpty(lngFreeSpace) Then Drive_Space_Free = True
  End If
  If IsObject(objDisk) Then Set objDisk = Nothing
End Function

'***********************************************************************
Private Sub EventLog_Write(ByVal intLogType, ByVal strLogText)
  
  If Not g_dicSettings.Key("WriteToEventLog") Then Exit Sub
  
  Dim objWshShell
  
  If Not ObjectRef_Create(objWshShell, "Wscript.Shell") Then
    Exit Sub
  End If
  
  On Error Resume Next
  g_objError.Clear
  objWshShell.LogEvent intLogType, strLogText
  If g_objError.Check() Then
    Call LogIt("  Could not write to the event log. Event Type: " & intLogType & " Message: " & strLogText & " " & g_objError.Message, "EventLog_Write", LogTypeError)
  Else
    Call LogIt("  Wrote event log entry. Event Type: " & intLogType & " Message: " & strLogText, "EventLog_Write", LogTypeInfo + LogTypeVerbose)
  End If
  If IsObject(objWshShell) Then Set objWshShell = Nothing
End Sub

'***********************************************************************
Private Function IsAdmin()
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  IsAdmin = False
  
  Dim strComputerName
  Dim objFSO
  Dim objWshShell
  
  Set objWshShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  'This ends up being a bad check for admin access since the system could have share issues.
  strComputerName = objWshShell.ExpandEnvironmentStrings("%ComputerName%")
  
  If objFSO.FolderExists("\\" & strComputerName & "\Admin$\System32") Then IsAdmin = True
End Function

'***********************************************************************
Private Function IsNullOrEmpty(ByRef varVar)
  
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  IsNullOrEmpty = False
  
  Dim intCount
  
  If IsEmpty(varVar) Then _
    IsNullOrEmpty = True              'Uninitialized
  If IsNull(varVar) Then _
    IsNullOrEmpty = True              'No valid data
  Select Case TypeName(varVar)
    Case "Nothing"                    'Object variable that doesn't yet refer to an object instance
      IsNullOrEmpty = True
    Case "Variant()"
      On Error Resume Next
      intCount = UBound(varVar)
      If Err.Number = 9 Then _
        IsNullOrEmpty = True          'err 9 = 'Subscript out of range'
    Case "Object"
      If Not IsObject(varVar) Then _
        IsNullOrEmpty = True
    Case Else
      If Len(varVar) = 0 Then _
        IsNullOrEmpty = True
  End Select
End Function

'********************************************************************************
Private Function TrueOrFalse(ByVal varValue)
  On Error Resume Next
  Select Case TypeName(varValue)
    Case "Boolean"
      TrueOrFalse = varValue
    Case "Integer"
      If CInt(varValue) = 0 Then
        TrueOrFalse = False
      Else
        TrueOrFalse = True
      End If
    Case "String"
      If StrIn(1, "yes, true", varValue) > 0 Then
        TrueOrFalse = True
      ElseIf StrIn(1, "no, false", varValue) > 0 Then
          TrueOrFalse = False
      ElseIf CInt(varValue) = 0 Then
        TrueOrFalse = False
      ElseIf CInt(varValue) = -1 Then
        TrueOrFalse = True
      ElseIf CInt(varValue) = 1 Then
        TrueOrFalse = True
      End If
  End Select
End Function

'********************************************************************************
Private Function StrCompare(ByVal strString1, ByVal strString2)
  On Error Resume Next
  StrCompare = StrComp(strString1, strString2, vbTextCompare)
End Function

'********************************************************************************
Private Function StrIn(ByVal intStart, ByVal strToSearch, ByVal strToFind)
  On Error Resume Next
  StrIn = InStr(intStart, strToSearch, strToFind, vbTextCompare)
End Function

'********************************************************************************
Private Function StrPadLeft(ByVal strString, ByVal strCharacter, ByVal intLength)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  While Len(strString) < intLength
    strString = strCharacter & strString
  Wend
  StrPadLeft = strString
End Function

'********************************************************************************
Private Function StrPadRight(ByVal strString, ByVal strCharacter, ByVal intLength)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  While Len(strString) < intLength
    strString = strString & strCharacter
  Wend
  StrPadRight = strString
End Function

'********************************************************************************
Private Function StrReplace(ByVal strToSearch, ByVal strToFind, ByVal strReplacement, ByVal intStart, ByVal intCount)
  On Error Resume Next
  StrReplace = Replace(strToSearch, strToFind, strReplacement, intStart, intCount, vbTextCompare)
End Function

'********************************************************************************
Private Function StrSplit(ByVal strString, ByVal strDelimiter, ByVal intCount)
  On Error Resume Next
  StrSplit = Split(strString, strDelimiter, intCount, vbTextCompare)
End Function

'********************************************************************************
Private Function ObjectRef_Create(ByRef objObj, ByVal strProgID)

  On Error Resume Next
  
  ObjectRef_Create = False
  
  Call LogIt("Attempting to create an object reference to '" & strProgID & "'.", "ObjectRef_Create", LogTypeInfo + LogTypeVerbose)
  g_objError.Clear
  Set objObj = CreateObject(strProgID)
  If (Not IsObject(objObj)) Or (g_objError.Check()) Then
    Call LogIt("  Could not create an object reference to '" & strProgID & "'. " & g_objError.Message, "ObjectRef_Create", LogTypeError)
  Else
    Call LogIt("  Created an object reference to '" & strProgID & "'.", "ObjectRef_Create", LogTypeInfo + LogTypeVerbose)
    ObjectRef_Create = True
  End If
  If g_dicSettings.Key("Debug") Then On Error Goto 0
  
End Function

'********************************************************************************
Private Function ObjectRef_Get(ByRef objObj, ByVal strProgID)
  
  On Error Resume Next
  
  ObjectRef_Get = False
  
  Call LogIt("Attempting to get an object reference to '" & strProgID & "'.", "ObjectRef_Get", LogTypeInfo + LogTypeVerbose)
  g_objError.Clear
  Set objObj = GetObject(CStr(strProgID))
  If (Not IsObject(objObj)) Or (g_objError.Check()) Then
    Call LogIt("  Could not get an object reference to '" & strProgID & "'. " & g_objError.Message, "ObjectRef_Get", LogTypeError)
  Else
    Call LogIt("  Obtained an object reference to '" & strProgID & "'.", "ObjectRef_Get", LogTypeInfo + LogTypeVerbose)
    ObjectRef_Get = True
  End If
  If g_dicSettings.Key("Debug") Then On Error Goto 0
  
End Function
