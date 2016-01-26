Dim strFileScript : strFileScript = "libFileSystem.vbs"
'***********************************************************************
'File:     libFileSystem.vbs
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

'Constants used to specify file I/O
'These are only required when not running a WSF script
'  and "Scripting.FileSystemObject" is not set as an object
'  in the script.

'   Const ForReading           = 1
'   Const ForWriting           = 2
'   Const ForAppending         = 8

'   Const TristateUseDefault   = -2 'Opens the file using the system default.
'   Const TristateTrue         = -1 'Opens the file as Unicode.
'   Const TristateFalse        =  0 'Opens the file as ASCII.

'***********************************************************************
Private Sub GetNewestFileDate(ByVal strFolder, ByVal blnRecurse)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Dim objFolder
  Dim objSubFolder
  Dim objFile
  Dim intDateDiff
  
  Set objFolder = g_objFSO.GetFolder(strFolder)
  For Each objFile In objFolder.Files
    If IsNullOrEmpty(g_dicSettings.Key("FileDateLastModified")) Then
      g_dicSettings.Key("FileDateLastModified") = objFile.DateLastModified
    Else
      intDateDiff = CInt(DateDiff("d", objFile.DateLastModified, CDate(g_dicSettings.Key("FileDateLastModified"))))
      If intDateDiff < 0 Then
        g_dicSettings.Key("FileDateLastModified") = objFile.DateLastModified
      End If
    End If
  Next
  
  If blnRecurse Then
    For Each objSubFolder In objFolder.SubFolders
      Call GetNewestFileDate(objSubFolder.Path, blnRecurse)
    Next
  End If
End Sub

'***********************************************************************
Private Function Folder_Delete(ByVal strFolder)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Folder_Delete = False
  
  Dim intCountLoop
  
  intCountLoop = 0

  If StrIn(1, strFolder, "\\") > 0 Then strFolder = StrReplace(strFolder, "\\", "\", 1, -1)
  
  If g_objFSO.FolderExists(strFolder) Then
    Call LogIt("Initiating a delete of the " & strFolder & " folder.", "Folder_Delete", LogTypeInfo + LogTypeVerbose)
    
    Do
      On Error Resume Next
      g_objError.Clear
      g_objFSO.DeleteFolder strFolder, True
      If g_objError.Check() Then
        Call LogIt("  Could not delete the " & strFolder & " folder. " & g_objError.Message, "Folder_Delete", LogTypeError)
      Else
        If g_objFSO.FolderExists(strFolder) Then
          Call LogIt("  Failed to delete the " & strFolder & " folder.", "Folder_Delete", LogTypeError)
        Else
          Call LogIt("  The " & strFolder & " folder has been deleted.", "Folder_Delete", LogTypeInfo + LogTypeVerbose)
          Folder_Delete = True
          Exit Do
        End If
      End If
      
      intCountLoop = intCountLoop + 1
      
      Wscript.Sleep 2000
      
      'Try for 30 seconds
      If intCountLoop > 15 Then Exit Do
    Loop
  Else
    'nothing to do
  End If
End Function

'***********************************************************************
Private Function Folder_Move(ByVal strFolder_Source, ByVal strFolder_Dest)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Folder_Move = False
  
  Dim objFolder_Source
  
  If Not g_objFSO.FolderExists(strFolder_Source) Then
    Call LogIt("The " & strFolder_Source & " source folder does not exist.", "Folder_Move", LogTypeWarning)
    Exit Function
  End If
  
  If g_objFSO.FolderExists(strFolder_Dest) Then
    If Not Folder_Delete(strFolder_Dest) Then
      Exit Function
    End If
  End If
  
  On Error Resume Next
  g_objError.Clear
  g_objFSO.MoveFolder strFolder_Source, strFolder_Dest
  If g_objError.Check() Then
    Call LogIt("Failed to move the " & strFolder_Source & " folder to " & strFolder_Dest & ". " & g_objError.Message, "Folder_Move", LogTypeError)
    Exit Function
  Else
    Call LogIt("Moved the " & strFolder_Source & " folder to " & strFolder_Dest & ". ", "Folder_Move", LogTypeInfo + LogTypeVerbose)
    Folder_Move = True
  End If
End Function

'***********************************************************************
Private Function Drive_Space_Free_Get(ByVal strDrive, ByRef lngFreeSpace)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Drive_Space_Free_Get = False
  
  Dim objDisk
  
  lngFreeSpace = Empty
  
  Call LogIt("Checking free space on drive " & strDrive, "Drive_Space_Free_Get", LogTypeInfo)
  
  If IsNullOrEmpty(strDrive) Then
    Call LogIt("An invalid drive spec was passed.", "Drive_Space_Free_Get", LogTypeError)
    Exit Function
  End If
  
  On Error Resume Next
  Set objDisk = g_objFSO.GetDrive(strDrive)
  If g_objError.Check Then
    Call LogIt("An error occurred while getting an object reference to the drive." & g_objError.Message, "Drive_Space_Free_Get", LogTypeError)
  Else
    lngFreeSpace = objDisk.FreeSpace / 1048576
    If Not IsNullOrEmpty(lngFreeSpace) Then Drive_Space_Free_Get = True
  End If
  If IsObject(objDisk) Then Set objDisk = Nothing
End Function

'***********************************************************************
Private Function File_Copy(ByVal strFile_Source, ByVal strFile_Dest, ByVal blnFile_Backup)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  File_Copy = False
  
  Dim objFile_Source, objFile_Dest
  Dim blnCopyDoIt
  Dim intCountLoop
  Dim blnErrorOccurred
  
  blnCopyDoIt = False
  blnErrorOccurred = False
  
  If g_objFSO.FileExists(strFile_Source) Then
    g_objError.Clear
    Set objFile_Source = g_objFSO.GetFile(strFile_Source)
    If (Not IsObject(objFile_Source)) Or (g_objError.Check()) Then
      Call LogIt("  Could not instantiate an object reference to the " & strFile_Source & " file. " & g_objError.Message, "File_Copy", LogTypeError)
      Exit Function
    End If
  Else
    Call LogIt("The " & strFile_Source & " file is missing.", "File_Copy", LogTypeWarning)
    Exit Function
  End If
  
  If Right(strFile_Dest, 1) = "\" Then
    strFile_Dest = strFile_Dest & g_objFSO.GetBaseName(strFile_Source) & "." & g_objFSO.GetExtensionName(strFile_Source)
  End If
  
  If g_objFSO.FileExists(strFile_Dest) Then
    g_objError.Clear
    Set objFile_Dest = g_objFSO.GetFile(strFile_Dest)
    If (Not IsObject(objFile_Dest)) Or (g_objError.Check()) Then
      Call LogIt("  Could not instantiate an object reference to the " & strFile_Dest & " file. " & g_objError.Message, "File_Copy", LogTypeError)
      Exit Function
    End If
    If objFile_Source.DateLastModified <> objFile_Dest.DateLastModified Then
      blnCopyDoIt = True
    Else
      File_Copy = True
    End If
  Else
    'Call LogIt("The " & strFile_Dest & " file is missing.", "File_Copy", LogTypeWarning)
    blnFile_Backup = False
    blnCopyDoIt = True
  End If
  
  If IsObject(objFile_Dest) Then Set objFile_Dest = Nothing
  If IsObject(objFile_Source) Then Set objFile_Source = Nothing
  
  'Do we back up the currrent file before we copy over the new file?
  If blnCopyDoIt And blnFile_Backup Then
    'Set the path for the new backup file
    strFile_Dest_Bak = g_objFSO.GetParentFolderName(strFile_Dest) & "\" & g_objFSO.GetBaseName(strFile_Dest) & ".bak"
    
    'Delete the current back up file if it exists
    If g_objFSO.FileExists(strFile_Dest_Bak) Then
      Call LogIt("Deleting current backup file " & strFile_Dest_Bak, "File_Copy", LogTypeInfo + LogTypeVerbose)
      
      Call File_Delete(strFile_Dest_Bak)
      'What to do if this fails??
    End If
  
    Call LogIt("Backing up the old " & strFile_Dest & " file to " & strFile_Dest_Bak, "File_Copy", LogTypeInfo + LogTypeVerbose)
    
    'Rename the existing destination file to the back up file name
    On Error Resume Next
    g_objError.Clear
    g_objFSO.MoveFile strFile_Dest, strFile_Dest_Bak
    If g_objError.Check() Then
      Call LogIt("  Could not rename the " & strFile_Dest & " file to " & strFile_Dest_Bak & ". " & g_objError.Message, "File_Copy", LogTypeError)
      'What to do if this fails??
    End If
  End If
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  If blnCopyDoIt Then
    Call LogIt("Copying " & strFile_Source & " to " & strFile_Dest, "File_Copy", LogTypeInfo + LogTypeVerbose)
    
    intCountLoop = 0
    Do
      On Error Resume Next
      g_objError.Clear
      g_objFSO.CopyFile strFile_Source, strFile_Dest, True
      If g_objError.Check() Then
        Call LogIt("  Could not copy the " & strFile_Source & " file to " & strFile_Dest & ". " & g_objError.Message, "File_Copy", LogTypeError)
        blnErrorOccurred = True
      Else
        Call LogIt("  Copied " & strFile_Source & " to " & strFile_Dest, "File_Copy", LogTypeInfo + LogTypeVerbose)
        blnErrorOccurred = False
        Exit Do
      End If
      If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
      If intCountLoop > 3 Then Exit Do
      intCountLoop = intCountLoop + 1
      WScript.Sleep 1000
    Loop
    
    If Not blnErrorOccurred Then
      'Check if the file exists at the destination
      '  Should probably add verification of size and datelastmodified
      If g_objFSO.FileExists(strFile_Dest) Then
        Call LogIt("  Verified that the " & strFile_Source & " exists at " & strFile_Dest, "File_Copy", LogTypeInfo + LogTypeVerbose)
        File_Copy = True
      Else
        Call LogIt("  The " & strFile_Source & " does not exist at " & strFile_Dest, "File_Copy", LogTypeError)
      End If
    End If
  End If
End Function

'***********************************************************************
Private Function File_Move(ByVal strFile_Source, ByVal strFile_Dest)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  File_Move = False
  
  Dim intCount
  
  intCount = 0
  
  If File_Copy(strFile_Source, strFile_Dest, False) Then intCount = intCount + 1
  
  WScript.Sleep(1000)

  If File_Delete(strFile_Source) Then intCount = intCount + 1
  
  If intCount = 2 Then File_Move = True
End Function

'***********************************************************************
Private Function File_Delete(ByVal strFile)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  File_Delete = False
  
  Dim objFile
  
  Call LogIt("Deleting " & strFile, "File_Delete", LogTypeInfo)
  
  If g_objFSO.FileExists(strFile) Then
    g_objError.Clear
    On Error Resume Next
    Set objFile = g_objFSO.GetFile(strFile)
    If (Not IsObject(objFile)) Or (g_objError.Check()) Then
      Call LogIt("  Could not instantiate an object reference to " & strFile & ". " & g_objError.Message, "File_Delete", LogTypeError)
      Exit Function
    End If
    
    If objFile.Attributes And 1 Then
      Call LogIt("  Removing the read-only attribute on " & strFile, "File_Delete", LogTypeInfo + LogTypeVerbose)
      On Error Resume Next
      objFile.Attributes = objFile.Attributes Xor 1
      If g_objError.Check() Then
        Call LogIt("    Could not remove the read-only attribute from " & strFile & ". " & g_objError.Message, "File_Delete", LogTypeError)
        Exit Function
      End If
    End If
    
    Set objFile = Nothing
    
    On Error Resume Next
    g_objError.Clear
    g_objFSO.DeleteFile strFile, True
    If g_objError.Check() Then
      Call LogIt("  Could not delete the " & strFile & " file. " & g_objError.Message, "File_Delete", LogTypeError)
    Else
      Call LogIt("  The DeleteFile method succeeded.", "File_Delete", LogTypeInfo + LogTypeVerbose)
      If g_objFSO.FileExists(strFile) Then
        Call LogIt("    The " & strFile & " has NOT been deleted.", "File_Delete", LogTypeError)
      Else
        Call LogIt("    Verified the " & strFile & " has been deleted.", "File_Delete", LogTypeInfo + LogTypeVerbose)
        File_Delete = True
      End If
    End If
  Else
    Call LogIt("  The " & strFile & " does not exist.", "File_Delete", LogTypeWarning + LogTypeVerbose)
    File_Delete = True
  End If
End Function

'***********************************************************************
Private Function File_Version(ByVal strFileName, ByRef strFileVersion)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  File_Version = False
  
  strFileVersion = Empty
  
  If Not g_objFSO.FileExists(strFileName) Then
    Exit Function
  End If
  
  g_objError.Clear
  strFileVersion = g_objFSO.GetFileVersion(strFileName)
  If (IsNullOrEmpty(strFileVersion)) Or (g_objError.Check()) Then
    Call LogIt("  Could not get file version information for the " & strFileName & " file. " & g_objError.Message, "File_Version", LogTypeError)
  Else
    File_Version = True
  End If
End Function
  
'***********************************************************************
Private Function TextFile_Open(ByRef objFile, ByVal strFileName, ByVal intMode)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  TextFile_Open = False
  
  On Error Resume Next
  g_objError.Clear
  Set objFile = g_objFSO.OpenTextFile(strFileName, intMode, True)
  If (Not IsObject(objFile)) Or (g_objError.Check()) Then
    Call LogIt("  Could not instantiate an object reference to file " & strFileName & ". " & g_objError.Message, "TextFile_Open", LogTypeError)
  Else
    TextFile_Open = True
  End If
End Function
  

'*******************************************************************
' Function  Folder_Path_Verify
'
' Purpose:  Verifies that a file system path exists. Any missing
'           path elements are created.
'
' Input:    strPath - The path to verify.
'
' Returns:  True - The procedure succeeded.
'           False - The procedure failed.
'
'*******************************************************************
Private Function Folder_Path_Verify(ByVal strPath)
  If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
  
  Folder_Path_Verify = False
  
  Dim blnIsPathLocal
  Dim arrFolders
  Dim intStart
  Dim strPathToCheck
  Dim intErrCount
  Dim strError
  
  blnIsPathLocal = True
  intErrCount = 0
  
  'Is the path a UNC?
  If StrIn(1, strPath, "\\") = 1 Then
    blnIsPathLocal = False
    'Strip leading backslashes
    strPath = Mid(strPath, 3)
  End If
  'Strip double backslashes
  If StrIn(1, strPath, "\\") > 0 Then
    strPath = StrReplace(strPath, "\\", "\", 1, -1)
  End If
  
  'Split the path into an array
  arrFolders = StrSplit(strPath, "\", -1)
  
  'Set initial path root
  If Not blnIsPathLocal Then
    strPathToCheck = "\\" & arrFolders(0) & "\" & arrFolders(1) & "\"
    intStart = 2
  Else
    strPathToCheck = Empty
    intStart = 0
  End If
  
  'Loop through the array verifying the path fully exists.  Missing folders will be created
  For i = intStart To UBound(arrFolders)
    'Append path elements
    If Len(strPathToCheck) = 0 Then
      strPathToCheck = arrFolders(i) '& "\"
    Else
      strPathToCheck = strPathToCheck & "\" & arrFolders(i)
    End If
    
    If Not g_objFSO.FolderExists(strPathToCheck) Then
      Call LogIt("Create folder: " & strPathToCheck, "Folder_Path_Verify", LogTypeInfo + LogTypeVerbose)
      On Error Resume Next
      g_objError.Clear
      g_objFSO.CreateFolder(strPathToCheck)
      If g_objError.Check() Then
        Call LogIt("  Failed to create folder. " & g_objError.Message, "Folder_Path_Verify", LogTypeError)
        intErrCount = intErrCount + 1
        Exit For
      Else
        Call LogIt("  Created folder. ", "Folder_Path_Verify", LogTypeInfo + LogTypeVerbose)
      End If
      If g_dicSettings.Key("Debug") Then On Error GoTo 0 Else On Error Resume Next
    End If
  Next
  If intErrCount = 0 Then Folder_Path_Verify = True
End Function
