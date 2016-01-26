Dim strFileScript : strFileScript = "clsDictionary.vbs"
'***********************************************************************
'File:     clsDictionary.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 6/18/2014
'
'          This script file is based on the Configuration Manager Health Check Tool
'          (http://configmgrclienthtc.codeplex.com/)
'
'Notes:    This class encapsulates the standard dictionary object and provides additional
'          functionality.
'
'Requires: g_objError
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


'***********************************************************************
Class clsDictionary
  
  Private m_dicDictionary
  Private m_strSeparator
  
 '***********************************************************************
  Private Sub Class_Initialize()
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    m_strSeparator = "#!#"
    
    Call ObjectRef_Create(m_dicDictionary, "Scripting.Dictionary")
    If IsObject(m_dicDictionary) Then m_dicDictionary.CompareMode = 1 'vbTextCompare
  End Sub
  
  '***********************************************************************
  Private Sub Class_Terminate()
    If IsObject(m_dicDictionary) Then Set m_dicDictionary = Nothing
  End Sub
  
  '********************************************************************************
  'Adds a key/item pair to the dictionary
  Public Sub Add(ByVal strKey, ByVal varValue)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    If IsObject(g_objError) Then g_objError.Clear
    If m_dicDictionary.Exists(strKey) Then
      m_dicDictionary(strKey) = varValue
    Else
      m_dicDictionary.Add strKey, varValue
    End If
    If IsObject(g_objError) Then _
      If g_objError.Check() Then _
        Call LogIt("Failed to update " & strKey & "|" & varValue & ". " & g_objError.Message, "Add", LogTypeError)
  End Sub
  
  '********************************************************************************
  'Sets the dictionary CompareMode
  Public Property Let CompareMode(ByVal intCompareMode)
    'Compare is a value representing the comparison mode. Acceptable values are
    ' 0 (Binary), 1 (Text), 2 (Database). Values greater than 2 can be used to
    ' refer to comparisons using specific Locale IDs (LCID).
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    m_dicDictionary.CompareMode = intCompareMode
  End Property
  
  '********************************************************************************
  'Returns the current dictionary CompareMode
  Public Property Get CompareMode()
    CompareMode = m_dicDictionary.CompareMode
  End Property
  
  '********************************************************************************
  'Returns the count of keys in the dictionary
  Public Function Count()
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    Count = m_dicDictionary.Count
  End Function
  
  '********************************************************************************
  'Returns True or False depending if the specified key was found
  Public Function Exists(ByVal strKey)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    If m_dicDictionary.Exists(strKey) Then Exists = True Else Exists = False
  End Function
  
  '********************************************************************************
  'Returns True or False depending if the specified item was found
  Public Function Item_Exists(ByVal varItemToFind)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    Dim varItem
    
    Item_Exists = False
    
    For Each varItem In m_dicDictionary.Items
      Select Case VarType(varItem)
        Case vbString
          If StrCompare(varItem, varItemToFind) = 0 Then
            Item_Exists = True
            Exit For
          End If
        Case Else
          If varItem = varItemToFind Then
            Item_Exists = True
            Exit For
          End If
      End Select
    Next
  End Function
  
  '********************************************************************************
  'Removes all keys that contain the specified item
  Public Function Item_Remove(ByVal varItemToRemove)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    Dim strKey
    Dim varItem
    Dim blnDoRemove
    
    For Each strKey In m_dicDictionary.Keys
      varItem = m_dicDictionary.Item(strKey)
      blnDoRemove = False
      Select Case VarType(varItem)
        Case vbString
          If StrCompare(varItem, varItemToRemove) = 0 Then blnDoRemove = True
        Case Else
          If varItem = varItemToRemove Then blnDoRemove = True
      End Select
      If blnDoRemove Then
        g_objError.Clear
        m_dicDictionary.Remove strKey
        If g_objError.Check() Then
          Call LogIt("Failed to remove " & varItemToRemove & ". " & g_objError.Message, "Item_Remove", LogTypeError)
          Item_Remove = False
        Else
          Item_Remove = True
        End If
      Else
        Item_Remove = True
      End If
    Next
  End Function
  
  '********************************************************************************
  'Returns the item for the specified key
  Public Property Get Item(ByVal strKeyToFind)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    Dim strKey
    Item = Empty
    If m_dicDictionary.Exists(strKeyToFind) Then Item = m_dicDictionary.Item(strKeyToFind)
  End Property
  
  '********************************************************************************
  'Sets the item for the specified key
  Public Property Let Item(ByVal strKey, ByVal varValue)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    m_dicDictionary.Item(strKey) = varValue
  End Property
  
  '********************************************************************************
  'Returns an array of items
  Public Function Items()
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    Items = m_dicDictionary.Items
  End Function
  
  '********************************************************************************
  'Returns an array of keys where the specified item can be found
  Public Function ItemKeys(ByVal varItemToFind)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    Dim strItems, arrItems
    Dim strKey
    
    For Each strKey In m_dicDictionary.Keys
      If StrCompare(varItemToFind, m_dicDictionary.Item(strKey)) = 0 Then
        If IsNullOrEmpty(strItems) Then
          strItems = strKey
        Else
          strItems = strItems & "#!#" & strKey
        End If
      End If
    Next
    If StrIn(1, strItems, "#!#") = 0 Then
      arrItems = Array(strItems)
    Else
      arrItems = StrSplit(strItems, "#!#", -1)
    End If
    ItemKeys = arrItems
  End Function
  
  '********************************************************************************
  'Returns the item for the specified key
  Public Property Get Key(ByVal strKey)
    On Error Resume Next
    Key = Empty
    If IsObject(m_dicDictionary) Then
      If m_dicDictionary.Exists(strKey) Then Key = m_dicDictionary.Item(strKey)
    End If
  End Property
  
  '********************************************************************************
  'Sets the item for the specified key
  Public Property Let Key(ByVal strKey, ByVal varValue)
    Call Add(strKey, varValue)
  End Property
  
  '********************************************************************************
  'Returns an array of keys
  Public Function Keys()
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    Keys = m_dicDictionary.Keys
  End Function
  
  '********************************************************************************
  'Returns the value of the specified list item name in the specified list item key
  '
  '   WScript.Echo dictionary.ListItem("test1", "a")
  Public Property Get ListItem(ByVal strKey, ByVal strValueName)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    ListItem = Empty
    
    Dim arrItems, strItem
    Dim arrItem
    
    If Not m_dicDictionary.Exists(strKey) Then Exit Property

    arrItems = m_dicDictionary.Item(strKey)
    
    If IsNullOrEmpty(arrItems) Then Exit Property
    
    For Each strItem In arrItems
      arrItem = StrSplit(strItem, m_strSeparator, 2)
      If StrCompare(strValueName, arrItem(0)) = 0 Then
        ListItem = arrItem(1)
        Exit For
      End If
    Next
  End Property
  
  '********************************************************************************
  'Sets the value of the specified list item name in the specified list item key
  '
  'Single item in the list:
  '   Key = Array of items: ItemName#!#ItemValue
  
  'Multiple items in the list:
  '   Key = Array of items: ItemName#!#ItemValue
  '                         ItemName#!#ItemValue
  '                         ItemName#!#ItemValue
  
  '  dictionary.ListItem("test1", "a") = 1
  Public Property Let ListItem(ByVal strKey, ByVal varValueName, ByVal varValue)
    On Error goto 0 'Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    Dim arrItems
    Dim blnFound
    Dim varItemName, varItemValue
    Dim i
    
    'If the list doesn't exist, create the list with the item and exit
    If Not m_dicDictionary.Exists(strKey) Then
      m_dicDictionary.Add strKey, Array(varValueName & m_strSeparator & varValue)
      Exit Property
    End If
    
    'If we get here, the list exists.  Get the array of list items
    arrItems = m_dicDictionary.Item(strKey)
    
    'If we get here, there is at least one item in the array of list items. Compare each value name and update or add new
    blnFound = False
    For i = LBound(arrItems) To UBound(arrItems)
      varItemName = Left(arrItems(i), StrIn(1, arrItems(i), m_strSeparator) - 1)
      'varItemValue = Mid(arrItems(i), StrIn(1, arrItems(i), m_strSeparator) + 3)
      If StrCompare(varItemName, varValueName) = 0 Then
        arrItems(i) = varValueName & m_strSeparator & varValue
        blnFound = True
        Exit For
      End If
    Next
    
    If Not blnFound Then
      ReDim Preserve arrItems(UBound(arrItems) + 1)
      arrItems(UBound(arrItems)) = varValueName & m_strSeparator & varValue
    End If
    
    m_dicDictionary(strKey) = arrItems
  End Property
    
  '********************************************************************************
  'Removes the specified list item name and its value from the specified list item key
  '
  '  dictionary.ListItemRemove "test1", "a"
  '  dictionary.ListItemRemove "test1", "*"
  Public Sub ListItemRemove(ByVal strKey, ByVal varValueName)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    Dim arrItemsCurrent, arrItemsNew()
    Dim blnFound
    Dim varItemName, varItemValue
    Dim i
    Dim j
    
    'Exit if the list doesn't exist
    If Not m_dicDictionary.Exists(strKey) Then
      Exit Sub
    End If
    
    ReDim arrItemsNew(0)
    
    If varValueName = "*" Then
      blnFound = True
    Else
      'If we get here, the list exists.  Get the array of list items
      arrItemsCurrent = m_dicDictionary.Item(strKey)
      
      'If we get here, there is at least one item in the array of list items. Compare each value name and remove
      blnFound = False
      For i = LBound(arrItemsCurrent) To UBound(arrItemsCurrent)
        varItemName = Left(arrItemsCurrent(i), StrIn(1, arrItemsCurrent(i), m_strSeparator) - 1)
        varItemValue = Mid(arrItemsCurrent(i), StrIn(1, arrItemsCurrent(i), m_strSeparator) + 3)
        If StrCompare(varItemName, varValueName) = 0 Then
          blnFound = True
        Else
          j = UBound(arrItemsNew)
          arrItemsNew(j) = varItemName & m_strSeparator & varItemValue
          ReDim Preserve arrItemsNew(j + 1)
        End If
      Next
      'The ReDim statement in the For/Next loop leaves an extra unused upper element in the array.
      'ReDim again to remove that unused upper element
      ReDim Preserve arrItemsNew(UBound(arrItemsNew) - 1)
    End If
    
    If blnFound Then
      m_dicDictionary(strKey) = arrItemsNew
    End If
  End Sub

  '********************************************************************************
  '  Returns a standard dictionary object containing key/value pairs of items in the specified list.
  '  Set dicListItems = dictionary.ListItems("TEST1")
  Public Function ListItems(ByVal strKey)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    ListItems = Empty
    
    Dim arrItems, strItem
    Dim arrItem
    Dim dicDictionary
    
    If Not m_dicDictionary.Exists(strKey) Then Exit Function

    arrItems = m_dicDictionary.Item(strKey)
    
    If IsNullOrEmpty(arrItems) Then Exit Function
    
    Set dicDictionary = CreateObject("Scripting.Dictionary")
    dicDictionary.CompareMode = vbTextCompare
    
    For Each strItem In arrItems
      'If Not IsNullOrEmpty(strItem) Then
        arrItem = StrSplit(strItem, m_strSeparator, 2)
        dicDictionary.Add arrItem(0), arrItem(1)
        arrItem = Empty
      'End If
    Next
    
    Set ListItems = dicDictionary
    Set dicDictionary = Nothing
  End Function
  
  '********************************************************************************
  'Removes the specified key from the dictionary
  Public Function Remove(ByVal strRemoveKey)
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error Goto 0
    
    Dim strKey
    Dim strKeyName
    
    strRemoveKey = UCase(Trim(CStr(strRemoveKey)))
    
    For Each strKey In m_dicDictionary.Keys
      strKeyName = UCase(Trim(CStr(strKey)))
      If strKeyName = strRemoveKey Then
        g_objError.Clear
        m_dicDictionary.Remove strKey
        If g_objError.Check() Then
          Call LogIt("Failed to remove " & strRemoveKey & ". " & g_objError.Message, "Remove", LogTypeError)
          Remove = False
        Else
          Remove = True
        End If
      Else
        Remove = True
      End If
    Next
  End Function
  
  '********************************************************************************
  'Removes all keys from the dictionary
  Public Sub RemoveAll()
    On Error Resume Next
    If g_dicSettings.Key("Debug") Then On Error GoTo 0
    m_dicDictionary.RemoveAll
  End Sub
  
  '********************************************************************************
  ' Sort
  '
  '********************************************************************************
  Public Function Sort()
    If g_dicSettings.Key("Debug") Then On Error Goto 0 Else On Error Resume Next
    
    Sort = False
    
    Dim objRecordSet
    Dim dicSorted
    Dim strKey
    Dim intCountRecords
    
    ' Create a detached recordset to store some stuff so we can sort it easily
    If Not ObjectRef_Create(objRecordSet, "ADOR.Recordset") Then Exit Function
    
    Do
      Set dicSorted = CreateObject("Scripting.Dictionary")
      If Not IsObject(dicSorted) Then Exit Do
      
      ' Add some fields to the recordset
      g_objError.Clear
      objRecordSet.Fields.Append "Column1", adVarChar, adMaxCharacters
      
      ' Open the recordset
      objRecordSet.Open
      
      If g_objError.Check() Then
        Call LogIt(g_objError.Message, "Sort", LogTypeError)
        Exit Do
      End If
      
      If objRecordSet.State <> adStateOpen Then Exit Do
      
      For Each strKey In m_dicDictionary.Keys
        ' Add key to the recordset
        g_objError.Clear
        objRecordSet.AddNew
        objRecordSet("Column1") = strKey
        
        ' Update the recordset
        objRecordSet.Update
      Next
      
      If objRecordSet.RecordCount = 0 Then Exit Do
      
      objRecordSet.Sort = "Column1 ASC"
      
      ' Move to the first record in the recordset
      If Not objRecordSet.BOF Then objRecordSet.MoveFirst
      
      ' Loop through the recordset and add the key/item pairs to our new dictionary
      Do While Not objRecordSet.EOF
        ' Get the sorted key from the recordset
        strKey = objRecordSet.Fields.Item("Column1")
        
        'Add the key to the sorted dictionary, retrieving the item from the unsorted dictionary
        dicSorted(strKey) = m_dicDictionary.Item(strKey)
        
        ' Go to the next record
        objRecordSet.MoveNext
      Loop
      Sort = True
      Exit Do
    Loop
    Set m_dicDictionary = dicSorted
    
    If objRecordSet.State <> adStateClosed Then objRecordSet.Close
    If IsObject(objRecordSet) Then Set objRecordSet = Nothing
    If IsObject(dicSorted) Then Set dicSorted = Nothing
  End Function
  
End Class
