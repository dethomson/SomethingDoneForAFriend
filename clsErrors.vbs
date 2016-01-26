Dim strFileScript : strFileScript = "clsErrors.vbs"
'***********************************************************************
'File:     clsErrors.vbs
'
'Comments: Developed by Dan Thomson (dethomson@hotmail.com)
'          Last modified on 3/31/2011
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
Class clsError
  
  Private m_strSource
  Private m_strErrDec
  Private m_strErrHex
  Private m_strErrDesc
  Private m_strWMIDesc
  Private m_strWMIProvider
  Private m_strWMIOperation
  Private m_strWMIParameterInfo
  Private m_strWMIStateCode
  Private m_blnHasWMIError
  Private m_strMessage
  
  '***********************************************************************
  Private Sub Class_Initialize()
  End Sub
  
  '***********************************************************************
  Private Sub Class_Terminate()
  End Sub

  '***********************************************************************
  Public Property Get HasWMIError()
    HasWMIError = m_blnHasWMIError
  End Property

  '***********************************************************************
  Public Property Get Source()
    Source = m_strSource
  End Property

  '***********************************************************************
  Public Property Get Desc()
    Desc = m_strDesc
  End Property

  '***********************************************************************
  Public Property Get ErrDec()
    If IsNullOrEmpty(m_strErrDec) Then
      ErrDec = vbNull
    Else
      ErrDec = CDbl(m_strErrDec)
    End If
  End Property

  '***********************************************************************
  Public Property Get ErrHex()
    If IsNullOrEmpty(m_strErrHex) Then
      ErrHex = vbNull
    Else
      ErrHex = CDbl(m_strErrHex)
    End If
  End Property

  '***********************************************************************
  Public Property Get Operation()
    Operation = m_strOperation
  End Property
  
  '***********************************************************************
  Public Property Get StateCode()
    StateCode = m_strStateCode
  End Property

  '***********************************************************************
  Public Property Get Provider()
    Provider = m_strProvider
  End Property

  '***********************************************************************
  Public Property Get Message()
    Message = m_strMessage
  End Property

  '***********************************************************************
  Public Property Get ParameterInfo()
    ParameterInfo = m_strWMIParameterInfo
  End Property
  
  '***********************************************************************
  Public Sub Clear()
    Err.Clear
    m_strSource              = Empty
    m_strErrDec              = Empty
    m_strErrHex              = Empty
    m_strErrDesc             = Empty
    m_strWMIDesc             = Empty
    m_strWMIProvider         = Empty
    m_strWMIOperation        = Empty
    m_strWMIParameterInfo    = Empty
    m_strWMIStateCode        = Empty
    m_strMessage             = Empty
    m_blnHasWMIError         = False
  End Sub
  
  '***********************************************************************
  Public Function Check()
    
    If (Err.Number <> 0) Then
      m_strMessage = _
               "Err Src: " & Err.Source & _
               " || Num (dec): " & Err.Number & _
               " || Num (hex): &H" & Hex(Err.Number) & _
               " || Desc: " & Err.Description
      
      m_strErrDec = Err.Number
      m_strErrHex = "&H" & Hex(Err.Number)
      m_strErrDesc = "Err: " & Err.Description
      
      Check = True
    Else
      Check = False
    End If
  
    On Error Resume Next
    
    'Instantiate SWbemLastError object.
    Dim objError, strWMI
    Set objError = CreateObject("WbemScripting.SwbemLastError")
    
    If objError Then
      m_blnHasWMIError = True
      
      strWMI                = Empty
      m_strWMIDesc          = Empty
      m_strWMIProvider      = Empty
      m_strWMIOperation     = Empty
      m_strWMIParameterInfo = Empty
      m_strWMIStateCode     = Empty
      
      strWMI = " || Provider: " & objError.ProviderName & _
               " || Operation: " & objError.Operation
      
      m_strWMIDesc          = objError.Description
      m_strWMIProvider      = objError.ProviderName
      m_strWMIOperation     = objError.Operation
      m_strWMIParameterInfo = objError.ParameterInfo
      m_strWMIStateCode     = objError.StatusCode
  
      If (m_strWMIDesc <> "") Then
        strWMI = strWMI & " || Description: " & m_strWMIDesc
        m_strErrDesc = m_strErrDec & " WMI: " & m_strWMIDesc
      End If
  
      If (m_strWMIParameterInfo <> "") Then
        strWMI = strWMI & " || Parameter information: " & m_strWMIParameterInfo
      End If
  
      If (m_strWMIStateCode <> "") Then
        strWMI = strWMI & " || Status: " & m_strWMIStateCode
      End If
  
      Set objError = Nothing
  
      Check = True
    End If
    
    m_strMessage = m_strMessage & strWMI
    On Error GoTo 0
  End Function
    
End Class
