VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Implements Iutility

Private mPoheader As PoHeader

'Private mPoitem As POITEM

Private mPoItemS As POITEMS

Private mPoremarkS As POREMARKS

Private mPoClauseS As POClauses

Private mPoReceipientS As PORECEIPIENTS

Private mConnection As ADODB.Connection

Private mOpenConnection As Boolean

Private MNamespace As String * 5

Private mLooKups As LookUps
Private mpoUnits As PoUnits
Private mFQA As FQA
Public Sub Configure(Namespace As String, Cn As ADODB.Connection)

MNamespace = Namespace
Set mConnection = Cn

End Sub

'This Function Should Return an array Filled with Ponumbers.
Public Sub GetAllPonumb()

End Sub

'It returns an object of type POITEMS containing all the poitems which are R/Q/B
Public Function GetRQBItems() As POITEMS

End Function

Public Function CanUserEditDocType(UserId, Doctype) As Boolean

End Function

Public Property Get PoHeader() As PoHeader
  
  If mPoheader Is Nothing Then
  
    Set mPoheader = New PoHeader
    mPoheader.Configure Me
    
  End If
  
    Set PoHeader = mPoheader
    
End Property

Public Property Get POITEMS() As POITEMS
'
'  If Not mPoItemS Is Nothing Then Set mPoItemS = Nothing
'    Set mPoItemS = New POITEMS
'
'    mPoItemS.Configure Me
'
' ' End If
'  Set POITEMS = mPoItemS

Set POITEMS = GetPoItem

End Property

Public Property Get FQA() As FQA
'
'  If Not mPoItemS Is Nothing Then Set mPoItemS = Nothing
'    Set mPoItemS = New POITEMS
'
'    mPoItemS.Configure Me
'
' ' End If
'  Set POITEMS = mPoItemS

Set FQA = GetFQA

End Property
Public Property Get POREMARKS() As POREMARKS
''
''  If mPoremarkS Is Nothing Then Set mPoremarkS = Nothing
''    Set mPoremarkS = New POREMARKS
''      mPoremarkS.Configure Me
''
''  Set POREMARKS = mPoremarkS

 Set POREMARKS = GetPoRemarks


End Property

Public Property Get POClauses() As POClauses
''  If mPoClauseS Is Nothing Then Set mPoClauseS = Nothing
''    Set mPoClauseS = New POClauses
''      mPoClauseS.Configure Me
''  Set POClauses = mPoClauseS
  
  Set POClauses = GetPoClause
  
End Property

Public Property Get PORECEIPIENTS() As PORECEIPIENTS
  
''   If mPoReceipientS Is Nothing Then Set mPoReceipientS = Nothing
''    Set mPoReceipientS = New PORECEIPIENTS
''           mPoReceipientS.Configure Me
''  Set PORECEIPIENTS = mPoReceipientS

  Set PORECEIPIENTS = GetPoRecepients

End Property

Private Property Get Iutility_Connection() As ADODB.Connection
  
  If mConnection Is Nothing Then
     
     If mOpenConnection = True Then
     'We can add code to Open a new connection HERE.
     End If
     
     Err.Raise vbObjectError + 100, "IMSPODLL.MAINPO.Iutility_Connection", "Either Connection is not set or the Connection is not Valid"
     
  End If
     
     Set Iutility_Connection = mConnection
       
End Property

Public Property Get OpenConnection() As Boolean
OpenConnection = mOpenConnection
End Property

Public Property Let OpenConnection(ByVal vNewValue As Boolean)
mOpenConnection = OpenConnection
End Property

Private Property Get Iutility_Namespace() As String
Iutility_Namespace = MNamespace
End Property

Public Property Get LookUps() As LookUps
If mLooKups Is Nothing Then
    Set mLooKups = New LookUps
    mLooKups.Configure Me
 End If
 
    Set LookUps = mLooKups
    
End Property


Public Property Get PoUnits() As PoUnits
If mpoUnits Is Nothing Then
  Set mpoUnits = New PoUnits
   mpoUnits.Configure Me
 End If
   
   Set PoUnits = mpoUnits
   
End Property
Private Function GenerateAutoNumber(Doctype As String, ServCode As String) As String   'AM
Dim Cmd As New ADODB.Command 'AM
Dim Retval As String 'AM
ServCode = Trim(ServCode)
With Cmd 'AM
    .CommandType = adCmdStoredProc 'AM
    .CommandText = "GETAUTONUMBERFORPO" 'AM
    .ActiveConnection = mConnection 'AM
    .Parameters.Append .CreateParameter("@CHRONOCODE", adVarChar, adParamInput, 2, GChronoCodeForPO) 'AM
    .Parameters.Append .CreateParameter("@DOCUTYPE", adVarChar, adParamInput, 2, Doctype) 'AM
    .Parameters.Append .CreateParameter("@NPECODE", adVarChar, adParamInput, 10, MNamespace) 'AM
    .Parameters.Append .CreateParameter("@SERVCODE", adVarChar, adParamInput, 15, ServCode) 'AM
    .Parameters.Append .CreateParameter("@RETVAL", adVarChar, adParamOutput, 15, Retval) 'AM
    
    .Execute 'AM

    GenerateAutoNumber = Trim(.Parameters("@RETVAL").Value) & "" 'AM
    
    End With
    Set Cmd = Nothing
End Function
'SetPONUMBforAllPoObjects = 0 fOR SUCCESS
'SetPONUMBforAllPoObjects =1 FOR UNIDENTIFIEED ERROR
'SetPONUMBforAllPoObjects = 2 FOR NO AUTONUMBERING AVAILABLE FOR THE DOCTYPE
Public Function SetPONUMBforAllPoObjects(Ponumber As String) As Integer

On Error GoTo Errhandler

SetPONUMBforAllPoObjects = 1

'If a number is not received then a number is generated

If Len(Trim(Ponumber)) = 0 Then

    Ponumber = GenerateAutoNumber(mPoheader.docutype, mPoheader.srvccode)
    
    'INDICATES IF AUTO NUMBERING IS NOT AVIALABLE FOR THIS PERTICULAR DOCTYPE
    
    If Len(Trim(Ponumber)) = 0 Then
    
        SetPONUMBforAllPoObjects = 2
        
        Exit Function
        
    End If
    
End If

If Not mPoheader Is Nothing Then mPoheader.POnumb = Ponumber

If Not mPoItemS Is Nothing Then
    
   Call mPoItemS.Replace("PONUMB", Ponumber)
        
End If

If Not mPoReceipientS Is Nothing Then

    Call mPoReceipientS.Replace("PONUMB", Ponumber)
    
End If

If Not mPoremarkS Is Nothing Then
    
    Call mPoremarkS.Replace("PONUMB", Ponumber)
    
End If

If Not mPoClauseS Is Nothing Then

    Call mPoClauseS.Replace("PONUMB", Ponumber)

    
End If

If Not mFQA Is Nothing Then

    Call mFQA.Replace("PONUMB", Ponumber)

    
End If

SetPONUMBforAllPoObjects = 0

Exit Function

Errhandler:

'n

End Function

''Public Function ConvertRequisition(RequisitionNumber As String) As Boolean
''Dim Error As Boolean
''Dim TemporaryPo As String
''On Error GoTo ErrHandler
''ConvertRequisition = False
''
''Call GetPoItem
''Call GetPoClause
''Call GetPoRecepients
''Call GetPoRemarks
''
'''Call POITEMS
'''Call POClauses
'''Call PORECEIPIENTS
'''Call POREMARKS
''
''TemporaryPo = mPoheader.POnumb
''
''Error = mPoheader.LoadFromRequsition(RequisitionNumber, TemporaryPo)
''If Error = False Then Error = mPoItemS.LoadFromRequsition(RequisitionNumber, TemporaryPo)
''If Error = False Then Error = mPoClauseS.LoadFromRequsition(RequisitionNumber, TemporaryPo)
''If Error = False Then Error = mPoReceipientS.LoadFromRequsition(RequisitionNumber, TemporaryPo)
''If Error = False Then Error = mPoremarkS.LoadFromRequsition(RequisitionNumber, TemporaryPo)
''
''
''Exit Function
''ErrHandler:
''
''
''Err.Clear
''End Function

Private Function GetPoItem() As POITEMS

  If Not mPoItemS Is Nothing Then Set mPoItemS = Nothing
    Set mPoItemS = New POITEMS
    
    mPoItemS.Configure Me
  
 ' End If
  Set GetPoItem = mPoItemS

End Function
Private Function GetFQA() As FQA

  If Not mFQA Is Nothing Then Set mFQA = Nothing
    
    Set mFQA = New FQA
    
    mFQA.Configure Me
    
    Set GetFQA = mFQA

End Function
Private Function GetPoClause() As POClauses
  If mPoClauseS Is Nothing Then Set mPoClauseS = Nothing
    Set mPoClauseS = New POClauses
      mPoClauseS.Configure Me
  Set GetPoClause = mPoClauseS
End Function
Private Function GetPoRemarks() As POREMARKS
  
  If mPoremarkS Is Nothing Then Set mPoremarkS = Nothing
    Set mPoremarkS = New POREMARKS
      mPoremarkS.Configure Me
  'End If
  Set GetPoRemarks = mPoremarkS
End Function
Private Function GetPoRecepients() As PORECEIPIENTS
  
   If mPoReceipientS Is Nothing Then Set mPoReceipientS = Nothing
    Set mPoReceipientS = New PORECEIPIENTS
           mPoReceipientS.Configure Me
  Set GetPoRecepients = mPoReceipientS
End Function
