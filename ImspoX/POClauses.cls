VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "POClauses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"POCLAUSE"
Attribute VB_Ext_KEY = "Member0" ,"POCLAUSE"


Option Explicit

Private mIsItemLoaded As Boolean
''
''Public Enum Status
''  ADSTATUSUnknow = 0
''  ADSTATUSMODIFIED = 1
''  ADSTATUSINSERT = 2
''End Enum

Private mRsPOClause As ADODB.Recordset

Private MStatus  As Status

Private MNamespace As String

Private MPonumb As String

Private mCount As Integer

Private mValueChanged As Boolean

Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteClause(cancel As Boolean)
Public Event AfterDeleteClause()

Private miutility As Iutility
Public mCancel As Boolean
Private mErrors As Errors
Private mError As ErrorDetl
Public Sub Class_Initialize()
  
   MStatus = ADSTATUSINSERT
   End Sub

Public Property Get MyStatus() As Status
 MyStatus = MStatus
End Property

Public Property Let MyStatus(ByVal indata As Status)
MStatus = indata
End Property

Public Property Get POnumb() As String
'Ponumb = mRsPOClause!poc_ponumb
POnumb = MPonumb
End Property
Public Property Let POnumb(indata As String)
If Trim$(indata) <> MPonumb Then mIsItemLoaded = False
  mRsPOClause!poc_ponumb = Trim$(indata)
  MPonumb = Trim$(indata)
End Property
Public Property Get Namespace() As String
Namespace = mRsPOClause!poc_npecode
End Property
Public Property Let Namespace(ByVal indata As String)
mRsPOClause!poc_npecode = Trim$(indata)
End Property

Public Property Get lineNumb() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
lineNumb = mRsPOClause!poc_linenumb
End Property

Public Property Let lineNumb(ByVal indata As Integer)
 mRsPOClause!poc_linenumb = (Trim$(indata))
End Property

Public Function AddNew() As Boolean
Dim cancel As Boolean
AddNew = False
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
       If cancel = False Then
           mRsPOClause.AddNew
           SetTheNewHeader
           AddNew = True
       End If
  cancel = True
  
  
  Exit Function

handler:
     

    Err.Clear
End Function

Public Function DeleteCurrentLI() As Boolean

End Function

Public Function Update() As Boolean


Update = False
On Error GoTo handler
mRsPOClause.UpdateBatch
Update = True
Exit Function

handler:
  Err.Raise "12312", , Err.Description
  Err.Clear


End Function

Public Function Configure(Iutilty As Iutility)
Set miutility = Iutilty
End Function


Private Function LoadCollection() As Boolean


Dim RecordCount As Integer
Dim query As String
LoadCollection = False
On Error GoTo handler
 
  query = "SELECT poc_ponumb,poc_npecode, poc_linenumb, poc_clau FROM POCLAUSE WHERE "
  query = query & "(poc_ponumb = '" & MPonumb & "') AND (poc_npecode = '" & miutility.Namespace & "')"
       Set mRsPOClause = New ADODB.Recordset
       
       mRsPOClause.ActiveConnection = miutility.Connection
       mRsPOClause.Source = query
       mRsPOClause.CursorLocation = adUseClient
       mRsPOClause.LockType = adLockBatchOptimistic
       mRsPOClause.Open , , adOpenKeyset

  LoadCollection = True
  mIsItemLoaded = True
  Exit Function
   
handler:
  Err.Raise Err.Number, "IMSPODLL.POREMARKS.Loadcollection", Err.Description
  Err.Clear
End Function


Public Function Move(POnumb As String) As Boolean
On Error GoTo handler
Move = False
If Trim$(POnumb) <> MPonumb Then mIsItemLoaded = False
MPonumb = Trim$(POnumb)

If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection

If mIsItemLoaded = True Then
   mRsPOClause.MoveFirst

   
   If mRsPOClause.Eof = False Then
   Move = True

   End If
End If
   
   
   Exit Function
   
handler:

   Err.Clear
End Function

Public Property Get Count() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Count = mRsPOClause.RecordCount

End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler
  If Not mRsPOClause.Bof Then
   
     mRsPOClause.MoveFirst
     If Not mRsPOClause.Bof Then
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsPOClause.Eof Then
   
     mRsPOClause.MoveNext
      If mRsPOClause.AbsolutePosition = adPosEOF And mRsPOClause.Eof = True And mRsPOClause.Bof <> True Then mRsPOClause.MoveLast
     If Not mRsPOClause.Eof Then
        MoveNext = True
     End If
  End If
  Exit Function
handler:
  Err.Clear
End Function

Public Function MovePrevious() As Boolean
MovePrevious = False
On Error GoTo handler
  If Not mRsPOClause.Bof Then
   
     mRsPOClause.MovePrevious
     If mRsPOClause.AbsolutePosition = adPosBOF And mRsPOClause.Eof <> True And mRsPOClause.Bof = True Then mRsPOClause.MoveFirst
     If Not mRsPOClause.Bof Then
      MovePrevious = True
     End If
  End If
 Exit Function
handler:
  Err.Clear
End Function

Public Function MoveLast() As Boolean

MoveLast = False
On Error GoTo handler
  If Not mRsPOClause.Eof Then
   
     mRsPOClause.MoveLast
     If Not mRsPOClause.Eof Then  'FillmeWithTheCurrentRecord (mRsHeader)
       MoveLast = True
     End If
  End If
  Exit Function
handler:
  Err.Clear
End Function

Private Sub SetTheNewHeader()

   MStatus = ADSTATUSINSERT
   mCancel = False
   
End Sub

Public Property Get Errors() As Errors
Set Errors = mErrors
End Property



Private Sub HandleErrors(ErrNumber As Long, ErrDesc As String, ErrLocation As String)
 
 If mErrors Is Nothing Then Set mErrors = New Errors
   
       Set mError = New ErrorDetl
    
       mError.ErrorDesc = ErrDesc
       mError.ErrorLocation = ErrLocation
       mError.ErrorNumber = ErrNumber
       
  mErrors.Add mError
       
End Sub

Public Property Get Eof() As Boolean
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
If mIsItemLoaded = True Then
  Eof = mRsPOClause.Eof
End If
End Property

Public Property Get Bof() As Boolean
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Bof = mRsPOClause.Bof
End Property


Public Property Get Clause() As String
Clause = mRsPOClause!poc_clau
End Property

Public Property Let Clause(ByVal vNewValue As String)
mRsPOClause!poc_clau = Trim$(vNewValue)
End Property


Public Property Get EditMode() As Integer


If mRsPOClause.AbsolutePosition = adPosBOF Then mRsPOClause.MoveFirst
If mRsPOClause.AbsolutePosition = adPosEOF Then mRsPOClause.MoveLast
On Error GoTo handler
If mRsPOClause.State = 1 Then
   EditMode = mRsPOClause.EditMode
End If
Exit Function
handler:


Err.Raise Err.Number, , Err.Description
Err.Clear



End Property


Public Function CancelUpdate()
On Error GoTo handler

If mRsPOClause.AbsolutePosition = adPosBOF And mRsPOClause.Eof <> True And mRsPOClause.Bof = True Then mRsPOClause.MoveFirst
If mRsPOClause.AbsolutePosition = adPosEOF And mRsPOClause.Bof <> True And mRsPOClause.Eof = True Then mRsPOClause.MoveLast
If mRsPOClause.AbsolutePosition = adPosUnknown And mRsPOClause.RecordCount > 0 Then mRsPOClause.MoveLast
If mRsPOClause.EditMode = adEditAdd Then
   mRsPOClause.CancelUpdate
End If

If mRsPOClause.AbsolutePosition = adPosBOF And mRsPOClause.Eof <> True And mRsPOClause.Bof = True Then mRsPOClause.MoveFirst
If mRsPOClause.AbsolutePosition = adPosEOF And mRsPOClause.Bof <> True And mRsPOClause.Eof = True Then mRsPOClause.MoveLast
If mRsPOClause.AbsolutePosition = adPosUnknown And mRsPOClause.RecordCount > 0 Then mRsPOClause.MoveLast

Exit Function
   
handler:
   Err.Clear
End Function
Public Function DeleteAllClause(PO_ponumb As String) As Boolean
Dim cancel As Boolean
DeleteAllClause = False
On Error GoTo handler

RaiseEvent BeforeDeleteClause(cancel)

        
If cancel = False Then
   
   mRsPOClause.Filter = "por_ponumb='" & PO_ponumb & "'"
   If mRsPOClause.RecordCount > 0 Then
       mRsPOClause.Delete adAffectGroup
   End If
        
        RaiseEvent AfterDeleteClause
        
Else

   cancel = False
           
End If
DeleteAllClause = True
Exit Function
handler:
MsgBox "Error Occurred During Deleting the PO Recepients.Err.desc -- " & Err.Description
Err.Clear
End Function

Public Function Replace(PropertyName As String, Value As String) As Boolean

On Error GoTo Errhandler

Replace = False

PropertyName = Trim$(PropertyName)

If mRsPOClause.RecordCount > 0 Then

    mRsPOClause.MoveFirst
    
Else
    
    Replace = True
    
    Exit Function

End If

Select Case PropertyName

    Case "PONUMB"
          
          mRsPOClause.MoveFirst
          
          Do While Not mRsPOClause.Eof
          
                mRsPOClause!poc_ponumb = Value
                
                MPonumb = Value
                
                mRsPOClause.MoveNext
                
          Loop
      
End Select

Replace = True

Exit Function

Errhandler:

'n

End Function


Public Function LoadFromRequsition(RequsitionNo As String, TemporaryPoNumb As String, ErrorsIfany As String) As Boolean
Dim RsRequisition As ADODB.Recordset
Dim field As ADODB.field
On Error GoTo Errhandler
LoadFromRequsition = True

Set RsRequisition = Nothing

If Move(RequsitionNo) = True Then

Set RsRequisition = mRsPOClause.Clone

If RsRequisition.RecordCount > 0 Then

    Set mRsPOClause = Nothing
    
    Move TemporaryPoNumb
    
    RsRequisition.MoveFirst
    
    Do While Not RsRequisition.Eof
    
        mRsPOClause.AddNew
        
        For Each field In RsRequisition.Fields
        
                mRsPOClause(field.Name) = field.Value
        
        Next
       
       mRsPOClause("poc_ponumb") = TemporaryPoNumb
       
       RsRequisition.MoveNext
       
     Loop
       
     End If
       
    Set RsRequisition = Nothing

End If

LoadFromRequsition = False

Exit Function
Errhandler:

ErrorsIfany = Err.Description

Err.Clear
End Function
