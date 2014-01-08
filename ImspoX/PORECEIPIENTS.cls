VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PORECEIPIENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Poreceipient"
Attribute VB_Ext_KEY = "Member0" ,"Poreceipient"
Option Explicit

Private mIsItemLoaded As Boolean
''
''Public Enum Status
''  ADSTATUSUnknow = 0
''  ADSTATUSMODIFIED = 1
''  ADSTATUSINSERT = 2
''End Enum

Private mRsPOReceipients As ADODB.Recordset

Private MStatus  As Status

Private MNamespace As String

Private MPonumb As String

Private mCount As Integer

Private mValueChanged As Boolean
Private mSupplier As String
Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteRecepients(cancel As Boolean)
Public Event AfterDeleteRecepients()

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
'Ponumb = MPonumb
'If mRsPOReceipients.RecordCount > 0 And mRsPOReceipients.AbsolutePosition = adPosEOF Then mRsPOReceipients.MoveFirst
'PONUMB = mRsPOReceipients!PORC_PONUMB
POnumb = MPonumb
End Property
Public Property Let POnumb(indata As String)
'If Trim$(indata) <> MPonumb Then mIsItemLoaded = False
 mRsPOReceipients!PORC_PONUMB = Trim$(indata)
 MPonumb = Trim$(indata)
End Property
Public Property Get Namespace() As String
Namespace = mRsPOReceipients!PORC_NPECODE
End Property
Public Property Let Namespace(ByVal indata As String)
mRsPOReceipients!PORC_NPECODE = Trim$(indata)
End Property
Public Property Get lineNumb() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
lineNumb = mRsPOReceipients!porc_recpnumb
End Property

Public Property Let lineNumb(ByVal indata As Integer)
 mRsPOReceipients!porc_recpnumb = (Trim$(indata))
End Property

Public Function AddNew() As Boolean
Dim cancel As Boolean
AddNew = False
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
       If cancel = False Then
           mRsPOReceipients.AddNew
           SetTheNewHeader
           AddNew = True
       End If
  cancel = True
  
  
  Exit Function

handler:

    Err.Clear
End Function

Public Function DeleteCurrentLI(text As String) As Boolean
 On Error GoTo handler
 DeleteCurrentLI = False
 mRsPOReceipients.MoveFirst
 mRsPOReceipients.Find "porc_rec='" & text & "'", , adSearchForward
 
 If Not mRsPOReceipients.AbsolutePosition = adPosEOF Then
     mRsPOReceipients.Delete adAffectCurrent
     mRsPOReceipients.MoveFirst
     Do While Not mRsPOReceipients.Eof
        mRsPOReceipients!porc_recpnumb = mRsPOReceipients.AbsolutePosition
        mRsPOReceipients.MoveNext
     Loop
  End If
  DeleteCurrentLI = True
  Exit Function
handler:
   Err.Clear
End Function

Public Function Update() As Boolean
Update = False
On Error GoTo handler
mRsPOReceipients.UpdateBatch
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
 
  query = "SELECT porc_ponumb, porc_npecode, porc_rec, porc_recpnumb FROM POREC WHERE "
  query = query & "(porc_ponumb = '" & MPonumb & "') AND (porc_npecode = '" & miutility.Namespace & "') ORDER BY porc_recpnumb "
       Set mRsPOReceipients = New ADODB.Recordset
       
       mRsPOReceipients.ActiveConnection = miutility.Connection
       mRsPOReceipients.Source = query
       mRsPOReceipients.CursorLocation = adUseClient
       mRsPOReceipients.LockType = adLockBatchOptimistic
       mRsPOReceipients.Open , , adOpenKeyset, adLockBatchOptimistic

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
   mRsPOReceipients.MoveFirst

   
   If mRsPOReceipients.Eof = False Then
   Move = True

   End If
End If
   
   
   Exit Function
   
handler:

   Err.Clear
End Function

Public Property Get Count() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Count = mRsPOReceipients.RecordCount

End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler

  If Not mRsPOReceipients.Bof Then
   
     mRsPOReceipients.MoveFirst
     If Not mRsPOReceipients.Bof Then
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  Err.Clear
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsPOReceipients.Eof Then
   
     mRsPOReceipients.MoveNext
     If Not mRsPOReceipients.Eof Then
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
  If Not mRsPOReceipients.Bof Then
   
     mRsPOReceipients.MovePrevious
     If mRsPOReceipients.Bof Then
      MovePrevious = False
     Else
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
  If Not mRsPOReceipients.Eof Then
   
     mRsPOReceipients.MoveLast
     If Not mRsPOReceipients.Eof Then 'FillmeWithTheCurrentRecord (mRsHeader)
       MoveLast = False
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
  Eof = mRsPOReceipients.Eof
End If
End Property

Public Property Get Bof() As Boolean
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Bof = mRsPOReceipients.Bof
End Property



Public Property Get Receipient() As String
Receipient = mRsPOReceipients!porc_rec
End Property

Public Property Let Receipient(ByVal vNewValue As String)
mRsPOReceipients!porc_rec = Trim$(vNewValue)
End Property

Public Property Get EditMode() As Integer
If mRsPOReceipients.AbsolutePosition = adPosBOF Then mRsPOReceipients.MoveFirst
If mRsPOReceipients.AbsolutePosition = adPosEOF Then mRsPOReceipients.MoveLast
On Error GoTo handler
If mRsPOReceipients.State = 1 Then
   EditMode = mRsPOReceipients.EditMode
End If
Exit Function
handler:
Err.Raise Err.Number, , Err.Description
Err.Clear

End Property

Public Function SubmitSupplier(Supplier As String, POnumb As String) As Boolean

On Error GoTo handler
SubmitSupplier = False


If Len(Trim$(mSupplier)) > 0 Then

    If mRsPOReceipients.RecordCount > 0 Then
       
       mRsPOReceipients.MoveFirst
       mRsPOReceipients.Find "porc_rec='" & mSupplier & "'"
       
         If Not mRsPOReceipients.AbsolutePosition = adPosEOF Then
        
         
              If Len(Trim$(Supplier)) > 0 Then
               mRsPOReceipients!porc_rec = Supplier
              Else
              
               mRsPOReceipients.CancelUpdate
              
              ' mRsPOReceipients.Delete adAffectCurrent
              
              End If
         Else
                mRsPOReceipients.AddNew
                mRsPOReceipients!PORC_PONUMB = Trim$(POnumb)
                mRsPOReceipients!PORC_NPECODE = miutility.Namespace
                mRsPOReceipients!porc_recpnumb = mRsPOReceipients.RecordCount
                mRsPOReceipients!porc_rec = Supplier
                
              
         End If
    
    Else
      mRsPOReceipients.AddNew
    mRsPOReceipients!PORC_PONUMB = Trim$(POnumb)
    mRsPOReceipients!PORC_NPECODE = miutility.Namespace
    mRsPOReceipients!porc_recpnumb = mRsPOReceipients.RecordCount
    mRsPOReceipients!porc_rec = Supplier
    
       
    End If
Else
    If Len(Trim$(Supplier)) > 0 Then
    
            mRsPOReceipients.AddNew
            mRsPOReceipients!PORC_PONUMB = Trim$(POnumb)
            mRsPOReceipients!PORC_NPECODE = miutility.Namespace
            mRsPOReceipients!porc_recpnumb = mRsPOReceipients.RecordCount
            mRsPOReceipients!porc_rec = Supplier
    End If
End If

mSupplier = Supplier
SubmitSupplier = True
Exit Function
handler:
MsgBox "Error occurred while addting the supplier.Error description " & Err.Description, vbCritical, "Imswin"
Err.Clear
End Function

Public Function DeleteRecepients(PO_ponumb As String) As Boolean
Dim cancel As Boolean
DeleteRecepients = False
On Error GoTo handler

RaiseEvent BeforeDeleteRecepients(cancel)
If cancel = False Then
   
   mRsPOReceipients.Filter = "porc_ponumb='" & PO_ponumb & "'"
   If mRsPOReceipients.RecordCount > 0 Then
       mRsPOReceipients.Delete adAffectGroup
   End If
        
        RaiseEvent AfterDeleteRecepients
        
Else

   cancel = False
           
End If
DeleteRecepients = True
Exit Function
handler:
MsgBox "Error Occurred During Deleting the PO Recepients.Err.desc -- " & Err.Description
Err.Clear
End Function


Public Function Replace(PropertyName As String, Value As String) As Boolean

On Error GoTo ErrHandler

Replace = False

PropertyName = Trim$(PropertyName)

If mRsPOReceipients.RecordCount > 0 Then

    mRsPOReceipients.MoveFirst
    
Else
    
    Replace = True
    
    Exit Function

End If

Select Case PropertyName

    Case "PONUMB"
          
          mRsPOReceipients.MoveFirst
          
          Do While Not mRsPOReceipients.Eof
          
                 mRsPOReceipients!PORC_PONUMB = Value
                
                MPonumb = Value
                
                mRsPOReceipients.MoveNext
                
          Loop
      
End Select

Replace = True

Exit Function

ErrHandler:

'n

End Function



Public Function LoadFromRequsition(RequsitionNo As String, TemporaryPoNumb As String, ErrorIfany As String) As Boolean
Dim RsRequisition As ADODB.Recordset
Dim field As ADODB.field
On Error GoTo ErrHandler
LoadFromRequsition = True

Set RsRequisition = Nothing

If Move(RequsitionNo) = True Then

Set RsRequisition = mRsPOReceipients.Clone

If RsRequisition.RecordCount > 0 Then

    Set mRsPOReceipients = Nothing
    
    Move TemporaryPoNumb
    
    RsRequisition.MoveFirst
    
    Do While Not RsRequisition.Eof
    
        mRsPOReceipients.AddNew
        
        For Each field In RsRequisition.Fields
        
                mRsPOReceipients(field.Name) = field.Value
        
        Next
       
        mRsPOReceipients("porc_ponumb") = TemporaryPoNumb
       
        RsRequisition.MoveNext
       
     Loop
       
     End If
       
    Set RsRequisition = Nothing

End If

LoadFromRequsition = False

Exit Function
ErrHandler:
ErrorIfany = Err.Description

Err.Clear
End Function
