VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "POREMARKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"POREMARK"
Attribute VB_Ext_KEY = "Member0" ,"POREMARK"
Option Explicit

Private mIsItemLoaded As Boolean

'''Public Enum Status
'''  ADSTATUSUnknow = 0
'''  ADSTATUSMODIFIED = 1
'''  ADSTATUSINSERT = 2
'''End Enum

Private mRsPOREMARKs As ADODB.Recordset

Private MStatus  As Status

Private MNamespace As String

Private MPonumb As String

Private mCount As Integer

Private mValueChanged As Boolean

Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteRemarks(cancel As Boolean)
Public Event AfterDeleteRemarks()
  
  
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
POnumb = MPonumb
End Property
Public Property Let POnumb(indata As String)
If Trim$(indata) <> MPonumb Then mIsItemLoaded = False
 mRsPOREMARKs!por_ponumb = Trim$(indata)
  MPonumb = Trim$(indata) 'AM
End Property
Public Property Get Namespace() As String
Namespace = mRsPOREMARKs!por_npecode
End Property
Public Property Let Namespace(ByVal indata As String)
mRsPOREMARKs!por_npecode = Trim$(indata)
End Property
Public Property Get lineNumb() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
lineNumb = mRsPOREMARKs!por_linenumb
End Property

Public Property Let lineNumb(ByVal indata As Integer)
 mRsPOREMARKs!por_linenumb = (Trim$(indata))
End Property

Public Function AddNew() As Boolean
Dim cancel As Boolean
AddNew = False
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
       If cancel = False Then
           mRsPOREMARKs.AddNew
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
mRsPOREMARKs.UpdateBatch
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
 
  query = "SELECT por_ponumb, por_npecode, por_linenumb, por_remk FROM POREM WHERE "
  query = query & "(por_ponumb = '" & MPonumb & "') AND (por_npecode = '" & miutility.Namespace & "') ORDER BY por_linenumb "
       Set mRsPOREMARKs = New ADODB.Recordset
       
       mRsPOREMARKs.ActiveConnection = miutility.Connection
       mRsPOREMARKs.Source = query
       mRsPOREMARKs.CursorLocation = adUseClient
       mRsPOREMARKs.LockType = adLockBatchOptimistic
       mRsPOREMARKs.Open , , adOpenKeyset, adLockBatchOptimistic

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
   mRsPOREMARKs.MoveFirst

   
   If mRsPOREMARKs.Eof = False Then
   Move = True

   End If
End If
   
   
   Exit Function
   
handler:

   Err.Clear
End Function

Public Property Get Count() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Count = mRsPOREMARKs.RecordCount

End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler

  If Not mRsPOREMARKs.Bof Then
   
     mRsPOREMARKs.MoveFirst
     If Not mRsPOREMARKs.Bof Then
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsPOREMARKs.Eof Then
   
     mRsPOREMARKs.MoveNext
      If mRsPOREMARKs.AbsolutePosition = adPosEOF And mRsPOREMARKs.Eof = True And mRsPOREMARKs.Bof <> True Then mRsPOREMARKs.MoveLast
     If Not mRsPOREMARKs.Eof Then
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
  If Not mRsPOREMARKs.Bof Then
   
     mRsPOREMARKs.MovePrevious
       If mRsPOREMARKs.AbsolutePosition = adPosBOF And mRsPOREMARKs.Eof <> True And mRsPOREMARKs.Bof = True Then mRsPOREMARKs.MoveFirst
     If mRsPOREMARKs.Bof Then
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
  If Not mRsPOREMARKs.Eof Then
   
     mRsPOREMARKs.MoveLast
     If Not mRsPOREMARKs.Eof Then 'FillmeWithTheCurrentRecord (mRsHeader)
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
  Eof = mRsPOREMARKs.Eof
End If
End Property

Public Property Get Bof() As Boolean
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Bof = mRsPOREMARKs.Bof
End Property



Public Property Get Remarks() As String
Remarks = mRsPOREMARKs!por_remk
End Property

Public Property Let Remarks(ByVal vNewValue As String)
mRsPOREMARKs!por_remk = Trim$(vNewValue)
End Property

Public Property Get EditMode() As Integer
If mRsPOREMARKs.AbsolutePosition = adPosBOF Then mRsPOREMARKs.MoveFirst
If mRsPOREMARKs.AbsolutePosition = adPosEOF Then mRsPOREMARKs.MoveLast
On Error GoTo handler
If mRsPOREMARKs.State = 1 Then
   EditMode = mRsPOREMARKs.EditMode
End If
Exit Property
handler:
Err.Raise Err.Number, , Err.Description
Err.Clear

End Property
Public Function CancelUpdate()
On Error GoTo handler
'''If mRsPOREMARKs.AbsolutePosition = adPosBOF Then mRsPOREMARKs.MoveFirst
'''If mRsPOREMARKs.AbsolutePosition = adPosEOF Then mRsPOREMARKs.MoveLast
'''If mRsPOREMARKs.AbsolutePosition = adPosUnknown Then mRsPOREMARKs.MoveLast
'''
'''   mRsPOREMARKs.CancelUpdate



If mRsPOREMARKs.AbsolutePosition = adPosBOF And mRsPOREMARKs.Eof <> True And mRsPOREMARKs.Bof = True Then mRsPOREMARKs.MoveFirst
If mRsPOREMARKs.AbsolutePosition = adPosEOF And mRsPOREMARKs.Bof <> True And mRsPOREMARKs.Eof = True Then mRsPOREMARKs.MoveLast
If mRsPOREMARKs.AbsolutePosition = adPosUnknown And mRsPOREMARKs.RecordCount > 0 Then mRsPOREMARKs.MoveLast
If mRsPOREMARKs.EditMode = adEditAdd Then
   mRsPOREMARKs.CancelUpdate
End If

If mRsPOREMARKs.AbsolutePosition = adPosBOF And mRsPOREMARKs.Eof <> True And mRsPOREMARKs.Bof = True Then mRsPOREMARKs.MoveFirst
If mRsPOREMARKs.AbsolutePosition = adPosEOF And mRsPOREMARKs.Bof <> True And mRsPOREMARKs.Eof = True Then mRsPOREMARKs.MoveLast
If mRsPOREMARKs.AbsolutePosition = adPosUnknown And mRsPOREMARKs.RecordCount > 0 Then mRsPOREMARKs.MoveLast



Exit Function
   
handler:
   Err.Clear
End Function
Public Function DeleteAllRemarks(PO_ponumb As String) As Boolean
Dim cancel As Boolean
DeleteAllRemarks = False
On Error GoTo handler

RaiseEvent BeforeDeleteRemarks(cancel)
  
If cancel = False Then
   
   mRsPOREMARKs.Filter = "por_ponumb='" & PO_ponumb & "'"
   If mRsPOREMARKs.RecordCount > 0 Then
       mRsPOREMARKs.Delete adAffectGroup
   End If
        
        RaiseEvent AfterDeleteRemarks
        
Else

   cancel = False
           
End If
DeleteAllRemarks = True
Exit Function
handler:
MsgBox "Error Occurred During Deleting the PO Recepients.Err.desc -- " & Err.Description
Err.Clear
End Function

Public Function Replace(PropertyName As String, Value As String) As Boolean

On Error GoTo ErrHandler

Replace = False

PropertyName = Trim$(PropertyName)

If mRsPOREMARKs.RecordCount > 0 Then

    mRsPOREMARKs.MoveFirst
    
Else
    
    Replace = True
    
    Exit Function

End If

Select Case PropertyName

    Case "PONUMB"
          
          mRsPOREMARKs.MoveFirst
          
          Do While Not mRsPOREMARKs.Eof
          
                 mRsPOREMARKs!por_ponumb = Value
                
                MPonumb = Value
                
                mRsPOREMARKs.MoveNext
                
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

Set RsRequisition = mRsPOREMARKs.Clone

If RsRequisition.RecordCount > 0 Then

    Set mRsPOREMARKs = Nothing
    
    Move TemporaryPoNumb
    
    RsRequisition.MoveFirst
    
    Do While Not RsRequisition.Eof
    
        mRsPOREMARKs.AddNew
        
        For Each field In RsRequisition.Fields
        
                mRsPOREMARKs(field.Name) = field.Value
        
        Next
       
        mRsPOREMARKs("por_ponumb") = TemporaryPoNumb
       
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

