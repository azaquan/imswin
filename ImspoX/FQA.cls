VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

''Public FromCompany As String
''Public FromLocation As String
''Public FromUSChart As String
''Public FromCamChart As String
''Public FromStockType As String
Private MPonumb As String
Public User As String
Private mRsHeader As ADODB.Recordset
'Private mRsHeader1 As adodb.Recordset
Private miutility As Iutility
Public Function Configure(Iutilty As Iutility)

    Set miutility = Iutilty

End Function
Public Function GetFQAInfo(TransactioNo As String) As Boolean

'Dim mRsheader As adodb.Recordset

On Error GoTo ErrHand
If MPonumb <> TransactioNo Then

MPonumb = TransactioNo

Set mRsHeader = New ADODB.Recordset
mRsHeader.Source = "SELECT * from poFQA where ponumb ='" & TransactioNo & "' and Npce_code ='" & miutility.Namespace & "' order by itemno"
mRsHeader.Open , miutility.Connection, adOpenKeyset, adLockBatchOptimistic

End If
GetFQAInfo = True
Exit Function
ErrHand:
Err.Clear
End Function

Public Function AddNew() As Boolean

AddNew = False

On Error GoTo handler
  
 If mRsHeader Is Nothing Then Call GetFQAInfo("#$%W#$#dd")
 
    mRsHeader.AddNew
    mRsHeader!npce_code = miutility.Namespace
    mRsHeader!CREADATE = Now
    mRsHeader!modiDATE = Now
    mRsHeader!creauser = "DBO"
    mRsHeader!modiuser = "DBO"
    mRsHeader!TBS = 1
 
    If mRsHeader.RecordCount = 1 Then
    
        mRsHeader!ItemNo = 0
''        mRsheader!FromCompany = "X"
''        mRsheader!FromLocation = "X"
''        mRsheader!FromUsChar = "X"
''        mRsheader!FromStockType = "X"
''        mRsheader!FromCamChar = "X"
        mRsHeader!ToCompany = "x"
        mRsHeader!Tolocation = "x"
        mRsHeader!ToUsChar = "x"
        mRsHeader!ToStockType = "x"
        mRsHeader!ToCamChar = "x"
            
    End If
 
 AddNew = True
  
 Exit Function

handler:
     

 Err.Clear

End Function

''Public Function AddNewLine() As Boolean
''
''AddNewLine = False
''
''On Error GoTo handler
''
'' mRsheader.AddNew
''         mRsheader!CREADATE = Now
''         mRsheader!TBS = 1
''         mRsheader!TransactionType = "PO"
''         mRsheader!npce_code = miutility.Namespace
'' AddNewLine = True
''
'' Exit Function
''
''handler:
''
''
'' Err.Clear
''
''End Function

Public Function DeleteLine() As Boolean
Dim varBookmark  As Variant
On Error GoTo handler
DeleteCurrentLI = False


If Not mRsHeader.AbsolutePosition = adPosEOF Then
     
  mRsHeader.Delete adAffectCurrent
  
  If mRsHeader.RecordCount > 0 Then
            mRsHeader.MoveNext
            
            If mRsHeader.AbsolutePosition = adPosEOF Then
               mRsHeader.MoveFirst
            End If
        
            varBookmark = mRsHeader.Bookmark
            mRsHeader.MoveFirst
            
            Do While Not mRsHeader.Eof
               mRsHeader!ItemNo = mRsHeader.AbsolutePosition - 1
               mRsHeader.MoveNext
            Loop
            mRsHeader.Bookmark = varBookmark
     
  End If
  
End If
DeleteCurrentLI = True
Exit Function
handler:
   MsgBox "Error while deleting the line item." & Err.Description
   Err.Clear
End Function

''Public Function Save(Ponumb As String) As Boolean
''
'' On Error GoTo handler
'' Save = False
''
''
''     mRsheader.UpdateBatch
''
''     If mRsheader.EditMode = 2 Then mRsHeader1("CreaUser") = User
''     mRsHeader1("ModiUser") = User
''
''     mRsHeader1.Update
''
''     If mRsheader.Eof = False Then
''     End If
''
''
''     Save = True
''     Exit Function
''handler:
''
''        Err.Clear
''
''End Function

Public Property Get FromLocation() As Variant


If mRsHeader.Eof = False Then
    FromLocation = mRsHeader("fromlocation")
Else
    FromLocation = ""
End If


End Property

Public Property Let FromLocation(ByVal vNewValue As Variant)
mRsHeader("fromlocation") = vNewValue
End Property

Public Property Get FromCompany() As Variant


If mRsHeader.Eof = False Then
    FromCompany = mRsHeader("fromCompany")
Else
    FromCompany = ""
End If


End Property

Public Property Let FromCompany(ByVal vNewValue As Variant)
mRsHeader("fromCompany") = vNewValue
End Property

Public Property Get ToCompany() As Variant


If mRsHeader.Eof = False Then
    ToCompany = mRsHeader("toCompany")
Else
    ToCompany = ""
End If


End Property

Public Property Let ToCompany(ByVal vNewValue As Variant)
mRsHeader("toCompany") = vNewValue
End Property
Public Property Get Tolocation() As Variant

 

If mRsHeader.Eof = False Then
     Tolocation = mRsHeader("toLocation")
Else
    Tolocation = ""
End If


 
End Property
Public Property Let ItemNo(ByVal vNewValue As Variant)
mRsHeader("ItemNo") = vNewValue
End Property
Public Property Get ItemNo() As Variant

If mRsHeader.Eof = False Then
     ItemNo = mRsHeader("ItemNo")
Else
    ItemNo = ""
End If

End Property

Public Property Let Tolocation(ByVal vNewValue As Variant)
mRsHeader("toLocation") = vNewValue
End Property

Public Property Get FromUSChart() As Variant
If mRsHeader.Eof = False Then
     FromUSChart = mRsHeader("fromUsChar")
Else
    FromUSChart = ""
End If
End Property

Public Property Let FromUSChart(ByVal vNewValue As Variant)
mRsHeader("fromUsChar") = vNewValue
End Property


Public Property Get ToUSChart() As Variant

If mRsHeader.Eof = False Then
     ToUSChart = mRsHeader("toUsChar")
Else
    ToUSChart = ""
End If

End Property

Public Property Let ToUSChart(ByVal vNewValue As Variant)
mRsHeader("toUsChar") = vNewValue
End Property

Public Property Get FromCamChart() As Variant
If mRsHeader.Eof = False Then
    FromCamChart = mRsHeader("fromCamChar")
Else
    FromCamChart = ""
End If


End Property

Public Property Let FromCamChart(ByVal vNewValue As Variant)
mRsHeader("fromCamChar") = vNewValue
End Property


Public Property Get ToCamChart() As Variant
If mRsHeader.Eof = False Then
    ToCamChart = mRsHeader("toCamChar")
Else
    ToCamChart = ""
End If


End Property

Public Property Let ToCamChart(ByVal vNewValue As Variant)
mRsHeader("toCamChar") = vNewValue
End Property

Public Property Get FromStockType() As Variant


If mRsHeader.Eof = False Then
    FromStockType = mRsHeader("FromStockType")
Else
    FromStockType = ""
End If


End Property

Public Property Let FromStockType(ByVal vNewValue As Variant)
mRsHeader("fromStockType") = vNewValue
End Property

Public Property Get ToStockType() As Variant


If mRsHeader.Eof = False Then
    ToStockType = mRsHeader("toStockType")
Else
    ToStockType = ""
End If


End Property

Public Property Let ToStockType(ByVal vNewValue As Variant)
mRsHeader("toStockType") = vNewValue
End Property

''Public Property Get TransactionNo() As Variant
''TransactionNo = mRsHeader("TransactionNo")
''End Property
''
''Public Property Let TransactionNo(ByVal vNewValue As Variant)
''mRsHeader("TransactionNo") = vNewValue
''End Property


''Public Property Let ToItemNo(ByVal vNewValue As Variant)
''mRsheader("ItemNo") = vNewValue
''End Property
''
''Public Property Get ToItemNo() As Variant
''
''
''
''If mRsheader.Eof = False Then
''    ToItemNo = mRsheader("ItemNo")
''Else
''    ToItemNo = ""
''End If
''
''
''End Property

Public Property Let ToTransactionType(ByVal vNewValue As Variant)
mRsHeader("StockType") = vNewValue
End Property
Public Property Get ToTransactionType() As Variant



If mRsHeader.Eof = False Then
   ToTransactionType = mRsHeader("TransactionType")
Else
   ToTransactionType = ""
End If


End Property
Public Property Get creauser() As Variant

If mRsHeader.Eof = False Then
  creauser = mRsHeader("CreaUser")
Else
  creauser = ""
End If


End Property
Public Property Let creauser(ByVal vNewValue As Variant)
mRsHeader("CreaUser") = vNewValue
End Property

Public Property Get modiuser() As Variant



If mRsHeader.Eof = False Then
 modiuser = mRsHeader("ModiUser")
Else
 modiuser = ""
End If



End Property
Public Property Let modiuser(ByVal vNewValue As Variant)
mRsHeader("ModiUser") = vNewValue
End Property

Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsHeader.Eof Then
   
     mRsHeader.MoveNext
     
    ' If mRsHeader.AbsolutePosition = adPosEOF Then mRsHeader.MoveLast
     If Not mRsHeader.Eof Then
        MoveNext = True
     End If
  End If
  Exit Function
handler:
  
End Function

Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler
  If Not mRsHeader.Bof Then
   
     mRsHeader.MoveFirst
     If mRsHeader.Bof = True Then
       MoveFirst = False
     Else
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  Err.Clear
End Function

Public Function MoveLast() As Boolean
MoveLast = False
On Error GoTo handler
  'If Not mRsHeader.Eof Then
   
    If mRsHeader.RecordCount > 0 Then mRsHeader.MoveLast
     
     If mRsHeader.Eof = True Then
       MoveLast = False
     Else
       MoveLast = True
     End If
     
  'End If
  Exit Function
handler:
  Err.Clear
End Function

Public Function MovePrevious() As Boolean
MovePrevious = False
On Error GoTo handler
  If Not mRsHeader.Bof Then
   
     mRsHeader.MovePrevious
     If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
     If mRsHeader.Bof Then
      MovePrevious = False
     Else
      MovePrevious = True
     End If
  End If
 Exit Function
handler:
  
End Function

Public Function Mode() As Integer
   
On Error GoTo ErrHand

Mode = 0

Mode = mRsHeader.EditMode

Exit Function

ErrHand:
Err.Clear

End Function

Public Function GetFQAForTransaction(TranNo As String) As Boolean


On Error GoTo ErrHand

Set mRsHeader = New ADODB.Recordset
mRsHeader.Source = "SELECT * from poFQA where TransactionNo ='" & TranNo & "' and Npce_code ='" & miutility.Namespace & "'"
mRsHeader.Open , miutility.Connection, adOpenStatic, adLockOptimistic
GetTOFQAForTransaction = True

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to get the FQA details for the transaction." & Err.Description, vbCritical, "Ims"
Err.Clear
End Function

'''Public Function GetFROMFQAForTransaction(TranNo As String) As Boolean
'''
'''
'''On Error GoTo ErrHand
'''
'''Set mRsHeader1 = New ADODB.Recordset
'''
'''mRsHeader1.Source = "SELECT * from FromFQA where TransactionNo ='" & TranNo & "' and Npce_code ='" & miutility.Namespace & "'"
'''mRsHeader1.Open , miutility.Connection, adOpenStatic, adLockOptimistic
'''
'''If mRsHeader1.Eof = False Then
'''
''' FromCompany = mRsHeader1("FromCompany")
''' FromLocation = mRsHeader1("FromLocation")
''' FromUSChart = mRsHeader1("FromUsChar")
''' FromCamChart = mRsHeader1("FromCamChar")
''' FromStockType = mRsHeader1("FromStockType")
'''
'''Else
'''
'''    CleanFromValues
'''
'''End If
'''
'''GetFROMFQAForTransaction = True
'''
'''Exit Function
'''ErrHand:
'''
'''MsgBox "Errorsheader1 occurred while trying to get the FQA details for the transaction. " & Err.Description, vbCritical, "Ims"
'''Err.Clear
'''End Function

Public Function MoveLineTo(LineNo As Integer) As Boolean

On Error GoTo Errhandler

MoveLineTo = False

mRsHeader.MoveFirst

mRsHeader.Find "ItemNo ='" & LineNo & "'"

MoveLineTo = True

Exit Function

Errhandler:

Err.Clear

End Function

Public Function Replace(PropertyName As String, Value As String) As Boolean

On Error GoTo Errhandler

Replace = False

mRsHeader!POnumb = Value

PropertyName = Trim$(UCase(PropertyName))

If mRsHeader.RecordCount = 0 Then
 
    Replace = True
    
    Exit Function

End If

Select Case PropertyName

    Case "PONUMB"
          
          mRsHeader.MoveFirst
          
          Do While Not mRsHeader.Eof
          
                mRsHeader!POnumb = Value
                
                MPonumb = Value
                
                mRsHeader.MoveNext
                
          Loop
      
End Select

Replace = True

Exit Function

Errhandler:

'n

Err.Clear

End Function

Public Function Count() As Integer

If mRsHeader Is Nothing Then

    Count = 0

Else

    Count = mRsHeader.RecordCount

End If

End Function

Public Function Eof() As Boolean
Eof = mRsHeader.Eof
End Function

Public Property Get POnumb() As String
If mRsHeader.Eof = False Then
  POnumb = mRsHeader("ponumb")
Else
   POnumb = ""
End If
End Property

Public Property Let POnumb(ByVal vNewValue As String)

mRsHeader("ponumb") = vNewValue

MPonumb = vNewValue

End Property

Public Function Update() As Boolean
On Error GoTo ErrHand

If mRsHeader.RecordCount > 0 Then
    mRsHeader.UpdateBatch
    mRsHeader.MoveFirst
End If
    
    Update = True
Exit Function
ErrHand:

Err.Clear
End Function
Public Function CancelUpdateline() As Boolean
On Error GoTo handler
If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Bof <> True And mRsHeader.Eof = True Then mRsHeader.MoveLast
If mRsHeader.AbsolutePosition = adPosUnknown And mRsHeader.RecordCount > 0 Then mRsHeader.MoveLast
If mRsHeader.EditMode = adEditAdd Then
   mRsHeader.CancelUpdate
   
   If mRsHeader.RecordCount > 0 Then
   
        mRsHeader.MoveFirst
        
        Do While Not mRsHeader.Eof
        
             mRsHeader("ItemNo") = mRsHeader.AbsolutePosition - 1
             
             mRsHeader.MoveNext
        Loop
        
    End If
    
End If

If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Bof <> True And mRsHeader.Eof = True Then mRsHeader.MoveLast
If mRsHeader.AbsolutePosition = adPosUnknown And mRsHeader.RecordCount > 0 Then mRsHeader.MoveLast

'mPonumb = ""

Exit Function
   
handler:
   Err.Clear
End Function

Public Function CancelUpdate() As Boolean
On Error GoTo handler
   mRsHeader.CancelUpdate
Exit Function
   
handler:
   Err.Clear
End Function

