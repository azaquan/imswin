VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StockHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNamespace As String

Private mConnection  As ADODB.Connection

Private mRsStock As ADODB.Recordset
Private mStocknumber As String
Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteALLPOITEM(cancel As Boolean)
Public Event AfterDeletePOITEMS()

Public Property Get StockNumber() As String
StockNumber = mRsStock!stk_stcknumb
End Property

Public Property Let StockNumber(ByVal indata As String)
'If Not mRsStock!poi_ponumb = (Trim$(indata)) Then mValueChanged = True
mRsStock!stk_stcknumb = (Trim$(indata))
'MPonumb = mRsStock!stk_stcknumb 'AM
End Property

Public Property Get NameSpace() As String

NameSpace = mNamespace
End Property
Public Property Let NameSpace(ByVal indata As String)
mNamespace = indata
' mRsStock!stk_npecode = indata
End Property

Public Function AddNew() As Boolean
Dim cancel As Boolean
AddNew = False
  
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
  
  If cancel = False Then
           
           'If mRsStock Is Nothing Then Set mRsStock = New ADODB.Recordset
           mStocknumber = ""
           Call LoadCollection
           
           mRsStock.AddNew
           SetTheNewHeader
   End If
   
  cancel = True
  
  AddNew = True
  
  Exit Function

handler:
     

  Err.Clear
End Function

Public Function Update() As String

 On Error GoTo handler
 
     Update = "UnIndetified Errors."
 
     mRsStock.Update
   
     If mRsStock.Eof = False Then
     
     End If
 
     Update = ""
     
     Exit Function
     
handler:

    If Err.Number = "-2147217864" Then
        Update = ""
        Err.Clear
        mConnection.Errors.Clear
    Else
        Update = Err.Description
        'MsgBox "Errors Occured in Saving POITEMS  & " & Err.Description
        Err.Clear
    End If

'Update = Err.Description

Err.Clear
    
End Function

Private Function LoadCollection() As Boolean
LoadCollection = False
On Error GoTo handler

Set mRsStock = Nothing

If mRsStock Is Nothing Then Set mRsStock = New ADODB.Recordset
''mRsStock.Source = "select *, e.eccn_no, p.source from stockmaster stk"
''mRsStock.Source = mRsStock.Source & " left outer join eccn e on e.eccnid = stk.stk_eccnid "
''mRsStock.Source = mRsStock.Source & " left outer join picklist p on p.sourceid = stk.stk_eccnsourceid "
''mRsStock.Source = mRsStock.Source & " where stk_npecode='" & mNamespace & "' and stk_stcknumb ='" & mStocknumber & "'"

mRsStock.Source = "select * from stockmaster stk"
mRsStock.Source = mRsStock.Source & " where stk_npecode='" & mNamespace & "' and stk_stcknumb ='" & mStocknumber & "'"


mRsStock.Open , mConnection, adOpenDynamic, adLockOptimistic

   LoadCollection = True
   
   Exit Function

handler:
   Err.Raise Err.Number, "IMSPODLL.POITEMS.Loadcollection", Err.Description
   Err.Clear

End Function


Public Function MoveToStocknumber(StockNumber As String, Optional Force As Boolean) As Integer
'Returns 0 for Success
'Returns 1 for Unidentified Error
'Returns 2 for No Records

On Error GoTo handler

MoveToStocknumber = 1

If Trim(UCase((mStocknumber))) <> Trim(UCase(StockNumber)) Or Force = True Then

    mStocknumber = Trim(UCase(StockNumber))

    If LoadCollection = False Then Exit Function

    If mRsStock.state = 0 Then MoveToStocknumber = 2: Exit Function
    
    If mRsStock.state = 1 And mRsStock.RecordCount = 0 Then MoveToStocknumber = 2: Exit Function

    If mRsStock.RecordCount > 0 Then mRsStock.MoveFirst
            
End If
   
   
   MoveToStocknumber = 0
   
   Exit Function
   
handler:

   Err.Clear
   
End Function

Public Property Get Count() As Integer

Count = mRsStock.RecordCount
End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler
  If Not mRsStock.Bof Then
   
     mRsStock.MoveFirst
     If mRsStock.Bof = True Then 'FillmeWithTheCurrentRecord (mRsStock)
       MoveFirst = False
     Else
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsStock.Eof Then
   
     mRsStock.MoveNext
    ' If mRsStock.AbsolutePosition = adPosEOF And mRsStock.Eof = True And mRsStock.Bof <> True Then mRsStock.MoveLast 'AM
     
     If mRsStock.AbsolutePosition = adPosEOF Then mRsStock.MoveLast 'AM
     If Not mRsStock.Eof Then
        MoveNext = True
     End If
  End If
  Exit Function
handler:
  
End Function

Public Function MovePrevious() As Boolean
MovePrevious = False
On Error GoTo handler
  If Not mRsStock.Bof Then
   
     mRsStock.MovePrevious
     If mRsStock.AbsolutePosition = adPosBOF And mRsStock.Eof <> True And mRsStock.Bof = True Then mRsStock.MoveFirst
     If mRsStock.Bof Then
      MovePrevious = False
     Else
      MovePrevious = True
     End If
  End If
 Exit Function
handler:
  
End Function

Public Function MoveLast() As Boolean

MoveLast = False
On Error GoTo handler
  If Not mRsStock.Eof Then
   
     mRsStock.MoveLast
     If mRsStock.Eof = True Then
       MoveLast = False
     Else
       MoveLast = True
     End If
  End If
  Exit Function
handler:
  Err.Clear
End Function

Private Sub SetTheNewHeader()
    mRsStock!stk_npecode = mNamespace
    'Juan 2010-11-9 Default values when creating a new recordd
'    mRsStock!stk_mini = 0
'    mRsStock!stk_maxi = 0
'    mRsStock!stk_characctcode = "N/A"
'    mRsStock!stk_stcknumb = "NEW"
'    mRsStock!stk_desc = "NEW"
End Sub


Public Property Get Eof() As Boolean


  Eof = mRsStock.Eof

End Property

Public Property Get Bof() As Boolean

Bof = mRsStock.Bof
End Property


Public Function CancelUpdate() As Boolean
On Error GoTo handler
If mRsStock.AbsolutePosition = adPosBOF And mRsStock.Eof <> True And mRsStock.Bof = True Then mRsStock.MoveFirst
If mRsStock.AbsolutePosition = adPosEOF And mRsStock.Bof <> True And mRsStock.Eof = True Then mRsStock.MoveLast
If mRsStock.AbsolutePosition = adPosUnknown And mRsStock.RecordCount > 0 Then mRsStock.MoveLast
If mRsStock.EditMode = adEditAdd Then
   mRsStock.CancelUpdate
   
'ElseIf mRsStock.EditMode = 1 Then


   
End If

If mRsStock.AbsolutePosition = adPosBOF And mRsStock.Eof <> True And mRsStock.Bof = True Then mRsStock.MoveFirst
If mRsStock.AbsolutePosition = adPosEOF And mRsStock.Bof <> True And mRsStock.Eof = True Then mRsStock.MoveLast
If mRsStock.AbsolutePosition = adPosUnknown And mRsStock.RecordCount > 0 Then mRsStock.MoveLast

Exit Function
   
handler:
   Err.Clear
End Function

Public Function CancelBatch()
On Error GoTo handler
If mRsStock.EditMode = adEditAdd Then
   mRsStock.CancelBatch
End If
Exit Function
   
handler:
   Err.Clear
End Function


Public Property Get EditMode() As String
If mRsStock.state = 1 And mRsStock.RecordCount > 0 Then
  EditMode = mRsStock.EditMode
Else
   EditMode = "-1"
End If
End Property


Public Function Requery()
On Error GoTo handler
   mRsStock.Requery
   Exit Function
   
handler:

  MsgBox "Error When Poitem tried to Requery Database"
  Err.Clear
End Function


Public Function AnyChangesBeforeSaving() As Boolean
On Error GoTo handler

AnyChangesBeforeSaving = False

AnyChangesBeforeSaving = True
Exit Function
handler:

   Err.Clear
End Function


Public Property Get Description() As String
Description = mRsStock!stk_desc & ""
End Property

Public Property Let Description(ByVal vNewValue As String)
mRsStock!stk_desc = vNewValue
End Property

Public Property Get DescFlag() As Boolean
DescFlag = mRsStock!stk_descflag
End Property

Public Property Let DescFlag(ByVal vNewValue As Boolean)
mRsStock!stk_descflag = vNewValue
End Property

Public Property Get PrimUOfMeasure() As String
PrimUOfMeasure = mRsStock!stk_primuon & ""
End Property

Public Property Let PrimUOfMeasure(ByVal vNewValue As String)
mRsStock!stk_primuon = vNewValue
End Property
Public Property Get SecoUOfMeasure() As String
 SecoUOfMeasure = mRsStock!stk_secouom & ""
End Property

Public Property Let SecoUOfMeasure(ByVal vNewValue As String)
mRsStock!stk_secouom = vNewValue
End Property

Public Property Get StockType() As String
StockType = mRsStock!stk_stcktype & ""
End Property

Public Property Let StockType(ByVal vNewValue As String)
mRsStock!stk_stcktype = vNewValue
End Property

Public Property Get CategoryCode() As String
CategoryCode = mRsStock!stk_catecode & ""
End Property

Public Property Get ratio1() As Double
ratio1 = mRsStock!stk_ratio1
End Property

Public Property Let CategoryCode(ByVal vNewValue As String)
mRsStock!stk_catecode = vNewValue
End Property

Public Property Let ratio1(ByVal ratio1Value As Double)
mRsStock!stk_ratio1 = ratio1Value
End Property
Public Property Get ratio2() As Double
    ratio2 = mRsStock!stk_ratio2
End Property

Public Property Let ratio2(ByVal ratio2Value As Double)
    mRsStock!stk_ratio2 = ratio2Value
End Property
Public Property Get ComputationFactor() As Double
'ComputationFactor = IIf(IsNull(mRsStock!stk_compfctr) = True, CDbl("0.00"), CDbl(mRsStock!stk_compfctr & ""))

If IsNull(mRsStock!stk_compfctr) = True Then

    ComputationFactor = 0

Else

    ComputationFactor = CDbl(mRsStock!stk_compfctr)

End If

End Property

Public Property Let ComputationFactor(ByVal vNewValue As Double)
mRsStock!stk_compfctr = vNewValue
End Property
Public Property Get PoolOrSpecific() As String
PoolOrSpecific = mRsStock!stk_poolspec
End Property
Public Property Let PoolOrSpecific(ByVal vNewValue As String)
mRsStock!stk_poolspec = vNewValue
End Property

Public Property Get hazmatclau() As String
hazmatclau = mRsStock!stk_hazmatclau & ""
End Property

Public Property Let hazmatclau(ByVal vNewValue As String)
mRsStock!stk_hazmatclau = vNewValue
End Property

Public Property Get Minimum() As String
Minimum = mRsStock!stk_mini & ""
End Property

Public Property Let Minimum(ByVal vNewValue As String)
mRsStock!stk_mini = vNewValue
End Property

Public Property Get Maximum() As String
Maximum = mRsStock!stk_maxi & ""
End Property

Public Property Let Maximum(ByVal vNewValue As String)
mRsStock!stk_maxi = vNewValue
End Property

Public Property Get characctcode() As String
characctcode = mRsStock!stk_characctcode & ""
End Property

Public Property Let characctcode(ByVal vNewValue As String)
mRsStock!stk_characctcode = IIf(Len(Trim(vNewValue)) = 0, Null, vNewValue)   '07/02/02 tHIS IS HOW IT IS STORED IN THE PREVIOUS VERSION OF STOCKMASTER
End Property

Public Property Get stdrcost() As String
stdrcost = mRsStock!stk_stdrcost & ""
End Property

Public Property Let stdrcost(ByVal vNewValue As String)

If Len(Trim(vNewValue & "")) = 0 Then

    mRsStock!stk_stdrcost = Null

Else

    mRsStock!stk_stdrcost = CDbl(vNewValue)

End If

'mRsStock!stk_stdrcost = IIf(Len(Trim(vNewValue)) > 0, CDbl(vNewValue), Null)
End Property

Public Property Get estmprice() As String
estmprice = mRsStock!stk_estmprice & ""
End Property

Public Property Let estmprice(ByVal vNewValue As String)

If Len(Trim(vNewValue & "")) = 0 Then

    mRsStock!stk_estmprice = Null

Else

    mRsStock!stk_estmprice = CDbl(vNewValue)

End If

'mRsStock!stk_estmprice = IIf(Len(Trim(vNewValue)) = 0, Null, CDbl(vNewValue))
End Property

Public Property Get Grpe() As String
Grpe = mRsStock!stk_grpe & ""
End Property

Public Property Let Grpe(ByVal vNewValue As String)
mRsStock!stk_grpe = IIf(Len(Trim(vNewValue)) = 0, Null, vNewValue) '07/02/02 tHIS IS HOW IT IS STORED IN THE PREVIOUS VERSION OF STOCKMASTER
End Property

Public Property Get Techspec() As String
Techspec = mRsStock!stk_techspec & ""
End Property

Public Property Let Techspec(ByVal vNewValue As String)
mRsStock!stk_techspec = IIf(Len(Trim(vNewValue)) = 0, Null, vNewValue) '07/02/02 tHIS IS HOW IT IS STORED IN THE PREVIOUS VERSION OF STOCKMASTER
End Property

Public Property Get Activeflag() As Boolean
Activeflag = mRsStock!stk_flag
End Property

Public Property Let Activeflag(ByVal vNewValue As Boolean)
mRsStock!stk_flag = vNewValue
End Property

Public Property Get Connection() As ADODB.Connection
Set Connection = mConnection
End Property

Public Property Let Connection(ByVal vNewValue As ADODB.Connection)
Set mConnection = vNewValue
End Property

Public Property Let CreateUser(ByVal vNewValue As String)
mRsStock!stk_creauser = vNewValue
End Property

Public Property Let ModiUser(ByVal vNewValue As String)
mRsStock!stk_modiuser = vNewValue
End Property

Public Property Let Eccnid(ByVal vNewValue As String)
mRsStock!stk_eccnid = vNewValue
End Property

Public Property Get Eccnid() As String
Eccnid = mRsStock!stk_eccnid & ""
End Property

''Public Property Let Eccnno(ByVal vNewValue As String)
''mRsStock!eccn_no = vNewValue
''End Property
''Public Property Get Eccnno() As String
''Eccnno = mRsStock!eccn_no & ""
''End Property

Public Property Let Eccnlicsreq(ByVal vNewValue As Boolean)
mRsStock!stk_Eccnlicsreq = vNewValue
End Property
Public Property Get Eccnlicsreq() As Boolean
Eccnlicsreq = IIf(IsNull(mRsStock!stk_Eccnlicsreq), 0, mRsStock!stk_Eccnlicsreq)
End Property

Public Property Let Eccnsource(ByVal vNewValue As Integer)
mRsStock!stk_eccnsourceid = vNewValue
End Property

Public Property Get Eccnsource() As Integer
Eccnsource = IIf(IsNull(mRsStock!stk_eccnsourceid), 0, mRsStock!stk_eccnsourceid)
End Property

''Public Property Let EccnsourceName(ByVal vNewValue As String)
''mRsStock!Source = vNewValue
''End Property
''
''Public Property Get EccnsourceName() As String
''EccnsourceName = mRsStock!Source & ""
''End Property

