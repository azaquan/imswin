VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LookUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mNamespace As String
Private mConnection As ADODB.Connection

Public Property Get NameSpace() As String
NameSpace = mNamespace
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
mNamespace = vNewValue
End Property

Public Property Get Connection() As ADODB.Connection
Set Connection = mConnection
End Property

Public Property Let Connection(ByVal vNewValue As ADODB.Connection)
Set mConnection = vNewValue
End Property

Public Function GetManufacturers(Rs As ADODB.Recordset) As Integer

Dim RsManufacturer As New ADODB.Recordset

On Error GoTo Errhandler

GetManufacturers = 1

RsManufacturer.Source = "SELECT man_code, man_name FROM MANUFACTURER WHERE (man_npecode = '" & mNamespace & "') ORDER BY man_code, man_name"

RsManufacturer.Open , mConnection, adOpenKeyset

Set Rs = RsManufacturer

GetManufacturers = 0

Exit Function
Errhandler:

Err.Clear

End Function

Public Function GetStocknumbers() As ADODB.Recordset

On Error GoTo Errhandler

Dim RsStockMaster As New ADODB.Recordset

RsStockMaster.Source = "SELECT stk_stcknumb, stk_desc FROM STOCKMASTER WHERE (stk_npecode = '" & mNamespace & "')  order by stk_stcknumb"

RsStockMaster.Open , mConnection, adOpenKeyset, adLockPessimistic

Set GetStocknumbers = RsStockMaster

Exit Function

Errhandler:

Err.Clear

End Function

Public Function DoesStockExist(StockNumber As String) As Integer

On Error GoTo Errhandler

Dim RsStockMaster As New ADODB.Recordset

DoesStockExist = 1
StockNumber = Trim(StockNumber)
RsStockMaster.Source = "SELECT count(*) X from stockmaster where stk_npecode ='" & mNamespace & "' and stk_stcknumb ='" & StockNumber & "'"

RsStockMaster.Open , mConnection, adOpenKeyset, adLockPessimistic

DoesStockExist = RsStockMaster("x")

Exit Function

Errhandler:

Err.Clear

End Function

Public Function HowManyManufacturers(StockNumber As String) As Integer

Dim RsManuCount As New ADODB.Recordset

On Error GoTo Errhandler

HowManyManufacturers = 0

RsManuCount.Source = " select count(*) X from stockmanufac where stm_stcknumb ='" & StockNumber & "' and stm_npecode ='" & mNamespace & "'"

RsManuCount.Open , mConnection

HowManyManufacturers = RsManuCount("x")

Exit Function

Errhandler:

Err.Clear

End Function
'All for 0
'Active for 1
Public Function GetListofEccns(AllORActive As Integer) As ADODB.Recordset

On Error GoTo Errhandler

Dim RsEccn As New ADODB.Recordset

RsEccn.Source = "SELECT eccnid, eccn_no, eccn_desc FROM eccn WHERE (eccn_npecode = '" & mNamespace & "')"
If AllORActive = 1 Then RsEccn.Source = RsEccn.Source & " and eccn_active =1"
RsEccn.Source = RsEccn.Source & " order by eccn_no "
RsEccn.Open , mConnection, adOpenKeyset, adLockPessimistic

Set GetListofEccns = RsEccn

Exit Function

Errhandler:

Err.Clear

End Function
Public Function GetListofEccnSource(AllORActive As Integer) As ADODB.Recordset

On Error GoTo Errhandler

Dim RsEccn As New ADODB.Recordset

RsEccn.Source = "SELECT sourceid, source FROM picklist WHERE (npecode = '" & mNamespace & "')"
If AllORActive = 1 Then RsEccn.Source = RsEccn.Source & " and source_active =1"
RsEccn.Open , mConnection, adOpenKeyset, adLockPessimistic

Set GetListofEccnSource = RsEccn

Exit Function

Errhandler:

Err.Clear

End Function
Public Function GetStockTypes() As ADODB.Recordset

On Error GoTo Errhandler

Dim RsStockT As New ADODB.Recordset

RsStockT.Source = "SELECT sty_stcktype FROM stocktype WHERE (sty_npecode = '" & mNamespace & "')"

RsStockT.Open , mConnection, adOpenKeyset, adLockPessimistic

Set GetStockTypes = RsStockT

Exit Function

Errhandler:

Err.Clear

End Function



