VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mStockNumber As String
Private mPriUnit As String
Private mSecUnit As String
Private mCompFactor As Double
Private miutility As Iutility
Private mDesc As String
Private mRatio As Integer
Public Property Get ratioValue() As Integer
    ratioValue = mRatio
End Property

'Private mIsObjectLoaded As Boolean
Public Property Get StockNumber() As String
StockNumber = mStockNumber
End Property

Public Property Let StockNumber(ByVal vNewValue As String)
mStockNumber = Trim$(vNewValue)
Call LoadObject
  
End Property

Public Property Get PrimaryUnit() As String
PrimaryUnit = mPriUnit
End Property


Public Property Get SecondaryUnit() As String
SecondaryUnit = mSecUnit
End Property


Public Property Get ComPutationFactor() As String
ComPutationFactor = mCompFactor
End Property

Public Property Get Description() As String
Description = mDesc
End Property

Public Function Configure(Iutilty As Iutility)
Set miutility = Iutilty
End Function

Public Function LoadObject() As Boolean
LoadObject = False
On Error GoTo handler
Dim Rs As New ADODB.Recordset
  Rs.Source = "select stk_stcknumb, stk_desc,stk_primuon, stk_secouom, stk_compfctr, stk_ratio2  from stockmaster where stk_stcknumb='" & mStockNumber & "' and stk_npecode='" & miutility.Namespace & "'"
  Rs.ActiveConnection = miutility.Connection
  Rs.CursorType = adOpenKeyset
  Rs.Open
  
  If Not Rs.Eof Then
     mPriUnit = Trim$(Rs!stk_primuon)
     mSecUnit = Trim$(Rs!stk_secouom)
     mCompFactor = IIf(IsNull(Rs!stk_compfctr), 0, (Rs!stk_compfctr))
     mDesc = Trim$(Rs!stk_desc)
     mRatio = Rs!stk_ratio2
  End If
  LoadObject = True
  Exit Function
handler:

   Err.Clear
   
End Function




