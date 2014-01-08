VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMDIActiveX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub ShowForm1(transaction As kindTRANSAC)
Dim f As frmWarehouse
Dim WH As Warehouse
  Set f = New frmWarehouse
  f.Show
  'Set WH = New Warehouse
  'Call WH.Loading(transaction)
End Sub
