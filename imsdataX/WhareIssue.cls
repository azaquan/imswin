VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsWhareIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A750372"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=38955A760016
Private FTranscationNumber As Variant
'##ModelId=38955A760020
Private FUser As Variant
'##ModelId=38955A760034
Private FTransactionDate As Variant
'##ModelId=38955A76003E
Private FCompany As Variant
'##ModelId=38955A760052
Private FNameSpace As Variant
'##ModelId=38955A76005C
Private FStockNumber As Variant
'##ModelId=38955A760070
Private FCondition As Variant
'##ModelId=38955A76007A
Private FConnection As Variant
'##ModelId=38955A76008E
Private FIssueTo As Variant
'##ModelId=38955A760098
Private FWhareHouse As Variant

'##ModelId=38955A7600A2
Public Property Let WhareHouse(ByVal vData As Variant)
    FWhareHouse = vData
End Property

'##ModelId=38955A7600C0
Public Property Get WhareHouse() As Variant
    WhareHouse = FWhareHouse
End Property

'##ModelId=38955A7600CA
Public Property Let IssueTo(ByVal vData As Variant)
    FIssueTo = vData
End Property

'##ModelId=38955A7600E8
Public Property Get IssueTo() As Variant
    IssueTo = FIssueTo
End Property

'##ModelId=38955A7600FC
Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set FConnection = vData
End Property

'##ModelId=38955A760110
Public Property Get Connection() As ADODB.Connection
    Connection = FConnection
End Property

'##ModelId=38955A760124
Public Property Let Condition(ByVal vData As Variant)
    FCondition = vData
End Property

'##ModelId=38955A760138
Public Property Get Condition() As Variant
    Condition = FCondition
End Property

'##ModelId=38955A76014C
Public Property Let StockNumber(ByVal vData As Variant)
    FStockNumber = vData
End Property

'##ModelId=38955A760161
Public Property Get StockNumber() As Variant
    StockNumber = FStockNumber
End Property

'##ModelId=38955A760175
Public Property Let NameSpace(ByVal vData As Variant)
    FNameSpace = vData
End Property

'##ModelId=38955A760189
Public Property Get NameSpace() As Variant
    NameSpace = FNameSpace
End Property

'##ModelId=38955A76019D
Public Property Let Company(ByVal vData As Variant)
    FCompany = vData
End Property

'##ModelId=38955A7601B1
Public Property Get Company() As Variant
    Company = FCompany
End Property

'##ModelId=38955A7601C5
Public Property Get TransactionDate() As Variant
    TransactionDate = FTransactionDate
End Property

'##ModelId=38955A7601D9
Public Property Let User(ByVal vData As Variant)
    FUser = vData
End Property

'##ModelId=38955A7601ED
Public Property Get User() As Variant
    User = FUser
End Property

'##ModelId=38955A760201
Public Property Let TranscationNumber(ByVal vData As Variant)
    FTranscationNumber = vData
End Property

'##ModelId=38955A760215
Public Property Get TranscationNumber() As Variant
Attribute TranscationNumber.VB_UserMemId = 0
    TranscationNumber = FTranscationNumber
End Property
