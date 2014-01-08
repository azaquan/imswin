VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsInvoiceDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A74033E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable(s) to hold property value(s)
'##ModelId=38955A7403A2
Private fPONumber As String
Attribute fPONumber.VB_VarDescription = "local variable(s) to hold property value(s)"
'##ModelId=38955A7403B6
Private FNameSpace As String
'##ModelId=38955A7403C1
Private fInvoiceNumber As String
'##ModelId=38955A7403D5
Private fLineItem As String
'##ModelId=38955A7403DF
Private fQuantity As Double
'##ModelId=38955A75000B
Private fTotalPrice As Double
'##ModelId=38955A750015
Private fUnitPrice As Double

'##ModelId=38955A75001F
Public Property Let UnitPrice(ByVal vData As Double)
    fUnitPrice = vData
End Property

'##ModelId=38955A75003D
Public Property Get UnitPrice() As Double
    UnitPrice = fUnitPrice
End Property

'##ModelId=38955A750047
Public Property Let TotalPrice(ByVal vData As Double)
    fTotalPrice = vData
End Property

'##ModelId=38955A750065
Public Property Get TotalPrice() As Double
    TotalPrice = fTotalPrice
End Property

'##ModelId=38955A75006F
Public Property Let Quantity(ByVal vData As Double)
    fQuantity = vData
End Property

'##ModelId=38955A75008D
Public Property Get Quantity() As Double
    Quantity = fQuantity
End Property

'##ModelId=38955A750097
Public Property Let LineItem(ByVal vData As String)
    fLineItem = vData
End Property

'##ModelId=38955A7500E7
Public Property Get LineItem() As String
    LineItem = fLineItem
End Property

'##ModelId=38955A75010F
Public Property Let InvoiceNumber(ByVal vData As String)
    fInvoiceNumber = vData
End Property

'##ModelId=38955A75012D
Public Property Get InvoiceNumber() As String
Attribute InvoiceNumber.VB_UserMemId = 0
Attribute InvoiceNumber.VB_MemberFlags = "200"
    InvoiceNumber = fInvoiceNumber
End Property

'##ModelId=38955A75014B
Public Property Let NameSpace(ByVal vData As String)
    FNameSpace = vData
End Property

'##ModelId=38955A750169
Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property

'##ModelId=38955A75017D
Public Property Let PONumber(ByVal vData As String)
    fPONumber = vData
End Property


'##ModelId=38955A750191
Public Property Get PONumber() As String
    PONumber = fPONumber
End Property



