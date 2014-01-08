VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImsRecipients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private FUser As String
Private FNameSpace As String
Private FReceipient As String
Private FRecpNumb As Integer

Public Property Let User(ByVal vData As String)
    FUser = vData
End Property

Public Property Get User() As String
    User = FUser
End Property

Public Property Let RecpNumb(ByVal vData As Integer)
    FRecpNumb = vData
End Property

Public Property Get RecpNumb() As Integer
    RecpNumb = FRecpNumb
End Property

Public Property Let Receipient(ByVal vData As String)
    FReceipient = vData
End Property

Public Property Get Receipient() As String
    Receipient = FReceipient
End Property

Public Property Let NameSpace(ByVal vData As String)
    FNameSpace = vData
End Property

Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property
