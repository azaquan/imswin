VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsInvoiceDetls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"InvoiceDetl"
Attribute VB_Ext_KEY = "Member0" ,"InvoiceDetl"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A75021D"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable to hold collection
'##ModelId=38955A750277
Private mCol As Collection
Attribute mCol.VB_VarDescription = "local variable to hold collection"

'##ModelId=38955A75028B
Public Function Add(PONumber As String, NameSpace As String, InvoiceNumber As String, LineItem As String, Quantity As Long, TotalPrice As Double, UnitPrice As Double, Optional sKey As String) As imsInvoiceDetl
    'create a new object
    Dim objNewMember As imsInvoiceDetl
    Set objNewMember = New imsInvoiceDetl


    'set the properties passed into the method
    objNewMember.PONumber = PONumber
    objNewMember.NameSpace = NameSpace
    objNewMember.InvoiceNumber = InvoiceNumber
    objNewMember.LineItem = LineItem
    objNewMember.Quantity = Quantity
    objNewMember.TotalPrice = TotalPrice
    objNewMember.UnitPrice = UnitPrice
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    If Err Then Err.Clear
End Function

'##ModelId=38955A7502B4
Public Property Get Item(vntIndexKey As Variant) As imsInvoiceDetl
Attribute Item.VB_UserMemId = 0
On Error Resume Next
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
  
  If Err Then Err.Clear
End Property



'##ModelId=38955A7502C8
Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mCol.count
End Property


'##ModelId=38955A7502DC
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


'##ModelId=38955A7502E7
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


'##ModelId=38955A7502FA
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


'##ModelId=38955A750304
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

'##ModelId=38955A750318
Public Sub RemoveAll()
Dim i As Long, x As Long

    i = mCol.count
    
    For x = 1 To i
        mCol.Remove (1)
    Next
    
End Sub
