VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackingListDetls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"imsPackinListDetl"
Attribute VB_Ext_KEY = "Member0" ,"imsPackinListDetl"
Option Explicit

'local variable to hold collection


Private mCol As Collection

Public Event SaveError(sError As String, bContinue As Boolean)

Public Function Add(MainfestNumber As String, NameSpace As String, ManiFestSerialNumb As Integer, PoNumber As String, LineNumber As String, BoxNumber As String, RequestedQty As Double, UnitPrice As Currency, TotalPrice As Double, Tobeship As Double, User As String, Optional sKey As String) As imsPackinListDetl
    'create a new object
    Dim objNewMember As imsPackinListDetl
    Set objNewMember = New imsPackinListDetl


    'set the properties passed into the method
    objNewMember.MainfestNumber = MainfestNumber
    objNewMember.NameSpace = NameSpace
    objNewMember.ManiFestSerialNumb = mCol.Count + 1
    objNewMember.PoNumber = PoNumber
'    If IsObject(PoNumber) Then
' '       Set objNewMember.PoNumber = PoNumber
'    Else
'        objNewMember.PoNumber = PoNumber
'    End If
    objNewMember.LineNumber = LineNumber
    objNewMember.BoxNumber = BoxNumber
    
'    objNewMember.RequestedQty = RequestedQty
'    objNewMember.Tobeship = Tobeship
    
    If IsObject(RequestedQty) Then
        objNewMember.RequestedQty = RequestedQty
    Else
        objNewMember.Tobeship = Tobeship
    End If
    objNewMember.UnitPrice = UnitPrice
    objNewMember.TotalPrice = TotalPrice
    objNewMember.Tobeship = Tobeship
    objNewMember.User = User
    
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As imsPackinListDetl
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Sub AddimsPackingListDetl(pl As imsPackinListDetl, Optional Key As String)
    If Len(Key) Then
        mCol.Add pl, Key
    Else
        mCol.Add pl
    End If
End Sub

'call function to insert record to database

Public Sub UpdateAll()
On Error Resume Next
Dim pl As Object
Dim bContinue As Boolean

    For Each pl In mCol
        Call pl.Update
        
        If Err Then
            bContinue = True
            RaiseEvent SaveError(Err.Description, bContinue)
            
            Err.Clear
            If Not bContinue Then Exit Sub
        End If
        
    Next
    
End Sub
