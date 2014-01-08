VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsPackingListRecp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Store packing list receipients"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"imsPackingListRec"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Event UpdateError(sError As String, bContinue As Boolean)

Public Function Add(NameSpace As String, Manfnumb As String, Receipient As String, User As String, Optional sKey As String) As imsPackingListRec
    'create a new object
    Dim objNewMember As imsPackingListRec
    Set objNewMember = New imsPackingListRec


    'set the properties passed into the method
    objNewMember.NameSpace = NameSpace
    objNewMember.ManifestNumber = Manfnumb
    objNewMember.Receipient = Receipient
    objNewMember.RecpNumb = mCol.Count + 1
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

Public Property Get Item(vntIndexKey As Variant) As imsPackingListRec
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub AddimsPackingRecp(Rep As imsPackingListRec, Optional Key As String)
    If Len(Key) Then
        mCol.Add Rep, Key
    Else
        mCol.Add Rep
    End If
End Sub

'call function to insert data to database

Public Sub UpdateAll(cn As ADODB.Connection)
On Error Resume Next
Dim Rep As Object
Dim bContinue As Boolean

    For Each Rep In mCol
        Call Rep.Update(cn)
        
        If Err Then
            bContinue = True
            RaiseEvent UpdateError(Err.Description, bContinue)
            
            Err.Clear
            If Not bContinue Then Exit Sub
        End If
        
    Next
    
End Sub

