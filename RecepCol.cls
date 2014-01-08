VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecepCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"imsReception"
Attribute VB_Ext_KEY = "Member0" ,"imsReception"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Event UpdateError(sError As String, bContinue As Boolean)

Public Function Add(Npecode As String, Recpnumb As String, Recponumb As String, Linenumb As Integer, RecepCol As RecepCol, Optional sKey As String) As imsReception
    'create a new object
    Dim objNewMember As imsReception
    Set objNewMember = New imsReception


    'set the properties passed into the method
    objNewMember.Npecode = Npecode
    objNewMember.Recpnumb = Recpnumb
 
    objNewMember.Recponumb = Recponumb
    objNewMember.Linenumb = Linenumb
    
'    Set objNewMember.RecepCol = RecepCol
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As imsReception
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

Public Sub AddimsReception(Recep As imsReception, Optional Key As String)
    If Len(Key) Then
        mCol.Add Recep, Key
    Else
        mCol.Add Recep
    End If
End Sub

Public Sub UpdateAll()
On Error Resume Next
Dim Recep As Object
Dim bContinue As Boolean

    For Each Recep In mCol
        Call Recep.Update
        
        If Err Then
            bContinue = True
            RaiseEvent UpdateError(Err.Description, bContinue)
            
            Err.Clear
            If Not bContinue Then Exit Sub
        End If
        
    Next
    
End Sub



