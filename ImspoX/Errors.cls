VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Errors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mErrors As Collection
Private mErrorobj As ErrorDetl



Public Sub Add(ErrorDetl As ErrorDetl)
mErrors.Add ErrorDetl
End Sub

Public Sub Clear()
Dim i
For i = 0 To mErrors.Count
  mErrors.Remove (i)
Next i
End Sub

Public Function Count() As Boolean
Count = mErrors.Count
End Function

Public Function Item(Key As Integer) As ErrorDetl
 
   Item = mErrors.Item(Key)
   
End Function

Public Property Get NumEnum() As IUnknown
Attribute NumEnum.VB_Description = "It Enables "
Attribute NumEnum.VB_UserMemId = -4
Attribute NumEnum.VB_MemberFlags = "40"

   Set NumEnum = mErrors.[_numenum]

End Property


