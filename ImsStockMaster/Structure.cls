VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteALLPOITEM(cancel As Boolean)
Public Event AfterDeletePOITEMS()


Public Property Get NameSpace() As String
End Property
Public Property Let NameSpace(ByVal indata As String)
End Property

Public Function AddNew() As Boolean
End Function

Public Function DeleteCurrentLI() As Boolean
   
End Function

Public Function DeleteAllPOITEM(PO_ponumb As String) As Boolean
End Function

Public Function Update() As Boolean
End Function


Private Function LoadCollection() As Boolean
End Function


Public Function Move(POnumb As String) As Boolean
End Function

Public Property Get Count() As Integer
End Property
Public Function MoveFirst() As Boolean
End Function



Public Function MoveNext() As Boolean
End Function

Public Function MovePrevious() As Boolean
End Function

Public Function MoveLast() As Boolean

End Function

Private Sub SetTheNewHeader()

End Sub


Public Property Get Eof() As Boolean
End Property

Public Property Get Bof() As Boolean
End Property


Public Function CancelUpdate() As Boolean
End Function

Public Function CancelBatch()
End Function


Public Property Get EditMode() As String
End Property


Public Function Requery()
End Function


Public Function AnyChangesBeforeSaving() As Boolean
End Function






