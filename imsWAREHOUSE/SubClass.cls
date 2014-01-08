VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colMessage As New Collection
Private mHwnd As Long
Private mProcess As MessageProcess

Public Event WndProc(Msg As Long, wParam As Long, lParam As Long, nResult As Long)

Public Enum MessageProcess
    mpSendAndProcess = 0
    mpProcessAndSend
    mpProcessAndEat
End Enum

Property Let MessageProcessing(nValue As MessageProcess)
    mProcess = nValue
End Property

Property Get MessageProcessing() As MessageProcess
    MessageProcessing = mProcess
End Property


Property Let hWnd(hWndValue As Long)
    If hWndValue <> 0 Then
        mHwnd = hWndValue
        mSubClass.SubClass Me
    End If
End Property

Property Get hWnd() As Long
    hWnd = mHwnd
End Property

Public Sub AttachMessage(Msg As Long)

Dim i As Long

    On Error GoTo ERRTAG
    If Message(Msg) Then
        Exit Sub
    Else
        colMessage.Add Msg, "key" & Msg
    End If
        
    Exit Sub
ERRTAG:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Resume ERREND
ERREND:

End Sub


Public Function Message(Msg As Long) As Boolean

Dim i As Integer
    
    For i = 1 To colMessage.Count
        If colMessage(i) = Msg Then
            Message = True
            Exit Function
        End If
    Next
    
End Function


Private Sub Class_Terminate()

    Do Until colMessage.Count = 0
        colMessage.Remove 1
    Loop
    Set colMessage = Nothing
    
End Sub

Friend Function RaiseWndProc(Msg As Long, wParam As Long, lParam As Long) As Long

Dim lRet As Long

    RaiseEvent WndProc(Msg, wParam, lParam, lRet)
    RaiseWndProc = lRet
    
End Function
