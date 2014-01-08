Attribute VB_Name = "mSubClass"
Option Explicit

'This is common module, so we have to keep track of each
'cSubClass instance to call correct Window Procedure
'Do not call this procedures outside cSubClass

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Const GWL_WNDPROC = (-4)

Private Type SCInfo
    ProcOld As Long
    cSC As cSubClass
End Type

Private arrSubClassInfo() As SCInfo
Private arrSubClassInfoCount As Long

Public Sub SubClass(cSC As cSubClass)
'Do not use outside off cSubClass
Dim i As Long

    For i = 0 To arrSubClassInfoCount - 1
        If arrSubClassInfo(i).cSC.hWnd = cSC.hWnd Then
            'Already subclassed
            Exit Sub
        End If
    Next
            
    arrSubClassInfoCount = arrSubClassInfoCount + 1
    ReDim Preserve arrSubClassInfo(arrSubClassInfoCount)
    
    With arrSubClassInfo(arrSubClassInfoCount - 1)
        Set .cSC = cSC
        .ProcOld = GetWindowLong(.cSC.hWnd, GWL_WNDPROC)
        SetWindowLong .cSC.hWnd, GWL_WNDPROC, AddressOf MyProc
    End With
    
End Sub
Private Function MyProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim i As Long

    For i = 0 To arrSubClassInfoCount - 1
        With arrSubClassInfo(i)
            If .cSC.hWnd = hWnd Then
                If .cSC.Message(Msg) Then
                    If .cSC.MessageProcessing = mpSendAndProcess Then
                        'Send original message to window before processing
                        MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                    End If
                
                    'Fire WndProc event of cSC (Custom processing)
                    MyProc = .cSC.RaiseWndProc(Msg, wParam, lParam)
                    
                    If .cSC.MessageProcessing = mpProcessAndSend Then
                        'Send original message to window after processing
                        MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                    End If
                
                Else
                    'Call original window procedure
                    MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                End If
                
                Exit For
            End If
        End With
    Next
    
End Function

