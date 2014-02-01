VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTreeTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Properties
Public ShowIconsInScrollTips As Boolean

Private mTree As MSComctlLib.TreeView

Private WithEvents cSub As cSubClass
Attribute cSub.VB_VarHelpID = -1


'General API
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

'Messages & Flags
Private Const WM_VSCROLL = &H115
Private Const SB_ENDSCROLL = 8
Private Const SB_THUMBTRACK = 5

'Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Neccessary Treeview Definitions
Private Const TV_FIRST      As Long = &H1100
Private Const TVM_GETNEXTITEM  As Long = (TV_FIRST + 10)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVGN_FIRSTVISIBLE = &H5

Public Property Set Tree(TreeView1 As MSComctlLib.TreeView)
    Set mTree = TreeView1
    Set cSub = New cSubClass
    cSub.hWnd = mTree.hWnd
    cSub.AttachMessage WM_VSCROLL
End Property

Private Sub Class_Terminate()

    If Not mTree Is Nothing Then
        Set cSub = Nothing
        Set mTree = Nothing
    End If
    
End Sub

Private Function LowWord(ByVal Value As Long) As Integer
    CopyMemory LowWord, Value, 2
End Function

Private Function HiWord(ByVal Value As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(Value) + 2, 2
End Function
Private Function TreeView_GetNextItem(hWnd As Long, hItem As Long, Flag As Long) As Long
   TreeView_GetNextItem = SendMessageAny(hWnd, TVM_GETNEXTITEM, Flag, ByVal hItem)
End Function
Private Function GetFirstVisibleNode() As MSComctlLib.Node
Dim hItem As Long
Dim rc As RECT
Dim ret As Boolean
Dim nodX As Node
Dim fItemRect As Long
Dim Flag As Long
    'Get First Visible Item
    hItem = TreeView_GetNextItem(mTree.hWnd, 0, TVGN_FIRSTVISIBLE)
    If hItem = 0 Then
        Exit Function
    End If

    rc.Left = hItem
    fItemRect = 1
    
    ret = SendMessageAny(mTree.hWnd, TVM_GETITEMRECT, ByVal fItemRect, rc)
    If ret Then
        Set GetFirstVisibleNode = mTree.HitTest((rc.Left) * Screen.TwipsPerPixelX, (rc.Top) * Screen.TwipsPerPixelY)
       
    End If
End Function

Private Sub cSub_WndProc(Msg As Long, wParam As Long, lParam As Long, nResult As Long)
Static bVisible As Boolean
Dim nodX As Node
Dim TipText As String
    On Error Resume Next
                 
        If Msg = WM_VSCROLL Then
            Select Case LowWord(wParam)
                'Drag na scrollbar
                Case SB_THUMBTRACK, 1
                    Set nodX = GetFirstVisibleNode()
                    If Not nodX Is Nothing Then
                        TipText = nodX.key
                    End If
            End Select
        End If
        If nodeONtop <> nodX.Index Then
            rePositionThings (nodeONtop) 'Juan 2014-01-29, for re-position after scrolling
            nodeONtop = nodX.Index
            'Call workBOXESlist("FIX")
        End If
End Sub
