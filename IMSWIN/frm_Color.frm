VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Color 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Scheme"
   ClientHeight    =   3090
   ClientLeft      =   30
   ClientTop       =   2400
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Tag             =   "05010000"
   Begin VB.CommandButton cmd_Ok 
      Caption         =   "Ok"
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Help 
      Caption         =   "Help"
      Height          =   300
      Left            =   4110
      TabIndex        =   11
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Apply 
      Caption         =   "Apply"
      Height          =   300
      Left            =   2760
      TabIndex        =   10
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Change1 
      Caption         =   "Change"
      Height          =   396
      Index           =   3
      Left            =   2688
      TabIndex        =   7
      Top             =   2064
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Change1 
      Caption         =   "Change"
      Height          =   396
      Index           =   2
      Left            =   2688
      TabIndex        =   5
      Top             =   1488
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Change1 
      Caption         =   "Change"
      Height          =   396
      Index           =   1
      Left            =   2688
      TabIndex        =   3
      Top             =   912
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Change1 
      Caption         =   "Change"
      Height          =   396
      Index           =   0
      Left            =   2688
      TabIndex        =   1
      Top             =   384
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cdl_Color 
      Left            =   1980
      Top             =   780
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   4260
      X2              =   5235
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   5220
      X2              =   5220
      Y1              =   1980
      Y2              =   2370
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   4260
      X2              =   5235
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4260
      X2              =   4260
      Y1              =   1980
      Y2              =   2370
   End
   Begin VB.Label txt_Htextbox 
      BackColor       =   &H8000000D&
      Height          =   390
      Left            =   4260
      TabIndex        =   15
      Top             =   1980
      Width           =   975
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   4260
      X2              =   5235
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   5220
      X2              =   5220
      Y1              =   1440
      Y2              =   1830
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   4260
      X2              =   5235
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4260
      X2              =   4260
      Y1              =   1440
      Y2              =   1830
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   5220
      X2              =   5220
      Y1              =   900
      Y2              =   1290
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   4260
      X2              =   5235
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   4260
      X2              =   5235
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4260
      X2              =   4260
      Y1              =   900
      Y2              =   1290
   End
   Begin VB.Label txt_WBackground 
      Height          =   390
      Left            =   4260
      TabIndex        =   13
      Top             =   900
      Width           =   975
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   4224
      X2              =   5199
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5186
      X2              =   5186
      Y1              =   384
      Y2              =   774
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   4224
      X2              =   5199
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4224
      X2              =   4224
      Y1              =   384
      Y2              =   774
   End
   Begin VB.Label txt_Background 
      BackColor       =   &H8000000C&
      Height          =   396
      Left            =   4224
      TabIndex        =   12
      Top             =   384
      Width           =   972
   End
   Begin VB.Label lbl_Htextbox 
      BackStyle       =   0  'Transparent
      Caption         =   "Highlighted Text Boxes"
      Height          =   390
      Left            =   240
      TabIndex        =   6
      Top             =   2070
      Width           =   2460
   End
   Begin VB.Label lbl_textbox 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Boxes"
      Height          =   390
      Left            =   240
      TabIndex        =   4
      Top             =   1485
      Width           =   2460
   End
   Begin VB.Label lbl_WBackground 
      BackStyle       =   0  'Transparent
      Caption         =   "Window Background"
      Height          =   390
      Left            =   240
      TabIndex        =   2
      Top             =   915
      Width           =   2460
   End
   Begin VB.Label lbl_Background 
      BackStyle       =   0  'Transparent
      Caption         =   "Workspace Background"
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   2460
   End
   Begin VB.Label txt_textbox 
      BackColor       =   &H80000005&
      Height          =   390
      Left            =   4260
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frm_Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Apply_Click()
On Error Resume Next
    
    Call SaveColor
    ChangeFormsColor
    
End Sub

Private Sub cmd_cancel_Click()
    Hide
End Sub

Private Sub cmd_Change1_Click(Index As Integer)
On Error GoTo ErrHandler
    
    cdl_Color.Flags = cdlCCRGBInit
    
    Select Case Index
        Case 0
            cdl_Color.Color = txt_Background.BackColor
        Case 1
            cdl_Color.Color = txt_WBackground.BackColor
        Case 2
            cdl_Color.Color = txt_textbox.BackColor
        Case 3
            cdl_Color.Color = txt_Htextbox.BackColor
    End Select
    
    cdl_Color.ShowColor
    
    Select Case Index
        Case 0
            txt_Background.BackColor = cdl_Color.Color
        Case 1
            txt_WBackground.BackColor = cdl_Color.Color
        Case 2
            txt_textbox.BackColor = cdl_Color.Color
        Case 3
            txt_Htextbox.BackColor = cdl_Color.Color
    End Select
    
    Exit Sub
    
ErrHandler:

    If Err = cdlCancel Then
        Err.Clear
    Else: MsgBox Err.Description
    End If
    Exit Sub
    
End Sub

Private Sub cmd_Help_Click()
    Call Winhelp(HWND, App.HelpFile, HELP_CONTEXT, 0)
End Sub

Private Sub cmd_ok_Click()
On Error Resume Next
    Screen.MousePointer = vbHourglass
    Call cmd_Apply_Click: Hide
    Screen.MousePointer = vbArrow
End Sub

Private Sub Form_Load()
On Error Resume Next

    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Color")
    '------------------------------------------

    Hide
    ZOrder
    Call GetColors
    Caption = Caption + " - " + Tag
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

    Call SaveColor
End Sub

Public Sub SaveColor()
On Error Resume Next
Dim colors(3) As String
Dim str As String

    colors(2) = CStr(txt_textbox.BackColor)
    colors(3) = CStr(txt_Htextbox.BackColor)
    colors(0) = CStr(txt_Background.BackColor)
    colors(1) = CStr(txt_WBackground.BackColor)
    
    str = CurrentUser
    If str = "" Then str = "Default"
    Call SaveSetting(App.Title, str, "Back Color", colors(0))
    Call SaveSetting(App.Title, str, "Window Color", colors(1))
    Call SaveSetting(App.Title, str, "TextBox Color", colors(2))
    Call SaveSetting(App.Title, str, "HighLite Color", colors(3))
End Sub

Public Sub GetColors()
On Error Resume Next
Dim colors(3) As String
Dim str As String
    
    str = CurrentUser
    If str = "" Then str = "Default"
    txt_Htextbox.BackColor = GetSetting(App.Title, str, "HighLite Color", vbHighlight)
    txt_WBackground.BackColor = GetSetting(App.Title, str, "Window Color", vbButtonFace)
    txt_textbox.BackColor = GetSetting(App.Title, str, "TextBox Color", vbWindowBackground)
    txt_Background.BackColor = GetSetting(App.Title, str, "Back Color", vbApplicationWorkspace)
End Sub

Public Sub ChangeFormsColor()
On Error Resume Next
Dim li_curr_frm As Form
Dim li_curr_ctl As Control

    For Each li_curr_frm In Forms
    
            '(li_curr_frm Is frm_Login) Or
            
        'Modified by Juan Gonzalez (8/29/2000) for Translation fixes
        'If ((li_curr_frm Is frm_bkgnd)) Then
            'li_curr_frm.BackColor = Me.txt_Background.BackColor

        'Else:
            'li_curr_frm.BackColor = Me.txt_WBackground.BackColor

        'End If
        '------------------------------------------------------------

        For Each li_curr_ctl In li_curr_frm.Controls

            If (li_curr_frm Is Me) Then
                If Not (TypeOf li_curr_ctl Is textBOX) Then _
                    Call gsb_fade_to_black(li_curr_ctl)

            Else:
                Call gsb_fade_to_black(li_curr_ctl)

        End If

        Next li_curr_ctl

    Next li_curr_frm

    'change the current workspace background
    MDI_IMS.BackColor = txt_Background.BackColor

    GetColors
End Sub

Public Function WorkSpaceColor() As OLE_COLOR
    WorkSpaceColor = txt_Background.BackColor
End Function

Public Function WindowColor() As OLE_COLOR
    WindowColor = txt_WBackground.BackColor
End Function

Public Function TextColor() As OLE_COLOR
    TextColor = txt_textbox.BackColor
End Function

Public Function HiliteColor() As OLE_COLOR
    HiliteColor = txt_Htextbox.BackColor
End Function
'

Private Sub Form_Unload(Cancel As Integer)
    Set frm_Color = Nothing
End Sub

