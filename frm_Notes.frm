VERSION 5.00
Begin VB.Form frm_Notes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notes"
   ClientHeight    =   3060
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5880
   Begin VB.TextBox txt_Remarks 
      Height          =   1272
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1140
      Width           =   4272
   End
   Begin VB.TextBox txt_Description 
      Height          =   288
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   4
      Top             =   840
      Width           =   4272
   End
   Begin VB.ComboBox cbo_Notes 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   510
      Width           =   2112
   End
   Begin VB.Label lbl_Notes2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5670
   End
   Begin VB.Label lbl_Remarks 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   396
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1400
   End
   Begin VB.Label lbl_Description 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   228
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1400
   End
   Begin VB.Label lbl_Notes 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   1400
   End
End
Attribute VB_Name = "frm_Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnNewClick()
End Sub

Private Sub NavBar1_OnSaveClick()
End Sub

Private Sub Form_Load()
    Dim li_x As Integer
    
    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Trans("frm_Notes")
    '------------------------------------------
    
'    Me.BackColor = frm_Color.txt_WBackground.BackColor
    For li_x = 0 To (Controls.Count - 1)
        'If Not (TypeOf Controls(li_x) Is Toc) Then Call gsb_fade_to_black(Controls(li_x))
    Next li_x
    Caption = Caption + " - " + Tag
    
    With frm_Notes
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub txt_Remarks_GotFocus()
Call HighlightBackground(txt_Remarks)
End Sub

Private Sub txt_Remarks_LostFocus()
Call NormalBackground(txt_Remarks)
End Sub
