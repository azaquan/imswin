VERSION 5.00
Begin VB.Form FrmShowApproving 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1035
   ClientLeft      =   15
   ClientTop       =   2010
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Approving PO ......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2595
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Please Wait "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2595
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmShowApproving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Left = Int((MDI_IMS.Width - Me.Width) / 2)
    Me.Top = Int((MDI_IMS.Height - Me.Height) / 2) - 500
End Sub

