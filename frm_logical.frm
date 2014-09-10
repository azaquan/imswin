VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_logical 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Tag             =   "01030500"
   Begin VB.TextBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      FirstEnabled    =   0   'False
      FirstVisible    =   0   'False
      LastEnabled     =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintVisible    =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid logwar 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   285
      AllowBigSelection=   0   'False
      GridLinesFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Visualization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   6600
      TabIndex        =   2
      Top             =   3840
      Width           =   2460
   End
   Begin VB.Label lbl_Logicals 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logical Warehouse"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frm_logical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Screen.MousePointer = vbHourglass
msg1 = translator.Trans("M00126")
Modify = IIf(msg1 = "", "Modification", msg1)
msg1 = translator.Trans("M00092")
Visualize = IIf(msg1 = "", "Visualization", msg1)
msg1 = translator.Trans("M00125")
Create = IIf(msg1 = "", "Creation", msg1)
Screen.MousePointer = vbHourglass
Me.BackColor = frm_Color.txt_WBackground.BackColor
Visible = True
Screen.MousePointer = vbDefault
Caption = Caption + " - " + Tag
NVBAR_EDIT = NavBar1.EditEnabled
NVBAR_ADD = NavBar1.NewEnabled
NVBAR_SAVE = NavBar1.SaveEnabled
NavBar1.EditEnabled = True
NavBar1.EditVisible = True
NavBar1.CancelEnabled = False
NavBar1.SaveEnabled = False
NavBar1.CloseEnabled = True
NavBar1.Width = 5050
''''Call DisableButtons(Me, NavBar1)
SSDBLogical.AllowUpdate = False

With frm_Logicals
    .Left = Round((Screen.Width - .Width) / 2)
    .Top = Round((Screen.Height - .Height) / 2)
End With
End Sub
