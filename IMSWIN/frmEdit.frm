VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Template"
   ClientHeight    =   2235
   ClientLeft      =   465
   ClientTop       =   2490
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   4020
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Ok 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
'   DlgResult = -1
'   Unload FrmEdit
End Sub

'load form get list

Private Sub Form_Load()
    Dim i As Integer
    Dim li_x As Integer

    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("FrmEdit")
    '------------------------------------------

   For i = 0 To Forms.Count - 1
      List1.AddItem Forms(i).Name
   Next

   List1.ListIndex = 0  ' select the first item

'    color the controls and form backcolor
'    Me.BackColor = frm_Color.txt_WBackground.BackColor
'    For li_x = 0 To (Controls.count - 1)
''        Debug.Print Controls(li_x).name
'        'If TOC then deny it a call to gsb_fade_to_black
'        If Not (TypeOf Controls(li_x) Is Toc) Then Call gsb_fade_to_black(Controls(li_x))
'    Next li_x

    FrmEdit.Cancel = FrmEdit.Cancel + " - " + FrmEdit.Tag
End Sub

'Private Sub List1_DblClick()
'   ok_click
'End Sub
'
'Private Sub ok_click()
'   DlgResult = List1.ListIndex
'
'   Unload FrmEdit
'
'End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub
