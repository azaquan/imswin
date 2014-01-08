VERSION 5.00
Begin VB.Form frmPrintDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Options"
   ClientHeight    =   1155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPrint 
      Caption         =   "Print"
      Height          =   795
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1935
      Begin VB.OptionButton optprintCurrent 
         Caption         =   "Print Current"
         CausesValidation=   0   'False
         Height          =   192
         Left            =   60
         TabIndex        =   4
         Top             =   480
         Width           =   1755
      End
      Begin VB.OptionButton optPrintAll 
         Caption         =   "Print All"
         CausesValidation=   0   'False
         Height          =   192
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'set print option menu

Public Enum PrintOpts
    poNone = 0
    poPrintAll = 1
    poPrintCurrent = 2
    poPrintSelection
End Enum

Dim FOpts As PrintOpts

'cancel print option buttom

Private Sub CancelButton_Click()
    Hide
    FOpts = poNone
End Sub

'set print values

Private Sub Form_Load()

    'Added by Juan (9/15/2000) for Multilingual
    Call Translate_Forms("frmPrintDialog")
    '------------------------------------------

    FOpts = poNone
    optprintCurrent.Value = True
'    frmPrintDialog.CancelButton = frmPrintDialog.CancelButton + " - " + frmPrintDialog.Tag
End Sub

'set print option values

Public Function Result() As PrintOpts
    Result = FOpts
End Function

Private Sub Form_Unload(Cancel As Integer)
'If open_forms <= 5 Then ShowNavigator
End Sub

'set print option values

Private Sub OKButton_Click()
    Hide
    
    If optPrintAll.Value Then
        FOpts = poPrintAll
        
    ElseIf optprintCurrent.Value Then
        FOpts = poPrintCurrent
    
    'ElseIf optprintSel.Value Then
        'FOpts = poPrintSelection
        
    Else:
        FOpts = poNone
    
    End If
    
End Sub

'call function to print all

Private Sub optPrintAll_Click()
    If optPrintAll.Value Then FOpts = poPrintAll
End Sub

'get print option current value

Private Sub optprintCurrent_Click()
    If optprintCurrent.Value Then FOpts = poPrintCurrent
End Sub

