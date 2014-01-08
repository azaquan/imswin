VERSION 5.00
Begin VB.Form LogOff 
   Caption         =   "Exit"
   ClientHeight    =   2025
   ClientLeft      =   5340
   ClientTop       =   4785
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3930
   Begin VB.CommandButton CmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton OptExit 
         Caption         =   "Exit?"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton OptLogin 
         Caption         =   "Login to a different namespace?"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Are you sure you want to:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      Picture         =   "LogOff.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "LogOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
If MsgBox("Are you sure you want to exit?", vbCritical + vbYesNo, "Ims") = vbYes Then

End
 
End If
 
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdNo_Click()
Unload Me
End Sub

Private Sub CmdYes_Click()

If OptLogin.Value = True Then
    
    'If open_forms > 5 Then MsgBox "Froms open": Exit Sub
    If VB.Forms.count > 4 Then
        MsgBox " Please close all the forms before exiting the namespace.", vbInformation, "Ims"
        Exit Sub
     End If
    
    
    MDI_IMS.tmrStateMonitor.Enabled = False
    MDI_IMS.tmrPeriod.Enabled = False
    Unload MDI_IMS
    Set deIms = Nothing
    Set deIms = New deIms
    Call Main


ElseIf OptExit.Value = True Then

    End

End If

End Sub

Private Sub Form_Load()
LogOff.Height = 2430
LogOff.Width = 4050
LogOff.Top = 3480
LogOff.Left = 5595
OptExit.Value = True

End Sub



Public Function StoreTheUserName()
Dim Sa As Scripting.FileSystemObject
Dim t As Scripting.TextStream
Dim CompletePath As String
Dim i As Integer
CompletePath = App.Path & "\ImsAutomaticLogin.Ims"
Set Sa = New Scripting.FileSystemObject

If Sa.FileExists(CompletePath) = False Then
          
    Sa.CreateTextFile CompletePath
    
End If

   Set t = Sa.OpenTextFile(CompletePath, ForWriting, True)
    
   t.WriteLine CurrentUser
    
      
End Function
