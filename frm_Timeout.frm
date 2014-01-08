VERSION 5.00
Begin VB.Form Timeout 
   Caption         =   "Logged Out"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Your session has expired.  Reopen the form and try again."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Timeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yes As Boolean, no As Boolean

Private Sub Form_Load()
    Timeout.Left = Int((MDI_IMS.Width - Timeout.Width) / 2)
    Timeout.Top = Int((MDI_IMS.Height - Timeout.Height) / 2) - 500
    
    
    On Error Resume Next
Timeout.Show vbModal

End Sub



Private Sub OK_Click()
MDI_IMS.IdleStateDisengaged (0)
Unload Me
End Sub

