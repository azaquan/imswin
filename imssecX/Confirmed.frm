VERSION 5.00
Begin VB.Form frmPersPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personal Password"
   ClientHeight    =   1770
   ClientLeft      =   6465
   ClientTop       =   4695
   ClientWidth     =   4125
   Icon            =   "Confirmed.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   1200
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2700
      TabIndex        =   5
      Top             =   1200
      Width           =   1125
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2580
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2580
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Confirmed Personal Password"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "&Personal Password"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2235
   End
End
Attribute VB_Name = "frmPersPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim User As String
Dim cn As ADODB.Connection

'before change new password check password user length
'SQL statement to update password

Private Sub cmdOk_Click()
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    str = InputBox("Enter Current Password", "Change Password")
    
    If Len(Trim$(str)) = 0 Then Exit Sub
    
    
    If StrComp(txtPWD(0), txtPWD(1)) <> 0 Then
    
        'Modified by Juan Gonzalez (8/29/2000) for Translation fixes
        msg1 = Trans("M00140")
        MsgBox IIf(msg1 = "", "Password does not confirm", msg1)
        '-----------------------------------------------------------
        
        Exit Sub
    End If
        
    
    If Len(txtPWD(0)) < GetMinPasswordLength(np, User, cn) Then
        
        'Modified by Juan Gonzalez (8/29/2000) for Translation fixes
        msg1 = Trans("M00214")
        MsgBox IIf(msg1 = "", "Password length is too short", msg1): Exit Sub
        '-----------------------------------------------------------
        
    ElseIf PassWordExist(np, User, txtPWD(0), cn) Then
        
        'Modified by Juan Gonzalez (8/29/2000) for Translagion fixes
        msg1 = Trans("M00170")
        MsgBox IIf(msg1 = "", "You cannot reuse this password", msg1): Exit Sub
        '-----------------------------------------------------------
        
    Else
        
        With cmd
            .CommandType = adCmdText
            Set .ActiveConnection = cn
            
            .CommandText = "UPDATE XUSERPROFILE SET usr_pswd = '" & txtPWD(0) & "'"
            .CommandText = .CommandText & " WHERE usr_userid = '" & User & "'"
            .CommandText = .CommandText & " AND usr_npecode = '" & np & "'"
            
            .Execute
        End With
    End If
        
    Set cmd = Nothing
End Sub

'assign new name space

Public Property Let NameSpace(ByVal vNewValue As String)
    np = vNewValue
End Property

'assign new user value

Public Property Let UserName(ByVal vNewValue As String)
    User = vNewValue
End Property

'assign new conncetion value

Public Property Set Connection(vNewValue As ADODB.Connection)
    Set cn = vNewValue
End Property

'set form size

Private Sub Form_Load()

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmMenuTemp")
    '--------------------------------------------------

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Caption = Caption + " - " + Tag
    
End Sub

