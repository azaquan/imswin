VERSION 5.00
Begin VB.Form frmpass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2115
   ClientLeft      =   6390
   ClientTop       =   4695
   ClientWidth     =   4470
   HelpContextID   =   1000
   Icon            =   "Frmpass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   Tag             =   "04010400"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1000
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1000
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   1125
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      HelpContextID   =   1000
      Left            =   2760
      TabIndex        =   10
      Top             =   1560
      Width           =   1125
   End
   Begin VB.TextBox txtPwd 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      HelpContextID   =   1000
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   780
      WhatsThisHelpID =   1000
      Width           =   2235
   End
   Begin VB.TextBox txtPwd 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      HelpContextID   =   1000
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1110
      Width           =   2235
   End
   Begin VB.TextBox txtPwd 
      Height          =   315
      HelpContextID   =   1000
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   450
      Width           =   2235
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      HelpContextID   =   1000
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&New Password"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirm Password"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1110
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   450
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjXevents As ImsXevents
Dim np As String
Dim User As String
Dim cn As ADODB.Connection

'unload form

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'call function to check user name and password
'mimximun password lenght

Private Sub cmdOk_Click()

Dim lng As Double
Dim str As String
Dim cmd As ADODB.Command
 
    Set cmd = New ADODB.Command
    
    If txtPWD(1).Enabled = False Then
        
       If VerifyUserNameAndPassword = False Then MsgBox "Wrong Password Entered.", , "Imswin"
        
        Exit Sub
        
    End If
    
    str = txtPWD(1)
    If Len(txtPWD(2)) = 0 Then Exit Sub
    
    If StrComp(str, txtPWD(2)) <> 0 Then
    
        'Modified by Juan Gonzalez (8/29/200) for Translation fix
        msg1 = Trans("M00140")
        MsgBox IIf(msg1 = "", "Password does not confirm", msg1)
        '--------------------------------------------------------
      txtPWD(1) = ""
      txtPWD(2) = ""
      txtPWD(1).SetFocus
    
    Exit Sub

    End If
        
    
    lng = GetMinPasswordLength(np, User, cn)
    
    If Len(str) < lng Then
        MsgBox ReplaceNewLine(ReplaceDecimal(LoadResString(122), lng)): Exit Sub
        
    ElseIf PassWordExist(np, User, Encrypt(str), cn) Then
    
        'Modified by Juan Gonzalez (8/29/2000) for Translation fix
        msg1 = Trans("M00170")
        MsgBox IIf(msg1 = "", "You cannot reuse this password", msg1): Exit Sub
        '---------------------------------------------------------
        
    Else
        
        Hide
        Call UpdateUserPassword(np, User, Encrypt(str), cn)
        
        
        If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
              'This is the case of an personal password change
                    ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserName
                    ObjXevents.HisLoginId = txtUserName
                    ObjXevents.NameSpace = np
                    ObjXevents.EventDetail = "User " & Trim$(txtUserName) & " changed personal password" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
        
        
'        With cmd
'            .CommandType = adCmdText
'            Set .ActiveConnection = cn
'
'            .CommandText = "UPDATE XUSERPROFILE SET usr_pswd = '" & str & "'"
'            .CommandText = .CommandText & " WHERE usr_userid = '" & user & "'"
'            .CommandText = .CommandText & " AND usr_npecode = '" & np & "'"
'
'            .Execute
'        End With

    End If
        
    Set cmd = Nothing
End Sub

'set new name space value

Public Property Let NameSpace(ByVal vNewValue As String)
    np = vNewValue
End Property

'set new user value

Public Property Let UserName(ByVal vNewValue As String)
    User = vNewValue
End Property

'set new conncetion value

Public Property Set Connection(vNewValue As ADODB.Connection)
    Set cn = vNewValue
End Property

'load form and get caption text

Private Sub Form_Load()

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmMenuTemp")
    '--------------------------------------------------

    Caption = Caption + " - " + Tag
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

'call function to verity user name and password

Private Sub txtPWD_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
    
        Select Case Index
        
            Case 0
                VerifyUserNameAndPassword
                If txtPWD(1).Enabled Then txtPWD(1).SetFocus
            
            Case 1
            
                If txtPWD(2).Enabled Then txtPWD(2).SetFocus
                
            Case 3
                
                cmdOk_Click
                
        End Select

    End If
    
End Sub

'function verify user password and name

Public Function VerifyUserNameAndPassword() As Boolean
    
    If IsCurrentPassword(np, txtUserName, Encrypt(txtPWD(0)), cn) Then
    
        txtPWD(1).Enabled = True
        txtPWD(1).ForeColor = vbWindowText
        txtPWD(1).BackColor = vbWindowBackground
        
        txtPWD(2).Enabled = True
        txtPWD(2).ForeColor = vbWindowText
        txtPWD(2).BackColor = vbWindowBackground
        
        txtPWD(1).SetFocus
        VerifyUserNameAndPassword = True
        
    Else
    
        VerifyUserNameAndPassword = False
    
    End If
    
End Function

