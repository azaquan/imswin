VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmTempPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Temporary Password"
   ClientHeight    =   2310
   ClientLeft      =   5100
   ClientTop       =   4485
   ClientWidth     =   5235
   Icon            =   "TempPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   Tag             =   "04010300"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3900
      TabIndex        =   10
      Top             =   1680
      Width           =   1125
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboUserName 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   900
      Width           =   2655
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GroupHeaders    =   0   'False
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "TempPassword.frx":000C
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "TempPassword.frx":0028
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   5477
      Columns(0).Caption=   "User Name"
      Columns(0).Name =   "UserName"
      Columns(0).DataField=   "Column 0"
      Columns(0).FieldLen=   256
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   5292
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "UserId"
      Columns(1).Name =   "UserId"
      Columns(1).DataField=   "Column 1"
      Columns(1).FieldLen=   256
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   4445
      Columns(2).Caption=   "Minimum Password Length"
      Columns(2).Name =   "MiniPwdLength"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).HeadStyleSet=   "ColHeader"
      Columns(2).StyleSet=   "RowFont"
      _ExtentX        =   4683
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   16711680
      BackColor       =   16777152
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtAdminPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   540
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   1125
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1125
   End
   Begin VB.TextBox txtTempPassword 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1260
      Width           =   2655
   End
   Begin VB.Label lblAdminUserName 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   180
      Width           =   2655
   End
   Begin VB.Label lblAdminName 
      Caption         =   "&Admin's User Name"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
   Begin VB.Label Label2 
      Caption         =   "Admin's &Password"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2100
   End
   Begin VB.Label Label4 
      Caption         =   "&User Name"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "&New Temp Password"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1260
      Width           =   2100
   End
End
Attribute VB_Name = "frmTempPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public User As String
Dim UserLevel As Integer
Public NameSpace As String
Public cn As ADODB.Connection
Public FirstTimeUser As Boolean

'close form

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'check enter password and show message


Private Sub cmdOk_Click()
Dim pn As Long
Dim ObjXevents As ImsXevents
    
    If ssdcboUserName.Enabled Then
    
        pn = ssdcboUserName.Columns("MiniPwdLength").Text
        
        If Len(Trim$(txtTempPassword)) < pn Then
            Call MsgBox(ReplaceNewLine(ReplaceDecimal(LoadResString(122), CDbl(pn))))
        Else
           'cOMMENTED OUT BY MUZAMMIL 04/08/01
           'reASON - THIS CODE WAS CHECKING THE XPASSWORDX OF THE USER LOGGED IN INSTEAD OF CHECKING IT
           'FOR THE USER WHOSE PASSWORD IS BEING CHECKED.fRANCOIS SAID TO COMMENT IT OUT.
           '
            'If PassWordExist(NameSpace, User, Encrypt(txtTempPassword), cn) Then
            '    Call MsgBox(ReplaceNewLine(LoadResString(120)))
            
            'Else
                Call UpdatePassword:
                
                If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
              'This is the case of an personal password change
                    ObjXevents.AddNew
                    ObjXevents.MyLoginId = lblAdminUserName
                    ObjXevents.HisLoginId = Trim$(ssdcboUserName.Columns("UserId").Text)
                    ObjXevents.NameSpace = NameSpace
                    ObjXevents.EventDetail = " User " & Trim$(ssdcboUserName) & " has a temporary " & IIf(UserLevel = 2, "Owner", "Admin") & " password assigned by  " & lblAdminUserName.Caption & "."
                    ObjXevents.STAs = "A"
                    ObjXevents.update
                    Set ObjXevents = Nothing
        
                   
                Unload Me
            'End If
            
        End If
        
    ElseIf IsCurrentPassword(NameSpace, User, Encrypt(txtAdminPassword), cn) Then
    
        OpenRecordset
          If Len(ssdcboUserName) > 0 Then
            If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
              'This is the case of an personal password change
                    ObjXevents.AddNew
                    ObjXevents.MyLoginId = lblAdminUserName
                    ObjXevents.HisLoginId = ssdcboUserName
                    ObjXevents.NameSpace = NameSpace
                    ObjXevents.EventDetail = "User " & Trim$(ssdcboUserName) & " has a temporary password assigned by " & lblAdminUserName.Caption & "."
                    ObjXevents.STAs = "A"
                    ObjXevents.update
                    Set ObjXevents = Nothing
         End If
        txtAdminPassword.Enabled = False
    Else
    
        'Modified by Juan Gonzalez (8/29/2000) for Translation fix
        msg1 = Trans("M00099")
        msg2 = Trans("M00139")
        MsgBox IIf(msg1 = "", "Invalid Password", msg1) & vbCrLf & IIf(msg1 = "", "Please note the password is case Sensitive", msg1)
        '---------------------------------------------------------
    End If
End Sub

'load form and call function to check user stauts

Private Sub Form_Load()
Dim rst As ADODB.Recordset
Dim datax As New ADODB.Recordset
Dim sql

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmTempPass")
    '--------------------------------------------------

    lblAdminUserName = User
    ssdcboUserName.FieldSeparator = Chr(1)
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    If CheckUserStatus(NameSpace, User, cn, rst) = 1 Then Exit Sub
    
    sql = "SELECT * FROM XUSERPROFILE WHERE usr_userid = '" + User + "' AND usr_npecode = '" + NameSpace + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        UserLevel = datax!usr_leve
        If UserLevel > 2 Then Unload Me
    Else
        Unload Me
    End If
'    Caption = Caption + " - " + Tag
    

End Sub

'set database conncetion and get user information

Private Sub OpenRecordset()
Dim cmd As ADODB.Command

    
    If cn Is Nothing Then _
        Err.Raise 10001, "imssec", "invalid connection"
    
    If cn.State And adStateOpen <> adStateOpen Then _
        Err.Raise 10001, "imssec", "invalid connection"
        
    
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        
        
        .CommandText = "SELECT usr_npecode, usr_userid, usr_username," & vbCrLf
        .CommandText = .CommandText & " usr_minipswdleng , usr_tempswdadmn, usr_pswd" & vbCrLf
        .CommandText = .CommandText & " FROM XUSERPROFILE" & vbCrLf
        If UserLevel = 0 Then
            .CommandText = .CommandText & " WHERE usr_npecode = '" & NameSpace & "'"
        Else
            .CommandText = .CommandText & " WHERE usr_leve > 2 AND usr_npecode = '" & NameSpace & "'"
        End If
        If FirstTimeUser Then
            .CommandText = .CommandText & " AND usr_stas = 'A'"
            .CommandText = .CommandText & " AND usr_pswd is null "
            
            If UserLevel < 1 Then
                .CommandText = .CommandText & " AND usr_tempswdadmn is null "
            ElseIf UserLevel = 2 Then
                .CommandText = .CommandText & " AND usr_temppswdownr is null "
            End If
        End If
        
        .CommandText = .CommandText & " ORDER BY usr_username, usr_userid"
        
        
        EnableControls
        Call AddUserInfo(.Execute)
    End With
End Sub

'set return key equal to command click ok

Private Sub txtAdminPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOk_Click
End Sub

'populate user info data grid

Public Sub AddUserInfo(rs As ADODB.Recordset)
Dim sep As String
Dim str As String

    sep = Chr(1)
    ssdcboUserName.RemoveAll
    
    If rs Is Nothing Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    
    Do While Not rs.EOF
        str = rs!usr_username & "" & sep & rs!usr_userid & ""
        str = str & sep & rs!usr_minipswdleng & ""
        
        ssdcboUserName.AddItem str
        rs.MoveNext
    Loop
     
    Set rs = Nothing
     
End Sub

'enable controls and set back ground color

Private Sub EnableControls()
    
    With ssdcboUserName
        .Enabled = True
        .ForeColor = vbWindowText
        .BackColor = vbWindowBackground
    End With
    
    With txtTempPassword
        .Enabled = True
        .ForeColor = vbWindowText
        .BackColor = vbWindowBackground
    End With
End Sub

'SQL statement to update xuserprofile table

Private Sub UpdateTempAdminPwd()
On Error GoTo Cancelled

Dim pwd As String
Dim UserId As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        
        .CommandType = adCmdText
        
        .CommandText = "UPDATE XUSERPROFILE SET " & vbCrLf & vbTab
        .CommandText = .CommandText & " usr_pswd = Null ," & vbCrLf & vbTab
        
        If UserLevel = 2 Then
            .CommandText = .CommandText & " usr_temppswdownr = ?" & vbCrLf
        
        ElseIf UserLevel <= 1 Then
            .CommandText = .CommandText & " usr_tempswdadmn = ?" & vbCrLf
        End If
        
        If Not FirstTimeUser Then _
            .CommandText = .CommandText & ", usr_stas = 'S'" & vbCrLf
        
        .CommandText = .CommandText & " Where usr_userid = ? AND" & vbCrLf & vbTab
        .CommandText = .CommandText & " usr_npecode = ?"
        
        pwd = Encrypt(txtTempPassword.Text)
        UserId = ssdcboUserName.Columns("UserId").Text
        .Execute , Array(pwd, UserId, NameSpace), adExecuteNoRecords
        Call UPDATEXPASSWORD(NameSpace, UserId, pwd, "TEMPORARY ADMIN", "USER", cn)
    End With
    
    Set cmd = Nothing
    
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetComplete
    Exit Sub
    
Cancelled:
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetAbort
End Sub

'set store procedure parameters and call it

Private Sub UpdatePassword()
On Error Resume Next
Dim UserId As String
'Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
    'Set cmd = New ADODB.Command
    
    Set rs = New ADODB.Recordset
    UserId = ssdcboUserName.Columns("UserId").Text
    rs.Source = "select * from xuserprofile where usr_npecode='" & NameSpace & "' and usr_userid='" & UserId & "'"
    rs.ActiveConnection = cn
    rs.Open , , , adLockOptimistic
  
    If UserLevel = 1 Then
  
  
                  '  With cmd
                  '      Set .ActiveConnection = cn
                  '      .CommandType = adCmdStoredProc
                  '      .CommandText = "UPDATEADMPASSWORD"
                  
                '  If .Parameters.Count = 0 Then
                '      Call .Parameters.Append(.CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5))
                '      Call .Parameters.Append(.CreateParameter("@USERID", adVarChar, adParamInput, 15))
                      
                '      Call .Parameters.Append(.CreateParameter("@USERLEVEL", adInteger, adParamInput, 4))
                '      Call .Parameters.Append(.CreateParameter("@PASSWORD", adVarChar, adParamInput, 15))
                '  End If
        
 
                'Added by Muzammil
                'reason - the sotred procedure below semms no good.
                '---------------------------------------------------------
                
                
                
                If Trim$(rs!usr_stas) = "A" Then
                
                   rs!USR_pswd = Encrypt(txtTempPassword)
                   rs!USR_temppswdownr = Encrypt(txtTempPassword)
                 '  rs!usr_tempswdadmn = Null
                 '  rs!usr_temppswdownr = Null
                   
                ElseIf Trim$(rs!usr_stas) = "S" Then
                
                  'rs!usr_pswd = Null
                   rs!usr_tempswdadmn = Encrypt(txtTempPassword)
                
                End If
                
                rs.update
                
    ElseIf UserLevel = 2 Or UserLevel = 0 Then
            If UserLevel = 0 Then
                rs!USR_pswd = Encrypt(txtTempPassword)
                rs!USR_temppswdownr = Encrypt(txtTempPassword)
            Else
                rs!USR_temppswdownr = Encrypt(txtTempPassword)
            End If

               rs.update
        
    End If
    
    rs.Close
    Set rs = Nothing
    
     'Commented out by muzammil 04/06/01
     'REason add the code above
        
      '  .Parameters("@USERID") = UserId
      '  .Parameters("@NAMESPACE") = NameSpace
      '  .Parameters("@USERLEVEL") = UserLevel
      '  .Parameters("@PASSWORD") = Encrypt(txtTempPassword)
        
      '  Call .Execute(Options:=adExecuteNoRecords)
        
         'Call UPDATEXPASSWORD(NameSpace, UserId, .Parameters("@PASSWORD"), "TEMPORARY ADMIN", "USER", cn)
        'If UserLevel <> 0 Then
        Call UPDATEXPASSWORD(NameSpace, UserId, Encrypt(txtTempPassword), "TEMPORARY ADMIN", "USER", cn)
         
         
   'End With
        
        
End Sub


