VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Logon Information"
   ClientHeight    =   3855
   ClientLeft      =   5850
   ClientTop       =   3600
   ClientWidth     =   4530
   HelpContextID   =   1000
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Tag             =   "03050300"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtUserID 
      Height          =   315
      HelpContextID   =   1001
      Left            =   2370
      MaxLength       =   15
      TabIndex        =   2
      Top             =   720
      WhatsThisHelpID =   1001
      Width           =   1905
   End
   Begin VB.ComboBox cmbName 
      Height          =   315
      HelpContextID   =   1000
      ItemData        =   "frmLogon.frx":0442
      Left            =   2370
      List            =   "frmLogon.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      WhatsThisHelpID =   1000
      Width           =   1905
   End
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      HelpContextID   =   1000
      ItemData        =   "frmLogon.frx":0446
      Left            =   2400
      List            =   "frmLogon.frx":0448
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   1000
      Width           =   1905
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      HelpContextID   =   1002
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2070
      Visible         =   0   'False
      WhatsThisHelpID =   1002
      Width           =   1905
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      HelpContextID   =   1003
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2460
      Visible         =   0   'False
      WhatsThisHelpID =   1003
      Width           =   1905
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   150
      Top             =   1560
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3180
      TabIndex        =   6
      Top             =   3270
      WhatsThisHelpID =   1000
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2010
      TabIndex        =   5
      Top             =   3270
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3270
      Width           =   1125
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Name Space:"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   13
      Top             =   1110
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Login ID:"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   780
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   5
      X1              =   840
      X2              =   4290
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   4
      X1              =   840
      X2              =   4290
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      Height          =   195
      Index           =   4
      Left            =   870
      TabIndex        =   11
      Top             =   240
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   3
      X1              =   840
      X2              =   4290
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   840
      X2              =   4290
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " Login Process "
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1890
      TabIndex        =   10
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   840
      X2              =   4290
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   840
      X2              =   4290
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Personal Password:"
      Height          =   195
      Index           =   2
      Left            =   870
      TabIndex        =   0
      Top             =   2130
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirm Password:"
      Height          =   195
      Index           =   3
      Left            =   870
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image imgLight 
      Height          =   480
      Index           =   3
      Left            =   180
      Picture         =   "frmLogon.frx":044A
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLight 
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "frmLogon.frx":088C
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLight 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "frmLogon.frx":0CCE
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLight 
      Height          =   480
      Index           =   0
      Left            =   150
      Picture         =   "frmLogon.frx":1110
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sec As imsSecMod
Dim rst As ADODB.Recordset
Dim WithEvents cn As ADODB.Connection
Attribute cn.VB_VarHelpID = -1

Dim stat As Integer
Dim np As String
Dim NameSpaces As Collection
Dim LoginType As LoginStatus
Dim FResult As LoginSuccess

Dim recordchange As String
Dim fUser As String
'Dim NameSpace As String
Dim UserLevel As Integer
Dim AddingNew As Boolean
'Dim cn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim ObjXevents As ImsXevents
Dim TRANSACTIONNUBMER As Integer
Dim hasrecordchange As Integer
'............

Public Event InfoMessage(msg As String)

'Added by Juan Gonzalez (8/29/200) for Translation fix
Dim languages As ADODB.Recordset


'set password entry status and show passwords text boxse

Private Sub SetPasswordEntry(Caller As String)
    
    Select Case Caller
    
        Case "L"
        
            lblTit(2) = ""
            lblTit(3) = ""
            cmbName.Enabled = True
            txtUserID.Enabled = True
            txtPwd(0).Visible = False
            txtPwd(1).Visible = False
            
            txtUserID.SetFocus
            
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("L00162")
            lblStatus = IIf(msg1 = "", "Login Process", msg1): Exit Sub
            '---------------------------------------------------------

        Case "F"
        
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("M00104")
            lblStatus = IIf(msg1 = "", "First Time Login", msg1)
            msg1 = Trans("M00108")
            lblTit(2) = IIf(msg1 = "", "Admin Password:", msg1)
            msg1 = Trans("M00110")
            lblTit(3) = IIf(msg1 = "", "Owner Password:", msg1)
            '---------------------------------------------------------
            
        Case "N"
        
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("M00113")
            lblStatus = IIf(msg1 = "", "Initial/New Personal Password", msg1)
            msg1 = Trans("L00160")
            lblTit(2) = IIf(msg1 = "", "Personal Password:", msg1)
            msg1 = Trans("L00161")
            lblTit(3) = IIf(msg1 = "", "Confirm Password:", msg1)
            '---------------------------------------------------------
        txtPwd(0) = ""
        txtPwd(1) = ""
        txtPwd(0).SetFocus
            
    
        Case "R"
            
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("M00114")
            lblStatus = IIf(msg1 = "", "Regular Login", msg1)
            msg1 = Trans("L00160")
            lblTit(2) = IIf(msg1 = "", "Personal Password:", msg1)
            lblTit(3) = ""
            '---------------------------------------------------------
            
            txtPwd(0).WhatsThisHelpID = 1003
        Case "V"
        
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("M00119")
            lblStatus = IIf(msg1 = "", "Check Admin Password", msg1)
            msg1 = Trans("M00108")
            lblTit(2) = IIf(msg1 = "", "Admin Password:", msg1)
            '---------------------------------------------------------
            
            lblTit(3) = ""
        
        Case Else
                    
            lblTit(2) = ""
            lblTit(3) = ""
            
            'Modified by Juan Gonzalez (8/29/200) for Transalation fix
            msg1 = Trans("L00162")
            lblStatus = IIf(msg1 = "", "Login Process", msg1): Exit Sub
            '---------------------------------------------------------
            
    End Select
    
    cmbName.Enabled = False
    txtUserID.Enabled = False
    
    txtPwd(0) = ""
    txtPwd(1) = ""
    lblTit(2).Visible = IIf(InStr("L", Caller) <> 0, False, True)
    lblTit(3).Visible = IIf(InStr("LRV", Caller) <> 0, False, True)
    txtPwd(0).Visible = IIf(InStr("L", Caller) <> 0, False, True)
    txtPwd(1).Visible = IIf(InStr("LRV", Caller) <> 0, False, True)
    
    If txtPwd(0).Visible Then txtPwd(0).SetFocus
    End Sub

Private Sub cmbLanguage_Click()
'Procedure to select language by Juan Gonzalez (8/29/200)
    languages.MoveFirst
    languages.Find "lan_desc = '" + cmbLanguage + "'"
    TR_LANGUAGE = languages!lan_code
    If languages.EOF Then
        msg1 = Trans("M00100")
        MsgBox IIf(msg1 = "", "Language unavailable", msg1)
    Else
        Translate_Forms ("frmLogon")
    End If
End Sub

'cancel recordset update
Private Sub cmdCancel_Click()
    Hide
    FResult = lgsCancelled
End Sub

'call function to get help files

Private Sub cmdHelp_Click()
    Call imsutilsx.Winhelp(Hwnd, App.HelpFile, imsutilsx.HELP_CONTENTS, 1001)
End Sub

'check login user login status and insert a record to database


Private Sub cmdOk_Click()
Dim str As String
Dim Status As LoginStatus
Dim ObjXevents As ImsXevents
Dim datax As New ADODB.Recordset
Dim sql
Dim JustTurnedVisible As Boolean
    
        'If cmbName.Visible = False Then  'M 08/21/02
                JustTurnedVisible = cmbName.Visible
            If cmbName.Visible = False Then Call GetNameSpacesForUser(txtUserID, cn, cmbName, NameSpaces) 'M 08/21/02
                lblTit(0).Visible = cmbName.Visible
              If JustTurnedVisible = False Or cmbName.Visible = False Then Exit Sub 'M 08/21/02
        
        'End If 'M 08/21/02


    If Not txtPwd(0).Visible Then
        
            Status = HandleUserName 'M 08/21/02
        
        If Status = lsUnknown Then
        
        Exit Sub
     'This is the Case when the User has tried Logging in with an Unknown UserId
        
        ElseIf Status = lsSuccess Then
           Call UpdateMaxAttempts(NameSpace, txtUserID, cn)
        
        End If
  
    Else
        
        If LoginType = lsSuccess Then
            
            If handleregularlogin = lgsCancelled Then
                
                Exit Sub
            
            End If
            
            'This is the case of a Successfull Login
        
        ElseIf LoginType = lsFirstLogon Then
            
            If HandleFirstLogon = lgsCancelled Then Exit Sub
            
            'This is the case of a First time Login,which is a login with Admin and Owner Password
            
            If Len(Trim$(txtPwd(1).Text)) > 0 And Len(Trim$(txtPwd(0).Text)) > 0 Then
                
                If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
                
                ObjXevents.AddNew
                ObjXevents.MyLoginId = txtUserID
                ObjXevents.HisLoginId = txtUserID
                ObjXevents.NameSpace = NameSpaces(cmbName)
                ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " connected first time"
                ObjXevents.STAs = "A"
                ObjXevents.update
                Set ObjXevents = Nothing
            
            End If
        
        Else
            'This is the case of Suspended or Inactive Login
            Call HandleSuspendedUser
        
        End If
        
        If FResult = lgsSuccess Then
            
            If rst!DAYSTOEXPIRE > 0 Then
                Hide
                str = NameSpaces(cmbName)
                Call UpdateLastLogon(str, txtUserID, cn)
                Call ImsDatax.InsertIntoXLogin(str, "LOGON", txtUserID, cn)
                
                'Added by Juan 11/15/01 to inactivate user level 0 after the first logon
                sql = "SELECT * FROM XUSERPROFILE WHERE usr_userid = '" + txtUserID + "' AND usr_npecode = '" + NameSpace + "'"
                Set datax = New ADODB.Recordset
                datax.Open sql, cn, adOpenForwardOnly
                If datax.RecordCount > 0 Then
                    If datax!usr_leve = 0 Then
                        cn.Execute "UPDATE XUSERPROFILE SET usr_stas = 'I' WHERE usr_userid = '" + txtUserID + "'"
                    End If
                End If
                '-----------------------------------------------------------------------
                
            Else
                
                MsgBox ReplaceNewLine(LoadResString(124))
            
                If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
                'This is the case of an Expired Login
                ObjXevents.AddNew
                ObjXevents.MyLoginId = txtUserID
                ObjXevents.HisLoginId = txtUserID
                ObjXevents.NameSpace = NameSpaces(cmbName)
                ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with an an expired login" & "."
                ObjXevents.STAs = "A"
                ObjXevents.update
                Set ObjXevents = Nothing
            End If
        End If
    End If
End Sub

Private Sub Combo1_Change()

End Sub

'load form and set connection to database and
'call function to check error status

Private Sub Form_Load()
On Error Resume Next

Dim lng As Long
Dim str As String
Dim rs As ADODB.Recordset
Dim ap As imsutilsx.imsApis
Dim iData As ImsDatax.imsspInt
Dim lineITEM
    
    Hide
    Set ap = New imsutilsx.imsApis
    
    'Set cn = New ADODB.Connection  'M
    Set NameSpaces = New Collection
    Set iData = New ImsDatax.imsspInt
    
    FResult = lgsCancelled
    cn.Mode = adModeShareDenyNone
    
    str = "Attempting to connect to server."
    'str = str & vbCrLf & "This may take a few minutes"
    
    
    Call sec.RaiseErrorEvent(str, 1)
    
    cn.CursorLocation = adUseClient
    
    Set cn = sec.Connection  'M
    
    If Err Then
        
        Call LogErr("FrmLogon::Form_Load", Err.Description, Err, True)
        
        If cn.State = adStateClosed Then
            str = "Error Connecting to SQL Server." & vbCrLf
            str = str & "Make sure SQL Server is Running." & vbCrLf
            str = str & "Quitting ..........."
            
            Call sec.RaiseErrorEvent(str, 1)
            Err.Clear: Unload Me: FResult = lgsCancelled: Exit Sub
        Else
            If Err Then MsgBox Err.Description: Err.Clear
            Call sec.RaiseErrorEvent(Err.Description, 1)
        End If
    End If
    
    
    On Error Resume Next
    
    'Modified by Juan Gonzalez (8/28/2000) for Translation fix
    Dim sql
    Set languages = New ADODB.Recordset
    sql = "SELECT * FROM TRLANGUAGE ORDER BY lan_desc"
    languages.Open sql, cn, adOpenStatic
    Do While Not languages.EOF
        If Err.Number > 0 Then Exit Sub
        cmbLanguage.AddItem languages!lan_desc
        If languages!lan_desc = "US English" Then lineITEM = cmbLanguage.ListCount - 1
        languages.MoveNext
    Loop
    cmbLanguage.ListIndex = lineITEM
    TR_LANGUAGE = "US"
    '------------------------------------------------------
    
    Set TR_MESSAGES = New ADODB.Recordset
    sql = "SELECT TRLANGUAGE.lan_code, TRLANGUAGE.lan_desc, TRMESSAGE.msg_numb , TRMESSAGE.msg_text " _
        & "FROM TRLANGUAGE LEFT OUTER JOIN TRMESSAGE ON TRLANGUAGE.lan_code = TRMESSAGE.msg_lang ORDER BY TRLANGUAGE.lan_desc"
    TR_MESSAGES.Open sql, cn, adOpenStatic
    
    Set TR_CONTROLS = New ADODB.Recordset
    sql = "SELECT TRMESSAGE.msg_lang, TRTRANSLATION.trs_enty, TRTRANSLATION.trs_obj, " _
        & "TRMESSAGE.msg_text FROM TRTRANSLATION LEFT OUTER JOIN TRMESSAGE ON TRTRANSLATION.trs_mesgnum = TRMESSAGE.msg_numb"
    TR_CONTROLS.Open sql, cn, adOpenStatic
    
    msg1 = Trans("")
    Call sec.RaiseErrorEvent("Preparing to login", 1)
    
    
    Call sec.RaiseErrorEvent("0", 1)
    Call ap.StayOnTop(Hwnd, True)
    
    Set ap = Nothing
    Set rs = Nothing
    Set iData = Nothing
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    On Error Resume Next
    
    If Err Then Err.Clear
    
End Sub

Private Sub Form_Resize()
    GetTheUserName
End Sub

'unload form free name space memory

Private Sub Form_Unload(Cancel As Integer)
    
    Set NameSpaces = Nothing

End Sub





''Public Function GenerateXevents()
''Dim field As ADODB.field
''Dim fieldname As String
''If rs.EditMode = adEditInProgress Or rs.EditMode = adEditAdd Then
''
''      If rs.EditMode = adEditAdd Then
''
''            If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
''
''                  ObjXevents.AddNew
''            ObjXevents.OldVAlue = rs.Fields("usr_leve").OriginalValue
''            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
''
''            ObjXevents.NewVAlue = txtusr_leve.Text
''            ObjXevents.HisLoginId = Trim$(txtusr_userid)
''            ObjXevents.MyLoginId = CurrentUser
''            ObjXevents.NameSpace = NameSpace
''            ObjXevents.STAs = "A"
''            ObjXevents.EventDetail = " A New User  " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
''
''            Exit Function
''      End If
''
''      If Not IsNull(rs.Fields("usr_leve").OriginalValue) And txtusr_leve.Text <> rs.Fields("usr_leve").OriginalValue Then
''
''              If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
''
''
''                  ObjXevents.AddNew
''            ObjXevents.OldVAlue = rs.Fields("usr_leve").OriginalValue
''            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
''
''            ObjXevents.NewVAlue = txtusr_leve.Text
''            ObjXevents.HisLoginId = Trim$(txtusr_userid)
''            ObjXevents.MyLoginId = CurrentUser
''            ObjXevents.NameSpace = NameSpace
''            ObjXevents.STAs = "A"
''
''            ObjXevents.NewVAlue = txtusr_leve.texr
''            ObjXevents.HisLoginId = Trim$(txtusr_userid)
''            ObjXevents.MyLoginId = CurrentUser
''            objevents.NameSpace = NameSpace
''            ObjXevents.STAs = "b"
''
''
''
''
''   If Not IsNull(rs.Fields("usr_expidate").OriginalValue) And DTPicker1.Value <> rs.Fields("usr_expidate").OriginalValue Then
''
''              If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
''
''
''             ObjXevents.AddNew
''             ObjXevents.OldVAlue = rs.Fields("usr_expidate").OriginalValue
''            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
''
''            ObjXevents.NewVAlue = DTPicker1.Value
''            ObjXevents.HisLoginId = Trim$(txtusr_userid)
''            ObjXevents.MyLoginId = CurrentUser
''            ObjXevents.NameSpace = NameSpace
''            ObjXevents.STAs = "A"
''
''
''        ObjXevents.EventDetail = " user" & Trim$(txtusr_userid) & " tried to connect with an unknown login " & "."
''            If rs.EditMode = 1 Then
''            ObjXevents.EventDetail = " The Expiry Date of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been changed by " & CurrentUser & " from " & rs.Fields("usr_expidate").OriginalValue & " to " & DTPicker1.Value & "."
''            End If
''
''  End If
''
''   If rs.EditMode = 2 Then
''   ObjXevents.EventDetail = " user " & Trim$(txtuser_userid) & " tried to connect with an unknown login " & "."
''
''    For Each field In rs.Fields
''       Debug.Print field.Name
''
''
''         If field.OriginalValue <> field.Value Or (IsNull(field.OriginalValue) And IsNull(field.Value) = False) Then
''            If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
''         If UCase$(Trim$(field.Name)) = UCase$("usr_leve") Or UCase$(Trim$(field.Name)) = UCase$("usr_expidate") Then GoTo NEXTFIELD
''            ObjXevents.AddNew
''            ObjXevents.OldVAlue = field.OriginalValue
''            ObjXevents.NewVAlue = field.Value
''            ObjXevents.HisLoginId = Trim$(txtusr_userid)
''            ObjXevents.MyLoginId = CurrentUser
''            ObjXevents.NameSpace = NameSpace
''            ObjXevents.STAs = "A"
''
'''..........................................................................................................
''
''
'''set timer frequency
''
''Private Sub Timer1_Timer()
''Static Index As Integer
''
''    Index = Index Mod 3
''    Index = Index + 1
''    imgLight(0).Picture = imgLight(Index).Picture
''End Sub

'start timer

Private Sub txtPWD_Change(Index As Integer)
    Timer1.Enabled = True
End Sub

'set password text boxse to entry data

Private Sub txtPWD_GotFocus(Index As Integer)
    Call SelectWhenFocus(txtPwd(Index))
End Sub

'show confirm password text box
Private Sub txtPWD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Stop
    If Index = 0 Then
        If txtPwd(1).Visible Then
            If KeyAscii = 13 Then txtPwd(1).SetFocus
            End If
    End If
End Sub

'set userid text box to entry data

Private Sub txtUserID_GotFocus()
    Call SelectWhenFocus(txtUserID)
End Sub

'set return key equal to cmdOK_click

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOk_Click
    Else
        Call SetPasswordEntry("L")
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

'call function to check login user status

Private Function GetUserStatus() As LoginStatus
Dim MsgId As Long
Dim rs As ADODB.Recordset
Dim sql
On Error GoTo Handled

    stat = 0
    Set rst = Nothing
    'Set rst = New ADODB.Recordset
    GetUserStatus = CheckUserStatus(NameSpaces(cmbName), txtUserID, cn, rst)
    MsgId = 100 + GetUserStatus
    
    Select Case GetUserStatus
        
        Case lsInActive, lsSuspended, lsTempSuspended, _
              lsUnknown, lsUnknownUser, lsFirstLogonWithAdmin, _
              lsFirstLogonWithOwner, lsFirstLogonWithOutOwnerOrAdmin
              
              
              
            If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
             
             If GetUserStatus = lsInActive Then
             
                    ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with an inactive login" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
               ElseIf GetUserStatus = lsSuspended Then
               
               ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with a suspended login" & "."
                    ObjXevents.STAs = "S"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
               ElseIf GetUserStatus = lsTempSuspended Then
               
               ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with a temporary suspended login" & "."
                    ObjXevents.STAs = "T"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
               ElseIf GetUserStatus = lsUnknown Then
                    
               ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with an unknown login" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                ElseIf GetUserStatus = lsUnknownUser Then
               
               ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with an unknown user login" & "."
                    ObjXevents.STAs = "S"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                ElseIf GetUserStatus = lsFirstLogonWithOwner Then
                
                    ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect first time without temporary owner password" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                                                            
                ElseIf GetUserStatus = lsFirstLogonWithAdmin Then
                
                ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect first time without temporary passwords" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                ElseIf GetUserStatus = lsFirstLogon Then
                 
                 ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " connected first time" & "."
                    ObjXevents.STAs = "A"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                ElseIf GetUserStatus = lsFirstLogonWithOutOwnerOrAdmin Then
                 
                 ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " connected first time without admin and owner temporary passwords " & "."
                    ObjXevents.STAs = "A"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                    
                    
                         
               
               'elseif GetUserStatus =
               
             End If
             
                    Call MsgBox(Replace(LoadResString(MsgId), "/n", vbCrLf))
                    GetUserStatus = lsUnknown
                      
              
               Case Else
             
            End Select

Handled:
    If Err Then MsgBox Err.Description: Err.Clear
End Function

'set data to class collection

Public Function Result() As LoginSuccess
    Result = FResult
End Function

'set data to class collection

Public Function User() As String
    User = txtUserID
End Function

'check name space and get new one

Public Function NameSpace() As String
On Error Resume Next

    If NameSpaces.count Then _
        NameSpace = NameSpaces(cmbName)

    If Err Then Err.Clear
End Function

'set login status and call function to set password
'entry text boxse

Private Function HandleUserName() As LoginStatus
Dim str As String

    LoginType = GetUserStatus
    HandleUserName = LoginType
    
    Select Case HandleUserName
        
        Case lsUnknown
            txtUserID.SetFocus
            Call txtUserID_GotFocus: Exit Function
            
            
        Case lsSuccess
            str = "R"
            
        Case lsFirstLogon
            str = "F"
            
        Case lsSuspendedWithPasssword, lsTempSuspendedWithPasssword
            str = "V"
            
    End Select
    
    
    Call SetPasswordEntry(str)
    Call txtPwd(0).SetFocus

End Function

'function to handle regular login, call function to check password
'exist or not, insert record to xevent

Private Function handleregularlogin() As LoginSuccess
Dim Message As Integer
Dim men As String
Dim Attachments(0) As String
   
   If StrComp(Trim$(Encrypt(rst!USR_pswd & "")), txtPwd(0), vbTextCompare) <> 0 Then
        Message = Getremattadnum - 1
        'MsgBox LoadResString(120)
        If Message > 0 Then
            'Modified by Juan Gonzalez (8/29/200) for Transaction fixes
            msg1 = Trans("M00096")
            MsgBox IIf(msg1 = "", "Invalid password, Remaining attempts: '", msg1) + " '" & Message & "'  "
            '----------------------------------------------------------
            If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
            ObjXevents.AddNew
            ObjXevents.MyLoginId = txtUserID
            ObjXevents.HisLoginId = txtUserID
            ObjXevents.NameSpace = NameSpaces(cmbName)
            ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with a wrong personal password."
            ObjXevents.STAs = "A"
            ObjXevents.update
            Set ObjXevents = Nothing
        ElseIf Message = 0 Then
            'Modified by Juan Gonzalez (8/29/200) for Transaction fixes
            msg1 = Trans("M00096")
            MsgBox IIf(msg1 = "", "Invalid password.", msg1)
            '----------------------------------------------------------
            
            If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
            ObjXevents.AddNew
            ObjXevents.MyLoginId = txtUserID
            ObjXevents.HisLoginId = txtUserID
            ObjXevents.NameSpace = NameSpaces(cmbName)
            ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " tried to connect with a wrong personal password."
            ObjXevents.STAs = "A"
            ObjXevents.update
            Set ObjXevents = Nothing

            If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
            ObjXevents.AddNew
            ObjXevents.MyLoginId = txtUserID
            ObjXevents.HisLoginId = txtUserID
            ObjXevents.NameSpace = NameSpaces(cmbName)
            ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " is suspended."
            ObjXevents.STAs = "S"
            ObjXevents.update
                
            'No events to be tracked for user IMSUSA
            If Not UCase(Trim$(txtUserID)) = "IMSUSA" Then
                Attachments(0) = ""
                Call SendEmails("User " & Trim$(txtUserID) & " is suspended for Namespace " & cmbName, "User Suspension", _
                   "LO", NameSpaces(cmbName), cn, Attachments)
            End If
            Set ObjXevents = Nothing
        Else
            msg1 = Trans("M00096")
            MsgBox IIf(msg1 = "", "Wrong password. Please try again.", msg1), vbInformation, "IMS"
            '----------------------------------------------------------
        End If
        
        Call AddInvalidAttempt(txtUserID, NameSpace, cn)
        Call UNKNOWPASS(NameSpace, txtUserID, 1, 1, cn)
    
         
        'Call XEVENT_INSERT(NameSpace, Trim(txtUserID), Trim(txtUserID), LoadResString(1), rst!usr_stas & "", cn)
        Call SetPasswordEntry("L")
    Else
        FResult = lgsSuccess
        Dim tempPASS As Boolean
        
        If Not IsNull(rst!USR_temppswdownr) Then
            If rst!USR_temppswdownr <> "" Then
                tempPASS = True
                stat = True
                LoginType = lsFirstLogon
                Call SetPasswordEntry("N")
                handleregularlogin = lgsCancelled: Exit Function
            End If
        End If
        If Not tempPASS Then
            If rst!MAXPWDREACHED Then
                stat = True
                LoginType = lsFirstLogon
                Call SetPasswordEntry("N")
                handleregularlogin = lgsCancelled: Exit Function
            End If
        End If
     End If
    
    handleregularlogin = lgsSuccess
End Function

'function to check first login user

Public Function HandleFirstLogon() As LoginSuccess
Dim msg As String
Dim Style As VbMsgBoxStyle
Dim enc As ImsSecX.imsCryptoClass

    HandleFirstLogon = lgsSuccess
    Set enc = New ImsSecX.imsCryptoClass
    
    If stat = 0 Then
    
        If IsStringEqual(Trim$(enc.Encrypttext(rst!usr_tempswdadmn & "", CryptKey)), txtPwd(0)) And _
           IsStringEqual(Trim(enc.Encrypttext(rst!USR_temppswdownr & "", CryptKey)), txtPwd(1)) Then
       
            stat = 1
            Call SetPasswordEntry("N")
        Else
            msg = "Password is incorrect"
            HandleFirstLogon = lgsCancelled
            txtPwd(0) = ""
             txtPwd(1) = ""
             txtPwd(0).SetFocus
            
            
        End If
        
    Else
        
        If IsStringEqual(txtPwd(0), txtPwd(1)) Then
        
            If ComparePasswordLength Then
            
           'Added by muzammil.04/06/01
           'Reason -
           'In a case where the user is prompted to enter a new password (when the screnn shows two
           ' text boxes to confirms the passwords) ,if the use enters an old password which is already
           'there in the xpassword list,even though the message is shown that the password is valid for
           '35 days but the password is not saved and the same screen appears when the user tried loggin
           'in the second time.
           
           'Added this line
           
                If PassWordExist(NameSpace, txtUserID, Encrypt(txtPwd(0).Text), cn) Then
                  MsgBox "The password already exist please use a different one.", vbInformation, "Imswin"
                  HandleFirstLogon = lgsCancelled
                    txtPwd(0) = ""
                    txtPwd(1) = ""
                    txtPwd(0).SetFocus
                    Exit Function
                  
                  
                End If
                    
                    
                
                If Not SavePassword Then
                    HandleFirstLogon = lgsCancelled: Exit Function
                End If
                
                FResult = lgsSuccess
                HandleFirstLogon = lgsSuccess
                Call InsertIntoXevent(LoadResString(2))
            Else
                Style = vbExclamation
                HandleFirstLogon = lgsCancelled
                msg = ReplaceNewLine(ReplaceDecimal(LoadResString(122), rst!usr_minipswdleng))
                
            End If
            
        Else
            Style = vbCritical
            MsgBox ReplaceNewLine(LoadResString(123))
            HandleFirstLogon = lgsCancelled
             txtPwd(0) = ""
             txtPwd(1) = ""
             txtPwd(0).SetFocus
             Exit Function
        End If
        
    End If
   ' HandleFirstLogon = lgsSuccess
    If Len(msg) Then Call MsgBox(msg, Style)
End Function

'function to check suuspended user login

Public Function HandleSuspendedUser() As LoginSuccess
Dim msg As String
Dim Style As VbMsgBoxStyle
Dim sPassword As String

    
    HandleSuspendedUser = lgsCancelled
    sPassword = Encrypt(rst!usr_tempswdadmn & "")
   
    
    If stat = 0 Then

        If IsStringEqual(txtPwd(0).Text, sPassword) Then
        
            stat = 1
            Call SetPasswordEntry("N")
            HandleSuspendedUser = lgsSuccess
            
        Else
        
            Style = vbCritical
            msg = "Wrong Password. Please try again."
            txtPwd(0).Text = "" 'M 08/22
        End If
        
    Else
        If IsStringEqual(txtPwd(0), txtPwd(1)) Then
        
            If (ComparePasswordLength) Then
            
                If (Not (PassWordExist(NameSpace, txtUserID, Encrypt(txtPwd(0).Text), cn))) Then
                
                    If SavePassword Then
                        FResult = lgsSuccess
                        HandleSuspendedUser = lgsSuccess
                        'This is the Case when the User is Suspended with a temp password
                   If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
                     ObjXevents.AddNew
                    ObjXevents.MyLoginId = txtUserID
                    ObjXevents.HisLoginId = txtUserID
                    ObjXevents.NameSpace = NameSpaces(cmbName)
                    ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " Connected first time" & "."
                    ObjXevents.STAs = "I"
                    ObjXevents.update
                    Set ObjXevents = Nothing
                    
                        
                        'Call InsertIntoXevent(LoadResString(3))
                    End If
        
                        
                Else
                    msg = "Cannot reuse a password. Please enter a different one."
                    txtPwd(0).Text = ""
                    txtPwd(1).Text = ""
                End If
                
            Else
                Style = vbExclamation
                msg = ReplaceNewLine(ReplaceDecimal(LoadResString(122), rst!usr_minipswdleng))
                
            End If
            
        Else
            Style = vbCritical
            HandleSuspendedUser = lgsCancelled
            msg = ReplaceNewLine(LoadResString(123))
            txtPwd(0) = ""
             txtPwd(1) = ""
             txtPwd(0).SetFocus
            
            
            
        End If
        
    End If
    
    If Len(msg) Then Call MsgBox(msg, Style)
End Function

'call function to check password length

Private Function ComparePasswordLength() As Boolean
    ComparePasswordLength = CompareStringLength(txtPwd(0), rst!usr_minipswdleng) <> -1
End Function

'call function to update user xuserprofile table

Private Function SavePassword() As Boolean

    SavePassword = True
    Call UpdateUserPassword(NameSpace, txtUserID, Encrypt(txtPwd(0)), cn)
    
    If cn.Errors.count Then
        SavePassword = False
        MsgBox ReplaceNewLine(cn.Errors(0).Description)
        
    Else
        MsgBox ReplaceNewLine(ReplaceDecimal(LoadResString(121), rst!usr_maxidayspswd))
    End If
   ' ObjXevents.EventDetail = " user" & Trim$(txtusr_userid) & " tried to connect with an invalid login "
    
End Function

'insert record to xevent table

Public Function InsertIntoXevent(sEvent As String) As Boolean
    InsertIntoXevent = XEVENT_INSERT(NameSpace, txtUserID, txtUserID, sEvent, rst!usr_stas & "", cn)
End Function

'SQL statement to get remand attempt number

Private Function Getremattadnum() As Integer
On Error Resume Next
Dim str As String
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    
     With cmd
        .CommandType = adCmdText
         Set .ActiveConnection = cn
         
        .CommandText = "SELECT usr_numbremaatte "
        .CommandText = .CommandText & " From XUSERPROFILE "
        .CommandText = .CommandText & " Where usr_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " and usr_autoinacflag = 1"
        .CommandText = .CommandText & " AND usr_userid = '" & txtUserID & "'"
        
    '  .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        Getremattadnum = rst!usr_numbremaatte
    End With
        
        Set cmd = Nothing
    Set rst = Nothing
  If Err Then Call LogErr(Name & "::Getremattadnum", Err.Description, Err, True)
End Function

'set value to anme space description

Public Property Get namespacedescription() As String
    namespacedescription = cmbName.Text
    If Err Then Call LogErr(Name & "::getremattadnum", Err.Description, Err, True)
End Property
 
 

Public Function UpdateMaxAttempts(NameSpace As String, UserId As String, cn As ADODB.Connection)
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdText
        
        .CommandText = "UPDATE XUSERPROFILE SET "
        .CommandText = .CommandText & " usr_numbinvaatte = 0, "
        .CommandText = .CommandText & " usr_numbremaatte = usr_maxiatte"
        
        .CommandText = .CommandText & " WHERE usr_userid = ?"
        .CommandText = .CommandText & " AND usr_npecode = ?"
        .CommandText = .CommandText & " and  datediff(hh,usr_lastinvaattedate,getdate())>1 "
        .Execute , Array(UserId, NameSpace), adExecuteNoRecords
        
    End With
    
   ' If ObjXevents Is Nothing Then Call InitializeXevents(ObjXevents, cn)
    '                ObjXevents.AddNew
     '               ObjXevents.MyLoginId = txtUserID
      '              ObjXevents.HisLoginId = txtUserID
       '             ObjXevents.NameSpace = NameSpaces(cmbName)
        '            ObjXevents.EventDetail = "User " & Trim$(txtUserID) & " logged in Successfully after an hour from invalid attempts.His Attempts left has been reset to the Maximum Attmepts (RESET TIME OUT)."
         '           ObjXevents.STAs = "A"
          '          ObjXevents.update
           '         Set ObjXevents = Nothing

    
    Set cmd = Nothing
   
    Exit Function
    
End Function

Public Function GetTheUserName()
Dim completepath As String
Dim i As Integer

  If Len(Trim(CurrentUser)) > 0 Then
    txtUserID = CurrentUser
    Call cmdOk_Click
  End If
    
End Function
