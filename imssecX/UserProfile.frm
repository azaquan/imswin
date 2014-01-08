VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmUserProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Profile"
   ClientHeight    =   5535
   ClientLeft      =   4080
   ClientTop       =   2985
   ClientWidth     =   9120
   Icon            =   "UserProfile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Tag             =   "04010200"
   Begin VB.CheckBox chkusr_autoinacflag 
      Alignment       =   1  'Right Justify
      Caption         =   "Automatic Inactivation"
      DataField       =   "usr_autoinacflag"
      Height          =   230
      Left            =   240
      TabIndex        =   16
      Top             =   3060
      Width           =   2130
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cbousr_stas 
      DataField       =   "usr_stas"
      Height          =   315
      ItemData        =   "UserProfile.frx":000C
      Left            =   8385
      List            =   "UserProfile.frx":001C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1980
      Width           =   510
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "usr_expidate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   1860
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   20447235
      CurrentDate     =   55104
   End
   Begin VB.TextBox txtusr_tempswdadmn 
      DataField       =   "usr_tempswdadmn"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7035
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   29
      Top             =   4500
      Width           =   1860
   End
   Begin MSComCtl2.UpDown udUsrLevel 
      Height          =   270
      Left            =   4095
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2250
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   476
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtusr_leve"
      BuddyDispid     =   196615
      OrigLeft        =   3766
      OrigTop         =   2580
      OrigRight       =   3961
      OrigBottom      =   2895
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtusr_leve 
      DataField       =   "usr_leve"
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   2220
      Width           =   2160
   End
   Begin VB.ComboBox cboNameSpace 
      DataField       =   "usr_npecode"
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Text            =   "cboNameSpace"
      Top             =   960
      Visible         =   0   'False
      Width           =   2160
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DisableSaveOnSave=   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin VB.TextBox txtusr_maxiatte 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_maxiatte"
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   3420
      Width           =   1620
   End
   Begin VB.TextBox txtusr_resttmot 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_resttmot"
      DataMember      =   "UserProfile"
      DataSource      =   "deims"
      Height          =   315
      Left            =   8235
      TabIndex        =   23
      Top             =   3780
      Width           =   660
   End
   Begin VB.TextBox txtusr_username 
      DataField       =   "usr_username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   2955
   End
   Begin VB.TextBox txtusr_userid 
      DataField       =   "usr_userid"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   1140
      Width           =   2160
   End
   Begin VB.TextBox txtusr_numbprevpswd 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_numbprevpswd"
      DataMember      =   "UserProfile"
      DataSource      =   "deims"
      Height          =   315
      Left            =   8235
      TabIndex        =   10
      Top             =   2340
      Width           =   660
   End
   Begin VB.TextBox txtusr_minipswdleng 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_minipswdleng"
      DataMember      =   "UserProfile"
      DataSource      =   "deims"
      Height          =   315
      Left            =   8235
      TabIndex        =   20
      Top             =   3420
      Width           =   660
   End
   Begin VB.TextBox txtusr_minidayspswd 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_minidayspswd"
      DataMember      =   "UserProfile"
      DataSource      =   "deims"
      Height          =   315
      Left            =   8235
      TabIndex        =   14
      Top             =   2700
      Width           =   660
   End
   Begin VB.TextBox txtusr_menuleve 
      DataField       =   "usr_menuleve"
      Height          =   315
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2640
      Width           =   225
   End
   Begin VB.TextBox txtusr_maxidayspswd 
      Alignment       =   1  'Right Justify
      DataField       =   "usr_maxidayspswd"
      DataMember      =   "UserProfile"
      DataSource      =   "deims"
      Height          =   315
      Left            =   8235
      TabIndex        =   18
      Top             =   3060
      Width           =   660
   End
   Begin VB.Label Labeldate 
      Height          =   315
      Left            =   1920
      TabIndex        =   46
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label MenuLevelDesc 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2400
      TabIndex        =   45
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Created by:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   44
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Expiration Date:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   43
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User Menu Level:"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   42
      Top             =   2720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   41
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   40
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User Level:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   39
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Last Invalid Attempt Date:"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   38
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Number Invalid Attempt:"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   37
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User Max attempt:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   36
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Number Remain Attempt:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   35
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_creauser"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2160
      TabIndex        =   34
      Top             =   1500
      Width           =   2160
   End
   Begin VB.Label lblusr_datelastlogn 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_datelastlogn"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   7035
      TabIndex        =   26
      Top             =   4140
      Width           =   1860
   End
   Begin VB.Label lblusr_numbinvaatte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_numbinvaatte"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      TabIndex        =   24
      Top             =   3780
      Width           =   1620
   End
   Begin VB.Label lblusr_numbremaatte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_numbremaatte"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      TabIndex        =   27
      Top             =   4140
      Width           =   1620
   End
   Begin VB.Label lblusr_lastinvaattedate 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_lastinvaattedate"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      TabIndex        =   30
      Top             =   4500
      Width           =   1620
   End
   Begin VB.Label lblusr_creadate 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "usr_creadate"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label lblFieldLabel 
      Caption         =   "First Time Login Password"
      Height          =   195
      Index           =   21
      Left            =   4620
      TabIndex        =   28
      Top             =   4560
      Width           =   2400
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Last Login Date:"
      Height          =   195
      Index           =   3
      Left            =   4620
      TabIndex        =   25
      Top             =   4200
      Width           =   2400
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Reset Time Out:"
      Height          =   195
      Index           =   16
      Left            =   4620
      TabIndex        =   22
      Top             =   3840
      Width           =   3600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "User Profile Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8715
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User ID:"
      Height          =   195
      Index           =   18
      Left            =   -540
      TabIndex        =   33
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "User Status:"
      Height          =   195
      Index           =   17
      Left            =   6960
      TabIndex        =   32
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Number Previous Password to keep:"
      Height          =   195
      Index           =   14
      Left            =   4620
      TabIndex        =   9
      Top             =   2400
      Width           =   3525
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Namespace:"
      Height          =   195
      Index           =   12
      Left            =   6720
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Min. Password Length:"
      Height          =   195
      Index           =   11
      Left            =   4620
      TabIndex        =   19
      Top             =   3480
      Width           =   3600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Min. days to keep password:"
      Height          =   195
      Index           =   10
      Left            =   4620
      TabIndex        =   13
      Top             =   2760
      Width           =   3600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Max. days to keep password:"
      Height          =   195
      Index           =   8
      Left            =   4620
      TabIndex        =   17
      Top             =   3120
      Width           =   3600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "On:"
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   5
      Top             =   2040
      Width           =   435
   End
End
Attribute VB_Name = "frmUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim recordchange As String
Dim fUser As String
Dim NameSpace As String
Dim UserLevel As Integer
Dim AddingNew As Boolean
Dim cn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim ObjXevents As ImsXevents
Dim TRANSACTIONNUBMER As Integer
Dim hasrecordchange As Integer



Function checkFIELDS() As Boolean
Screen.MousePointer = 11
    Call txtusr_tempswdadmn_KeyPress(13)
    If txtusr_username = "" Then
        MsgBox "User Name Can Not be empty"
        txtusr_username.SetFocus
        Exit Function
    End If
    If txtusr_userid = "" Then
        MsgBox "Invalid User ID"
        txtusr_userid.SetFocus
        Exit Function
    End If
    If DTPicker1.Value <= Now - 1 Then
        MsgBox "Invalid Expiration Date"
        DTPicker1.SetFocus
        Exit Function
    End If
    If txtusr_leve = "" Then
        MsgBox "Ivalid User Level"
        txtusr_leve.SetFocus
        Exit Function
    End If
    If chkusr_autoinacflag.Value Then
        If Not IsNumeric(txtusr_maxiatte) Then
            MsgBox "Invalid User Max Attemp Number"
            txtusr_maxiatte.SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtusr_numbprevpswd) Then
        MsgBox "Invalid Number of Previous Password to keep"
        txtusr_numbprevpswd.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtusr_minidayspswd) Then
        MsgBox "Invalid Number of Days"
        txtusr_minidayspswd.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtusr_maxidayspswd) Then
        MsgBox "Invalid Number of Days"
        txtusr_maxidayspswd.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtusr_minipswdleng) Then
        MsgBox "Invalid Password Lenght"
        txtusr_minipswdleng.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtusr_resttmot) Then
        MsgBox "Invalid Reset Time"
        txtusr_resttmot.SetFocus
        Exit Function
    End If
    
    If Len(Trim$(txtusr_menuleve)) = 0 Then
        MsgBox "Menu level can not be left empty.", vbInformation, "imswin"
        txtusr_menuleve.SetFocus
        Exit Function
    End If
    
    
    If rs.EditMode = 2 Then
       
        If ISUSERIDVALID(txtusr_userid) = False Then
        
           MsgBox "User id already exists.Please use a different one."
           txtusr_userid.SetFocus
           Exit Function
           
        End If
        
   End If
    
        
        If Len(Trim$(txtusr_tempswdadmn)) < CInt(txtusr_minipswdleng) Then
            If txtusr_tempswdadmn <> "" Then
                msg1 = Trans("M00133")
                MsgBox IIf(msg1 = "", "Temporary password is too short", msg1)
                txtusr_tempswdadmn.SetFocus
                txtusr_tempswdadmn.BackColor = &HC0FFFF
                Exit Function
            End If
        End If

    'If txtusr_tempswdadmn = "" Then
        'MsgBox "Invalid Password"
        'txtusr_tempswdadmn.SetFocus
        'Exit Function
    'End If
    checkFIELDS = True
Screen.MousePointer = 0
End Function


Sub lockTHINGS()
    txtusr_username.Enabled = False
    txtusr_userid.Enabled = False
    DTPicker1.Enabled = False
    txtusr_leve.Enabled = False
    txtusr_menuleve.Enabled = False
    chkusr_autoinacflag.Enabled = False
    txtusr_maxiatte.Enabled = False
    cboNameSpace.Enabled = False
    cbousr_stas.Enabled = False
    cbousr_stas.Enabled = False
    txtusr_numbprevpswd.Enabled = False
    txtusr_minidayspswd.Enabled = False
    txtusr_maxidayspswd.Enabled = False
    txtusr_minipswdleng.Enabled = False
    txtusr_resttmot.Enabled = False
    txtusr_tempswdadmn.Enabled = False
End Sub

Private Sub cboNameSpace_GotFocus()
    cboNameSpace.BackColor = &HC0FFFF
End Sub


Private Sub cboNameSpace_LostFocus()
    cboNameSpace.BackColor = vbWhite
End Sub


Private Sub cbousr_stas_Change()
 hasrecordchange = hasrecordchange + 1
End Sub

Private Sub cbousr_stas_GotFocus()
    cbousr_stas.BackColor = &HC0FFFF
End Sub


Private Sub cbousr_stas_LostFocus()
    cbousr_stas.BackColor = vbWhite
End Sub


'set text boxse back ground color

Private Sub chkusr_autoinacflag_Click()
 hasrecordchange = hasrecordchange + 1
    If chkusr_autoinacflag Then
    
        txtusr_maxiatte.Enabled = True
        txtusr_maxiatte.ForeColor = vbWindowText
        txtusr_maxiatte.BackColor = vbWindowBackground
    
    Else
    
        rs!usr_maxiatte = Null
        txtusr_maxiatte.Enabled = False
        txtusr_maxiatte.ForeColor = lblusr_creadate.ForeColor
        txtusr_maxiatte.BackColor = lblusr_creadate.BackColor
    
    End If
End Sub





Private Sub chkusr_autoinacflag_GotFocus()
chkusr_autoinacflag.BackColor = &HC0FFFF
End Sub

Private Sub chkusr_autoinacflag_LostFocus()
chkusr_autoinacflag.BackColor = &H8000000F
End Sub

Private Sub DTPicker1_Change()
 hasrecordchange = hasrecordchange + 1
End Sub

Private Sub DTPicker1_GotFocus()
Labeldate.BackColor = &HC0FFFF
End Sub

Private Sub DTPicker1_LostFocus()
'DTPicker1.CalendarBackColor = vbWhite
Labeldate.BackColor = &H8000000F
End Sub

'set form size and caption

Private Sub Form_Load()
'NEXTFIELDOn Error Resume Next

    NavBar.SaveEnabled = Getmenuuser(NameSpace, CurrentUser, Me.Tag, cn)
    If NavBar.SaveEnabled = False Then Call lockTHINGS
    'navbar.Recordset.Filter =

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmUserProfile")
    '--------------------------------------------------

     'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Caption = Caption + " - " + Tag
                                  
    If Err Then Call LogErr("frmUserProfile::Form_Load", Err.Description, Err, True)
End Sub


Private Sub Form_Paint()
    If NavBar.SaveEnabled = False Then Call lockTHINGS
End Sub

Private Sub Form_Resize()
hasrecordchange = 0
End Sub

'close form and free memory

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set frmUserProfile = Nothing
End Sub





'set adding new value

Private Sub NavBar_BeforeNewClick()
    AddingNew = True
End Sub

'before save recordset validate recordset data format

Private Sub NavBar_BeforeSaveClick()
Dim EncptVAlue As String
On Error Resume Next
Screen.MousePointer = 11
Me.Refresh
Dim Cancel As Boolean

    If Not checkFIELDS Then
        NavBar.SaveEnabled = True
        Screen.MousePointer = 0
        NavBar.AllowUpdate = False
        Exit Sub
    End If
    Screen.MousePointer = 11

    ValidateControls
    Screen.MousePointer = 11
    rs("usr_modiuser") = fUser
    rs("USR_NPECODE") = NameSpace
    
    
    If Not IsNull(rs!usr_tempswdadmn) Then
        
        If Len(Trim$(rs!usr_tempswdadmn & "")) = 0 Then rs!usr_tempswdadmn = Null
    End If
    
    If Len(txtusr_tempswdadmn.Tag) = 0 Then
        If Len(txtusr_tempswdadmn) Then
            
            rs("usr_tempswdadmn") = Encrypt(txtusr_tempswdadmn)
         '   txtusr_tempswdadmn = EncptVAlue: txtusr_tempswdadmn.Tag = 1
              txtusr_tempswdadmn = rs("usr_tempswdadmn"): txtusr_tempswdadmn.Tag = 1
            
        End If
    End If
    
    Screen.MousePointer = 11

    NavBar.AllowUpdate = Not Cancel
    
    If NavBar.AllowUpdate Then
        If Err Then Err.Clear
        'Function Added By  12/25/00
        'Reason - To tracK Security Events
       
      GenerateXevents
        'If len(txtusr_tempswdadmn.Tag = 0 Then
        'If Len(txtusr_tempswdadmn) Then
        'txtusr_tempswwdadmn = Encrypt(txtusr_tempswdadmn): txtusr_tempswdadmn.Tag = 1
        
 rs.Move (0)
      
       ' MsgBox rs.EditMode
        Err.Clear
        cn.Errors.Clear 'M
        cn.BeginTrans 'M
        
        Dim UPDATEMENUUSER As Boolean
        
        If rs.EditMode <> 2 And (rs("USR_MENULEVE").OriginalValue = rs("USR_MENULEVE").Value) Then
            UPDATEMENUUSER = False
        Else
            UPDATEMENUUSER = True
        End If
        
        rs.update
        rs.UpdateBatch
'        rs.Update
        ' Added By  12/25/00
        'Reason - To tracK Security Events
        
        If Err.Number <> 0 Then GoTo ErrHandler
        
        If UPDATEMENUUSER = False Then GoTo COMMITTRANS
        
        
        
        If Err.Number = 0 Then 'Or Err.Number = -2147217864 Then
       
            
            Dim cmd As ADODB.Command
            Dim idofuser As String
            Dim menulevelofuser As String
            Set cmd = New ADODB.Command
            
            idofuser = Trim$(txtusr_userid)
            menulevelofuser = Trim$(txtusr_menuleve)
            
            cmd.ActiveConnection = cn
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "REFRESHMeNUUSER"
           
            cmd.Parameters.Append cmd.CreateParameter("@NPECODE", adVarChar, adParamInput, 5, NameSpace)
            cmd.Parameters.Append cmd.CreateParameter("@USERID", adVarChar, adParamInput, 10, idofuser)
            cmd.Parameters.Append cmd.CreateParameter("@MENULEVEL", adVarChar, adParamInput, 15, menulevelofuser)
              
            cmd.Execute
           
            If Err.Number <> 0 Or cn.Errors.count > 0 Then GoTo ErrHandler
           
COMMITTRANS:
           
           
            'UPDATEMNUUSER = False
            If Err.Number = 0 And cn.Errors.count = 0 Then
            
               cn.COMMITTRANS
               MsgBox "Record saved successfully.", vbInformation, "Imswin"
               
             End If
            If IsNothing(ObjXevents) = False Then
              
              ObjXevents.update
              Set ObjXevents = Nothing
              
            End If
             
             
        Else
ErrHandler:
            Dim str  As String
    
            Dim ERROR As Object
                       
                For Each ERROR In cn.Errors
                  
                   str = str & ERROR.Description & vbCrLf
            
                Next
    
            MsgBox "Could not Save the Record.Errors Occurred -- " & str, vbCritical, "Imswin"
            
            cn.RollbackTrans
            
            Err.Clear
        End If
        
    Screen.MousePointer = 0
    End If
    
    If Err Then
        Screen.MousePointer = 0
        MsgBox Err.Description
        NavBar.AllowUpdate = False
        Err.Clear
    End If
    
End Sub

Private Sub NavBar_MoveComplete()
    If Not NavBar.SaveEnabled Then Call lockTHINGS
End Sub

Private Sub NavBar_OnCancelClick()
'
End Sub

'close form

Private Sub NavBar_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar_OnFirstClick()
'If hasrecordchange = True Then
   ' If MsgBox("Do you Want To Save the Changes ?", vbYesNo, "Imswin") = vbYes Then
        ' NavBar_BeforeSaveClick
   ' Else
    ''    NavBar_OnCancelClick
    'End If
    'hasrecordchange = False
'End If
End Sub

Private Sub NavBar_OnLastClick()
'If hasrecordchange = True Then
   ' If MsgBox("Do you Want To Save the Changes ?", vbYesNo, "Imswin") = vbYes Then
       '  NavBar_BeforeSaveClick
    'Else
       ' NavBar_OnCancelClick
    'End If
 '   hasrecordchange = False
'End If
 'if isnumeric( the name of the textbox ) then MsgBox "Please enter a valid entry ", , "ImsWin"
      '(name of the control).setfocus
      
  'End If

End Sub

'set values to recordset
Private Sub NavBar_OnNewClick()
   txtusr_userid.Enabled = True
    rs!usr_creauser = fUser
    rs!usr_expidate = Date
    rs!usr_numbinvaatte = 0
    txtusr_tempswdadmn.Tag = ""
    If UserLevel = 1 Then txtusr_leve = 3
    rs!usr_creadate = CDate(Format(Now(), "mm/dd/yyyy hh:nn:ss a"))
    txtusr_username.SetFocus
    txtusr_username.BackColor = &HC0FFFF
    
    
End Sub

Private Sub NavBar_OnNextClick()
''''''If HasRecordChanged Then
''''''    If MsgBox("Do you Want To Save the Changes ?", vbYesNo, "Imswin") = vbYes Then
''''''         NavBar_BeforeSaveClick
''''''    Else
''''''        NavBar_OnCancelClick
''''''    End If
''''''    'hasrecordchange = False
''''''End If
End Sub

Private Sub NavBar_OnPreviousClick()
'If hasrecordchange > 0 Then
'    If MsgBox("Do you Want To Save the Changes ?", vbYesNo, "Imswin") = vbYes Then
 '        NavBar_BeforeSaveClick
 ''   Else
  ''      NavBar_OnCancelClick
 ''   End If
 '   hasrecordchange = False
'End If

'NavBar.PreviousEnabled = True
'hasrecordchange = 0
End Sub

'get crystal report parametrs
'and application path

Private Sub NavBar_OnPrintClick()
Dim ddss As String
Dim ddss_name As String
Dim ddss_bolean As Boolean
Err.Clear
    
    With Me.CrystalReport1
        .ReportFileName = ReportPath + "indiuserprof.rpt"
        .ParameterFields(0) = "namespace;" + NameSpace + ";TRUE"
        .ParameterFields(1) = "userid;" + rs!usr_userid + ";TRUE"
        'Modified by Juan (10/23/00) for Multilingual 'J added
        msg1 = Trans("M00203") 'J added
        .WindowTitle = IIf(msg1 = "", "User Profile", msg1) 'J modified
        
        ddss = "indiuserprof.rpt"
        ddss_name = Me.Name
        ddss_bolean = True
        Call translate_reports(ddss_name, ddss, ddss_bolean, cn, Me.CrystalReport1) 'J added
        Dim repo As String 'J added
        Dim i As Integer 'J added
        
        For i = 0 To .GetNSubreports - 1 'J added
            repo = .GetNthSubreportName(i) 'J added
            .SubreportToChange = repo 'J added
            Call translate_reports(Me.Name, repo, False, cn, Me.CrystalReport1) 'J added
        Next 'J added
        .SubreportToChange = "" 'J added
        '--------------------------------------------------
        .Action = 1
    End With
    
    If Err Then _
        MsgBox Err.Description: _
        Call LogErr("frmUserProfile::NavBar_OnPrintClick", Err.Description, Err, True)
End Sub

' save recordset to check error cause and show description

Private Sub NavBar_OnSaveClick()
On Error Resume Next
    If Not NavBar.AllowUpdate Then
        NavBar.AllowUpdate = True
        Exit Sub
    End If
    Screen.MousePointer = 11
    
    
    Call rs.Move(0)
    
    If Err Then
    
    MsgBox Err.Description
    
    If Err = -2147217864 Then
        rs.CancelUpdate
        rs.Requery
        Set NavBar.Recordset = rs
    
        BindAll
    End If
    
        Err.Clear
    Else
       'Updating the Values to Xevents
       If Not ObjXevents Is Nothing Then
         ObjXevents.update
         Set ObjXevents = Nothing
        End If
        
    End If

    AddingNew = False
    If Err Then Call LogErr("frmUserProfile::NavBar_OnSaveClick", Err.Description, Err, True)
    txtusr_userid.Enabled = False
    Screen.MousePointer = 0
    'Unload Me
End Sub

'SQL statement to get xuserprofile recordset

Public Function OpenRecordset() As Boolean
On Error Resume Next
    If cn Is Nothing Then Err.Raise 1001, "ImsSec", "Invalid Connection"
    If Not GetUserInfo Then Exit Function
    
    Set rs = Nothing
    Set rs = New ADODB.Recordset

    
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    'rs.LockType = adLockOptimistic
    rs.Source = "Select * from xuserprofile"
    rs.Source = rs.Source & " WHERE usr_npecode = '" & NameSpace & "'"
    
    Select Case UserLevel
        Case 1
            rs.Source = rs.Source & " AND usr_leve = 3"
        Case 0
            rs.Source = rs.Source & " AND usr_leve in (1, 2)"
    End Select
    
    rs.Source = rs.Source & " ORDER BY usr_userid"
    
    Call rs.Open(, cn, adOpenStatic)
    
    Set NavBar.Recordset = rs
    Call AddNameSpaces(GetNameSpaces(cn), cboNameSpace, True)
    
    BindAll
    OpenRecordset = True
    If Err Then Call LogErr("frmUserProfile::OpenRecordset", Err.Description, Err, True)
End Function

'clear form data fields

Private Sub ClearAll()
On Error Resume Next

Dim str As String
Dim ctl As Control

    For Each ctl In Controls
    
        str = ctl.DataField
        If Len(str) <> 0 Then
            
            If TypeOf ctl Is TextBox Then
                ctl = ""
                
            ElseIf TypeOf ctl Is ComboBox Then
                ctl.ListIndex = -1
                
            ElseIf TypeOf ctl Is Label Then
                ctl = ""
                
            End If
            
        End If
            
        If Err Then Err.Clear
    Next ctl
    
    'datachange = False
End Sub

'assign data to text boxse

Private Sub BindAll()
On Error Resume Next

Dim str As String
Dim ctl As Control

    For Each ctl In Controls
    
        str = ctl.DataField
        
        If Len(str) <> 0 Then
        
            Set ctl.DataSource = NavBar
            
        End If
        If Err Then Err.Clear
    Next
    If Err Then Err.Clear
End Sub

'set value to class

Public Sub SetUser(User As String)
    fUser = User
End Sub

'call function to check user level and show messege

Public Function GetUserInfo() As Boolean
Dim rst As ADODB.Recordset
Dim datax As ADODB.Recordset
Dim sql

    GetUserInfo = True
    udUsrLevel.Max = -1
    udUsrLevel.Min = -1
    If CheckUserStatus(NameSpace, fUser, cn, rst) = 1 Then Exit Function
    
    sql = "SELECT * FROM XUSERPROFILE WHERE usr_userid = '" + fUser + "' AND usr_npecode = '" + NameSpace + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        UserLevel = datax!usr_leve
    Else
        Exit Function
    End If
    
    If UserLevel = 0 Then
        udUsrLevel.Min = 0
        udUsrLevel.Max = 2
        
    ElseIf UserLevel = 1 Then
        udUsrLevel.Min = 3
        udUsrLevel.Max = 3
        txtusr_leve.Locked = True
        
    Else
        'Unload Me
        GetUserInfo = False
        
        'Modified by Juan Gonzalez (8/29/2000) for Translation fix
        msg1 = Trans("M00121")
        MsgBox IIf(msg1 = "", "You don not have the rights to change or add a user", msg1)
        '---------------------------------------------------------
    End If
    
End Function

'set name space to current name space

Public Sub SetNameSpace(np As String)
    NameSpace = np
End Sub

'check recordset end of file and before of file
'and assign data values to text boxse

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If (((adReason > adRsnMoveFirst) And (adReason < adRsnMoveLast)) Or (adReason = adRsnMove)) Then
    
        If rs.EOF Or rs.BOF Then Exit Sub
        
        If Len(rs!usr_tempswdadmn & "") Then
            txtusr_tempswdadmn.Tag = 1
        Else
            txtusr_tempswdadmn.Tag = ""
        End If
    End If
            
End Sub

'check connection status and call function to disable button

Public Sub SetConnection(con As ADODB.Connection)
On Error Resume Next

    Set cn = con
    
    If Not cn Is Nothing Then _
        If cn.State And adStateOpen <> adStateOpen Then cn.Open
        
    Call DisableButtons(Me, NavBar, NameSpace, CurrentUser, con)
        
    If Err Then MsgBox Err.Description
    Err.Clear
    hasrecordchange = 0
End Sub

Private Sub txtusr_leve_Change()
 ' recordchange = txtusr_leve.Text
End Sub

Private Sub txtusr_leve_GotFocus()
    txtusr_leve.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_leve_LostFocus()
    txtusr_leve.BackColor = vbWhite
End Sub


'validate text boxse data format

Private Sub txtusr_leve_Validate(Cancel As Boolean)
On Error Resume Next

    txtusr_leve.SetFocus
    If Len(txtusr_leve) = 0 Then Exit Sub
    
    If Not IsNumeric(txtusr_leve) Then
        txtusr_leve = ""
        
        'Modified by Juan Gonzalez (29/8/2000) for Translation fix
        msg1 = Trans("M00122")
        MsgBox IIf(msg1 = "", "Invalid Value", msg1): Exit Sub
        '---------------------------------------------------------
    Else
        If txtusr_leve.DataChanged Then
            If CInt(txtusr_leve) > udUsrLevel.Max Then txtusr_leve = udUsrLevel.Max
            If CInt(txtusr_leve) < udUsrLevel.Min Then txtusr_leve = udUsrLevel.Min
        End If
    End If
    If txtusr_leve.DataChanged Then
        txtusr_menuleve = txtusr_leve
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub txtusr_maxiatte_Change()
 'recordchange = txtusr_maxiatte.Text
End Sub

Private Sub txtusr_maxiatte_GotFocus()
    txtusr_maxiatte.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_maxiatte_LostFocus()
    txtusr_maxiatte.BackColor = vbWhite
End Sub


'validate maximum attempt data field format

Private Sub txtusr_maxiatte_Validate(Cancel As Boolean)
On Error Resume Next

    Cancel = True
    If Len(txtusr_maxiatte) Then
    
        If Not IsNumeric(txtusr_maxiatte) Then
        
            'Modified by Juan Gonzalez (8/29/2000) for Translation fix
            msg1 = Trans("M00122")
            MsgBox IIf(msg1 = "", "Invalid value", msg1)
            txtusr_maxiatte.SetFocus
            '---------------------------------------------------------
            
        Else
            lblusr_numbremaatte = txtusr_maxiatte
        End If
    End If
    
    Cancel = False
    If Err Then Err.Clear
End Sub

Private Sub txtusr_maxidayspswd_Change()
 'recordchange = txtusr_maxidayspswd.Text
End Sub

Private Sub txtusr_maxidayspswd_GotFocus()
    txtusr_maxidayspswd.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_maxidayspswd_LostFocus()
    txtusr_maxidayspswd.BackColor = vbWhite
End Sub


Private Sub txtusr_menuleve_Change()
'If IsNumeric(txtusr_menuleve) Then
'    If Val(txtusr_menuleve) > 8 Then txtusr_menuleve = ""
'Else
'    txtusr_menuleve = ""
'End If
Dim MenuLevel As New ADODB.Recordset
Dim sql As String
    Set MenuLevel = New ADODB.Recordset
    sql = "SELECT ml_melvname FROM MENULEVEL WHERE " _
        & "ml_npecode = '" + NameSpace + "' AND " _
        & "ml_melvid = '" + txtusr_menuleve + "'"
    MenuLevel.Open sql, cn, adOpenForwardOnly
    If MenuLevel.RecordCount > 0 Then
        MenuLevelDesc = MenuLevel!ml_melvname
    Else
        MenuLevelDesc = "No Menu Level"
    End If
End Sub

Private Sub txtusr_menuleve_GotFocus()
    With txtusr_menuleve
        .BackColor = &HC0FFFF
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtusr_menuleve_LostFocus()
    txtusr_menuleve.BackColor = vbWhite
End Sub


Private Sub txtusr_minidayspswd_Change()
 ' recordchange = txtusr_minidayspswd.Text
End Sub

Private Sub txtusr_minidayspswd_GotFocus()
    txtusr_minidayspswd.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_minidayspswd_LostFocus()
    txtusr_minidayspswd.BackColor = vbWhite
End Sub


Private Sub txtusr_minipswdleng_Change()
 ' recordchange = txtusr_minipswdleng.Text
End Sub

Private Sub txtusr_minipswdleng_GotFocus()
    txtusr_minipswdleng.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_minipswdleng_LostFocus()
    txtusr_minipswdleng.BackColor = vbWhite
End Sub


'validate minimum password lenght data format

Private Sub txtusr_minipswdleng_Validate(Cancel As Boolean)
       
    If Len(txtusr_minipswdleng) = 0 Then Exit Sub
        
           Cancel = True
    If Not IsNumeric(txtusr_minipswdleng) Then
        
        'Modified by Juan Gonzalez (8/29/200) for Translation fix
        msg1 = Trans("M00122")
        MsgBox IIf(msg1 = "", "Invalid Value", msg1)
        '--------------------------------------------------------
        
        txtusr_minipswdleng = ""
        
        Exit Sub
        
    ElseIf txtusr_tempswdadmn <> "" Then
        
        If Len(txtusr_tempswdadmn) < CInt(txtusr_minipswdleng) Then
            If txtusr_tempswdadmn <> "" Then
                txtusr_tempswdadmn.SetFocus
    
                'Modified by Juan Gonzalez (8/29/2000) for Translation fix
                msg1 = Trans("M00133")
                MsgBox IIf(msg1 = "", "Temporary password is too short", msg1)
                '---------------------------------------------------------
            End If
        End If
        
    End If
        
    Cancel = False
End Sub

Private Sub txtusr_numbprevpswd_Change()
 ' recordchange = txtusr_numbprevpswd.Text
End Sub

Private Sub txtusr_numbprevpswd_GotFocus()
    txtusr_numbprevpswd.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_numbprevpswd_LostFocus()
    txtusr_numbprevpswd.BackColor = vbWhite
End Sub


Private Sub txtusr_resttmot_Change()
 ' recordchange = txtusr_resttmot.Text
End Sub

Private Sub txtusr_resttmot_GotFocus()
    txtusr_resttmot.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_resttmot_LostFocus()
    txtusr_resttmot.BackColor = vbWhite
End Sub


Private Sub txtusr_tempswdadmn_Change()
 ' recordchange = txtusr_tempswdadmn.Text
End Sub

Private Sub txtusr_tempswdadmn_GotFocus()
txtusr_tempswdadmn.BackColor = &HC0FFFF
End Sub

'clear temporary adminitration data field

Public Sub txtusr_tempswdadmn_KeyPress(KeyAscii As Integer)
    If Len(txtusr_tempswdadmn) And Len(txtusr_tempswdadmn.Tag) Then
        txtusr_tempswdadmn = ""
        txtusr_tempswdadmn.Tag = ""
    End If
        
End Sub

Private Sub txtusr_tempswdadmn_LostFocus()
 txtusr_tempswdadmn.BackColor = vbWhite
End Sub

'call function to validate temporary password lenght

Private Sub txtusr_tempswdadmn_Validate(Cancel As Boolean)
    Call txtusr_minipswdleng_Validate(Cancel)
End Sub

Private Sub txtusr_userid_Change()
 ' recordchange = txtusr_userid.Text
End Sub

Private Sub txtusr_userid_GotFocus()
    txtusr_userid.BackColor = &HC0FFFF
End Sub


Private Sub txtusr_userid_LostFocus()
    txtusr_userid.BackColor = vbWhite
End Sub


Private Sub txtusr_userid_Validate(Cancel As Boolean)
Dim X As String

If Len(Trim$(txtusr_userid)) = 0 Then Exit Sub
If rs.EditMode <> 2 Then Exit Sub

If ISUSERIDVALID(txtusr_userid) = False Then

   Cancel = True
   MsgBox "User id already exists.Please use a different one."
   txtusr_userid.SetFocus
   
End If
 
End Sub

Private Sub txtusr_username_Change()
' recordchange = txtusr_username.Text

'''If IsNothing(ObjXevents) Then InitializeXevents (ObjXevents)
'''ObjXevents.NewVAlue = Trim$(txtusr_username)
'''ObjXevents.OldVAlue = rs
'''IsUserNameChanged.Changed = True
'''IsUserNameChanged.RowNumb = ObjXevents.RowNumber
End Sub

'Function Added By  12/25/00
'Reason - To Keep track Of Security Events

Public Function GenerateXevents()
Dim field As ADODB.field
Dim FieldName As String
If rs.EditMode = adEditInProgress Or rs.EditMode = adEditAdd Then
   
      If rs.EditMode = adEditAdd Then
     
            If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
            
                  ObjXevents.AddNew
            ObjXevents.OldVAlue = rs.Fields("usr_leve").OriginalValue
            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
            
            ObjXevents.NewVAlue = txtusr_leve.Text
            ObjXevents.HisLoginId = Trim$(txtusr_userid)
            ObjXevents.MyLoginId = CurrentUser
            ObjXevents.NameSpace = NameSpace
            ObjXevents.STAs = "A"
            ObjXevents.EventDetail = " A New User  " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
             
             
            Exit Function
      End If
      
      If Not IsNull(rs.Fields("usr_leve").OriginalValue) And txtusr_leve.Text <> rs.Fields("usr_leve").OriginalValue Then
                  
              If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
       
                  
                  ObjXevents.AddNew
            ObjXevents.OldVAlue = rs.Fields("usr_leve").OriginalValue
            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
            
            ObjXevents.NewVAlue = txtusr_leve.Text
            ObjXevents.HisLoginId = Trim$(txtusr_userid)
            ObjXevents.MyLoginId = CurrentUser
            ObjXevents.NameSpace = NameSpace
            ObjXevents.STAs = "A"
            
            If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new User level of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
            ElseIf rs.EditMode = 1 Then
                  ObjXevents.EventDetail = " The User level of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been changed by " & CurrentUser & " from " & rs.Fields("usr_leve").OriginalValue & " to " & txtusr_leve.Text & "."
                  
                End If
            
    
                  
         End If
   
    
   If Not IsNull(rs.Fields("usr_expidate").OriginalValue) And DTPicker1.Value <> rs.Fields("usr_expidate").OriginalValue Then
                  
              If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
       
                  
             ObjXevents.AddNew
             ObjXevents.OldVAlue = rs.Fields("usr_expidate").OriginalValue
            'ObjXevents.NewVAlue = rs.Fields("usr_leve").Value
            
            ObjXevents.NewVAlue = DTPicker1.Value
            ObjXevents.HisLoginId = Trim$(txtusr_userid)
            ObjXevents.MyLoginId = CurrentUser
            ObjXevents.NameSpace = NameSpace
            ObjXevents.STAs = "A"
            
            
            
            If rs.EditMode = 1 Then
            ObjXevents.EventDetail = " The Expiry Date of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been changed by " & CurrentUser & " from " & rs.Fields("usr_expidate").OriginalValue & " to " & DTPicker1.Value & "."
            End If
                  
  End If
   
   
   
    For Each field In rs.Fields
       Debug.Print field.Name
         
                     
         If field.OriginalValue <> field.Value Or (IsNull(field.OriginalValue) And IsNull(field.Value) = False) Then
            If IsNothing(ObjXevents) Then Call InitializeXevents(ObjXevents, cn)
         If UCase$(Trim$(field.Name)) = UCase$("usr_leve") Or UCase$(Trim$(field.Name)) = UCase$("usr_expidate") Then GoTo NEXTFIELD
            ObjXevents.AddNew
            ObjXevents.OldVAlue = field.OriginalValue
            ObjXevents.NewVAlue = field.Value
            ObjXevents.HisLoginId = Trim$(txtusr_userid)
            ObjXevents.MyLoginId = CurrentUser
            ObjXevents.NameSpace = NameSpace
            ObjXevents.STAs = "A"
            
            FieldName = ""
            
            Select Case field.Name
                  
                  Case "usr_userid"
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The User name " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  End If
                  
                  
                  Case "usr_username"
                        If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The User name " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                        ElseIf rs.EditMode = 1 Then
                           ObjXevents.EventDetail = " The User Name of the User " & Trim$(txtusr_userid) & " has been changed from " & field.OriginalValue & " to " & field.Value & "."
                         End If
                         
                         
                Case "usr_expidate"
                    If rs.EditMode = 2 Then
                           ObjXevents.EventDetail = " The new expiration date of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                    ObjXevents.EventDetail = " The expiration date of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed  from " & field.OriginalValue & " to " & field.Value & "."
                        End If
                        
                  'Case "usr_creadate"
                  ' FieldName = " Create Date "
                       
                 ' Case "usr_creauser"
                 '  FieldName = " Creator "
                 '  ObjXevents.EventDetail = FieldName & " user " & txtusr_userid & "& txtusr_username " & "has been created by " & CurrentUser
                 Case "usr_stas"
                 If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New status " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The status of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                        End If
                        
                 'FieldName = "Status Of The User"
                  Case "usr_autoinacflag"
                  FieldName = " Active Flag "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New automatic inactivation change " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The automatic inactivation change of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                        End If
                  
                  Case "usr_maxiatte"
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New maximum attempts of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The maximum attempts of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                        End If
                  
                  Case "usr_numbinvaatte"
                  FieldName = " No Of InValid Attempts "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new number of invalid attempts of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The number of invalid attempts of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
                  
                  
                  Case "usr_numbprevpswd"
                  FieldName = " No Of Previous "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New number of  previous passwords of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                  ObjXevents.EventDetail = " The number of  previous passwords of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                  End If
                  
                  Case "usr_minidayspswd"
                  FieldName = " Minimum Days Password "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New mimimum days  to keep password of" & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                  ObjXevents.EventDetail = " The minimum days to keep password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                  End If
                
                  
                  Case "usr_maxidayspswd"
                  FieldName = " Maximum Days Password "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The New maximum days to keep password of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                  ObjXevents.EventDetail = " The maximum days to keep password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                  End If
                  
                  Case "usr_minipswdleng"
                  FieldName = " Minimum password Length "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new password length of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                 ElseIf rs.EditMode = 1 Then
                  ObjXevents.EventDetail = " The mimimum length of password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                  End If
                  
                  'FieldName = " Level "
                  Case "usr_tempswdadmn"
                  FieldName = " Temporary Admin Password "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new temporary password " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The temporary password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser '& " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
                  
                  Case "usr_temppswdownr"
                  FieldName = " Temporary Owner Password "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new temporary owner password " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The temporary owner password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser '& " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
                  
                  Case "usr_pswd"
                  FieldName = " Password of the User "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new password of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The password of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser '& " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
                     
                  Case "usr_menuleve"
                  FieldName = " Menulevel "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new menu level of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The menu level of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                        End If
                
                
                        
                  Case "usr_resttmot"
                  FieldName = ""
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new reset time of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The reset time out of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
                     
                  Case "usr_modiuser"
                  FieldName = " Profile Modified  "
                  If rs.EditMode = 2 Then
                            ObjXevents.EventDetail = " The new reset time of " & Trim$(txtusr_userid) & "-" & txtusr_username & " has been created by " & CurrentUser & "."
                  ElseIf rs.EditMode = 1 Then
                     ObjXevents.EventDetail = " The reset time out of " & Trim$(txtusr_userid) & "-" & Trim$(txtusr_username) & " has been changed by " & CurrentUser & " from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue) & "."
                     End If
            End Select
            
            'ObjXevents.EventDetail = FieldName & " The value of the user " & txtusr_userid & " Has Changed from " & Trim(ObjXevents.OldVAlue) & " to " & Trim(ObjXevents.NewVAlue)
           'If Len(ObjXevents.EventDetail) = 0 Then Stop
            
          End If
NEXTFIELD:
    Next field
    
 End If
    
End Function



Public Function HasRecordChanged() As Boolean
Dim field As ADODB.field
HasRecordChanged = False

If Not IsNull(rs.Fields("usr_leve").OriginalValue) And txtusr_leve.Text <> rs.Fields("usr_leve").OriginalValue Then GoTo Recordchanged
   If Not IsNull(rs.Fields("usr_expidate").OriginalValue) And DTPicker1.Value <> rs.Fields("usr_expidate").OriginalValue Then GoTo Recordchanged

For Each field In rs.Fields
              Debug.Print field.Name
         If field.OriginalValue <> field.Value Or (IsNull(field.OriginalValue) And IsNull(field.Value) = False) Then
         
             
             GoTo Recordchanged
         End If
         
               
Next
Exit Function
Recordchanged:
HasRecordChanged = True
End Function

Private Sub txtusr_username_GotFocus()
    txtusr_username.BackColor = &HC0FFFF
End Sub

Private Sub txtusr_username_LostFocus()
    txtusr_username.BackColor = vbWhite
End Sub

Private Sub udUsrLevel_DownClick()
    txtusr_leve.SetFocus
End Sub



Public Function ISUSERIDVALID(UserId As String) As Boolean
  
  Dim rs As New ADODB.Recordset
  
  rs.Source = "select count(*) countit from xuserprofile where usr_npecode='" & NameSpace & "' and usr_userid='" & UserId & "'"
  rs.ActiveConnection = cn
  rs.Open
  
  If rs("COUNTIT") = 1 Then
    ISUSERIDVALID = False
    
  ElseIf rs("COUNTIT") = 0 Then
    ISUSERIDVALID = True
    
  End If
  
End Function


