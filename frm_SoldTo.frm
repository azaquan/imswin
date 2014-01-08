VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_SoldTo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sold To:"
   ClientHeight    =   4485
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   5220
   Tag             =   "01020300"
   Begin VB.ComboBox cbo_code 
      DataField       =   "slt_code"
      DataMember      =   "SOLDTO"
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "slt_ctry"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2340
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "slt_phonnumb"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2640
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "slt_cont"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   11
      Top             =   3540
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "slt_mail"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   59
      TabIndex        =   10
      Top             =   3240
      Width           =   3024
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "slt_faxnumb"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2940
      Width           =   3024
   End
   Begin VB.TextBox txt_State 
      DataField       =   "slt_stat"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2040
      Width           =   408
   End
   Begin VB.TextBox txt_City 
      DataField       =   "slt_city"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1740
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "slt_adr2"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1440
      Width           =   3024
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "slt_adr1"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1140
      Width           =   3024
   End
   Begin VB.TextBox txt_Name 
      DataField       =   "slt_name"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   1
      Top             =   840
      Width           =   3024
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "slt_zipc"
      DataMember      =   "SOLDTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2040
      Width           =   1350
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   3960
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      Mode            =   0
      CommandType     =   0
      CursorLocation  =   0
      CommandType     =   0
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin VB.Label lbl_Sup_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   14
      Top             =   828
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   288
      Left            =   120
      TabIndex        =   22
      Top             =   2940
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   288
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   288
      Left            =   120
      TabIndex        =   24
      Top             =   3540
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   288
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   288
      Left            =   120
      TabIndex        =   20
      Top             =   2340
      Width           =   2000
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   288
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   288
      Left            =   120
      TabIndex        =   17
      Top             =   1740
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   288
      Left            =   120
      TabIndex        =   15
      Top             =   1140
      Width           =   2000
   End
   Begin VB.Label lbl_SoldTo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sold To"
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
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   4935
   End
   Begin VB.Label lbl_Sup_Code 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   13
      Top             =   528
      Width           =   2000
   End
End
Attribute VB_Name = "frm_SoldTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String 'jawdat

'get data to data grid

Private Sub cbo_Code_Click()
'added by shakir
On Error Resume Next
Dim str As String
    
    str = cbo_code
    NavBar1.CancelUpdate
    deIms.rsSOLDTO.CancelUpdate
    deIms.rsSOLDTO.CancelUpdate
    Call deIms.rsSOLDTO.Resync(adAffectCurrent, adResyncAllValues)
    Call RecordsetFind(deIms.rsSOLDTO, "slt_code = '" & cbo_code & "'")
    
    If Err Then Call LogErr(Name & "::cbo_Code_Click", Err.Description, Err.number, True)


End Sub

Private Sub cbo_code_DragDrop(Source As Control, x As Single, Y As Single)
'cbo_code.locked = False
End Sub

'unlock code combo

Private Sub cbo_Code_DropDown()
   cbo_code.locked = False
End Sub

'set back ground color
Private Sub cbo_Code_GotFocus()
    Call HighlightBackground(cbo_code)
End Sub

'do not allow edit new character

Private Sub cbo_Code_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then
KeyAscii = 0
End If
End Sub

'set back ground color

Private Sub cbo_Code_LostFocus()
    Call NormalBackground(cbo_code)
End Sub

Private Sub cbo_Code_Validate(Cancel As Boolean)

If Len(cbo_code) > 10 Then
MsgBox "Code number can not be greater than 10 characters."
Cancel = True
cbo_code.SetFocus
cbo_code = ""
End If
End Sub

'get data for datagrid and populate data grid

Private Sub Form_Load()

'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
    Else
    TableLocked = True
    End If
End If

'end copy




On Error Resume Next
Dim ctl As Control

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_SoldTo")
    '------------------------------------------
    'added by shakir
   
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.SoldTo(deIms.NameSpace)
    
    Call BindAll(Me, deIms)
    Call DisableButtons(Me, NavBar1)
    Set NavBar1.Recordset = deIms.rsSOLDTO
    Call PopuLateFromRecordSet(cbo_code, deIms.rsSOLDTO, "slt_code", True)

    Caption = Caption + " - " + Tag
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.number, True)
''''''''''

    
    
    
    
    
    
    
    
    'Screen.MousePointer = vbHourglass
    'Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    'For Each ctl In Controls
     '   Call gsb_fade_to_black(ctl)
    'Next ctl
    
    'Call DisableButtons(Me, NavBar1)
    'Call deIms.SoldTo(deIms.NameSpace)
    ''Call PopuLateFromRecordSet(cbo_code, deIms.rsSOLDTO, "slt_code", True)
    
    'Call BindAll(Me, deIms)
    'Set NavBar1.Recordset = deIms.rsSOLDTO
    'Screen.MousePointer = vbDefault
    
    'Caption = Caption + " - " + Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsSOLDTO.Update
    deIms.rsSOLDTO.CancelUpdate
    
    deIms.rsSOLDTO.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    
End Sub

'added by shakir

Private Sub NavBar1_BeforeSaveClick()
NavBar1.AllowUpdate = False
    If Len(Trim$(cbo_code)) = 0 Then
        
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cbo_code.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Name)) = 0 Then
        
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Name.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Address1)) = 0 Then
    
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00004") 'J added
        MsgBox IIf(msg1 = "", "The Address cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Address1.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Country)) = 0 Then
    
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "The Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Country.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    
    NavBar1.AllowUpdate = True
    deIms.rsSOLDTO!slt_modiuser = CurrentUser
    deIms.rsSOLDTO!slt_npecode = deIms.NameSpace
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    
    Unload Me
End Sub

Private Sub NavBar1_OnEditClick()
'FormMode = ChangeModeOfForm(lblStatus, mdModification)
'If FormMode= mdModification Then MakeReadOnly (True)
End Sub


Private Sub NavBar1_OnNewClick()
On Error Resume Next
    
    deIms.rsSOLDTO!slt_creauser = CurrentUser
    deIms.rsSOLDTO!slt_modiuser = CurrentUser
    If Err Then Call LogErr(Name & "::NavBar1_OnNewClick", Err.Description, Err.number, True)

End Sub

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    
    deIms.rsSOLDTO.UpdateBatch
    Call deIms.rsSOLDTO.Move(0)
    
    'Modified by Juan (8/29/2000) for Multilingual
    msg1 = translator.Trans("M00002") 'J added
    If Err Then MsgBox IIf(msg1 = "", "Error Saving changes", msg1): Err.Clear 'J modified
    '---------------------------------------------
    
    
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub

'set back ground color

Private Sub txt_Address1_GotFocus()
    Call HighlightBackground(txt_Address1)
End Sub

'set back ground color

Private Sub txt_Address1_LostFocus()
    Call NormalBackground(txt_Address1)
End Sub

'set back ground color

Private Sub txt_Address2_GotFocus()
    Call HighlightBackground(txt_Address2)
End Sub

'set back ground color

Private Sub txt_Address2_LostFocus()
    Call NormalBackground(txt_Address2)
End Sub

'set back ground color

Private Sub txt_BilltoName_GotFocus()
    Call HighlightBackground(txt_Name)
End Sub

'set back ground color

Private Sub txt_BilltoName_LostFocus()
    Call NormalBackground(txt_Name)
End Sub

'set back ground color

Private Sub txt_City_GotFocus()
    Call HighlightBackground(txt_City)
End Sub

'set back ground color

Private Sub txt_City_LostFocus()
    Call NormalBackground(txt_City)
End Sub

'set back ground color

Private Sub txt_Contact_GotFocus()
    Call HighlightBackground(txt_Contact)
End Sub

'set back ground color

Private Sub txt_Contact_LostFocus()
    Call NormalBackground(txt_Contact)
End Sub

'set back ground color

Private Sub txt_Country_GotFocus()
    Call HighlightBackground(txt_Country)
End Sub

'set back ground color

Private Sub txt_Country_LostFocus()
    Call NormalBackground(txt_Country)
End Sub

'set back ground color

Private Sub txt_Email_GotFocus()
    Call HighlightBackground(txt_Email)
    txt_Email.MaxLength = 50
End Sub

'set back ground color

Private Sub txt_Email_LostFocus()
    Call NormalBackground(txt_Email)
End Sub

'set back ground color

Private Sub txt_FaxNumber_GotFocus()
    Call HighlightBackground(txt_FaxNumber)
    txt_FaxNumber.MaxLength = 30
End Sub

'set back ground color
 
Private Sub txt_FaxNumber_LostFocus()
    Call NormalBackground(txt_FaxNumber)
End Sub


'set back ground color

Private Sub txt_Name_GotFocus()
    Call HighlightBackground(txt_Name)
    txt_Name.MaxLength = 35
End Sub

'set back ground color

Private Sub txt_Name_LostFocus()
    Call NormalBackground(txt_Name)
End Sub

'set back ground color

Private Sub txt_PhoneNumber_GotFocus()
    Call HighlightBackground(txt_PhoneNumber)
End Sub

'set back ground color

Private Sub txt_PhoneNumber_LostFocus()
    Call NormalBackground(txt_PhoneNumber)
End Sub

'set back ground color

Private Sub txt_State_GotFocus()
    Call HighlightBackground(txt_State)
End Sub

'set back ground color

Private Sub txt_State_LostFocus()
    Call NormalBackground(txt_State)
End Sub

'set back ground color

Private Sub txt_Zipcode_GotFocus()
    Call HighlightBackground(txt_Zipcode)
End Sub

'set back ground color

Private Sub txt_Zipcode_LostFocus()
    Call NormalBackground(txt_Zipcode)
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handled

Dim retval As PrintOpts

    Load frmPrintDialog
    With frmPrintDialog
        .Show 1
        retval = .Result
        
        'Modified by Juan (9/14/2000) for Multilingual (only checked)
        msg1 = translator.Trans("M00109") 'J added
        DoEvents: DoEvents
        If retval = poPrintCurrent Then
            With MDI_IMS.CrystalReport1
                .Reset
                .ReportFileName = FixDir(App.Path) & "CRreports\Soldto.rpt"
                .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
                .ParameterFields(1) = "soldtocode;" & cbo_code & ";TRUE"
                .WindowTitle = IIf(msg1 = "", "Sold To:", msg1 + ":") 'J modified
                Call translator.Translate_Reports("Soldto.rpt") 'J added
                .Action = 1: .Reset
            End With
            
        ElseIf retval = poPrintAll Then
            With MDI_IMS.CrystalReport1
                .Reset
                .ReportFileName = FixDir(App.Path) & "CRreports\Soldto.rpt"
                .ParameterFields(1) = "soldtocode;ALL;TRUE"
                .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
                .WindowTitle = IIf(msg1 = "", "Sold To:", msg1 + ":") 'J modified
                Call translator.Translate_Reports("Soldto.rpt") 'J added
                .Action = 1: .Reset
            End With
            
        Else
            Exit Sub
        
        End If
        '-------------------------------------------------------------
        
    End With
    Unload frmPrintDialog
    Set frmPrintDialog = Nothing

Handled:
    If Err Then MsgBox Err.Description
End Sub
