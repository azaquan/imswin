VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Billto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill to"
   ClientHeight    =   4395
   ClientLeft      =   750
   ClientTop       =   1065
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   5220
   Tag             =   "01020100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   870
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3900
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "blt_zipc"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2040
      Width           =   1350
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "blt_adr1"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1140
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "blt_adr2"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1440
      Width           =   3024
   End
   Begin VB.TextBox txt_City 
      DataField       =   "blt_city"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1740
      Width           =   3024
   End
   Begin VB.TextBox txt_State 
      DataField       =   "blt_stat"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2040
      Width           =   525
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "blt_faxnumb"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2940
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "blt_mail"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2040
      MaxLength       =   59
      TabIndex        =   10
      Top             =   3240
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "blt_cont"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   11
      Top             =   3540
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "blt_phonnumb"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2640
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "blt_ctry"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2340
      Width           =   3024
   End
   Begin VB.TextBox txt_BilltoName 
      DataField       =   "blt_name"
      DataMember      =   "BILLTO"
      Height          =   288
      Left            =   2055
      TabIndex        =   1
      Top             =   840
      Width           =   3024
   End
   Begin VB.ComboBox cbo_BilltoCode 
      DataField       =   "blt_code"
      DataMember      =   "BILLTO"
      Height          =   315
      Left            =   2055
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   3024
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
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   2000
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
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   285
      Left            =   2595
      TabIndex        =   19
      Top             =   2085
      Width           =   1125
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
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   288
      Left            =   120
      TabIndex        =   20
      Top             =   2340
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
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   288
      Left            =   120
      TabIndex        =   24
      Top             =   3540
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
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   288
      Left            =   120
      TabIndex        =   22
      Top             =   2940
      Width           =   2000
   End
   Begin VB.Label lbl_InterSupp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   4980
   End
   Begin VB.Label lbl_Billto_Code 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to: Code"
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
      Top             =   540
      Width           =   2000
   End
   Begin VB.Label lbl_Billto_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to: Name"
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
      Top             =   840
      Width           =   2000
   End
End
Attribute VB_Name = "frm_Billto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat
'get billto record set to combo bill to code
Private Sub cbo_BilltoCode_Click()
On Error Resume Next
Dim str As String
    
    str = cbo_BilltoCode
    NavBar1.CancelUpdate
    deIms.rsBILLTO.CancelUpdate
    deIms.rsBILLTO.CancelUpdate
    Call deIms.rsBILLTO.Resync(adAffectCurrent, adResyncAllValues)
    Call RecordsetFind(deIms.rsBILLTO, "blt_code = '" & cbo_BilltoCode & "'")
    
    If Err Then Call LogErr(Name & "::cbo_BilltoCode_Click", Err.Description, Err.number, True)
End Sub
'drop down bill to combo
Private Sub cbo_BilltoCode_DragDrop(Source As Control, x As Single, Y As Single)
  cbo_BilltoCode.locked = False
End Sub
'if bill to combo got focus, set background color
Private Sub cbo_BilltoCode_GotFocus()
    Call HighlightBackground(cbo_BilltoCode)
End Sub
'Set bill to combo allow enter from keyboard
Private Sub cbo_BilltoCode_KeyPress(KeyAscii As Integer)
    If NavBar1.NewEnabled = False Then
        KeyAscii = 0
    End If
End Sub
'if bill to combo lost focus, set background color
Private Sub cbo_BilltoCode_LostFocus()
    Call NormalBackground(cbo_BilltoCode)
End Sub

Private Sub cbo_BilltoCode_Validate(Cancel As Boolean)
If Len(cbo_BilltoCode) > 10 Then
MsgBox "Code number can not be greater than 10 characters."
Cancel = True
cbo_BilltoCode.SetFocus
cbo_BilltoCode = "EXXONEFTE"
End If
End Sub

'unload bil to form
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Hide
    deIms.rsBILLTO.Close
    Set frm_Billto = Nothing
    If open_forms <= 5 Then ShowNavigator
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
End Sub

Private Sub lbl_Address1_Click()

End Sub

Private Sub lbl_InterSupp_Click()

End Sub

'before save record to bill to table, validate each data field
Private Sub NavBar1_BeforeSaveClick()

    NavBar1.AllowUpdate = False
    
    If Len(Trim$(cbo_BilltoCode)) = 0 Then
        
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cbo_BilltoCode.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_BilltoName)) = 0 Then
        
        'Modified by Juan (8/29/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_BilltoName.SetFocus: Exit Sub
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
    deIms.rsBILLTO!blt_modiuser = CurrentUser
    deIms.rsBILLTO!blt_npecode = deIms.NameSpace
 

    
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
'load form get record for bill to form, and disable buttons
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

    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Billto")
    '------------------------------------------
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.BILLTO(deIms.NameSpace)
    
    Call BindAll(Me, deIms)
    Call DisableButtons(Me, NavBar1)
    Set NavBar1.Recordset = deIms.rsBILLTO
    Call PopuLateFromRecordSet(cbo_BilltoCode, deIms.rsBILLTO, "blt_code", True)

    Caption = Caption + " - " + Tag
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.number, True)
    With frm_Billto
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub
'add new a record to bill to table, set create user to current user
Private Sub NavBar1_OnNewClick()
On Error Resume Next
    deIms.rsBILLTO!blt_creauser = CurrentUser
    deIms.rsBILLTO!blt_modiuser = CurrentUser
    If Err Then Call LogErr(Name & "::NavBar1_OnNewClick", Err.Description, Err.number, True)
End Sub
'save a record to bill to table
Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    
    deIms.rsBILLTO.UpdateBatch
    Call deIms.rsBILLTO.Move(0)
    
    'Modified by Juan (8/29/2000) for Multilingual
    msg1 = translator.Trans("M00002") 'J added
    If Err Then MsgBox IIf(msg1 = "", "Error Saving changes", msg1): Err.Clear 'J modified
    '---------------------------------------------
    
    
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub
'set back ground color to this field
Private Sub txt_Address1_GotFocus()
    Call HighlightBackground(txt_Address1)
End Sub
'set back ground color to this field
Private Sub txt_Address1_LostFocus()
    Call NormalBackground(txt_Address1)
End Sub
'set back ground color to this field
Private Sub txt_Address2_GotFocus()
    Call HighlightBackground(txt_Address2)
End Sub
'set back ground color to this field
Private Sub txt_Address2_LostFocus()
    Call NormalBackground(txt_Address2)
End Sub
'set back ground color to this field
Private Sub txt_BilltoName_GotFocus()
    Call HighlightBackground(txt_BilltoName)
   txt_BilltoName.MaxLength = 35
End Sub

'set back ground color to this field
Private Sub txt_BilltoName_LostFocus()
    Call NormalBackground(txt_BilltoName)
End Sub
'set back ground color to this field
Private Sub txt_City_GotFocus()
    Call HighlightBackground(txt_City)
End Sub
'set back ground color to this field
Private Sub txt_City_LostFocus()
    Call NormalBackground(txt_City)
End Sub
'set back ground color to this field
Private Sub txt_Contact_GotFocus()
    Call HighlightBackground(txt_Contact)
End Sub
'set back ground color to this field
Private Sub txt_Contact_LostFocus()
    Call NormalBackground(txt_Contact)
End Sub
'set back ground color to this field
Private Sub txt_Country_GotFocus()
    Call HighlightBackground(txt_Country)
End Sub
'set back ground color to this field
Private Sub txt_Country_LostFocus()
    Call NormalBackground(txt_Country)
End Sub
'set back ground color to this field
Private Sub txt_Email_GotFocus()
    Call HighlightBackground(txt_Email)
    txt_Email.MaxLength = 40
End Sub
'set back ground color to this field
Private Sub txt_Email_LostFocus()
    Call NormalBackground(txt_Email)
End Sub
'set back ground color to this field
Private Sub txt_FaxNumber_GotFocus()
    Call HighlightBackground(txt_FaxNumber)
    txt_FaxNumber.MaxLength = 30
End Sub
'set back ground color to this field
Private Sub txt_FaxNumber_LostFocus()
    Call NormalBackground(txt_FaxNumber)
End Sub
'set back ground color to this field
Private Sub txt_PhoneNumber_GotFocus()
    Call HighlightBackground(txt_PhoneNumber)
End Sub
'set back ground color to this field
Private Sub txt_PhoneNumber_LostFocus()
    Call NormalBackground(txt_PhoneNumber)
End Sub
'set back ground color to this field
Private Sub txt_State_GotFocus()
    Call HighlightBackground(txt_State)
End Sub
'set back ground color to this field
Private Sub txt_State_LostFocus()
    Call NormalBackground(txt_State)
End Sub
'set back ground color to this field
Private Sub txt_Zipcode_GotFocus()
    Call HighlightBackground(txt_Zipcode)
End Sub
'set back ground color to this field
Private Sub txt_Zipcode_LostFocus()
    Call NormalBackground(txt_Zipcode)
End Sub
'print crystal report
Private Sub NavBar1_OnPrintClick()
On Error GoTo Handled

'Dim retval As PrintOpts

    Load frmPrintDialog
    With frmPrintDialog
        
        .Show 1

        DoEvents: DoEvents
        If .Result = poPrintCurrent Then
            Call BeforePrint(poPrintCurrent)

        ElseIf .Result = poPrintAll Then
            Call BeforePrint(poPrintAll)

        Else
            Exit Sub

        End If

    End With
    
    Unload frmPrintDialog
    Set frmPrintDialog = Nothing
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset
    
Handled:
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)
    End If
End Sub
'function to set path to print crystal report
Public Sub BeforePrint(iOption As PrintOpts)
On Error Resume Next

Dim Path As String
On Error GoTo ErrHandler

    Path = FixDir(App.Path) + "CRreports\"
    
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = Path & "Billto.rpt"
        If iOption = poPrintCurrent Then
            .ParameterFields(1) = "billtocode;" & cbo_BilltoCode & ";TRUE"
        Else
            .ParameterFields(1) = "billtocode;ALL;TRUE"
        End If
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("M00107") 'J added
        .WindowTitle = IIf(msg1 = "", "Bill to", msg1) 'J modified
        Call translator.Translate_Reports("Billto.rpt") 'J added
        '---------------------------------------------
        
    End With
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::BeforePrint", Err.Description, Err.number, True)
    End If
End Sub
