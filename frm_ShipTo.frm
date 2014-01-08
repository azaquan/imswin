VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_ShipTo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ship To:"
   ClientHeight    =   5640
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   5370
   Tag             =   "01020200"
   Begin VB.TextBox Txtphone 
      DataField       =   "sht_mail"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   59
      TabIndex        =   13
      Top             =   3840
      Width           =   3024
   End
   Begin VB.TextBox Txtaddress2 
      DataField       =   "sht_adr1"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1740
      Width           =   3024
   End
   Begin VB.TextBox Txtaddress2line 
      DataField       =   "sht_adr2"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   6
      Top             =   2040
      Width           =   3024
   End
   Begin VB.CheckBox chkflag 
      Caption         =   "Check1"
      DataField       =   "sht_actvflag"
      DataMember      =   "shipto"
      DataSource      =   "deIms"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   4800
      Width           =   255
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   5100
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
   Begin VB.TextBox txt_ShipToName 
      DataField       =   "sht_name"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   35
      TabIndex        =   2
      Top             =   840
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "sht_ctry"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   10
      Top             =   2940
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "sht_phonnumb"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   11
      Top             =   3240
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "sht_cont"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   15
      Top             =   4440
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "sht_mail"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   59
      TabIndex        =   14
      Top             =   4140
      Width           =   3024
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "sht_faxnumb"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3540
      Width           =   3024
   End
   Begin VB.TextBox txt_State 
      DataField       =   "sht_stat"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2640
      Width           =   408
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "sht_zipc"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   3840
      MaxLength       =   11
      TabIndex        =   9
      Top             =   2640
      Width           =   1350
   End
   Begin VB.TextBox txt_City 
      DataField       =   "sht_city"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2340
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "sht_adr2"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1440
      Width           =   3024
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "sht_adr1"
      DataMember      =   "SHIPTO"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1140
      Width           =   3024
   End
   Begin VB.ComboBox cbo_ShipToCode 
      Height          =   315
      Left            =   2175
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   3024
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Telex #"
      Height          =   285
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Width           =   2000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address2"
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   1740
      Width           =   2000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   2000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Flag"
      DataField       =   "shi_actvflag"
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   4800
      Width           =   2000
   End
   Begin VB.Label lbl_ShipToName 
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
      TabIndex        =   18
      Top             =   840
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   3540
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   285
      Left            =   120
      TabIndex        =   27
      Top             =   4140
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   2940
      Width           =   2000
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   285
      Left            =   2715
      TabIndex        =   23
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   2340
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address1"
      Height          =   288
      Left            =   120
      TabIndex        =   19
      Top             =   1140
      Width           =   2000
   End
   Begin VB.Label lbl_ShipTo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ship To:"
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
      Left            =   180
      TabIndex        =   0
      Top             =   45
      Width           =   4950
   End
   Begin VB.Label lbl_Code 
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
      TabIndex        =   17
      Top             =   540
      Width           =   2000
   End
End
Attribute VB_Name = "frm_ShipTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String 'jawdat

'load recordset

Private Sub cbo_ShipToCode_Click()
Dim str As String
    
    'kin add function

    If cbo_ShipToCode.ListIndex < 0 Then 'If cbo_ShipToCode.ListIndex <= 0 Then
        Call claersrceen
        Exit Sub
    Else
        Call GetshiptoListRecord(cbo_ShipToCode, deIms.NameSpace)
    End If
    
        Call EnableButtons
    

cbo_ShipToCode.locked = True   'M
End Sub

'unlock shipper code combo

Private Sub cbo_ShipToCode_DropDown()
    cbo_ShipToCode.locked = False
End Sub

'set back ground color

Private Sub cbo_ShipToCode_GotFocus()
    Call HighlightBackground(cbo_ShipToCode)
End Sub

'do not allow add new character

Private Sub cbo_ShipToCode_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then
KeyAscii = 0
End If
End Sub

'set back ground color

Private Sub cbo_ShipToCode_LostFocus()
    Call NormalBackground(cbo_ShipToCode)
End Sub

Private Sub cbo_ShipToCode_Validate(Cancel As Boolean)
If Len(cbo_ShipToCode) > 10 Then
MsgBox "Code number can not be greater than 10 characters."
Cancel = True
cbo_ShipToCode.SetFocus
End If
End Sub

'get recordset and populate combo and set buttom

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
chkflag.Enabled = False
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
    Call translator.Translate_Forms("frm_ShipTo")
    '------------------------------------------
    
     Call PopuLateFromRecordSet(cbo_ShipToCode, GetshipCode(deIms.NameSpace, deIms.cnIms), "sht_code", True)
'    Set rstcomp = GetCompanyCode(deIms.NameSpace, deIms.cnIms)

    If cbo_ShipToCode.ListCount > 0 Then cbo_ShipToCode.ListIndex = 0
    
    frm_ShipTo.Caption = frm_ShipTo.Caption + " - " + frm_ShipTo.Tag
        
    cbo_ShipToCode.locked = True  'M
    txt_ShipToName.MaxLength = 35
    
    Call DisableButtons(Me, NavBar1)
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form and close recordset

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsSHIPTO.Update
    deIms.rsSHIPTO.UpdateBatch
    deIms.rsSHIPTO.CancelBatch
    
    deIms.rsSHIPTO.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    
End Sub

'set create user and modify user equal to current user name

Private Sub NavBar1_BeforeSaveClick()
Dim num As Integer
Dim list As Integer

    list = cbo_ShipToCode.ListIndex
    'kin add function to check shipto code
    If list <> -1 Then
        If Len(Trim$(cbo_ShipToCode)) <> 0 Then
            If Checkshipcode(cbo_ShipToCode) <> cbo_ShipToCode Then
            
                'Modified by Juan (9/14/2000) for Multilingual
                msg1 = translator.Trans("M00313") 'J added
                MsgBox IIf(msg1 = "", " You can not change Ship To code, Please make new one", msg1) 'J modified
                '---------------------------------------------

                Exit Sub
            Else
                If Validateshiptodata Then InsertintoShipto
            End If
        End If

    End If


     If list = -1 Then
        If Len(Trim$(cbo_ShipToCode)) <> 0 Then
            If Countshiptocode(cbo_ShipToCode) Then
              
                'Modified by Juan (9/14/2000) for Multilingual
                msg1 = translator.Trans("M00314") 'J added
                MsgBox IIf(msg1 = "", "Ship To code exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            Else
                If Validateshiptodata Then InsertintoShipto
            End If
        End If
     End If
End Sub

Private Sub NavBar1_OnCancelClick()
cbo_ShipToCode.locked = True
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

Private Sub NavBar1_OnFirstClick()
    cbo_ShipToCode.ListIndex = 0
End Sub

Private Sub NavBar1_OnLastClick()
     cbo_ShipToCode.ListIndex = cbo_ShipToCode.ListCount - 1
End Sub

'set name space equal to current name space

Private Sub NavBar1_OnNewClick()
    
    cbo_ShipToCode.locked = False   'M
    Call claersrceen
'    deIms.rsSHIPTO!sht_npecode = deIms.NameSpace
End Sub

'kin add function to move recordset

Private Sub NavBar1_OnNextClick()
    If (cbo_ShipToCode.ListIndex + 1) = cbo_ShipToCode.ListCount Then
        Exit Sub
    Else
    cbo_ShipToCode.ListIndex = cbo_ShipToCode.ListIndex + 1
    End If
End Sub

'kin add function to move recordset

Private Sub NavBar1_OnPreviousClick()
     If cbo_ShipToCode.ListIndex = 0 Or cbo_ShipToCode.ListIndex = -1 Then
        Exit Sub
    Else
        cbo_ShipToCode.ListIndex = cbo_ShipToCode.ListIndex - 1
    End If
End Sub

'get crystal report parameters and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handled

Dim retval As PrintOpts

    Load frmPrintDialog
    With frmPrintDialog
        .Show 1
        retval = .Result
        
        'Modified by Juan (9/14/2000) for Multilingual (only checked)
        msg1 = translator.Trans("L00088") 'J added
        DoEvents: DoEvents
        If retval = poPrintCurrent Then
            With MDI_IMS.CrystalReport1
                .Reset
                .ReportFileName = FixDir(App.Path) & "CRreports\Shipto.rpt"
                .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
                .ParameterFields(1) = "shiptocode;" & cbo_ShipToCode & ";TRUE"
                .WindowTitle = IIf(msg1 = "", "Ship to", msg1) 'J modified
                Call translator.Translate_Reports("Shipto.rpt") 'J added
                .Action = 1: .Reset
            End With
            
        ElseIf retval = poPrintAll Then
            With MDI_IMS.CrystalReport1
                .Reset
                .ReportFileName = FixDir(App.Path) & "CRreports\Shipto.rpt"
                .ParameterFields(1) = "shiptocode;ALL;TRUE"
                .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
                .WindowTitle = IIf(msg1 = "", "Ship to", msg1) 'J modified
                Call translator.Translate_Reports("Shipto.rpt") 'J added
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

Private Sub NavBar1_OnSaveClick()
cbo_ShipToCode.locked = True
'On Error Resume Next

'    If Validateshiptodata Then InsertintoShipto
    
'    deIms.rsSHIPTO.UpdateBatch
'    Call deIms.rsSHIPTO.Move(0)
'    If IndexOf(cbo_ShipToCode, cbo_ShipToCode) = CB_ERR Then _
'        cbo_ShipToCode.AddItem (cbo_ShipToCode)
'
'    Call deIms.rsSHIPTO.Resync(adAffectCurrent, adResyncAllValues)

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
    Call HighlightBackground(txt_ShipToName)
End Sub

'set back ground color

Private Sub txt_BilltoName_LostFocus()
    Call NormalBackground(txt_ShipToName)
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
End Sub

'set back ground color

Private Sub txt_Email_LostFocus()
    Call NormalBackground(txt_Email)
End Sub

'set back ground color

Private Sub txt_FaxNumber_GotFocus()
    Call HighlightBackground(txt_FaxNumber)
End Sub

'set back ground color

Private Sub txt_FaxNumber_LostFocus()
    Call NormalBackground(txt_FaxNumber)
End Sub

'set back ground color

Private Sub txt_Name_GotFocus()
    Call HighlightBackground(txt_ShipToName)
End Sub

'set back ground color

Private Sub txt_Name_LostFocus()
    Call NormalBackground(txt_ShipToName)
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

Private Sub txt_ShipToName_GotFocus()
    Call HighlightBackground(txt_ShipToName)
End Sub

'set back ground color

Private Sub txt_ShipToName_LostFocus()
    Call NormalBackground(txt_ShipToName)
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

'kin add function to insert record  to shipto table

Private Sub InsertintoShipto()
On Error Resume Next
Dim cmd As ADODB.Command
Dim v As Variant
On Error GoTo Noinsert

     Set cmd = New ADODB.Command
  
    With cmd
        .CommandText = "UP_INS_SHIPTO"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, cbo_ShipToCode)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 35, txt_ShipToName)
        .parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, txt_Address1)
        
        v = txt_Address2
        If Len(Trim$(txt_Address2)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, v)
        
         v = Txtaddress2
        If Len(Trim$(Txtaddress2)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@adr3", adVarChar, adParamInput, 25, v)
        
         v = Txtaddress2line
        If Len(Trim$(Txtaddress2line)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@adr4", adVarChar, adParamInput, 25, v)
         
        .parameters.Append .CreateParameter("@city", adVarChar, adParamInput, 25, txt_City)
        
        v = txt_State
        If Len(Trim$(txt_State)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@stat", adVarChar, adParamInput, 2, v)
        
        v = txt_Zipcode
        If Len(Trim$(txt_Zipcode)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@zipc", adVarChar, adParamInput, 11, v)
        
        .parameters.Append .CreateParameter("@ctry", adVarChar, adParamInput, 25, txt_Country)
        .parameters.Append .CreateParameter("@phonnumb", adVarChar, adParamInput, 25, txt_PhoneNumber)
        
        v = txt_FaxNumber
        If Len(Trim$(txt_FaxNumber)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@faxnumb", adVarChar, adParamInput, 50, v)
        
        v = Txtphone
        If Len(Trim$(Txtphone)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, v)
        
        v = txt_Email
        If Len(Trim$(txt_Email)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, v)
        
         v = txt_Contact
        If Len(Trim$(txt_Contact)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, v)
        .parameters.Append .CreateParameter("@CreatedBY", adVarChar, adParamInput, 20, CurrentUser)
        .parameters.Append .CreateParameter("@flag", adBoolean, adParamInput, , chkflag.value = vbChecked)
        
'        .Parameters.Append .CreateParameter("@modiuser", adVarChar, adParamInput, 20, CurrentUser)
      
        .Execute , , adExecuteNoRecords
    
    End With
    
    If IndexOf(cbo_ShipToCode, cbo_ShipToCode) = CB_ERR Then
        cbo_ShipToCode.AddItem (cbo_ShipToCode)
    End If
    
    
       Set cmd = Nothing
       
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00315") 'J added
        MsgBox IIf(msg1 = "", "Insert into Ship To is completed succesfully", msg1) 'J modified
        '--------------------------------------------
        
    Exit Sub
Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00316") 'J added
        MsgBox IIf(msg1 = "", "Insert into Ship To is failure", msg1) 'J modified
        '---------------------------------------------
        
End Sub

'kin add function to validate data

Public Function Validateshiptodata() As Boolean
On Error Resume Next
    
    Validateshiptodata = False
    
    If Len(Trim$(cbo_ShipToCode)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00316") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cbo_ShipToCode.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_ShipToName)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_ShipToName.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_Address1)) = 0 Then
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00004") 'J added
        MsgBox IIf(msg1 = "", "The Address cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Address1.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_Country)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox "The Country cannot be left empty"
        '---------------------------------------------
        
        txt_Country.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Function
    End If
    
     Validateshiptodata = True
End Function


'SQL statement to get shipto name
'kin add function to get shipto code

Public Function GetshipCode(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select sht_code from shipto"
        .CommandText = .CommandText & " where sht_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " ORDER BY sht_code"
        
         Set GetshipCode = .Execute
    End With
    
End Function

'kin add function to get shipto record list

Public Function GetshiptoListRecord(Code As String, NameSpace As String)
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT sht_code,sht_name, sht_adr1, "
        .CommandText = .CommandText & " sht_adr2, sht_adr3,sht_adr4, sht_city, "
        .CommandText = .CommandText & " sht_stat, sht_zipc, "
        .CommandText = .CommandText & " sht_ctry, sht_phonnumb, "
        .CommandText = .CommandText & " sht_faxnumb, sht_mail, "
        .CommandText = .CommandText & " sht_telxnumb, sht_cont, sht_actvflag "
        .CommandText = .CommandText & " From SHIPTO "
        .CommandText = .CommandText & " WHERE sht_code =  '" & Code & "'"
        .CommandText = .CommandText & " AND sht_npecode = '" & NameSpace & "'"
        Set rst = .Execute
    End With
        cbo_ShipToCode = rst!sht_code & ""
        txt_ShipToName = rst!sht_name & ""
        txt_Address1 = rst!sht_adr1 & ""
        txt_Address2 = rst!sht_adr2 & ""
        Txtaddress2 = rst!sht_adr3 & ""
        Txtaddress2line = rst!sht_adr4 & ""
        txt_City = rst!sht_city & ""
        txt_State = rst!sht_stat & ""
        txt_Zipcode = rst!sht_zipc & ""
        txt_Country = rst!sht_ctry & ""
        txt_PhoneNumber = rst!sht_phonnumb & ""
        txt_FaxNumber = rst!sht_faxnumb & ""
        Txtphone = rst!sht_telxnumb & ""
        txt_Email = rst!sht_mail & ""
        txt_Contact = rst!sht_cont & ""
        chkflag.value = IIf(rst!sht_actvflag, 1, 0)
    
End Function

'kin add function to clear srceem

Public Sub claersrceen()

        cbo_ShipToCode.ListIndex = -1
        txt_ShipToName = ""
        txt_Address1 = ""
        txt_Address2 = ""
        Txtaddress2 = ""
        Txtaddress2line = ""
        txt_City = ""
        txt_State = ""
        txt_Zipcode = ""
        txt_Country = ""
        txt_PhoneNumber = ""
        txt_FaxNumber = ""
        Txtphone = ""
        txt_Email = ""
        txt_Contact = ""
'        chkflag.Value = chkflag.

End Sub


'SQL statement to check shipto code exist or not
'kin add function to check shipto code


Private Function Checkshipcode(Code As String) As String
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT sht_code"
        .CommandText = .CommandText & " From shipto "
        .CommandText = .CommandText & " Where sht_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sht_code = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        Checkshipcode = rst!sht_code
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::Checkshipcode", Err.Description, Err.number, True)
End Function


'SQL statement to check shipto code exist or not
'kin add function to check shipto code

Private Function Countshiptocode(Code As String) As String
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) rt"
        .CommandText = .CommandText & " From shipto "
        .CommandText = .CommandText & " Where sht_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sht_code = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        Countshiptocode = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::Countshiptocode", Err.Description, Err.number, True)
End Function

'kin add function
'set navbar button

Public Sub EnableButtons()
Dim i As Integer

    i = cbo_ShipToCode.ListIndex

    If cbo_ShipToCode.ListCount = 0 Then

        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False

        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False

        Exit Sub

    ElseIf i = cbo_ShipToCode.ListCount - 1 Then
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False

        NavBar1.FirstEnabled = True
        NavBar1.PreviousEnabled = True

    ElseIf i = 0 Then

        NavBar1.LastEnabled = True
        NavBar1.NextEnabled = True

        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
    Else

        NavBar1.LastEnabled = True
        NavBar1.NextEnabled = True

        NavBar1.FirstEnabled = True
        NavBar1.PreviousEnabled = True

    End If

    If cbo_ShipToCode.ListIndex = CB_ERR Then cbo_ShipToCode.ListIndex = 0
    If Err Then Err.Clear
End Sub

