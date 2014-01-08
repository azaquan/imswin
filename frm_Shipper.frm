VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Shipper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipper"
   ClientHeight    =   4815
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5355
   Tag             =   "01010200"
   Begin VB.CheckBox chkflag 
      Caption         =   "Check1"
      DataField       =   "shi_actvflag"
      DataMember      =   "shipper"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   3960
      Width           =   255
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   4380
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DisableSaveOnSave=   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "shi_adr1"
      DataMember      =   "SHIPPER"
      Height          =   315
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1230
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "shi_adr2"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1560
      Width           =   3024
   End
   Begin VB.TextBox txt_City 
      DataField       =   "shi_city"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1860
      Width           =   3024
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "shi_zipc"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   3840
      MaxLength       =   11
      TabIndex        =   5
      Top             =   2160
      Width           =   1370
   End
   Begin VB.TextBox txt_State 
      DataField       =   "shi_stat"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2160
      Width           =   408
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "shi_faxnumb"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3060
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "shi_mail"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   59
      TabIndex        =   9
      Top             =   3360
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "shi_cont"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   10
      Top             =   3660
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "shi_phonnumb"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2760
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "shi_ctry"
      DataMember      =   "SHIPPER"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   6
      Top             =   2460
      Width           =   3024
   End
   Begin VB.TextBox txt_ShipperName 
      DataField       =   "shi_name"
      DataMember      =   "SHIPPER"
      Height          =   315
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   0
      Top             =   900
      Width           =   3024
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo dcboShipCode 
      Height          =   315
      Left            =   2175
      TabIndex        =   27
      Top             =   570
      Width           =   3024
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   5334
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Flag"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   288
      Left            =   120
      TabIndex        =   14
      Top             =   1260
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   2000
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   288
      Left            =   120
      TabIndex        =   16
      Top             =   1860
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   195
      Left            =   2715
      TabIndex        =   18
      Top             =   2190
      Width           =   1125
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   288
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   288
      Left            =   120
      TabIndex        =   19
      Top             =   2460
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   288
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   288
      Left            =   120
      TabIndex        =   23
      Top             =   3660
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   288
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   288
      Left            =   120
      TabIndex        =   21
      Top             =   3060
      Width           =   2000
   End
   Begin VB.Label lbl_Shipper 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shipper"
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
      TabIndex        =   11
      Top             =   120
      Width           =   4965
   End
   Begin VB.Label lbl_Sup_Code 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipper Code"
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
      TabIndex        =   12
      Top             =   660
      Width           =   2000
   End
   Begin VB.Label lbl_Sup_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipper's Name"
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
      Top             =   960
      Width           =   2000
   End
End
Attribute VB_Name = "frm_Shipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsShipperCode As ADODB.Recordset
'set back ground color
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Sub cbo_ShipperCode_GotFocus()
    Call HighlightBackground(dcboShipCode)
End Sub

'set back ground color

Private Sub cbo_ShipperCode_LostFocus()
    Call HighlightBackground(dcboShipCode)
End Sub

'Private Sub dcboShipCode_Click(Area As Integer)
'On Error Resume Next
'Dim rs As ADODB.Recordset
'
'    If Area = 2 Then
'            Set rs = deIms.rsSHIPPER
'
'        rs.CancelUpdate
'        Call rs.CancelBatch(adAffectCurrent)
'
'        If Err Then Err.Clear
'        Call rs.Find("shi_code = '" & dcboShipCode & "'", 0, adSearchForward, adBookmarkFirst)
'    End If
'
'    If Err Then Err.Clear
'End Sub

'get shipper information and fill combo

Private Sub dcboShipCode_Click()
''''Dim code As String
''''On Error Resume Next
''''
''''        'code = Trim$(dcboShipCode)
''''        Call ClearScreen
''''        'dcboShipCode.text = code
''''        If Not Len(Trim(dcboShipCode)) Then
''''            Call Getshipperinfo(dcboShipCode)
''''        End If
''''
''''        If Err Then Err.Clear
'''' End Sub

If Not RsShipperCode.EditMode = 2 Then
Dim str As String

    If Len(Trim$(dcboShipCode)) > 0 And dcboShipCode.IsItemInList Then
                RsShipperCode.MoveFirst
                RsShipperCode.Find "shi_code='" & dcboShipCode & "'", , adSearchForward
                
                If Not RsShipperCode.AbsolutePosition = adPosEOF Then
                   FillTextBox
                End If
        'Set plist = plist.GetPhoneDirectorylist(CmbPhoneDir, deIms.NameSpace, deIms.cnIms)
      '  EnableButtons
    End If
End If
End Sub


Private Sub dcboShipCode_DropDown()
If RsShipperCode.EditMode = 2 Then dcboShipCode.DroppedDown = False
End Sub

Private Sub dcboShipCode_GotFocus()
Call HighlightBackground(dcboShipCode)
End Sub

'set shipper combo input character size

Private Sub dcboShipCode_KeyPress(KeyAscii As Integer)
    If Len(dcboShipCode) = 10 Then
        If KeyAscii >= vbKeySpace Then KeyAscii = 0
    End If
End Sub

Private Sub dcboShipCode_LostFocus()
Call NormalBackground(dcboShipCode)
End Sub

'unlock shipper combo

Private Sub dcboShipCode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' dcboShipCode.locked = False
End Sub

'if navbar not equal add new then lock shipper code combo

Private Sub dcboShipCode_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  'dcboShipCode.locked = Not NavBar1.NewEnabled
End Sub

Private Sub dcboShipCode_Validate(Cancel As Boolean)
If RsShipperCode.EditMode <> 2 And Not dcboShipCode.IsItemInList Then
   MsgBox "The code does not exist.Please select an existing one.", vbInformation, "Imsiwn"
   Cancel = True
   dcboShipCode.SetFocus
 End If
 
If RsShipperCode.EditMode = 2 And dcboShipCode.IsItemInList Then
   MsgBox "The code already exists,Please use a new one.", vbInformation, "Imswin"
   Cancel = True
   dcboShipCode.SetFocus
 End If
   
   
   
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsSHIPPER.CancelUpdate
    
    
    NavBar1.Recordset.Close
    Set NavBar1.Recordset = Nothing
    
    If deIms.rsSHIPPER.State = 1 Then deIms.rsSHIPPER.Close
    RsShipperCode.Close
    Set RsShipperCode = Nothing
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator

    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

End Sub

'before save recordset set name space equal to current name space
'user and modify user are equal to current user

Private Sub NavBar1_BeforeSaveClick()
Dim Numb As String

    'deIms.rsSHIPPER!shi_npecode = deIms.NameSpace
    'deIms.rsSHIPPER!shi_creauser = CurrentUser
'    deIms.rsSHIPPER.Update
    'deIms.rsSHIPPER!shi_modiuser = CurrentUser
    

End Sub

Private Sub NavBar1_OnCancelClick()
dcboShipCode = Trim$(dcboShipCode)
   
   
   If RsShipperCode.EditMode = 2 Then
   
   RsShipperCode.CancelUpdate
   If Not RsShipperCode.BOF Then RsShipperCode.MoveFirst
   If Not RsShipperCode.AbsolutePosition = adPosBOF Then
     FillTextBox
   Else
     
     Call ClearScreen
   End If
   
 Else
 
   If Len(dcboShipCode) > 0 And dcboShipCode.IsItemInList Then
        FillTextBox
   End If
  
 End If
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

'move resordset to first position

'''Private Sub NavBar1_OnFirstClick()
'''    dcboShipCode.ListIndex = 0
'''End Sub
'''
''''move recordset to last position
'''
'''Private Sub NavBar1_OnLastClick()
'''    dcboShipCode.ListIndex = dcboShipCode.ListCount - 1
'''End Sub

Private Sub NavBar1_OnFirstClick()
    
    If RsShipperCode.RecordCount > 0 Then
        RsShipperCode.MoveFirst
        FillTextBox
    End If
    
End Sub

'move recordset to lasst position

Private Sub NavBar1_OnLastClick()
    
    If RsShipperCode.RecordCount > 0 Then
        RsShipperCode.MoveLast
        FillTextBox
    End If
    
End Sub

'clear form screen

Private Sub NavBar1_OnNewClick()
   If RsShipperCode.EditMode = 2 Then
      MsgBox "Can not add another record before saving the previous one.", vbInformation, "Imswin"
   Else
   
    RsShipperCode.AddNew
    Call ClearScreen
    chkflag.value = 1
   End If
End Sub

'move recordset to next position

'''Private Sub NavBar1_OnNextClick()
'''    dcboShipCode.ListIndex = dcboShipCode.ListIndex + 1
'''End Sub
'''
''''move recordset to previous position
'''
'''Private Sub NavBar1_OnPreviousClick()
''' If dcboShipCode.ListIndex = -1 Then
'''        Exit Sub
'''    Else
'''        dcboShipCode.ListIndex = dcboShipCode.ListIndex - 1
'''    End If
'''End Sub


Private Sub NavBar1_OnNextClick()
On Error Resume Next

    
       If Not RsShipperCode.EOF Then RsShipperCode.MoveNext
       
       RsShipperCode.CancelUpdate
       If Not RsShipperCode.AbsolutePosition = adPosEOF Then
        FillTextBox
       Else
         RsShipperCode.MoveLast
       End If
    
    
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
     
     If Not RsShipperCode.BOF Then RsShipperCode.MovePrevious
       
       
       If Not RsShipperCode.AbsolutePosition = adPosBOF Then
        FillTextBox
       Else
         RsShipperCode.MoveFirst
       End If
    
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Shipper.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00067") 'J added
        .WindowTitle = IIf(msg1 = "", "Shipper", msg1) 'J modified
        Call translator.Translate_Reports("Shipper.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form set back ground color and populate combo

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
Dim query As String
    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_Shipper")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
       Call gsb_fade_to_black(ctl)
    Next ctl
    
    If (Not CBool(((deIms.rsSHIPPER.State) And (adStateOpen)))) Then
        Call deIms.Shipper(deIms.NameSpace)
    End If
    Set NavBar1.Recordset = deIms.rsSHIPPER
    
    
'    Set dcboShipCode.RowSource = deIms
    
    'Call BindAll(Me, deIms)
    Screen.MousePointer = vbDefault
    
     'M Call PopuLateFromRecordSet(dcboShipCode, GetshipperCode(deIms.NameSpace), "shi_npecode", True)
    
     Set RsShipperCode = GetShipperDetail()
      Call LoadShipperCombo
      
     RsShipperCode.MoveFirst
     dcboShipCode = Trim(RsShipperCode!shi_code)
     Call dcboShipCode_Click
    
    NavBar1.CancelLastSepVisible = False
    NavBar1.LastPrintSepVisible = False
    NavBar1.PrintSaveSepVisible = False
     
     'M If dcboShipCode.ListCount Then dcboShipCode.ListIndex = 0
    Call DisableButtons(Me, NavBar1)
    frm_Shipper.Caption = frm_Shipper.Caption + " - " + frm_Shipper.Tag
    dcboShipCode.Columns(0).Width = 3024
    
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'before save records, validate data formats

Private Sub NavBar1_OnSaveClick()
On Error Resume Next

Dim Numb As Integer
    If Len(Trim$(dcboShipCode)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        dcboShipCode.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_ShipperName)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_ShipperName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Address1)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Address cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Address1.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Country)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "The Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Country.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    
      'Numb = dcboShipCode.ListIndex
     
    If SaveToRecordset = True Then
       
       RsShipperCode.Update
       If Err.number = 0 Then
          MsgBox "Record saved successfully."
          dcboShipCode.AddItem Trim$(dcboShipCode)
       Else
          MsgBox "Errors Occurred while saving record." & vbCrLf & "Error description -- " & Err.Description
          
       End If
    
    End If
     
     
''''''    If Numb = -1 Then
''''''        If Len(Trim$(dcboShipCode)) <> 0 Then
''''''            If CheckshipperCode(dcboShipCode) Then
''''''
''''''                'Modified by Juan (9/14/2000) for Multilingual
''''''                msg1 = translator.Trans("M00311") 'J added
''''''                MsgBox IIf(msg1 = "", "Shipper Code exist, please make new one", msg1) 'J modified
''''''                '---------------------------------------------
''''''
''''''                Exit Sub
''''''            Else
''''''                deIms.rsSHIPPER!shi_creauser = CurrentUser
''''''                deIms.rsSHIPPER!shi_modiuser = CurrentUser
''''''                deIms.rsSHIPPER!shi_code = dcboShipCode
''''''                deIms.rsSHIPPER!shi_npecode = deIms.NameSpace
''''''                deIms.rsSHIPPER.UpdateBatch
''''''
''''''                'Modified by Juan (9/14/2000) for Multilingual
''''''                msg1 = translator.Trans("M00312") 'J added
''''''                MsgBox IIf(msg1 = "", "Insert into shipper is completed successfully", msg1) 'J modified
''''''                '---------------------------------------------
''''''
''''''            End If
''''''        End If
''''''    End If
''''''
''''''
''''''    If Numb <> -1 Then
''''''        If Len(Trim$(dcboShipCode)) <> 0 Then
''''''            If CheckshipperCodeexit(dcboShipCode) <> dcboShipCode Then
''''''                MsgBox "You can not change Shipper Code, please make new one"
''''''                Exit Sub
''''''            Else
''''''                deIms.rsSHIPPER!shi_code = dcboShipCode
''''''                deIms.rsSHIPPER!shi_creauser = CurrentUser
''''''                deIms.rsSHIPPER!shi_modiuser = CurrentUser
''''''
''''''                deIms.rsSHIPPER!shi_npecode = deIms.NameSpace
''''''                deIms.rsSHIPPER.UpdateBatch
''''''
''''''                'Modified by Juan (9/14/2000) for Multilingual
''''''                msg1 = translator.Trans("M00312") 'J added
''''''                MsgBox IIf(msg1 = "", "Insert into shipper is completed successfully", msg1) 'J modified
''''''                '---------------------------------------------
''''''
''''''            End If
''''''        End If
''''''    End If
''''''
''''''     Call deIms.rsSHIPPER.Move(0)
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

Private Sub txt_ShipperName_GotFocus()
    Call HighlightBackground(txt_ShipperName)
End Sub

'set back ground color

Private Sub txt_ShipperName_LostFocus()
    Call NormalBackground(txt_ShipperName)
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
    Call HighlightBackground(txt_ShipperName)
End Sub

'set back ground color

Private Sub txt_Name_LostFocus()
    Call NormalBackground(txt_ShipperName)
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

Private Sub txt_SupName_GotFocus()
    Call HighlightBackground(txt_ShipperName)
End Sub

'set back ground color

Private Sub txt_SupName_LostFocus()
    Call NormalBackground(txt_ShipperName)
End Sub

'set back ground color

Private Sub txt_Zipcode_GotFocus()
    Call HighlightBackground(txt_Zipcode)
End Sub

'set back ground color

Private Sub txt_Zipcode_LostFocus()
    Call NormalBackground(txt_Zipcode)
End Sub

'SQL statement to get shipper information

Public Sub Getshipperinfo(Code As String)
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT shi_code, shi_npecode, shi_name, "
        .CommandText = .CommandText & " shi_adr1, shi_adr2, shi_city, shi_stat, shi_zipc,"
        .CommandText = .CommandText & " shi_phonnumb, shi_faxnumb, shi_telxnumb, shi_mail,"
        .CommandText = .CommandText & " shi_cont , shi_actvflag"
        .CommandText = .CommandText & " From Shipper"
        .CommandText = .CommandText & " WHERE (shi_npecode = '" & deIms.NameSpace & "') "
        .CommandText = .CommandText & " AND (shi_code = '" & Code & "')"

        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
        txt_ShipperName = rst!shi_name & ""
        txt_Address1 = rst!shi_adr1 & ""
        txt_Address2 = rst!shi_adr2 & ""
        txt_City = rst!shi_city & ""
        txt_State = rst!shi_stat & ""
        txt_Zipcode = rst!shi_zipc
        txt_PhoneNumber = rst!shi_phonnumb & ""
        txt_FaxNumber = rst!shi_faxnumb & ""
        txt_Email = rst!shi_mail & ""
        txt_Contact = rst!shi_cont & ""
        chkflag = IIf(rst!shi_actvflag, 1, 0)
        
        Set cmd = Nothing
        Set rst = Nothing
        
        If Err Then Call LogErr(Name & "::Getshipperinfo", Err.Description, Err.number, True)
   
   
End Sub

Public Function GetshipperCode(Name As String) As ADODB.Recordset
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
'        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT shi_code "
        .CommandText = .CommandText & " From Shipper "
        .CommandText = .CommandText & " where shi_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " where shi_npecode = '" & Name & "'"
        .CommandText = .CommandText & " order by shi_code  "
         Set rst = .Execute
   End With
    
    If rst.RecordCount = 0 Then GoTo CleanUp

    rst.MoveFirst

            Do While ((Not rst.EOF))
            dcboShipCode.AddItem rst!shi_code
            rst.MoveNext
        Loop

       
CleanUp:
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err Then Call LogErr(Name & "::GetshipperCode", Err.Description, Err.number, True)
End Function

'function clear screen

Public Sub ClearScreen()
        'dcboShipCode.Tag = dcboShipCode.ListIndex - 1
        dcboShipCode = ""
        txt_ShipperName = ""
        txt_Address1 = ""
        txt_Address2 = ""
        txt_City = ""
        txt_State = ""
        txt_Zipcode = ""
        txt_PhoneNumber = ""
        txt_FaxNumber = ""
        txt_Email = ""
        txt_Contact = ""
        txt_Country = ""
'        chkFlag = ""
End Sub

'SQL statement get shipper code

Public Function CheckshipperCode(Name As String) As Boolean
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
'        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT count(*) rt"
        .CommandText = .CommandText & " From Shipper "
        .CommandText = .CommandText & " where shi_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and shi_code = '" & Name & "'"
'        .CommandText = .CommandText & " order by shi_code  "
      
   
        Set rst = .Execute
        CheckshipperCode = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckshipperCode", Err.Description, Err.number, True)
End Function



'SQL statement get shipper code

Public Function CheckshipperCodeexit(Name As String) As Boolean

On Error GoTo Handler
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
CheckshipperCodeexit = True
    
    Set cmd = New ADODB.Command
        
    With cmd
'        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT Count(*) Count"
        .CommandText = .CommandText & " From Shipper "
        .CommandText = .CommandText & " where shi_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and shi_code = '" & Name & "'"

      
   
        Set rst = .Execute
        If rst!Count = 0 Then CheckshipperCodeexit = False
        
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    Exit Function
Handler:
   
   MsgBox "Errors Occurred while Quering the Database to check if the code exists." & vbCrLf & "Error Description -- " & Err.Description
   Err.Clear
   
End Function


Public Sub FillTextBox()

dcboShipCode = RsShipperCode!shi_code
txt_ShipperName = RsShipperCode!shi_name & ""
txt_Address1 = RsShipperCode!shi_adr1 & ""
txt_Address2 = RsShipperCode!shi_adr2 & ""
txt_City = RsShipperCode!shi_city & ""
txt_State = RsShipperCode!shi_stat & ""
txt_Zipcode = RsShipperCode!shi_zipc & ""
txt_Country = RsShipperCode!shi_ctry & ""
txt_PhoneNumber = RsShipperCode!shi_phonnumb & ""
txt_FaxNumber = RsShipperCode!shi_faxnumb & ""
txt_Email = RsShipperCode!shi_mail & ""
txt_Contact = RsShipperCode!shi_cont & ""
chkflag = IIf(RsShipperCode!shi_actvflag = True, 1, 0)
End Sub

Public Function GetShipperDetail() As ADODB.Recordset
Dim query As String
Dim rs As New ADODB.Recordset
On Error Resume Next
        query = " SELECT shi_code, shi_name, shi_npecode,"
        query = query & " shi_adr1, shi_adr2, shi_city, shi_stat, shi_zipc,shi_ctry,"
        query = query & " shi_phonnumb, shi_faxnumb, shi_telxnumb, shi_mail,"
        query = query & " shi_cont , shi_actvflag"
        query = query & " From Shipper"
        query = query & " WHERE (shi_npecode = '" & deIms.NameSpace & "') "
        
        
   rs.Source = query
   rs.ActiveConnection = deIms.cnIms
   rs.Open , , adOpenKeyset, adLockOptimistic
   
   Err.Clear
   Set GetShipperDetail = rs
End Function

Public Function LoadShipperCombo() As Boolean
'Added By muzammil
On Error GoTo Handler
LoadShipperCombo = False
If RsShipperCode.RecordCount > 0 Then RsShipperCode.MoveFirst
    Do While Not RsShipperCode.EOF
       dcboShipCode.AddItem Trim$(RsShipperCode!shi_code)
       RsShipperCode.MoveNext
    Loop
    
    
    RsShipperCode.MoveFirst
    'mIsCodeComboLoaded = True
    LoadShipperCombo = True
    
 Exit Function
Handler:
     MsgBox "Could not Load the Phone Code combo.Please close and reopen the form.Error Description -- " & Err.Description, vbCritical, "Imswin"
     Err.Clear
    
End Function


Public Function SaveToRecordset() As Boolean
SaveToRecordset = False
On Error GoTo Handler
 RsShipperCode!shi_npecode = deIms.NameSpace
 RsShipperCode!shi_code = Trim$(dcboShipCode)
 RsShipperCode!shi_name = Trim$(txt_ShipperName)
 RsShipperCode!shi_adr1 = Trim$(txt_Address1)
 RsShipperCode!shi_adr2 = Trim$(txt_Address2)
 RsShipperCode!shi_city = Trim$(txt_City)
 RsShipperCode!shi_stat = Trim$(txt_State)
 RsShipperCode!shi_zipc = Trim$(txt_Zipcode)
 RsShipperCode!shi_ctry = Trim$(txt_Country)
 RsShipperCode!shi_phonnumb = Trim$(txt_PhoneNumber)
 RsShipperCode!shi_faxnumb = Trim$(txt_FaxNumber)
 RsShipperCode!shi_mail = Trim$(txt_Email)
 RsShipperCode!shi_cont = Trim$(txt_Contact)
 RsShipperCode!shi_actvflag = IIf(chkflag = 1, True, 0)
 SaveToRecordset = True
Exit Function

Handler:
  MsgBox "Errors Occured while saving the record." & vbCrLf & "Error Description -- " & Err.Description
  Err.Clear
End Function
