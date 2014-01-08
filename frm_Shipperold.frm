VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frm_Shipper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipper"
   ClientHeight    =   5400
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   4755
   Tag             =   "01010200"
   Begin VB.TextBox txttelxnumber 
      Height          =   315
      Left            =   1575
      MaxLength       =   25
      TabIndex        =   10
      Tag             =   "10"
      Top             =   3360
      Width           =   3024
   End
   Begin VB.ComboBox dcboShipCode 
      Height          =   315
      Left            =   1572
      TabIndex        =   0
      Tag             =   "0"
      Text            =   "dcboShipCode"
      Top             =   580
      Width           =   3015
   End
   Begin VB.CheckBox chkflag 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   4440
      Width           =   255
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   435
      Left            =   720
      TabIndex        =   26
      Top             =   4740
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   767
      AllowCustomize  =   0   'False
      CancelToolTipText=   "Undo the changes made to the current record"
      CloseToolTipText=   "Closes the current window"
      EMailEnabled    =   0   'False
      EmailToolTipText=   "Send current record via email"
      FirstToolTipText=   "Moves to the first record"
      LastToolTipText =   "Moves to the last record"
      NewEnabled      =   -1  'True
      NewToolTipText  =   "Adds a new record"
      NextToolTipText =   "Moves to the next record"
      PreviousToolTipText=   "Moves to the previous record"
      PrintToolTipText=   "Prints current record"
      SaveToolTipText =   "Save the changes made to the current record"
      DeleteToolTipText=   ""
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
   End
   Begin VB.TextBox txt_Address1 
      Height          =   315
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1230
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      Height          =   288
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1560
      Width           =   3024
   End
   Begin VB.TextBox txt_City 
      Height          =   288
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1860
      Width           =   3024
   End
   Begin VB.TextBox txt_Zipcode 
      Height          =   288
      Left            =   3000
      MaxLength       =   11
      TabIndex        =   6
      Tag             =   "6"
      Top             =   2160
      Width           =   1596
   End
   Begin VB.TextBox txt_State 
      Height          =   288
      Left            =   1572
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "5"
      Top             =   2160
      Width           =   408
   End
   Begin VB.TextBox txt_FaxNumber 
      Height          =   288
      Left            =   1572
      MaxLength       =   50
      TabIndex        =   9
      Tag             =   "9"
      Top             =   3060
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      Height          =   288
      Left            =   1572
      MaxLength       =   59
      TabIndex        =   11
      Tag             =   "11"
      Top             =   3720
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      Height          =   288
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   12
      Tag             =   "12"
      Top             =   4020
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      Height          =   288
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   8
      Tag             =   "8"
      Top             =   2760
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      Height          =   288
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   7
      Tag             =   "7"
      Top             =   2460
      Width           =   3024
   End
   Begin VB.TextBox txt_ShipperName 
      Height          =   315
      Left            =   1572
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "1"
      Top             =   900
      Width           =   3024
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Telx Number"
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Flag"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   1365
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   288
      Left            =   120
      TabIndex        =   16
      Top             =   1260
      Width           =   1368
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1368
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   288
      Left            =   120
      TabIndex        =   18
      Top             =   1860
      Width           =   1368
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   195
      Left            =   2115
      TabIndex        =   20
      Top             =   2190
      Width           =   645
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   288
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1368
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   288
      Left            =   120
      TabIndex        =   21
      Top             =   2460
      Width           =   1368
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   288
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   1368
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   4020
      Width           =   1365
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   288
      Left            =   120
      TabIndex        =   23
      Top             =   3060
      Width           =   1368
   End
   Begin VB.Label lbl_Shipper 
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
      Left            =   1853
      TabIndex        =   13
      Top             =   120
      Width           =   1005
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
      TabIndex        =   14
      Top             =   660
      Width           =   1368
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
      TabIndex        =   15
      Top             =   960
      Width           =   1368
   End
End
Attribute VB_Name = "frm_Shipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cbo_ShipperCode_GotFocus()
    Call HighlightBackground(dcboShipCode)
End Sub

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

Private Sub dcboShipCode_Click()
On Error Resume Next

        

        If Not Len(Trim(dcboShipCode)) Then
            Call Getshipperinfo(dcboShipCode)
        End If
        
        If Err Then Err.Clear
 End Sub

'Private Sub dcboShipCode_KeyPress(KeyAscii As Integer)
'    If Len(dcboShipCode) = 10 Then
'        If KeyAscii >= vbKeySpace Then KeyAscii = 0
'    End If
'End Sub

'Private Sub dcboShipCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' dcboShipCode.locked = False
'End Sub
'
'Private Sub dcboShipCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  dcboShipCode.locked = Not NavBar1.NewEnabled
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsSHIPPER.CancelUpdate
    
    
    NavBar1.Recordset.Close
    Set NavBar1.Recordset = Nothing
    
    deIms.rsSHIPPER.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
End Sub



Private Sub NavBar1_OnCancelClick()

    Call NavBar1_OnPreviousClick

End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnFirstClick()
On Error Resume Next

        dcboShipCode.ListIndex = 0
    

    If Err Then Call LogErr(Name & "::NavBar1_OnFirstClick", True)

End Sub

Private Sub NavBar1_OnLastClick()
On Error Resume Next

        dcboShipCode.ListIndex = dcboShipCode.ListCount - 1

    If Err Then Call LogErr(Name & "::NavBar1_OnLastClick", True)

End Sub

Private Sub NavBar1_OnNewClick()
    Call ClearScreen

End Sub

Private Sub NavBar1_OnNextClick()
On Error Resume Next

    dcboShipCode.ListIndex = dcboShipCode.ListIndex + 1

    If Err Then Call LogErr(Name & "::NavBar1_OnNextClick", True)
End Sub

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next
     If dcboShipCode.ListIndex = -1 Then
        Exit Sub
    Else
        dcboShipCode.ListIndex = dcboShipCode.ListIndex - 1
    End If

    If Err Then Call LogErr(Name & "::NavBar1_OnPreviousClick", True)

End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Shipper.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim ctl As Control
    
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
    
    Call BindAll(Me, deIms)
    Screen.MousePointer = vbDefault
    
     Call PopuLateFromRecordSet(dcboShipCode, GetshipperCode(deIms.NameSpace), "shi_npecode", True)
    If dcboShipCode.ListCount Then dcboShipCode.ListIndex = 0
'    Call DisableButtons(Me, NavBar1)
    frm_Shipper.Caption = frm_Shipper.Caption + " - " + frm_Shipper.Tag
    
End Sub

Private Sub NavBar1_OnSaveClick()
    If Len(Trim$(dcboShipCode)) = 0 Then
        MsgBox "The Code cannot be left empty"
        dcboShipCode.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_ShipperName)) = 0 Then
        MsgBox "The Name cannot be left empty"
        txt_ShipperName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Address1)) = 0 Then
        MsgBox "The Address cannot be left empty"
        txt_Address1.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
        MsgBox "The City cannot be left empty"
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Country)) = 0 Then
        MsgBox "The Country cannot be left empty"
        txt_Country.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
        MsgBox "The Phone Number cannot be left empty"
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    
    Call InsertShipper
'    deIms.rsSHIPPER.UpdateBatch
    
'    deIms.rsSHIPPER!shi_creauser = CurrentUser
'    deIms.rsSHIPPER!shi_modiuser = CurrentUser
'     Call deIms.rsSHIPPER.Move(0)
End Sub

Private Sub txt_Address1_GotFocus()
    Call HighlightBackground(txt_Address1)
End Sub

Private Sub txt_Address1_LostFocus()
    Call NormalBackground(txt_Address1)
End Sub


Private Sub txt_Address2_GotFocus()
    Call HighlightBackground(txt_Address2)
End Sub

Private Sub txt_Address2_LostFocus()
    Call NormalBackground(txt_Address2)
End Sub

Private Sub txt_ShipperName_GotFocus()
    Call HighlightBackground(txt_ShipperName)
End Sub

Private Sub txt_ShipperName_LostFocus()
    Call NormalBackground(txt_ShipperName)
End Sub

Private Sub txt_City_GotFocus()
    Call HighlightBackground(txt_City)
End Sub

Private Sub txt_City_LostFocus()
    Call NormalBackground(txt_City)
End Sub

Private Sub txt_Contact_GotFocus()
    Call HighlightBackground(txt_Contact)
End Sub

Private Sub txt_Contact_LostFocus()
    Call NormalBackground(txt_Contact)
End Sub

Private Sub txt_Country_GotFocus()
    Call HighlightBackground(txt_Country)
End Sub

Private Sub txt_Country_LostFocus()
    Call NormalBackground(txt_Country)
End Sub

Private Sub txt_Email_GotFocus()
    Call HighlightBackground(txt_Email)
End Sub

Private Sub txt_Email_LostFocus()
    Call NormalBackground(txt_Email)
End Sub

Private Sub txt_FaxNumber_GotFocus()
    Call HighlightBackground(txt_FaxNumber)
End Sub

Private Sub txt_FaxNumber_LostFocus()
    Call NormalBackground(txt_FaxNumber)
End Sub

Private Sub txt_Name_Change()

End Sub

Private Sub txt_Name_GotFocus()
    Call HighlightBackground(txt_ShipperName)
End Sub

Private Sub txt_Name_LostFocus()
    Call NormalBackground(txt_ShipperName)
End Sub

Private Sub txt_PhoneNumber_GotFocus()
    Call HighlightBackground(txt_PhoneNumber)
End Sub

Private Sub txt_PhoneNumber_LostFocus()
    Call NormalBackground(txt_PhoneNumber)
End Sub

Private Sub txt_State_GotFocus()
    Call HighlightBackground(txt_State)
End Sub

Private Sub txt_State_LostFocus()
    Call NormalBackground(txt_State)
End Sub

Private Sub txt_SupName_GotFocus()
    Call HighlightBackground(txt_ShipperName)
End Sub

Private Sub txt_SupName_LostFocus()
    Call NormalBackground(txt_ShipperName)
End Sub

Private Sub txt_Zipcode_GotFocus()
    Call HighlightBackground(txt_Zipcode)
End Sub

Private Sub txt_Zipcode_LostFocus()
    Call NormalBackground(txt_Zipcode)
End Sub

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
        .CommandText = .CommandText & " shi_adr1, shi_adr2, shi_city, shi_stat, shi_zipc,shi_ctry,"
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
        txt_Country = rst!shi_ctry & ""
        txt_PhoneNumber = rst!shi_phonnumb & ""
        txt_FaxNumber = rst!shi_faxnumb & ""
        txttelxnumber = rst!shi_telxnumb & ""
        txt_Email = rst!shi_mail & ""
        txt_Contact = rst!shi_cont & ""
        chkflag = IIf(rst!shi_actvflag, 1, 0)
        
        Set cmd = Nothing
        Set rst = Nothing
        
        If Err Then Call LogErr(Name & "::Getshipperinfo", True)
   
   
End Sub

Public Function GetshipperCode(Name As String) As ADODB.Recordset
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim Rstlist As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
'        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT shi_code "
        .CommandText = .CommandText & " From Shipper "
'        .CommandText = .CommandText & " where shi_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " where shi_npecode = '" & Name & "'"
        .CommandText = .CommandText & " order by shi_code  "
         Set Rstlist = .Execute
   End With
    
    If Rstlist.RecordCount = 0 Then GoTo CleanUp

    Rstlist.MoveFirst

            Do While ((Not Rstlist.EOF))
            dcboShipCode.AddItem Rstlist!shi_code
            Rstlist.MoveNext
        Loop

       
CleanUp:
    Rstlist.Close
    Set Rstlist = Nothing
    Set cmd = Nothing
    
    If Err Then Call LogErr(Name & "::GetshipperCode", True)
End Function

Public Sub ClearScreen()
        dcboShipCode.ListIndex = -1
        txt_ShipperName = ""
        txt_Address1 = ""
        txt_Address2 = ""
        txt_City = ""
        txt_State = ""
        txt_Zipcode = ""
        txt_Country = ""
        txt_PhoneNumber = ""
        txt_FaxNumber = ""
        txttelxnumber = ""
        txt_Email = ""
        txt_Contact = ""
'        chkFlag = ""
End Sub


Private Sub InsertShipper()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "UP_INS_SHIPPER"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .Parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, dcboShipCode)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 25, txt_ShipperName)
        .Parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, txt_Address1)
        .Parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, txt_Address2)
        .Parameters.Append .CreateParameter("@city", adVarChar, adParamInput, 25, txt_City)
        .Parameters.Append .CreateParameter("@stat", adVarChar, adParamInput, 2, txt_State)
        .Parameters.Append .CreateParameter("@zipc", adVarChar, adParamInput, 11, txt_Zipcode)
        .Parameters.Append .CreateParameter("@ctry", adVarChar, adParamInput, 25, txt_Country)
        .Parameters.Append .CreateParameter("@phonnumb", adVarChar, adParamInput, 25, txt_PhoneNumber)
        .Parameters.Append .CreateParameter("@faxnumb", adVarChar, adParamInput, 50, txt_FaxNumber)
        .Parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, txttelxnumber)
        .Parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, txt_Email)
        .Parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, txt_Contact)
        .Parameters.Append .CreateParameter("@CreatedBY", adVarChar, adParamInput, 20, CurrentUser)
        .Parameters.Append .CreateParameter("@flag", adBoolean, adParamInput, , chkflag.Value = vbChecked)
        .Execute , , adExecuteNoRecords
    
    End With
    
    Set cmd = Nothing
    If IndexOf(dcboShipCode, dcboShipCode) = CB_ERR Then
        dcboShipCode.AddItem (dcboShipCode)
    End If
    Exit Sub
    
Noinsert:
        If Err Then Err.Clear
        MsgBox "Insert into SHIPPER is failure "
        
End Sub

