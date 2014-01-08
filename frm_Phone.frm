VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Phone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phone Directory"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   5625
   Tag             =   "01030200"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbPhoneDir 
      Height          =   315
      Left            =   2310
      TabIndex        =   0
      Top             =   720
      Width           =   3075
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   5424
      _ExtentY        =   564
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.TextBox txtcontact 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   16
      Top             =   5300
      Width           =   3072
   End
   Begin VB.TextBox txtalternative 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   15
      Top             =   5000
      Width           =   3072
   End
   Begin VB.TextBox txthome 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   14
      Top             =   4700
      Width           =   3072
   End
   Begin VB.TextBox txtbeep 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   13
      Top             =   4400
      Width           =   3072
   End
   Begin VB.TextBox txt_Zipcode 
      Height          =   288
      Left            =   4200
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txt_State 
      Height          =   288
      Left            =   2325
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2280
      Width           =   528
   End
   Begin VB.TextBox txt_Address1 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1365
      Width           =   3072
   End
   Begin VB.TextBox txt_Address2 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1680
      Width           =   3072
   End
   Begin VB.TextBox txt_City 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1980
      Width           =   3072
   End
   Begin VB.TextBox txt_PhondirName 
      Height          =   288
      Left            =   2325
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1050
      Width           =   3072
   End
   Begin VB.TextBox txt_Cell 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   12
      Top             =   4080
      Width           =   3072
   End
   Begin VB.TextBox txt_FaxNumber 
      Height          =   288
      Left            =   2325
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3180
      Width           =   3072
   End
   Begin VB.TextBox txt_PhoneNumber 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2880
      Width           =   3072
   End
   Begin VB.TextBox txt_Country 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2580
      Width           =   3072
   End
   Begin VB.TextBox txtTelexnumber 
      Height          =   288
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   10
      Top             =   3480
      Width           =   3072
   End
   Begin VB.TextBox txt_Email 
      Height          =   288
      Left            =   2325
      MaxLength       =   59
      TabIndex        =   11
      Top             =   3780
      Width           =   3072
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   5760
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frm_Phone.frx":0000
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
      DisableSaveOnSave=   0   'False
   End
   Begin VB.Label Lblcont 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   210
      Left            =   240
      TabIndex        =   35
      Top             =   5300
      Width           =   2000
   End
   Begin VB.Label txtalt 
      BackStyle       =   0  'Transparent
      Caption         =   "Alternative"
      Height          =   210
      Left            =   240
      TabIndex        =   34
      Top             =   5000
      Width           =   2000
   End
   Begin VB.Label Lblhome 
      BackStyle       =   0  'Transparent
      Caption         =   "Home"
      Height          =   210
      Left            =   240
      TabIndex        =   33
      Top             =   4700
      Width           =   2000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Beeper"
      Height          =   210
      Left            =   240
      TabIndex        =   32
      Top             =   4395
      Width           =   2000
   End
   Begin VB.Label lbl_Phonedir 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Directory"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   285
      TabIndex        =   31
      Top             =   120
      Width           =   5130
   End
   Begin VB.Label lbl_Cell 
      BackStyle       =   0  'Transparent
      Caption         =   "Cellular"
      Height          =   210
      Left            =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   210
      Left            =   240
      TabIndex        =   29
      Top             =   3180
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   210
      Left            =   240
      TabIndex        =   28
      Top             =   2880
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   210
      Left            =   240
      TabIndex        =   27
      Top             =   2580
      Width           =   2000
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   210
      Left            =   240
      TabIndex        =   26
      Top             =   2280
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   210
      Left            =   3120
      TabIndex        =   25
      Top             =   2310
      Width           =   765
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   1980
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   210
      Left            =   240
      TabIndex        =   21
      Top             =   1365
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneName 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   1065
      Width           =   2000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telex Number"
      Height          =   210
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   210
      Left            =   240
      TabIndex        =   18
      Top             =   3780
      Width           =   2000
   End
End
Attribute VB_Name = "frm_Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim plist As imsPhoneDirectory
Dim rsPHONEDIR As ADODB.Recordset
Dim mIsCodeComboLoaded As Boolean
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Sub CmbPhoneDir_Click()


If Not rsPHONEDIR.EditMode = 2 Then
Dim str As String

    If Len(Trim$(CmbPhoneDir)) > 0 And CmbPhoneDir.IsItemInList Then
                rsPHONEDIR.MoveFirst
                rsPHONEDIR.Find "phd_code='" & CmbPhoneDir & "'", , adSearchForward
                
                If Not rsPHONEDIR.AbsolutePosition = adPosEOF Then
                   FillTextBox
                End If
        'Set plist = plist.GetPhoneDirectorylist(CmbPhoneDir, deIms.NameSpace, deIms.cnIms)
      '  EnableButtons
    End If
End If
End Sub

Private Sub CmbPhoneDir_DropDown()

If rsPHONEDIR.EditMode = 2 Then CmbPhoneDir.DroppedDown = False
If mIsCodeComboLoaded = False Then
   Call LoadPhoneDircombo
End If
End Sub

Private Sub CmbPhoneDir_GotFocus()
Call HighlightBackground(CmbPhoneDir)
End Sub

Private Sub CmbPhoneDir_KeyDown(KeyCode As Integer, Shift As Integer)
If rsPHONEDIR.EditMode = 2 Then
   CmbPhoneDir.DroppedDown = False
   If KeyCode = 40 Or KeyCode = 38 Then KeyCode = 0
Else
  If Not CmbPhoneDir.DroppedDown Then CmbPhoneDir.DroppedDown = True
End If
End Sub

Private Sub CmbPhoneDir_KeyPress(KeyAscii As Integer)
''If rsPHONEDIR.editmode = 2 Then
''   CmbPhoneDir.DroppedDown = False
''Else
''   CmbPhoneDir.DroppedDown = True
''End If
End Sub

Private Sub CmbPhoneDir_LostFocus()
Call NormalBackground(CmbPhoneDir)
End Sub

Private Sub CmbPhoneDir_Validate(Cancel As Boolean)
CmbPhoneDir = Trim$(CmbPhoneDir)

If Len(CmbPhoneDir) = 0 Then Exit Sub

If rsPHONEDIR.EditMode = 2 Then
    
   
   If CmbPhoneDir.IsItemInList Then
          Cancel = True
          MsgBox "The code Already Exist , Please use a different one.", vbInformation, "Imswin"
          CmbPhoneDir.SetFocus
          Exit Sub
    End If
         
    If Len(CmbPhoneDir) > 10 Then
          Cancel = True
          MsgBox "Code can not be more than 10 characters.", vbInformation, "Imswin"
          CmbPhoneDir.SetFocus
          Exit Sub
    End If
    
Else
    
   If Not CmbPhoneDir.IsItemInList Then
          Cancel = True
          MsgBox "The code does not exist.", vbInformation, "Imswin"
          Call Clearform
          CmbPhoneDir.SetFocus
          Exit Sub
   End If
   
End If


           
End Sub

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




Dim cmd As ADODB.Command
Dim cn As ADODB.Connection
Dim query As String

    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Translate_Forms("frm_Phone")
    '------------------------------------------

    'Set plist = New imsPhoneDirectory
    'Call PopuLateFromRecordSet(CmbPhoneDir, plist.GetPhonedirectoryCode(deIms.NameSpace, deIms.cnIms), "phd_code", True)
    
'   Set rs = plist.GetPhonedirectoryCode(deIms.NameSpace, deIms.cnIms)
  
        query = " SELECT  phd_code,phd_name, phd_adr1, phd_adr2, "
        query = query & " phd_city, phd_stat, phd_zipc, phd_phonnumb, "
        query = query & " phd_faxnumb, phd_telxnumb, phd_cellnumb, "
        query = query & " phd_homenumb, phd_beepnumb, phd_altnnumb, "
        query = query & " phd_mail , phd_cont, phd_ctry"
        query = query & " From PHONEDIR "
        query = query & " WHERE phd_npecode = '" & deIms.NameSpace & "'order by phd_code"
        'Query = Query & " AND phd_code = '" & Code & "'"
    Set rsPHONEDIR = New ADODB.Recordset
    rsPHONEDIR.Source = query
    rsPHONEDIR.ActiveConnection = deIms.cnIms
    rsPHONEDIR.Open , , adOpenKeyset, adLockOptimistic
    
        
    Call LoadPhoneDircombo
    
    rsPHONEDIR.MoveFirst
     CmbPhoneDir = Trim(rsPHONEDIR!phd_code)
    Call CmbPhoneDir_Click
    
    NavBar1.CancelLastSepVisible = False
    NavBar1.LastPrintSepVisible = False
    NavBar1.PrintSaveSepVisible = False
    CmbPhoneDir.Columns(0).Width = 3075
    frm_Phone.Caption = frm_Phone.Caption + " - " + frm_Phone.Tag
    
      Call DisableButtons(Me, NavBar1)
    
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form free memory

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set plist = Nothing
    If open_forms <= 5 Then ShowNavigator
    
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    
End Sub

'clear form

Private Sub NavBar1_OnCancelClick()
   'CmbPhoneDir.ListIndex = CmbPhoneDir.ListCount - 1
   CmbPhoneDir = Trim$(CmbPhoneDir)
   
   
   If rsPHONEDIR.EditMode = 2 Then
   
   rsPHONEDIR.CancelUpdate
   If Not rsPHONEDIR.BOF Then rsPHONEDIR.MoveFirst
   If Not rsPHONEDIR.AbsolutePosition = adPosBOF Then
     FillTextBox
   Else
     
     Call Clearform
   End If
   
 Else
 
   If Len(CmbPhoneDir) > 0 And CmbPhoneDir.IsItemInList Then
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

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    
    If rsPHONEDIR.RecordCount > 0 Then
        rsPHONEDIR.MoveFirst
        FillTextBox
    End If
    
End Sub

'move recordset to lasst position

Private Sub NavBar1_OnLastClick()
    
    If rsPHONEDIR.RecordCount > 0 Then
        rsPHONEDIR.MoveLast
        FillTextBox
    End If
    
End Sub

'set navbar button

Public Sub EnableButtons()
Dim i As Integer

  '  i = CmbPhoneDir.ListIndex

    If CmbPhoneDir.ListCount = 0 Then

        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False

        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False

        Exit Sub

    ElseIf i = CmbPhoneDir.ListCount - 1 Then
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

    If CmbPhoneDir.ListIndex = CB_ERR Then CmbPhoneDir.ListIndex = 0
    If Err Then Err.Clear
End Sub

'on new click clear form

Private Sub NavBar1_OnNewClick()
       
       rsPHONEDIR.CancelUpdate
       rsPHONEDIR.AddNew
       Call Clearform
       
       
       CmbPhoneDir.SetFocus
       Call CmbPhoneDir_GotFocus

End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
On Error Resume Next

    
       If Not rsPHONEDIR.EOF Then rsPHONEDIR.MoveNext
       
       rsPHONEDIR.CancelUpdate
       If Not rsPHONEDIR.AbsolutePosition = adPosEOF Then
        FillTextBox
       Else
         rsPHONEDIR.MoveLast
       End If
    
    
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
     
     If Not rsPHONEDIR.BOF Then rsPHONEDIR.MovePrevious
       
       
       If Not rsPHONEDIR.AbsolutePosition = adPosBOF Then
        FillTextBox
       Else
         rsPHONEDIR.MoveFirst
       End If
    
End Sub

'call store procedure to add a phone record

Private Sub InsertPhoneDire()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "UP_INS_PHONEDIR"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms

        .parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, CmbPhoneDir)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 35, txt_PhondirName)
        .parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, txt_Address1)
        .parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, txt_Address2)
        .parameters.Append .CreateParameter("@city", adVarChar, adParamInput, 25, txt_City)
        .parameters.Append .CreateParameter("@stat", adVarChar, adParamInput, 2, txt_State)
        .parameters.Append .CreateParameter("@zipc", adVarChar, adParamInput, 11, txt_Zipcode)
        .parameters.Append .CreateParameter("@ctry", adVarChar, adParamInput, 25, txt_Country)
        .parameters.Append .CreateParameter("@phonnumb", adVarChar, adParamInput, 25, txt_PhoneNumber)
        .parameters.Append .CreateParameter("@faxnumb", adVarChar, adParamInput, 50, txt_FaxNumber)
        .parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, txtTelexnumber)
        .parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, txt_Email)
        .parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, txtcontact)
        .parameters.Append .CreateParameter("@Cell", adVarChar, adParamInput, 25, txt_Cell)
        .parameters.Append .CreateParameter("@BEEPER", adVarChar, adParamInput, 25, txtbeep)
        .parameters.Append .CreateParameter("@HOME", adVarChar, adParamInput, 25, txthome)
        .parameters.Append .CreateParameter("@ALTE", adVarChar, adParamInput, 25, txtalternative)
        .parameters.Append .CreateParameter("@user", adVarChar, adParamInput, 20, CurrentUser)
        Call .Execute(Options:=adExecuteNoRecords)

    End With

    Set cmd = Nothing
    
    'If IndexOf(CmbPhoneDir, CmbPhoneDir) = CB_ERR Then
      If Err.number = 0 Then CmbPhoneDir.AddItem (CmbPhoneDir)
    'End If
    
    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("M00308") 'J added
    MsgBox IIf(msg1 = "", "Insert into Phone Directory is completed successfully ", msg1) 'J modified
    '---------------------------------------------

    
    Exit Sub

Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00309") 'J added
        MsgBox IIf(msg1 = "", "Insert into Phone Directory is failure ", msg1) 'J modified
        '---------------------------------------------

End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\phonedir.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00098") 'J added
        .WindowTitle = IIf(msg1 = "", "Phone Directory", msg1) 'J modified
        Call translator.Translate_Reports("phonedir.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With

Handler:
    If Err Then MsgBox Err.Description: Err.Clear
End Sub

'before save a record validate data format

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
Dim Numb As Integer

     If Len(Trim$(CmbPhoneDir)) = 0 Then
     
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        CmbPhoneDir.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhondirName)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox "The Name cannot be left empty"
        '---------------------------------------------
        
        txt_PhondirName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multiliangual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '----------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox "The Phone Number cannot be left empty"
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
     
     Numb = CmbPhoneDir.ListIndex
     
    If Numb = -1 Then
        If Len(Trim$(CmbPhoneDir)) <> 0 Then
            If Checkphonedir(CmbPhoneDir) Then
            
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00310") 'J added
                MsgBox IIf(msg1 = "", "Phone Directory exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            End If
        End If
    End If
    'UpdateRow
    
    Call InsertPhoneDire
    rsPHONEDIR.CancelUpdate
    rsPHONEDIR.Requery
    mIsCodeComboLoaded = False
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub

'function assign value to text box

Private Sub FillTextBox()

''    txt_PhondirName = plist.phonename
''    txt_Address1 = plist.address1
''    txt_Address2 = plist.address2
''    txt_City = plist.City
''    txt_State = plist.State
''    txt_Zipcode = plist.ZipCode
''    txt_Country = plist.Country
''    txt_PhoneNumber = plist.PhoneNumber
''    txt_FaxNumber = plist.Faxnumber
''    txt_Email = plist.Email
''    txtcontact = plist.Contact
''    txtTelexnumber = plist.TelexNumber
''    txt_Cell = plist.Cellular
''    txtbeep = plist.Beeper
''    txthome = plist.Home
''    txtalternative = plist.Alternative
    
    CmbPhoneDir = rsPHONEDIR!phd_code
    txt_PhondirName = rsPHONEDIR!phd_name
    txt_Address1 = rsPHONEDIR!phd_adr1 & ""
    txt_Address2 = rsPHONEDIR!phd_adr2 & ""
    txt_City = rsPHONEDIR!phd_city & ""
    txt_State = rsPHONEDIR!phd_stat & ""
    txt_Zipcode = rsPHONEDIR!phd_zipc & ""
    txt_Country = rsPHONEDIR!phd_ctry & ""
    txt_PhoneNumber = rsPHONEDIR!phd_phonnumb & ""
    txt_FaxNumber = rsPHONEDIR!phd_faxnumb & ""
    txt_Email = rsPHONEDIR!phd_mail & ""
    txtcontact = rsPHONEDIR!phd_cont & ""
    txtTelexnumber = rsPHONEDIR!phd_telxnumb & ""
    txt_Cell = rsPHONEDIR!phd_cellnumb & ""
    txtbeep = rsPHONEDIR!phd_beepnumb & ""
    txthome = rsPHONEDIR!phd_homenumb & ""
    txtalternative = rsPHONEDIR!phd_altnnumb & ""
    
End Sub

'function to clear text box

Public Sub Clearform()
        
        CmbPhoneDir = ""
        txt_PhondirName = ""
        txt_Address1 = ""
        txt_Address2 = ""
        txt_City = ""
        txt_State = ""
        txt_Zipcode = ""
        txt_Country = ""
        txt_PhoneNumber = ""
        txt_FaxNumber = ""
        txtTelexnumber = ""
        txt_Email = ""
        txt_Cell = ""
        txtbeep = ""
        txthome = ""
        txtalternative = ""
        txtcontact = ""
End Sub

'SQL statement to check phone code exist or not

Private Function Checkphonedir(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From phonedir "
        .CommandText = .CommandText & " Where phd_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND phd_code = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        Checkphonedir = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::Checkphonedir", Err.Description, Err.number, True)
End Function

'Added By muzammil
Public Function LoadPhoneDircombo() As Boolean

On Error GoTo Handler
LoadPhoneDircombo = False

Do While Not rsPHONEDIR.EOF
       CmbPhoneDir.AddItem Trim$(rsPHONEDIR!phd_code)
       rsPHONEDIR.MoveNext
    Loop
    
    rsPHONEDIR.MoveFirst
    mIsCodeComboLoaded = True
    LoadPhoneDircombo = True
 Exit Function
Handler:
     MsgBox "Could not Load the Phone Code combo.Please close and reopen the form.Error Description -- " & Err.Description, vbCritical, "Imswin"
     Err.Clear
    
End Function

Private Sub txt_PhondirName_GotFocus()
Call HighlightBackground(txt_PhondirName)
End Sub

Private Sub txt_PhondirName_LostFocus()
Call NormalBackground(txt_PhondirName)
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

Private Sub txt_City_GotFocus()
Call HighlightBackground(txt_City)
End Sub

Private Sub txt_City_LostFocus()
Call NormalBackground(txt_City)
End Sub

Private Sub txt_State_GotFocus()
Call HighlightBackground(txt_State)
End Sub

Private Sub txt_State_LostFocus()
Call NormalBackground(txt_State)
End Sub

Private Sub txt_Zipcode_GotFocus()
Call HighlightBackground(txt_Zipcode)
End Sub

Private Sub txt_Zipcode_LostFocus()
Call NormalBackground(txt_Zipcode)
End Sub

Private Sub txt_Country_GotFocus()
Call HighlightBackground(txt_Country)
End Sub

Private Sub txt_Country_LostFocus()
Call NormalBackground(txt_Country)
End Sub
Private Sub txt_PhoneNumber_GotFocus()
Call HighlightBackground(txt_PhoneNumber)
End Sub

Private Sub txt_PhoneNumber_LostFocus()
Call NormalBackground(txt_PhoneNumber)
End Sub

Private Sub txt_FaxNumber_GotFocus()
Call HighlightBackground(txt_FaxNumber)
End Sub

Private Sub txt_FaxNumber_LostFocus()
Call NormalBackground(txt_FaxNumber)
End Sub

Private Sub txtTelexnumber_GotFocus()
Call HighlightBackground(txtTelexnumber)
End Sub

Private Sub txtTelexnumber_LostFocus()
Call NormalBackground(txtTelexnumber)
End Sub
Private Sub txt_Email_GotFocus()
Call HighlightBackground(txt_Email)
End Sub

Private Sub txt_Email_LostFocus()
Call NormalBackground(txt_Email)
End Sub

Private Sub txt_Cell_GotFocus()
Call HighlightBackground(txt_Cell)
End Sub

Private Sub txt_Cell_LostFocus()
Call NormalBackground(txt_Cell)
End Sub

Private Sub txtbeep_GotFocus()
Call HighlightBackground(txtbeep)
End Sub

Private Sub txtbeep_LostFocus()
Call NormalBackground(txtbeep)
End Sub
Private Sub txthome_GotFocus()
Call HighlightBackground(txthome)
End Sub

Private Sub txthome_LostFocus()
Call NormalBackground(txthome)
End Sub

Private Sub txtalternative_GotFocus()
Call HighlightBackground(txtalternative)
End Sub

Private Sub txtalternative_LostFocus()
Call NormalBackground(txtalternative)
End Sub

Private Sub txtcontact_GotFocus()
Call HighlightBackground(txtcontact)
End Sub

Private Sub txtcontact_LostFocus()
Call NormalBackground(txtcontact)
End Sub
