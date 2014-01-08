VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_LocationSITE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location/SITE"
   ClientHeight    =   5505
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   5295
   Tag             =   "01030900"
   Begin VB.CheckBox chkFlag 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2100
      TabIndex        =   17
      Top             =   4560
      Width           =   255
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   840
      TabIndex        =   18
      Top             =   4920
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frm_LocationSITE.frx":0000
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
      DisableSaveOnSave=   0   'False
   End
   Begin VB.ComboBox cboComp 
      Height          =   315
      Left            =   2115
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.ComboBox cboCode 
      Height          =   315
      Left            =   2115
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   3015
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "loc_zipc"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2340
      Width           =   1410
   End
   Begin VB.TextBox txt_City 
      DataField       =   "loc_city"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   5
      Top             =   2040
      Width           =   3024
   End
   Begin VB.ComboBox cbo_LocationType 
      DataField       =   "loc_gender"
      DataMember      =   "LOCATION"
      DataSource      =   "deIms"
      Height          =   315
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4140
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "loc_ctry"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2640
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "loc_phonnumb"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2940
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "loc_cont"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   14
      Top             =   3840
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "loc_mail"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   59
      TabIndex        =   12
      Top             =   3540
      Width           =   3024
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "loc_faxnumb"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3240
      Width           =   3024
   End
   Begin VB.TextBox txt_State 
      DataField       =   "loc_stat"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2115
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2340
      Width           =   408
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "loc_adr2"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1740
      Width           =   3024
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "loc_adr1"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1440
      Width           =   3024
   End
   Begin VB.TextBox txt_SupName 
      DataField       =   "loc_name"
      DataMember      =   "LOCATION"
      Height          =   288
      Left            =   2100
      TabIndex        =   2
      Top             =   1140
      Width           =   3024
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Flag"
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   2000
   End
   Begin VB.Label lbl_LocationType 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Type"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   4140
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3540
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   2940
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   2000
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   2340
      Width           =   2000
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   1740
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   2000
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Name"
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
      Left            =   120
      TabIndex        =   22
      Top             =   1170
      Width           =   2000
   End
   Begin VB.Label lbl_Code 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Code"
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
      Left            =   120
      TabIndex        =   21
      Top             =   810
      Width           =   2000
   End
   Begin VB.Label lbl_CompanyCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   120
      TabIndex        =   20
      Top             =   492
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   285
      Left            =   2655
      TabIndex        =   27
      Top             =   2340
      Width           =   1125
   End
   Begin VB.Label lbl_Location 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location/SITE"
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
      Left            =   1665
      TabIndex        =   19
      Top             =   60
      Width           =   1920
   End
End
Attribute VB_Name = "frm_LocationSITE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fm As FormMode
Dim State, starting As Boolean
Dim loc As imsLocation
Dim NAV_NEW As Boolean
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Sub cbo_LocationType_Click()
Dim i
    If Not starting Then
        If UCase(Trim(cbo_LocationType)) <> "SITE" Then
            MsgBox "You can not manipulate this gender through this option"
            For i = 0 To cbo_LocationType.ListCount - 1
                If UCase(Trim(cbo_LocationType.list(i))) = "SITE" Then
                    cbo_LocationType.ListIndex = i
                    Exit For
                End If
            Next
            cbo_LocationType.SetFocus
        End If
    End If
End Sub


Private Sub cbo_LocationType_GotFocus()
Call HighlightBackground(cbo_LocationType)
End Sub

Private Sub cbo_LocationType_LostFocus()
Call NormalBackground(cbo_LocationType)
End Sub

'select combo start point

Private Sub cboCode_Change()
Dim i As Integer
    
    i = cboCode.SelStart
    Debug.Print i
    If Len(cboCode) > 10 Then cboCode = VBA.Left$(cboCode, 10)
    cboCode.SelStart = i
End Sub

'set combo values and navbar button

Private Sub cboCode_Click()
'    Call ClearControls
    GetValues
    
    EnableButtons
    
    State = True
End Sub

Private Sub cboCode_GotFocus()
Call HighlightBackground(cboCode)
End Sub

Private Sub cboCode_KeyPress(KeyAscii As Integer)
If NAV_NEW = False Then KeyAscii = 0
End Sub

Private Sub cboCode_LostFocus()
Call NormalBackground(cboCode)
End Sub

'call procedure get data and populate combo

Private Sub cboComp_Click()
On Error Resume Next

    ClearControls
    If cboComp.ListIndex <> CB_ERR Then
        cboCode.Clear
        Call PopuLateFromRecordSet(cboCode, _
            loc.LocationCodesOfSites(deIms.NameSpace, cboComp, deIms.cnIms), "Code", True)
                
        cboCode.Tag = ""
        cboComp.Tag = cboComp
        
        State = False
        cboCode.ListIndex = IIf(cboCode.ListCount > 0, 0, CB_ERR)
        
        State = True
        EnableButtons
    End If
    
'        Call ClearControls

    If Err Then Err.Clear
End Sub

Private Sub cboComp_GotFocus()
Call HighlightBackground(cboComp)
End Sub

Private Sub cboComp_KeyPress(KeyAscii As Integer)
If NAV_NEW = False Then KeyAscii = 0
End Sub

Private Sub cboComp_LostFocus()
Call NormalBackground(cboComp)
End Sub

Private Sub chkflag_GotFocus()
Call HighlightBackground(chkFlag)
End Sub

Private Sub chkflag_LostFocus()
Call NormalBackground(chkFlag)
End Sub

'call function get gender information and set button

Private Sub Form_Load()

'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar


Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
   
    cbo_LocationType.Enabled = False
   
chkFlag.Enabled = False
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
   ' TableLocked = True
    End If
End If
'end copy



On Error Resume Next

Dim str() As String
Dim i As Integer, x As Integer

    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Translate_Forms("frm_Location")
    '------------------------------------------
    starting = True
    State = False
    Set loc = New imsLocation

    'str = loc.GenderList

    'i = UBound(str)

    'For x = LBound(str) To i
        'Call cbo_LocationType.AddItem(str(x))
       
     'Next x

     Call cbo_LocationType.AddItem("SITE")
   
   
    Call PopuLateFromRecordSet(cboComp, _
        loc.CompanyList(deIms.NameSpace, deIms.cnIms), "Comp", True)

    cboComp.ListIndex = 0
    cboCode.ListIndex = 0
      Caption = Caption + " - " + Tag
    State = True
    
    If TableLocked = False Then
    
    Call DisableButtons(Me, NavBar1)
    NavBar1.NewEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    
  
  
    starting = False
    
    NAV_NEW = NavBar1.NewEnabled
    cboCode.locked = False
    cboComp.locked = False
     End If
    
    
    With frm_LocationSITE
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'enable control button

Private Sub EnableButtons()
Dim i As Integer

    i = cboCode.ListIndex
    
    If cboCode.ListCount = 0 Then
    
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
        
        
    ElseIf i = cboCode.ListCount - 1 Then
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
    
    If Err Then Err.Clear
    
    NavBar1.NewEnabled = NAV_NEW
End Sub

'call function get location values

Public Sub GetValues()
'    If AssignValues Then _
'        Set loc = loc.Find(deIms.NameSpace, cboCode, cboComp, deIms.cnIms)
    
    Set loc = loc.Find(deIms.NameSpace, cboCode, cboComp, deIms.cnIms)
    
    cboCode.Tag = cboCode
    FillTextBox
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
'    Call loc.Update(deIms.cnIms)
    Hide
    Set loc = Nothing
    If open_forms <= 5 Then ShowNavigator
    
    
'If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
' currentformname = Forms(3).Name ' 2011-7-28 Juan, this is not necessary, it is already set
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
'End If
        
    
End Sub

'set combo list to first position

Private Sub NavBar1_OnCancelClick()


  If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
   cbo_LocationType.Enabled = False
   chkFlag.Enabled = False
NavBar1.SaveEnabled = False

On Error Resume Next

    State = False
    fm = mdVisualization
    cboCode.ListIndex = CB_ERR
    cboCode.ListIndex = IndexOf(cboCode, loc.LastCode)
    EnableButtons
    
    State = True
    If Err Then Err.Clear
    
    End If
End Sub

'Close Form

Private Sub NavBar1_OnCloseClick()
    
If TableLocked = True Then    'jawdat
'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'' currentformname = Forms(3).Name ' 2011-7-28 Juan, this is not necessary, it is already set
'Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    Unload Me
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    cboCode.ListIndex = 0
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    cboCode.ListIndex = cboCode.ListCount - 1
End Sub

'set navbar button

Private Sub NavBar1_OnNewClick()
Dim str As String

 cbo_LocationType.Enabled = True
   chkFlag.Enabled = True
   NavBar1.CancelEnabled = True
   NavBar1.SaveEnabled = True
'    AssignValues
'    fm = mdCreation
'    str = cboCode.Tag
'
'    State = False
'    Set loc = loc.AddNew()
'    cboCode.ListIndex = CB_ERR
'
'    Set loc = loc.AddNew
'    cboCode.Tag = str
    
    Call ClearControls
    
    With NavBar1
        .NewEnabled = False
        .LastEnabled = False
        .NextEnabled = False
        .FirstEnabled = False
        .PreviousEnabled = False
    End With

    State = True
End Sub

'clear controls

Private Sub ClearControls()
Dim ctl As Control
        
    For Each ctl In Controls
        If TypeOf ctl Is textBOX Then
             ctl = ""
        ElseIf TypeOf ctl Is ComboBox Then
            starting = True
            If Not ctl Is cboComp Then ctl.ListIndex = CB_ERR
            starting = False
        End If
    Next ctl
    
    cboCode = ""
    chkFlag.value = vbChecked
End Sub

'set combo index

Private Sub NavBar1_OnNextClick()
On Error Resume Next
Dim i As Integer

    i = IndexOf(cboCode, cboCode.Tag)
    
    If (i < (cboCode.ListCount - 1)) Then cboCode.ListIndex = cboCode.ListIndex + 1
    
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
Dim i As Integer
On Error Resume Next

    i = IndexOf(cboCode, cboCode.Tag)
    'If (i > 1) Then cboCode.ListIndex = cboCode.ListIndex - 1
    If (i > 0) Then cboCode.ListIndex = cboCode.ListIndex - 1
    
End Sub

'fill data to text box

Private Sub FillTextBox()
    With loc
        txt_City = .City
        txt_Email = .Email
        txt_State = .State
        txt_Contact = .Contact
        txt_Country = .Country
        txt_Zipcode = .ZipCode
        txt_Address1 = .address1
        txt_Address2 = .address2
        txt_FaxNumber = .Faxnumber
        txt_SupName = .LocationName
        txt_PhoneNumber = .PhoneNumber
        chkFlag.value = IIf(.Flag, 1, 0)
        
        'cboComp.Tag = .CompanyCode
        'cboCode.Tag = .LocationCode
        cbo_LocationType.ListIndex = IndexOf(cbo_LocationType, .Gender)
    End With
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\location.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00115")
        .WindowTitle = IIf(msg1 = "", "Location", msg1)
        Call translator.Translate_Reports("location.rpt") 'J added
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

'before save check location code exist or not,if it exist show message

Private Sub NavBar1_OnSaveClick()
Dim Numb As Integer

    Numb = cboCode.ListIndex
    
    If Numb = -1 Then
        
        If Len(Trim$(cboCode)) <> 0 Then
            
            If CheckLocation(cboCode, cboComp) Then
            
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00273") 'J added
                MsgBox IIf(msg1 = "", "Location code exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            End If
            
        End If
    
    End If
    
    AssignValues
 
    cboCode.Tag = cboCode
    If IndexOf(cboCode, cboCode) = CB_ERR Then cboCode.AddItem (cboCode)
    NavBar1.NewEnabled = True
End Sub

'set values to class

Private Function AssignValues() As Boolean

    AssignValues = True
    If State = False Then Exit Function
    
    'AssignValues = loc.Validate
    
    If AssignValues = False Then Exit Function
    'If fm = mdVisualization Then Exit Function
    
    'cboCode.Tag = IIf(loc.InsertMode, cboCode, cboCode.Tag)
    
    'cboCode.Tag = IIf(Len(cboCode.Tag), cboCode.Tag, cboCode)
    'cboComp.Tag = IIf(Len(cboComp.Tag), cboComp.Tag, cboComp)
    
    
    'If loc.InsertMode Then loc.Validate
    
    With loc

        .City = txt_City
        .State = txt_State
        .Email = txt_Email
        .ZipCode = txt_Zipcode
        .Contact = txt_Contact
        .Country = txt_Country
        .CompanyCode = cboComp
        .LocationCode = cboCode
        .address1 = txt_Address1
        .address2 = txt_Address2
        .Gender = cbo_LocationType
        .Faxnumber = txt_FaxNumber
        .LocationName = txt_SupName
        .NameSpace = deIms.NameSpace
        .PhoneNumber = txt_PhoneNumber
        .Flag = chkFlag.value = vbChecked
        .User = CurrentUser
         
        If cboCode = "" Then cboCode = .LocationCode
        
        cboCode.Tag = .LocationCode
        If .DataChanged Then AssignValues = .Update(deIms.cnIms)
        
    End With
    
    If AssignValues = False Then Exit Function
    EnableButtons
End Function

'SQL statement check location code exist or not

Private Function CheckLocation(Code As String, Loca As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From Location "
        .CommandText = .CommandText & " Where loc_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND loc_locacode = '" & Code & "'"
        .CommandText = .CommandText & " AND loc_compcode = '" & Loca & "'"
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckLocation = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckLocation", Err.Description, Err.number, True)
End Function

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
Call HighlightBackground(txt_SupName)
End Sub

Private Sub txt_SupName_LostFocus()
Call NormalBackground(txt_SupName)
End Sub

Private Sub txt_Zipcode_GotFocus()
Call HighlightBackground(txt_Zipcode)
End Sub

Private Sub txt_Zipcode_LostFocus()
Call NormalBackground(txt_Zipcode)
End Sub
