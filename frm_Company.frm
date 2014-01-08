VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Company 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company"
   ClientHeight    =   5475
   ClientLeft      =   750
   ClientTop       =   1065
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   5220
   Tag             =   "01030800"
   Begin VB.CheckBox chkflag 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   4320
      Width           =   255
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   750
      TabIndex        =   27
      Top             =   4800
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frm_Company.frx":0000
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
      DisableSaveOnSave=   0   'False
   End
   Begin VB.TextBox txtTelexnumber 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   10
      Top             =   3120
      Width           =   3072
   End
   Begin VB.TextBox txt_Zipcode 
      Height          =   288
      Left            =   3780
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2220
      Width           =   1332
   End
   Begin VB.TextBox txt_State 
      Height          =   288
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2220
      Width           =   528
   End
   Begin VB.TextBox txt_Address1 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1310
      Width           =   3072
   End
   Begin VB.TextBox txt_Address2 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1620
      Width           =   3072
   End
   Begin VB.TextBox txt_City 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1920
      Width           =   3072
   End
   Begin VB.TextBox txt_CompanyName 
      Height          =   288
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1000
      Width           =   3072
   End
   Begin VB.ComboBox cbo_CompanyCode 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   670
      Width           =   3075
   End
   Begin VB.TextBox txt_Contact 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   13
      Top             =   4020
      Width           =   3072
   End
   Begin VB.TextBox txt_Email 
      Height          =   288
      Left            =   2040
      MaxLength       =   59
      TabIndex        =   12
      Top             =   3720
      Width           =   3072
   End
   Begin VB.TextBox txt_FaxNumber 
      Height          =   288
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3420
      Width           =   3072
   End
   Begin VB.TextBox txt_PhoneNumber 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2820
      Width           =   3072
   End
   Begin VB.TextBox txt_Country 
      Height          =   288
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2520
      Width           =   3072
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Flag"
      Height          =   210
      Left            =   315
      TabIndex        =   29
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telex Number"
      Height          =   210
      Left            =   315
      TabIndex        =   23
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label lbl_Company 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4980
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   210
      Left            =   315
      TabIndex        =   26
      Top             =   4005
      Width           =   900
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   210
      Left            =   315
      TabIndex        =   25
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   210
      Left            =   315
      TabIndex        =   24
      Top             =   3405
      Width           =   780
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   210
      Left            =   315
      TabIndex        =   22
      Top             =   2850
      Width           =   1140
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   210
      Left            =   315
      TabIndex        =   21
      Top             =   2580
      Width           =   780
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   210
      Left            =   315
      TabIndex        =   19
      Top             =   2235
      Width           =   540
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   210
      Left            =   2760
      TabIndex        =   20
      Top             =   2235
      Width           =   1005
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   210
      Left            =   315
      TabIndex        =   18
      Top             =   1980
      Width           =   540
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
      Height          =   210
      Left            =   315
      TabIndex        =   14
      Top             =   765
      Width           =   1380
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   210
      Left            =   315
      TabIndex        =   17
      Top             =   1665
      Width           =   1140
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   210
      Left            =   315
      TabIndex        =   16
      Top             =   1350
      Width           =   660
   End
   Begin VB.Label lbl_CompanyName 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   315
      TabIndex        =   15
      Top             =   1050
      Width           =   1380
   End
End
Attribute VB_Name = "frm_Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clist As imsCompany
Dim rstcomp As ADODB.Recordset
'Dim WithEvents rstcomp As clist.GetCompanyCode
Dim TableLocked As Boolean, currentformname As String   'jawdat



Private Sub cbo_CompanyCode_Click()
Dim cn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim str As String

    If Len(cbo_CompanyCode) Then
    
         Set clist = clist.GetCompanyList(cbo_CompanyCode, deIms.NameSpace, deIms.cnIms)
        
        FillTextBox
        EnableButtons

    End If
End Sub




Private Sub cbo_CompanyCode_GotFocus()
Call HighlightBackground(cbo_CompanyCode)
End Sub

Private Sub cbo_CompanyCode_LostFocus()
Call NormalBackground(cbo_CompanyCode)
End Sub

Private Sub chkflag_GotFocus()

Call NormalBackground(chkflag)
End Sub

Private Sub chkflag_LostFocus()

Call NormalBackground(chkflag)
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
   chkflag.Enabled = False
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
'Dim rst As ADODB.Recordset
'
    
    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Company")
    '------------------------------------------
    
    Set clist = New imsCompany
    Call PopuLateFromRecordSet(cbo_CompanyCode, GetCompanyCode(deIms.NameSpace, deIms.cnIms), "com_compcode", True)
    Set rstcomp = GetCompanyCode(deIms.NameSpace, deIms.cnIms)
    If cbo_CompanyCode.ListCount Then cbo_CompanyCode.ListIndex = 0

    Caption = Caption + " - " + Tag
    
  
    Call DisableButtons(Me, NavBar1)
    
    With frm_Company
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Public Function GetCompanyListRecord(Code As String, NameSpace As String) As imsCompany
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT com_compcode,com_name, com_adr1, "
        .CommandText = .CommandText & " com_adr2, com_city, "
        .CommandText = .CommandText & " com_stat, com_zipc, "
        .CommandText = .CommandText & " com_ctry, com_phonnumb, "
        .CommandText = .CommandText & " com_faxnumb, com_mail, "
        .CommandText = .CommandText & " com_telxnumb, com_cont, com_actvflag "
        .CommandText = .CommandText & " From COMPANY "
        .CommandText = .CommandText & " WHERE com_compcode =  '" & Code & "'"
        .CommandText = .CommandText & " AND com_npecode = '" & NameSpace & "'"
        
        Set rst = .Execute
    End With
        cbo_CompanyCode = rst!com_compcode & ""
        txt_CompanyName = rst!com_name & ""
        txt_Address1 = rst!com_adr1 & ""
        txt_Address2 = rst!com_adr2 & ""
        txt_City = rst!com_city & ""
        txt_State = rst!com_stat & ""
        txt_Zipcode = rst!com_zipc & ""
        txt_Country = rst!com_ctry & ""
        txt_PhoneNumber = rst!com_phonnumb & ""
        txt_FaxNumber = rst!com_faxnumb & ""
        txtTelexnumber = rst!com_telxnumb & ""
        txt_Email = rst!com_mail & ""
        txt_Contact = rst!com_cont & ""
        chkflag.value = IIf(rst!com_actvflag, 1, 0)
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set clist = Nothing
    If open_forms <= 5 Then ShowNavigator
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
End Sub

Private Sub NavBar1_OnCancelClick()
    Call NavBar1_OnPreviousClick
End Sub

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
    cbo_CompanyCode.Enabled = False
End Sub

Private Sub NavBar1_OnFirstClick()
    cbo_CompanyCode.ListIndex = 0
End Sub

Private Sub NavBar1_OnLastClick()
    cbo_CompanyCode.ListIndex = cbo_CompanyCode.ListCount - 1
End Sub

Public Sub EnableButtons()
Dim i As Integer

    i = cbo_CompanyCode.ListIndex
    
    If cbo_CompanyCode.ListCount = 0 Then
    
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
        
        Exit Sub
        
    ElseIf i = cbo_CompanyCode.ListCount - 1 Then
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
    
    If cbo_CompanyCode.ListIndex = CB_ERR Then cbo_CompanyCode.ListIndex = 0
    If Err Then Err.Clear
End Sub

Private Sub NavBar1_OnNewClick()

   
    Call PopuLateFromRecordSet(cbo_CompanyCode, clist.GetCompanyCode(deIms.NameSpace, deIms.cnIms), "com_compcode", True)
    Call CleanForm
    
'    If Len(Trim$(cbo_CompanyCode)) <> 0 Then
'        If CheckCompanyCode(cbo_CompanyCode) Then
'            MsgBox "Company Code exist, Please make new one"
'            Exit Sub
'        End If
'    End If
End Sub

Private Sub NavBar1_OnNextClick()
On Error Resume Next

    cbo_CompanyCode.ListIndex = cbo_CompanyCode.ListIndex + 1
End Sub

Private Sub NavBar1_OnPreviousClick()
    If cbo_CompanyCode.ListIndex = -1 Then
        Exit Sub
    Else
        cbo_CompanyCode.ListIndex = cbo_CompanyCode.ListIndex - 1
    End If
End Sub


Private Sub Insertcomarder()
On Error GoTo Noinsert
Dim cmd As ADODB.Command
Dim v As Variant

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "UP_INS_COMPANY"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, cbo_CompanyCode)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 50, txt_CompanyName)
        .parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, txt_Address1)
        
        v = txt_Address2
        If Len(Trim$(txt_Address2)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, v)
         
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
        
        v = txtTelexnumber
        If Len(Trim$(txtTelexnumber)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, v)
        
        v = txt_Email
        If Len(Trim$(txt_Email)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, v)
        
         v = txt_Contact
        If Len(Trim$(txt_Contact)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, v)
        .parameters.Append .CreateParameter("@CreatedBY", adVarChar, adParamInput, 20, CurrentUser)
        .parameters.Append .CreateParameter("@flag", adBoolean, adParamInput, , chkflag.value = vbChecked)
        .Execute , , adExecuteNoRecords
    
    End With
    
    Set cmd = Nothing
    If IndexOf(cbo_CompanyCode, cbo_CompanyCode) = CB_ERR Then
        cbo_CompanyCode.AddItem (cbo_CompanyCode)
    End If
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00008") 'J added
        MsgBox IIf(msg1 = "", "Insert into company was completed", msg1) 'J modified
        '---------------------------------------------
        
    Exit Sub
    
Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/11/2000) for Multilanguage
        msg1 = translator.Trans("M00009") 'J added
        MsgBox IIf(msg1 = "", "Insert into company is failure ", msg1) 'J modified
        '----------------------------------------------
        
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\company.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00041") 'J added
        .WindowTitle = IIf(msg1 = "", "Company", msg1) 'J modified
        Call translator.Translate_Reports("company.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
        
    End With

Handler:
    If Err Then MsgBox Err.Description: Err.Clear
End Sub

Private Sub NavBar1_OnSaveClick()
Dim Numb As Integer

    If Len(Trim$(cbo_CompanyCode)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cbo_CompanyCode.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_CompanyName)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_CompanyName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Address1)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00004") 'J added
        MsgBox IIf(msg1 = "", "The address cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Address1.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_Country)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "The Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_Country.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    
    Numb = cbo_CompanyCode.ListIndex
     
    If Numb = -1 Then
        If Len(Trim$(cbo_CompanyCode)) <> 0 Then
            If CheckCode(cbo_CompanyCode) Then
            
                'Modified by Juan (9/11/2000) for Multilingual
                msg1 = translator.Trans("M00010") 'J added
                MsgBox IIf(msg1 = "", "Company Code exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            End If
        End If
    End If
    
    Call Insertcomarder
End Sub

Private Sub FillTextBox()
    txt_CompanyName = clist.companyNAME
    txt_Address1 = clist.address1
    txt_Address2 = clist.address2
    txt_City = clist.City
    txt_State = clist.State
    txt_Zipcode = clist.ZipCode
    txt_Country = clist.Country
    txt_PhoneNumber = clist.PhoneNumb
    txt_FaxNumber = clist.Faxnumb
    txt_Email = clist.Email
    txt_Contact = clist.Contact
    txtTelexnumber = clist.Telexnumb
    chkflag.value = IIf(clist.Actvflag, 1, 0)
End Sub


Public Sub CleanForm()
    cbo_CompanyCode = ""
    txt_CompanyName = ""
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
    txt_Contact = ""
    chkflag.value = vbChecked
End Sub

'Private Function CheckCompanyCode(code As String) As Boolean
'On Error Resume Next
'Dim cmd As ADODB.Command
'Dim rst As ADODB.Recordset
'
'    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
'
'    With cmd
'        .CommandText = "SELECT count(*) RT"
'        .CommandText = .CommandText & " From company "
'        .CommandText = .CommandText & " Where com_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " AND com_compcode = '" & code & "'"
'
''        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
'
'        Set rst = .Execute
'        CheckCompanyCode = rst!rt
'    End With
'
'
'    Set cmd = Nothing
'    Set rst = Nothing
'    If Err Then Call LogErr(Name & "::CheckCompanyCode", Err.Description, Err.number, True)
'End Function

Public Function GetCompanyCode(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim rstcomp As ADODB.Recordset


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select com_compcode from COMPANY"
        .CommandText = .CommandText & " where com_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " ORDER BY com_compcode"
        
         Set rstcomp = .Execute
         Set GetCompanyCode = rstcomp
    End With
    
End Function

Private Function CheckCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From company "
        .CommandText = .CommandText & " Where com_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND com_compcode = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckCode = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckCode", Err.Description, Err.number, True)
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

Private Sub txt_CompanyName_GotFocus()
Call HighlightBackground(txt_CompanyName)
End Sub

Private Sub txt_CompanyName_LostFocus()
Call NormalBackground(txt_CompanyName)
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

Private Sub txt_Zipcode_GotFocus()
Call HighlightBackground(txt_Zipcode)

End Sub

Private Sub txt_Zipcode_LostFocus()
Call NormalBackground(txt_Zipcode)
End Sub

Private Sub txtTelexnumber_GotFocus()
Call HighlightBackground(txtTelexnumber)
End Sub

Private Sub txtTelexnumber_LostFocus()
Call NormalBackground(txtTelexnumber)
End Sub
