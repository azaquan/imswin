VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Manufacturer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manufacturer"
   ClientHeight    =   4665
   ClientLeft      =   750
   ClientTop       =   1110
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   5340
   Tag             =   "01011700"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   4200
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frm_Manufacturer.frx":0000
      EmailEnabled    =   -1  'True
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
      DisableSaveOnSave=   0   'False
   End
   Begin VB.TextBox txtTelenumb 
      DataField       =   "man_cont"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   10
      Top             =   2880
      Width           =   3024
   End
   Begin VB.TextBox txt_SupName 
      DataField       =   "man_name"
      Height          =   288
      Left            =   2175
      MaxLength       =   35
      TabIndex        =   2
      Top             =   780
      Width           =   3024
   End
   Begin VB.TextBox txt_Country 
      DataField       =   "man_ctry"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2280
      Width           =   3024
   End
   Begin VB.TextBox txt_PhoneNumber 
      DataField       =   "man_phonnumb"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2580
      Width           =   3024
   End
   Begin VB.TextBox txt_Contact 
      DataField       =   "man_cont"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   13
      Top             =   3800
      Width           =   3024
   End
   Begin VB.TextBox txt_Email 
      DataField       =   "man_mail"
      Height          =   288
      Left            =   2175
      MaxLength       =   59
      TabIndex        =   12
      Top             =   3500
      Width           =   3024
   End
   Begin VB.TextBox txt_FaxNumber 
      DataField       =   "man_faxnumb"
      Height          =   288
      Left            =   2175
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3200
      Width           =   3024
   End
   Begin VB.TextBox txt_State 
      DataField       =   "man_stat"
      Height          =   288
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1980
      Width           =   408
   End
   Begin VB.TextBox txt_City 
      DataField       =   "man_city"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1680
      Width           =   3024
   End
   Begin VB.TextBox txt_Address2 
      DataField       =   "man_adr2"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1380
      Width           =   3024
   End
   Begin VB.TextBox txt_Address1 
      DataField       =   "man_adr1"
      Height          =   288
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1080
      Width           =   3024
   End
   Begin VB.TextBox txt_Zipcode 
      DataField       =   "man_zipc"
      Height          =   288
      Left            =   3840
      MaxLength       =   11
      TabIndex        =   7
      Top             =   1980
      Width           =   1350
   End
   Begin VB.ComboBox cbo_Code 
      DataField       =   "man_code"
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   450
      Width           =   3024
   End
   Begin VB.Label LblTelex 
      BackStyle       =   0  'Transparent
      Caption         =   "TelexNumber"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   2000
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
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
      Top             =   780
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Tag             =   "01011700"
      Top             =   3240
      Width           =   2000
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   3540
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   288
      Left            =   120
      TabIndex        =   22
      Top             =   2580
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   288
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   2000
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   288
      Left            =   120
      TabIndex        =   19
      Top             =   1980
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   285
      Left            =   2715
      TabIndex        =   20
      Top             =   1995
      Width           =   1125
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   288
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   288
      Left            =   120
      TabIndex        =   17
      Top             =   1380
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   288
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   2000
   End
   Begin VB.Label lbl_Manufacturer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
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
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   5070
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
      Height          =   204
      Left            =   120
      TabIndex        =   14
      Top             =   456
      Width           =   2000
   End
End
Attribute VB_Name = "frm_Manufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Manulist As imsManufacturer
Dim mIsComboLoaded  As Boolean
Dim mIsItANewRecord As Boolean
Dim TableLocked As Boolean, currentformname As String   'jawdat
Private Sub cbo_Code_Click()
Dim cn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim str As String
    
'    cbo_Code.locked = False
    
    If Len(cbo_Code) Then
    
        Set Manulist = Manulist.GetManufacturerlist(cbo_Code, deIms.NameSpace, deIms.cnIms)
        
        FillTextBox
        EnableButtons

    End If
    mIsItANewRecord = False
'    cbo_Code.locked = True
End Sub
Public Sub EnableButtons()
Dim i As Integer

    i = cbo_Code.ListIndex
    
    If cbo_Code.ListCount = 0 Then
    
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
        
        Exit Sub
        
    ElseIf i = cbo_Code.ListCount - 1 Then
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
    
    If cbo_Code.ListIndex = CB_ERR Then cbo_Code.ListIndex = 0
    If Err Then Err.Clear
End Sub


Private Sub cbo_Code_DropDown()
Dim str As String
If mIsComboLoaded = False Then
    str = cbo_Code
    cbo_Code.Clear
    If Manulist Is Nothing Then Set Manulist = New imsManufacturer
        Call PopuLateFromRecordSet(cbo_Code, Manulist.GetManufacturerCode(deIms.NameSpace, deIms.cnIms), "man_code", True)
    mIsComboLoaded = True
     cbo_Code = str
End If
End Sub

Private Sub cbo_Code_GotFocus()
Call HighlightBackground(cbo_Code)
End Sub

Private Sub cbo_Code_KeyPress(KeyAscii As Integer)

If mIsItANewRecord = True Then

If Len(cbo_Code) > 10 And Not KeyAscii = 8 Then
   MsgBox "Code can not be more than 10 characters long.", vbInformation, "Imswin"
   KeyAscii = 0
   cbo_Code.SetFocus
End If

Else
 KeyAscii = 0
End If

End Sub

Private Sub cbo_Code_LostFocus()
Call NormalBackground(cbo_Code)
End Sub

Private Sub cbo_Code_Validate(Cancel As Boolean)

If mIsItANewRecord = True Then
       If Len(Trim$(cbo_Code)) <> 0 Then
            If CheckManuCode(cbo_Code) Then
            
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00277") 'J added
                MsgBox IIf(msg1 = "", "Manufacturer code exist, please Use a Different one", msg1) 'J modified
                '---------------------------------------------
                Cancel = True
                Exit Sub
            End If
        End If
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
    
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
End Sub

Private Sub NavBar1_OnCancelClick()
    mIsItANewRecord = False
    Call CleanForm
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

    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Translate_Forms("frm_Manufacturer")
    '------------------------------------------
    
    Set Manulist = New imsManufacturer
    Call PopuLateFromRecordSet(cbo_Code, Manulist.GetManufacturerCode(deIms.NameSpace, deIms.cnIms), "man_code", True)
    If cbo_Code.ListCount Then cbo_Code.ListIndex = 0
    
    cbo_Code.locked = False
    
    Caption = Caption + " - " + Tag
    mIsComboLoaded = True
    
      Call DisableButtons(Me, NavBar1)
    
    
    With frm_Manufacturer
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub NavBar1_OnFirstClick()
    cbo_Code.ListIndex = 0
End Sub

Private Sub NavBar1_OnLastClick()
    cbo_Code.ListIndex = cbo_Code.ListCount - 1
End Sub

Private Sub NavBar1_OnNewClick()
    Call PopuLateFromRecordSet(cbo_Code, Manulist.GetManufacturerCode(deIms.NameSpace, deIms.cnIms), "man_code", True)
   Call CleanForm
    'cbo_Code.locked = False
    mIsItANewRecord = True
End Sub

Private Sub NavBar1_OnNextClick()
    cbo_Code.ListIndex = cbo_Code.ListIndex + 1
End Sub

Private Sub NavBar1_OnPreviousClick()
    If cbo_Code.ListIndex = -1 Then
        Exit Sub
    Else
        cbo_Code.ListIndex = cbo_Code.ListIndex - 1
    End If
    
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Manu.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00101") 'J added
        .WindowTitle = IIf(msg1 = "", "Manufacturer", msg1) 'J modified
        Call translator.Translate_Reports("Manu.rpt") 'J added
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


Public Sub FillTextBox()
    txt_SupName = Manulist.ManufName
    txt_Address1 = Manulist.address1
    txt_Address2 = Manulist.address2
    txt_City = Manulist.City
    txt_State = Manulist.State
    txt_Zipcode = Manulist.ZipCode
    txt_Country = Manulist.Country
    txt_PhoneNumber = Manulist.PhoneNumb
    txtTelenumb = Manulist.Telexnumb
    txt_FaxNumber = Manulist.Faxnumb
    txt_Email = Manulist.Email
    txt_Contact = Manulist.Contact

End Sub

Public Sub InsertManufacturer()
    On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "UP_INS_MANUFACTURER"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, cbo_Code)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 35, txt_SupName)
        .parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, txt_Address1)
        .parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, txt_Address2)
        .parameters.Append .CreateParameter("@city", adVarChar, adParamInput, 25, txt_City)
        .parameters.Append .CreateParameter("@stat", adVarChar, adParamInput, 2, txt_State)
        .parameters.Append .CreateParameter("@zipc", adVarChar, adParamInput, 11, txt_Zipcode)
        .parameters.Append .CreateParameter("@ctry", adVarChar, adParamInput, 25, txt_Country)
        .parameters.Append .CreateParameter("@phonnumb", adVarChar, adParamInput, 25, txt_PhoneNumber)
        .parameters.Append .CreateParameter("@faxnumb", adVarChar, adParamInput, 50, txt_FaxNumber)
        .parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, txtTelenumb)
        .parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, txt_Email)
        .parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, txt_Contact)
        .parameters.Append .CreateParameter("@user", adVarChar, adParamInput, 20, CurrentUser)
        .Execute , , adExecuteNoRecords
    
    End With
    
    Set cmd = Nothing
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00275") 'J added
        MsgBox IIf(msg1 = "", "Insert into Manufacturer is completed", msg1) 'J modified
        '---------------------------------------------
        
    Exit Sub
    
Noinsert:

        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00276") 'J added
        MsgBox IIf(msg1 = "", "Insert into Manufacturer is failure ", msg1) 'J modified
        '---------------------------------------------
        
End Sub

Private Sub NavBar1_OnSaveClick()
Dim num As Integer
Dim Numb As Integer
    
    If Len(Trim$(cbo_Code)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cbo_Code.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_SupName)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_SupName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'j added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    


    Numb = cbo_Code.ListIndex
    
    If Numb = -1 Then
        If Len(Trim$(cbo_Code)) <> 0 Then
            If CheckManuCode(cbo_Code) Then
            
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00277") 'J added
                MsgBox IIf(msg1 = "", "Manufacturer code exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            End If
        End If
    End If
    
    Call InsertManufacturer
    mIsItANewRecord = False
    mIsComboLoaded = False
End Sub

Public Sub CleanForm()
    cbo_Code = ""
    txt_SupName = ""
    txt_Address1 = ""
    txt_Address2 = ""
    txt_City = ""
    txt_State = ""
    txt_Zipcode = ""
    txt_Country = ""
    txt_PhoneNumber = ""
    txtTelenumb = ""
    txt_FaxNumber = ""
    txt_Email = ""
    txt_Contact = ""

End Sub

Private Function CheckManuCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From MANUFACTURER "
        .CommandText = .CommandText & " Where man_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND man_code = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckManuCode = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckManuCode", Err.Description, Err.number, True)
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

Private Sub txtTelenumb_GotFocus()
Call HighlightBackground(txtTelenumb)
End Sub

Private Sub txtTelenumb_LostFocus()
Call NormalBackground(txtTelenumb)
End Sub
