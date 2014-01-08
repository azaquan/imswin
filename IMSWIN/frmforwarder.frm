VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmForwarder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forwarder"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5520
   Tag             =   "01011200"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   27
      Top             =   4560
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frmforwarder.frx":0000
      DeleteVisible   =   -1  'True
      EmailEnabled    =   -1  'True
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin VB.TextBox txt_Email 
      Height          =   288
      Left            =   2205
      MaxLength       =   59
      TabIndex        =   12
      Top             =   3660
      Width           =   3072
   End
   Begin VB.TextBox txtTelexnumber 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   10
      Top             =   3360
      Width           =   3072
   End
   Begin VB.ComboBox Cmbforwarder 
      Height          =   315
      Left            =   2175
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3100
   End
   Begin VB.TextBox txt_Country 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   7
      Top             =   2460
      Width           =   3072
   End
   Begin VB.TextBox txt_PhoneNumber 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2760
      Width           =   3072
   End
   Begin VB.TextBox txt_FaxNumber 
      Height          =   288
      Left            =   2205
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3060
      Width           =   3072
   End
   Begin VB.TextBox txt_Contact 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   13
      Top             =   3960
      Width           =   3072
   End
   Begin VB.TextBox txt_ForwarderName 
      Height          =   288
      Left            =   2205
      MaxLength       =   35
      TabIndex        =   1
      Top             =   940
      Width           =   3072
   End
   Begin VB.TextBox txt_City 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1860
      Width           =   3072
   End
   Begin VB.TextBox txt_Address2 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1560
      Width           =   3072
   End
   Begin VB.TextBox txt_Address1 
      Height          =   288
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1250
      Width           =   3072
   End
   Begin VB.TextBox txt_State 
      Height          =   288
      Left            =   2205
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2160
      Width           =   528
   End
   Begin VB.TextBox txt_Zipcode 
      Height          =   288
      Left            =   4080
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      Height          =   210
      Left            =   240
      TabIndex        =   26
      Top             =   3660
      Width           =   2000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telex Number"
      Height          =   210
      Left            =   240
      TabIndex        =   25
      Top             =   3360
      Width           =   2000
   End
   Begin VB.Label lbl_ForwarderName 
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarder Name"
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
      TabIndex        =   24
      Top             =   940
      Width           =   2000
   End
   Begin VB.Label lbl_Address1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   210
      Left            =   240
      TabIndex        =   23
      Top             =   1250
      Width           =   2000
   End
   Begin VB.Label lbl_Address2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address(Line2)"
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   2000
   End
   Begin VB.Label lbl_ForwarderCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarder Code"
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
      TabIndex        =   21
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label lbl_City 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   1860
      Width           =   2000
   End
   Begin VB.Label lbl_Zip 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   210
      Left            =   2925
      TabIndex        =   19
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label lbl_State 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   210
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   2000
   End
   Begin VB.Label lbl_Country 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   2460
      Width           =   2000
   End
   Begin VB.Label lbl_PhoneNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #"
      Height          =   210
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   2000
   End
   Begin VB.Label lbl_FaxNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
      Height          =   210
      Left            =   240
      TabIndex        =   15
      Top             =   3060
      Width           =   2000
   End
   Begin VB.Label lbl_Contact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   2000
   End
   Begin VB.Label lbl_Forwarder 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarder"
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
      Left            =   300
      TabIndex        =   11
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim flist As imsForwarder

'call function get data and populate text boxes

Private Sub Cmbforwarder_Click()
Dim cn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim str As String

    If Len(Cmbforwarder) Then
    
        Set flist = flist.GetForwarder(Cmbforwarder, deIms.NameSpace, deIms.cnIms)
        
        FillTextBox
        EnableButtons

    End If
End Sub

'call function get data and populate forwarder list combo

Private Sub Form_Load()
Dim cmd As ADODB.Command
Dim cn As ADODB.Connection

    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmForwarder")
    '------------------------------------------

    Set flist = New imsForwarder
    Call PopuLateFromRecordSet(Cmbforwarder, flist.GetForwarderCode(deIms.NameSpace, deIms.cnIms), "forw_code", True)
    If Cmbforwarder.ListCount Then Cmbforwarder.ListIndex = 0
    'Call DisableButtons(Me, NavBar1)
    
    frmForwarder.Caption = frmForwarder.Caption + " - " + frmForwarder.Tag
End Sub

'SQL statement get forwarder list and fill text boxes

Public Function GetForwarderList(Code As String, NameSpace As String) As imsForwarder
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT forw_name, forw_adr1, "
        .CommandText = .CommandText & " forw_adr2, forw_city, "
        .CommandText = .CommandText & " forw_stat, forw_zipc, "
        .CommandText = .CommandText & " forw_ctry, forw_phonnumb, "
        .CommandText = .CommandText & " forw_faxnumb, forw_mail, "
        .CommandText = .CommandText & " forw_telxnumb, forw_cont "
        .CommandText = .CommandText & " From FORWARDER "
        .CommandText = .CommandText & " WHERE forw_code =  '" & Code & "'"
        .CommandText = .CommandText & " AND forw_npecode = '" & NameSpace & "'"
        
        Set rst = .Execute
    End With
        
        txt_ForwarderName = rst!forw_name
        txt_Address1 = rst!forw_adr1
        txt_Address2 = rst!forw_adr2
        txt_City = rst!forw_city
        txt_State = rst!forw_stat
        txt_Zipcode = rst!forw_zipc
        txt_Country = rst!forw_ctry
        txt_PhoneNumber = rst!forw_phonnumb
        txt_FaxNumber = rst!forw_faxnumb
        txtTelexnumber = rst!forw_telxnumb
        txt_Email = rst!forw_mail
        txt_Contact = rst!forw_cont
    
End Function

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set flist = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'move recordset to previous position

Private Sub NavBar1_OnCancelClick()
    Call NavBar1_OnPreviousClick
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    Cmbforwarder.ListIndex = 0
End Sub

'move recordset to list position

Private Sub NavBar1_OnLastClick()
    Cmbforwarder.ListIndex = Cmbforwarder.ListCount - 1
End Sub

'set navbar buttom

Public Sub EnableButtons()
Dim i As Integer

    i = Cmbforwarder.ListIndex
    
    If Cmbforwarder.ListCount = 0 Then
    
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
        
        Exit Sub
        
    ElseIf i = Cmbforwarder.ListCount - 1 Then
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
    
    If Cmbforwarder.ListIndex = CB_ERR Then Cmbforwarder.ListIndex = 0
    If Err Then Err.Clear
End Sub

'clear form

Private Sub NavBar1_OnNewClick()
        
    Call PopuLateFromRecordSet(Cmbforwarder, flist.GetForwarderCode(deIms.NameSpace, deIms.cnIms), "forw_code", True)
    'If Cmbforwarder.ListCount Then Cmbforwarder.ListIndex = 0
    
        Cmbforwarder = ""
        txt_ForwarderName = ""
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

End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
On Error Resume Next

    Cmbforwarder.ListIndex = Cmbforwarder.ListIndex + 1
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
    If Cmbforwarder.ListIndex = -1 Then
        Exit Sub
    Else
        Cmbforwarder.ListIndex = Cmbforwarder.ListIndex - 1
    End If
End Sub

'call store procedure to insert a record to database

Private Sub InsertForwarder()
On Error GoTo Noinsert
Dim cmd As ADODB.Command
Dim va As Variant


    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "UP_INS_FORWARDER"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .Parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 10, Cmbforwarder)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 35, txt_ForwarderName)
        
         va = txt_Address1
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@adr1", adVarChar, adParamInput, 25, va)
        
         va = txt_Address2
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@adr2", adVarChar, adParamInput, 25, va)
        .Parameters.Append .CreateParameter("@city", adVarChar, adParamInput, 25, txt_City)
         
         va = txt_State
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@stat", adVarChar, adParamInput, 2, va)
        
        va = txt_Zipcode
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@zipc", adVarChar, adParamInput, 11, va)
        
        va = txt_Country
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@ctry", adVarChar, adParamInput, 25, va)
        
        .Parameters.Append .CreateParameter("@phonnumb", adVarChar, adParamInput, 25, txt_PhoneNumber)
        
        va = txt_Country
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@faxnumb", adVarChar, adParamInput, 50, va)
        
        va = txtTelexnumber
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@telxnumb", adVarChar, adParamInput, 25, va)
        
        va = txt_Email
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@mail", adVarChar, adParamInput, 59, va)
        
        va = txt_Contact
        If Len(va) = 0 Then va = Null
        .Parameters.Append .CreateParameter("@cont", adVarChar, adParamInput, 25, va)
        .Parameters.Append .CreateParameter("@user", adVarChar, adParamInput, 20, CurrentUser)
        Call .Execute(Options:=adExecuteNoRecords)
    
    End With
    
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00357") 'J added
    MsgBox IIf(msg1 = "", "Insert into Forwarder is completed successfully ", msg1) 'J modified
    '---------------------------------------------
    
    Set cmd = Nothing
    Exit Sub
    
Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00358") 'J added
        MsgBox IIf(msg1 = "", "Insert into Forwarder is failure ", msg1) 'J modified
        '---------------------------------------------
        
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\forwarder.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00080") 'J added
        .WindowTitle = IIf(msg1 = "", "Forwarder", msg1) 'J modified
        Call translator.Translate_Reports("forwarder.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With

Handler:
    If Err Then MsgBox Err.Description: Err.Clear
End Sub

'before save to validate data format

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
Dim Numb As Integer

    If Len(Trim$(Cmbforwarder)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00014") 'J added
        MsgBox IIf(msg1 = "", "The Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        Cmbforwarder.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_ForwarderName)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00001") 'J added
        MsgBox IIf(msg1 = "", "The Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        txt_ForwarderName.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_City)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00005") 'J added
        MsgBox IIf(msg1 = "", "The City cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_City.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txt_PhoneNumber)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00011") 'J added
        MsgBox IIf(msg1 = "", "The Phone Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txt_PhoneNumber.SetFocus: Exit Sub
    End If
    
    Numb = Cmbforwarder.ListIndex
     
    If Numb = -1 Then
        If Len(Trim$(Cmbforwarder)) <> 0 Then
            If CheckForwarder(Cmbforwarder) Then
            
                'Modified by Juan (9/15/2000) for Multilingual
                msg1 = translator.Trans("M00310") 'J added
                MsgBox IIf(msg1 = "", "Phone Directory exist, please make new one", msg1) 'J modified
                '---------------------------------------------
                
                Exit Sub
            End If
        End If
    End If
    
    Call InsertForwarder
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub

'fill text boxse

Private Sub FillTextBox()
    txt_ForwarderName = flist.forwname
    txt_Address1 = flist.address1
    txt_Address2 = flist.address2
    txt_City = flist.City
    txt_State = flist.State
    txt_Zipcode = flist.ZipCode
    txt_Country = flist.contury
    txt_PhoneNumber = flist.phonnumb
    txt_FaxNumber = flist.Faxnumb
    txt_Email = flist.Email
    txt_Contact = flist.Contact
    txtTelexnumber = flist.Telexnumb
End Sub

'SQL statement check forwarder exist or not

Private Function CheckForwarder(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From forwarder "
        .CommandText = .CommandText & " Where forw_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND forw_code = '" & Code & "'"
        
        
'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckForwarder = rst!rt
    End With
        

    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckForwarder", Err.Description, Err.number, True)
End Function
