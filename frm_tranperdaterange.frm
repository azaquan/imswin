VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_tranperdaterange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction per Date Range"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   4710
   Tag             =   "03030300"
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
      Width           =   1092
   End
   Begin MSComCtl2.DTPicker dtdate2 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20447233
      CurrentDate     =   36524
   End
   Begin MSComCtl2.DTPicker DTdate1 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20447235
      CurrentDate     =   36524
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_company 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      DividerStyle    =   0
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_ware 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2487
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4551
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1260
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label lbl_todate 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1980
      Width           =   2000
   End
   Begin VB.Label lbl_fromdate 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1620
      Width           =   2000
   End
   Begin VB.Label lbl_ware 
      Caption         =   "Location"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label lbl_company 
      Caption         =   "Company"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   2000
   End
End
Attribute VB_Name = "frm_tranperdaterange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHndlr
If (DTdate1.value > dtdate2.value) Then

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("M00003") 'J added
    msg2 = translator.Trans("L00318") 'J modified
    MsgBox IIf(msg1 = "", "Make sure the 'From date' is less than the 'To date'", msg1), , IIf(msg2 = "", "Date", msg2) 'J modified
    '---------------------------------------------
    
    DTdate1_Validate ("true")
    Else
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\transperdate.rpt"
        .ParameterFields(2) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(0) = "company;" + Trim$(SSOleDB_company.Text) + ";true"
        .ParameterFields(1) = "ware;" + Trim$(SSOleDB_ware.Text) + ";true"
        .ParameterFields(3) = "date1;date(" & Year(DTdate1.value) & "," & Month(DTdate1.value) & "," & Day(DTdate1.value) & ");true"
        .ParameterFields(4) = "date2;date(" & Year(dtdate2.value) & "," & Month(dtdate2.value) & "," & Day(dtdate2.value) & ");true"
        '.ParameterFields(5) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(5) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00183") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction per Date range", msg1) 'J modified
        Call translator.Translate_Reports("transperdate.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1
        .Reset
 End With
 End If
 Exit Sub
 
ErrHndlr:
 MsgBox Err.Description, , "Imswin"
 Err.Clear
End Sub

Private Sub DTdate1_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub

'SQL statement get company information, populate data grid

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_tranperdaterange")
    '------------------------------------------
    
'Me.Height = 3400
'Me.Width = 5000
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
SSOleDB_company.Text = "ALL"
SSOleDB_ware.Text = "ALL"
SSOleDB_company.FieldSeparator = Chr$(1)
SSOleDB_ware.FieldSeparator = Chr$(1)

    With rs
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "'"
        .Source = .Source & " order by com_compcode "
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
If get_status(rs) Then
SSOleDB_company.AddItem "ALL" & Chr$(1) & "ALL"
Do While (Not rs.EOF)
SSOleDB_company.AddItem (rs!com_compcode & Chr$(1) & rs!com_name & " ")
rs.MoveNext
Loop
Set rs = Nothing
End If


    Call GetCurrencylist
    SSOleDBCurrency = "USD"
'rs1.Source = "Select loc_locacode,loc_name from location where loc_npecode='" & deIms.NameSpace & "'"
'rs1.ActiveConnection = deIms.cnIms
'rs1.Open
'If get_status(rs1) Then
'Do While (Not rs1.EOF)
'SSOleDB_ware.AddItem (rs1!loc_locacode & Chr$(1) & rs1!loc_name & " ")
'rs1.MoveNext
'Loop
'Set rs1 = Nothing
'End If

    DTdate1.value = FirstOfMonth
    dtdate2.value = Now
    
    frm_tranperdaterange.Caption = frm_tranperdaterange.Caption + " - " + frm_tranperdaterange.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub


'SQL statement get location list for location combo

Private Sub GetlocationName(Company As String)
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT loc_locacode,loc_name "
        .CommandText = .CommandText & " From location "
        .CommandText = .CommandText & " WHERE loc_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and loc_compcode = '" & Company & "'"
        .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDB_ware.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    
    SSOleDB_ware.RemoveAll
    
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSOleDB_ware.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetlocationName", Err.Description, Err.number, True)
End Sub

'SQL statement get all location list for location combo

Private Sub GetalllocationName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT loc_locacode,loc_name "
        .CommandText = .CommandText & " From location "
        .CommandText = .CommandText & " WHERE loc_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDB_ware.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDB_ware.RemoveAll
    
    rst.MoveFirst
      
    SSOleDB_ware.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDB_ware.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub

'SQL statement get all currency list for currency combo

Private Sub GetCurrencylist()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT curr_code, curr_desc "
        .CommandText = .CommandText & " FROM CURRENCY "
        .CommandText = .CommandText & " WHERE curr_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by curr_code"
         Set rst = .Execute
    End With


    str = Chr$(1)
    SSOleDBCurrency.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDBCurrency.RemoveAll
    
    rst.MoveFirst
      
'    SSOleDBCurrency.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCurrency.AddItem rst!curr_code & str & (rst!curr_desc & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::Getcurrencylist", Err.Description, Err.number, True)
End Sub



' check recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
 End Function

'resize form

Private Sub Form_Resize()
If Not (Me.WindowState = vbMinimized) Then
'Me.Height = 3400
'Me.Width = 5000
End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'call function get company location

Private Sub SSOleDB_company_Click()
Dim str As String

    str = Trim$(SSOleDB_company.Columns(0).Text)
    If Trim$(SSOleDB_company.Columns(0).Text) = "ALL" Then
        SSOleDB_ware = ""
        SSOleDB_ware.RemoveAll
        Call GetalllocationName
    Else
        SSOleDB_ware = ""
        SSOleDB_ware.RemoveAll
        Call GetlocationName(str)
    End If
    
End Sub

'load company data grid

Private Sub SSOleDB_company_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_company.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_company.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDB_company.Columns(0).Width = 1500
    SSOleDB_company.Columns(1).Width = 2000
End Sub

Private Sub SSOleDB_company_GotFocus()
Call HighlightBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_company.DroppedDown Then SSOleDB_company.DroppedDown = True
End Sub

Private Sub SSOleDB_company_LostFocus()
Call NormalBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_company)) > 0 Then
    If SSOleDB_company.Rows > 0 Then
        If Not SSOleDB_company.IsItemInList Then
            SSOleDB_company.Text = ""
        End If
    End If
    If Len(Trim$(SSOleDB_company)) = 0 Then
        SSOleDB_company.SetFocus
        Cancel = True
    End If
End If
End Sub

'load warehouse data grid

Private Sub SSOleDB_ware_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_ware.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_ware.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDB_ware.Columns(0).Width = 900
    SSOleDB_ware.Columns(1).Width = 2000
End Sub

Private Sub SSOleDB_ware_GotFocus()
Call HighlightBackground(SSOleDB_ware)
End Sub

Private Sub SSOleDB_ware_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_ware.DroppedDown Then SSOleDB_ware.DroppedDown = True
End Sub

Private Sub SSOleDB_ware_LostFocus()
Call NormalBackground(SSOleDB_ware)
End Sub

'load currency data grid
'
'Private Sub SSOleDBCurrency()
'    SSOleDB_ware.Columns(0).Caption = "Code"
'    SSOleDB_ware.Columns(0).Width = 900
'    SSOleDB_ware.Columns(1).Width = 2000
'    SSOleDB_ware.Columns(1).Caption = "Name"
'End Sub
Private Sub SSOleDB_ware_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_ware)) > 0 Then
    If SSOleDB_ware.Rows > 0 Then
        If Not SSOleDB_ware.IsItemInList Then
            SSOleDB_ware.Text = ""
        End If
    End If
    If Len(Trim$(SSOleDB_ware)) = 0 Then
        SSOleDB_ware.SetFocus
        Cancel = True
    End If
End If
End Sub
