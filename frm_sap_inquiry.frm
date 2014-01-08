VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_sap_inquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sap inquiry"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleMode       =   0  'User
   ScaleWidth      =   4642.898
   Tag             =   "02050500"
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   1278
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1278
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBcompany 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      DataFieldList   =   "Column 1"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBlocation 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   720
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBstocknumber 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
      DataFieldList   =   "column 0"
      _Version        =   196617
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
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
      Columns(0).Width=   2408
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4683
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
      Left            =   360
      TabIndex        =   9
      Top             =   1140
      Width           =   1756
   End
   Begin VB.Label lbl_location 
      Caption         =   "Location"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   780
      Width           =   1756
   End
   Begin VB.Label lbl_company 
      Caption         =   "Company"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   420
      Width           =   1756
   End
   Begin VB.Label lbl_stock_number 
      Caption         =   "Stock Number"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1500
      Width           =   1756
   End
End
Attribute VB_Name = "frm_sap_inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'unload form

Private Sub cmd_cancel_Click()
    Unload Me
End Sub

'get crystal report form parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler



With MDI_IMS.CrystalReport1
    .Reset
    .ReportFileName = FixDir(App.Path) + "CRreports\sapinquiry.rpt"
    .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    .ParameterFields(1) = "compcode;" + IIf(Trim$(UCase(SSOleDBCompany.Text)) = "ALL", "ALL", UCase(SSOleDBCompany.Text)) + ";true"
    .ParameterFields(2) = "locacode;" + IIf(Trim$(UCase(SSOleDBLocation.Text)) = "ALL", "ALL", UCase(SSOleDBLocation.Text)) + ";true"
    .ParameterFields(3) = "stcknumb;" + IIf(Trim$(UCase(SSOleDBstocknumber.Text)) = "ALL", "ALL", UCase(SSOleDBstocknumber.Text)) + ";true"
    .ParameterFields(4) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'    .ParameterFields(4) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("M00157") 'J added
    .WindowTitle = IIf(msg1 = "", "SAP Inquiry", msg1) 'J modified
    Call translator.Translate_Reports("sapinquiry.rpt") 'J added
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

Private Sub Form_Activate()
Dim str As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

Screen.MousePointer = 11
Me.Refresh
'Added by Juan (9/14/2000) for Multilingual
Call translator.Translate_Forms("frm_sap_inquiry")
'------------------------------------------

'Me.Width = 4000
'Me.Height = 2900
 str = Chr$(1)
 Set cmd = New ADODB.Command
 frm_sap_inquiry.Caption = frm_sap_inquiry.Caption + " - " + frm_sap_inquiry.Tag
    
    
    
   With cmd
      
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        .CommandText = "select com_compcode,com_name from company "
        .CommandText = .CommandText & " Where com_npecode= '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by com_compcode "
        Set rs = .Execute
        
    End With
    
   If rs Is Nothing Then Exit Sub
   If rs.State And adStateOpen = adStateClosed Then Exit Sub

    If rs.EOF And rs.BOF Then GoTo CleanUp
    If rs.RecordCount = 0 Then GoTo CleanUp
    
    rs.MoveFirst
    SSOleDBCompany.FieldSeparator = str
    Call SSOleDBCompany.AddItem("ALL" & str & "ALL")
    
    Do While (Not rs.EOF)
        Call SSOleDBCompany.AddItem((rs!com_name & str) & rs!com_compcode & "")
        rs.MoveNext
    Loop
    SSOleDBCompany.Bookmark = 0
    SSOleDBCompany.Text = "ALL"

    Call GetCurrencylist
    Call GetalllocationName
    
    
    
    
    
    
'    With cmd
'        .ActiveConnection = deIms.cnIms
'        .CommandType = adCmdText
'        .CommandText = "SELECT stk_stcknumb,stk_desc FROM stockmaster"
'        .CommandText = .CommandText & " WHERE stk_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " UNION SELECT qs1_stcknumb, qs1_desc"
'        .CommandText = .CommandText & " FROM QTYST1 where qs1_stcknumb not in "
'        .CommandText = .CommandText & " (SELECT stk_stcknumb FROM stockmaster) "
'        .CommandText = .CommandText & " AND qs1_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " ORDER BY 1"
'        Set rs = .Execute
'    End With
'
'   If rs Is Nothing Then Exit Sub
'   If rs.State And adStateOpen = adStateClosed Then Exit Sub
'
'   If rs.EOF And rs.BOF Then GoTo CleanUp
'   If rs.RecordCount = 0 Then GoTo CleanUp
'
'   rs.MoveFirst

        With SSOleDBstocknumber
            Set .DataSourceList = deIms.Commands("getStockOnHandQTYST1").Execute(100, Array(0, deIms.NameSpace))
            .DataFieldToDisplay = "qs1_stcknumb"
            .DataFieldList = "qs1_stcknumb"
            .Refresh
        End With



'   SSOleDBstocknumber.FieldSeparator = str
'   SSOleDBstocknumber.AddItem (("ALL" & str) & "ALL" & str)
'   DoEvents
'   Do While (Not rs.EOF)
'        Call SSOleDBstocknumber.AddItem((Trim$(rs!stk_stcknumb) & str) & rs!stk_desc & "")
'        rs.MoveNext
'    Loop
    SSOleDBstocknumber.Bookmark = 0
    SSOleDBstocknumber.Text = "ALL"
    SSOleDBCompany.Text = "ALL"
    SSOleDBLocation.Text = "ALL"
    
CleanUp:
       rs.Close
       Set cmd = Nothing
       Set rs = Nothing
       
Screen.MousePointer = 0

End Sub

'SQL statement get stock master numbers, company name

Private Sub Form_Load()
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
    SSOleDBLocation.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    
    SSOleDBLocation.RemoveAll
    
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSOleDBLocation.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        rst.MoveNext
    Loop
    SSOleDBLocation.Bookmark = 0
    SSOleDBLocation.Text = "ALL"
      
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
         .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBLocation.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDBLocation.RemoveAll
    
    rst.MoveFirst
      
    SSOleDBLocation.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBLocation.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
    SSOleDBLocation.Bookmark = 0
    SSOleDBLocation.Text = "ALL"
      
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
        .CommandText = .CommandText & " and curr_code in "
        .CommandText = .CommandText & " (Select curd_code from currencydetl "
        .CommandText = .CommandText & " WHERE curr_npecode = '" & deIms.NameSpace & "' and "
        .CommandText = .CommandText & "curd_from  !> '" & Now & "' and curd_to !< '" & Now & "') "
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by curr_code"
         Set rst = .Execute
    End With


    str = Chr$(1)
    SSOleDBCurrency.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDBCurrency.RemoveAll
    
    rst.MoveFirst

    SSOleDBCurrency.AddItem (("USD" & str) & "US DOLLAR" & "")
    Do While ((Not rst.EOF))
         If Not UCase(rst!curr_code) = "USD" Then
        SSOleDBCurrency.AddItem rst!curr_code & str & (rst!curr_desc & "")
        End If
        rst.MoveNext
    Loop
      
      SSOleDBCurrency = "US DOLLAR"
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::Getcurrencylist", Err.Description, Err.number, True)
End Sub

'resize form

Private Sub Form_Resize()
'If Not Me.WindowState = vbMinimized Then
'Me.Width = 4200
'Me.Height = 2900
'End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'call function to get company information

Private Sub SSOleDBCompany_Click()
Dim str As String
 str = SSOleDBCompany.Columns(1).Text
 
 
        SSOleDBLocation = ""
        SSOleDBLocation.RemoveAll
        
    If Trim$(str) = "ALL" Then
'        SSOleDBlocation = ""
'        SSOleDBlocation.RemoveAll
        Call GetalllocationName
    Else
        Call GetlocationName(str)
    End If
    
    
'   If Len(str) Then
'       Call AddLocation(GetLocation(Trim$(str)))
'   End If
End Sub

Public Function GetLocation(CompanyCode As String) As ADODB.Recordset
'Dim cmd As ADODB.Command
'
'    Set cmd = New ADODB.Command
'
'    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
'
'    With cmd
'        .CommandText = "select loc_name, loc_locacode from location "
'        .CommandText = .CommandText & " where loc_npecode ='" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & IIf(IsStringEqual(CompanyCode, "ALL"), "", " and loc_compcode='" & CompanyCode & "'")
'
'        Set GetLocation = .Execute
'    End With
'
'    Set cmd = Nothing
End Function

Public Sub AddLocation(rs As ADODB.Recordset)
'Dim str As String
'
'   str = Chr$(1)
'   If rs Is Nothing Then Exit Sub
'   If rs.State And adStateOpen = adStateClosed Then Exit Sub
'
'    If rs.EOF And rs.BOF Then GoTo CleanUp
'    If rs.RecordCount = 0 Then GoTo CleanUp
'
'    rs.MoveFirst
'    SSOleDBlocation.FieldSeparator = str
'
'     SSOleDBlocation.AddItem ("ALL" & str & "ALL")
'
'    Do While (Not rs.EOF)
'
'        SSOleDBlocation.AddItem ((rs!loc_name & str) & rs!loc_locacode & "")
'        rs.MoveNext
'    Loop
'
'
'CleanUp:
'       rs.Close
End Sub

'set company combo

Private Sub SSOleDBCompany_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    msg2 = translator.Trans("L00028") 'J added
    SSOleDBCompany.Columns(0).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    SSOleDBCompany.Columns(1).Caption = IIf(msg2 = "", "Code", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBCompany.Columns(0).Width = 2000
    SSOleDBCompany.Columns(1).Width = 1500
End Sub

Private Sub SSOleDBCompany_GotFocus()
Call HighlightBackground(SSOleDBCompany)
End Sub

Private Sub SSOleDBCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCompany.DroppedDown Then SSOleDBCompany.DroppedDown = True
End Sub

Private Sub SSOleDBCompany_KeyPress(KeyAscii As Integer)
'SSOleDBcompany.MoveNext
End Sub

Private Sub SSOleDBCompany_LostFocus()
Call NormalBackground(SSOleDBCompany)
End Sub

Private Sub SSOleDBCompany_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCompany)) > 0 Then
         If Not SSOleDBCompany.IsItemInList Then
                SSOleDBCompany.Text = ""
            End If
            If Len(Trim$(SSOleDBCompany)) = 0 Then
            SSOleDBCompany.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBCurrency_GotFocus()
Call HighlightBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCurrency.DroppedDown Then SSOleDBCurrency.DroppedDown = True
End Sub

Private Sub SSOleDBCurrency_KeyPress(KeyAscii As Integer)
'SSOleDBCurrency.MoveNext
End Sub

Private Sub SSOleDBCurrency_LostFocus()
Call NormalBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_Validate(Cancel As Boolean)

If Len(Trim$(SSOleDBCurrency)) > 0 Then
         If Not SSOleDBCurrency.IsItemInList Then
                SSOleDBCurrency.Text = ""
            End If
            If Len(Trim$(SSOleDBCurrency)) = 0 Then
            SSOleDBCurrency.SetFocus
            Cancel = True
            End If
            End If
End Sub

'set location combo

Private Sub SSOleDBlocation_DropDown()
    
    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    msg2 = translator.Trans("L00028") 'J added
    SSOleDBLocation.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDBLocation.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBLocation.Columns(0).Width = 1000
    SSOleDBLocation.Columns(1).Width = 2000
End Sub

Private Sub SSOleDBlocation_GotFocus()
Call HighlightBackground(SSOleDBLocation)
End Sub

Private Sub SSOleDBlocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBLocation.DroppedDown Then SSOleDBLocation.DroppedDown = True
End Sub

Private Sub SSOleDBlocation_KeyPress(KeyAscii As Integer)
'SSOleDBlocation.MoveNext
End Sub

Private Sub SSOleDBlocation_LostFocus()
Call NormalBackground(SSOleDBLocation)
End Sub

Private Sub SSOleDBlocation_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBLocation)) > 0 Then
         If Not SSOleDBLocation.IsItemInList Then
                SSOleDBLocation.Text = ""
            End If
            If Len(Trim$(SSOleDBLocation)) = 0 Then
            SSOleDBLocation.SetFocus
            Cancel = True
            End If
            End If
End Sub

'set stock master ccombo

Private Sub SSOleDBstocknumber_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00538") 'J added
    msg2 = translator.Trans("L00029") 'J added
    SSOleDBstocknumber.Columns(0).Caption = IIf(msg1 = "", "Number", msg1) 'J modified
    SSOleDBstocknumber.Columns(1).Caption = IIf(msg2 = "", "Description", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBstocknumber.Columns(0).Width = 1600
    SSOleDBstocknumber.Columns(1).Width = 4000
    
    With SSOleDBstocknumber
        .MoveNext
    End With
End Sub

Private Sub SSOleDBstocknumber_GotFocus()
Call HighlightBackground(SSOleDBstocknumber)
End Sub

Private Sub SSOleDBstocknumber_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBstocknumber.DroppedDown Then SSOleDBstocknumber.DroppedDown = True
End Sub

Private Sub SSOleDBstocknumber_KeyPress(KeyAscii As Integer)
'SSOleDBstocknumber.MoveNext
End Sub

Private Sub SSOleDBstocknumber_LostFocus()
Call NormalBackground(SSOleDBstocknumber)
End Sub

Private Sub SSOleDBstocknumber_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBstocknumber)) > 0 Then
         If Not SSOleDBstocknumber.IsItemInList Then
                SSOleDBstocknumber.Text = ""
            End If
            If Len(Trim$(SSOleDBstocknumber)) = 0 Then
            SSOleDBstocknumber.SetFocus
            Cancel = True
            End If
            End If

End Sub
