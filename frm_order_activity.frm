VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_order_activity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Activity"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   4005
   Tag             =   "03020200"
   Begin MSComCtl2.DTPicker DTto 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      DateIsNull      =   -1  'True
      Format          =   60227585
      CurrentDate     =   36515
   End
   Begin MSComCtl2.DTPicker DTfrom 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      DateIsNull      =   -1  'True
      Format          =   60227585
      CurrentDate     =   36515
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3960
      Width           =   1092
   End
   Begin VB.CheckBox chk_list 
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBpostatus 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBsupplier 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBinventory 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBbuyer 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCompany 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
      DataFieldList   =   "Column 0"
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
      Columns(0).Width=   3016
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 1"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3916
      Columns(1).Caption=   "Company Name"
      Columns(1).Name =   "Company Name"
      Columns(1).DataField=   "Column 0"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
      DataFieldList   =   "Column 0"
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
      Columns(0).Width=   2990
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4630
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBDocType 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      FieldSeparator  =   ";"
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Lbl_DocType 
      Caption         =   "Document Type"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory Company"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1995
   End
   Begin VB.Label lbl_to 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label lbl_from 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   2000
   End
   Begin VB.Label lbl_listline 
      Caption         =   "List Line Item"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label lbl_Inventorylocation 
      Caption         =   "Inventory Location"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Label lbl_supplier 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label lbl_buyer 
      Caption         =   "Buyer"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lbl_POstatus 
      Caption         =   "PO status"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1995
   End
End
Attribute VB_Name = "frm_order_activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

Private Sub chk_list_GotFocus()
Call HighlightBackground(chk_list)
End Sub

Private Sub chk_list_LostFocus()
Call NormalBackground(chk_list)
End Sub

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

If Not (DTto.value < DTfrom.value) Then
With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\orderactivity.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "postatus;" + IIf(Trim$(SSOleDBpostatus.text) = "ALL", "ALL", SSOleDBpostatus.Columns(0).text) + ";true"
        .ParameterFields(2) = "fromdate;date(" & Year(DTfrom.value) & "," & Month(DTfrom.value) & "," & Day(DTfrom.value) & ");true"
        .ParameterFields(3) = "todate;date(" & Year(DTto.value) & "," & Month(DTto.value) & "," & Day(DTto.value) & ");true"
        
        .ParameterFields(4) = "buyer;" + IIf(Trim$(SSOleDBbuyer.text) = "ALL", "ALL", SSOleDBbuyer.text) + ";true"
        .ParameterFields(5) = "supplier;" + IIf(Trim$(SSOleDBsupplier.text) = "ALL", "ALL", SSOleDBsupplier.Columns(0).text) + ";true"
        .ParameterFields(6) = "invtloca;" + IIf(Trim$(SSOleDBinventory.text) = "ALL", "ALL", SSOleDBinventory.Columns(0).text) + ";true"
        .ParameterFields(7) = "listli;" + IIf(chk_list.value = 0, "N", "Y") + ";true"
         
        .ParameterFields(8) = "compcode;" + IIf(Trim$(SSOleDBCompany.text) = "ALL", "ALL", SSOleDBCompany.Columns("code").text) + ";TRUE"
        .ParameterFields(9) = "Doccode;" + Trim$(SSOleDBDocType.Columns(0).text) + ";TRUE"
'        .ParameterFields(9) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00205") 'J added
        .WindowTitle = IIf(msg1 = "", "Order Activity", msg1) 'J modified
        Call translator.Translate_Reports("orderactivity.rpt")  'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
End With
Else

'Modified by Juan (9/13/2000) for Multilingual
msg1 = translator.Trans("M00003") 'J added
msg2 = translator.Trans("L00318") 'J added
MsgBox IIf(msg1 = "", "Make Sure The 'To Date' is greater or equal to the 'From Date' ", msg1), , IIf(msg2 = "", "Date", msg2)
'---------------------------------------------

DTto_Validate ("true")
End If
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
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



'SQL statement get buyer,supplier, location information
'populate combo

'Private Sub combo_namespace_click()
'
'
'Dim str As String
'Dim seperator As String
'seperator = Chr$(1)
'
'
'
'
'    Call get_seperator(SSOleDBpostatus)
'    SSOleDBCompany.Enabled = True
'   'If Trim$(combo_namespace.Text) = "ALL" Then
'   ' SSOleDBpostatus.Text = ""
'   ' str = "select distinct sts_code,sts_name from status where sts_npecode in (select npce_code from namespace )"
'   'Else
'   ' str = "select sts_code,sts_name from status where sts_npecode ='" & combo_namespace.Text & "'"
'   'End If
'   'SSOleDBpostatus.RemoveAll
'   ' Set rs = get_recordset(str)
'   '     If get_status(SSOleDBpostatus, rs) Then
'   '     rs.MoveFirst
'   '      SSOleDBpostatus.AddItem (("ALL" & seperator) & "ALL" & "")
'   '     Do While (Not rs.EOF)
'   '     SSOleDBpostatus.AddItem ((rs!sts_code & seperator) & rs!sts_name & "")
'   '     rs.MoveNext
'   '     Loop
'   '    Call CleanUp
'   '   End If
'
'
'      Call get_seperator(SSOleDBbuyer)
''     If Trim$(combo_namespace.Text) = "ALL" Then
''     SSOleDBbuyer.Text = ""
''      str = "select distinct buy_username from buyer where buy_npecode in (select npce_code from namespace )"
''     Else
''      str = "select buy_username from buyer where buy_npecode ='" & combo_namespace.Text & "'"
''     End If
'
'      str = "select buy_username from buyer where buy_npecode ='" & deIms.NameSpace & "'"
'      SSOleDBbuyer.RemoveAll
'      Set rs = get_recordset(str)
'
'      If get_status(SSOleDBbuyer, rs) Then
'             rs.MoveFirst
'              SSOleDBbuyer.AddItem (("ALL" & seperator) & "ALL" & "")
'             Do While (Not rs.EOF)
'             SSOleDBbuyer.AddItem (rs!buy_username & "")
'             rs.MoveNext
'             Loop
'            Call CleanUp
'      End If
'
'
'   If SSOleDBbuyer.Enabled = False Then
'            SSOleDBpostatus.Enabled = False
'            SSOleDBpostatus.Text = ""
'    Else
'         SSOleDBpostatus.RemoveAll
'         SSOleDBpostatus.Enabled = True
'        SSOleDBpostatus.AddItem (("ALL" & seperator) & "ALL" & "")
'        SSOleDBpostatus.AddItem (("CA" & seperator) & "CANCELLED" & "")
'        SSOleDBpostatus.AddItem (("CL" & seperator) & "CLOSED" & "")
'        SSOleDBpostatus.AddItem (("OP" & seperator) & "OPEN" & "")
'        SSOleDBpostatus.AddItem (("OH" & seperator) & "ON HAND" & "")
'
'    End If
'
'      Call get_seperator(SSOleDBsupplier)
'
''    If Trim$(combo_namespace.Text) = "ALL" Then
''        str = "select distinct sup_code,sup_name from supplier where sup_npecode in (select npce_code from namespace )"
''         SSOleDBsupplier.Text = ""
''   Else
''        str = "select sup_code, sup_name from supplier where sup_npecode ='" & combo_namespace.Text & "'"
''    End If
'
'
'    str = "select sup_code, sup_name from supplier where sup_npecode ='" & deIms.NameSpace & "'"
'
'    SSOleDBsupplier.RemoveAll
'    Set rs = get_recordset(str)
'        If get_status(SSOleDBsupplier, rs) Then
'        rs.MoveFirst
'         SSOleDBsupplier.AddItem (("ALL" & seperator) & "ALL" & "")
'        Do While (Not rs.EOF)
'        SSOleDBsupplier.AddItem ((rs!sup_code & seperator) & rs!sup_name & "")
'        rs.MoveNext
'        Loop
'
'    Call CleanUp
'      End If
'
'      SSOleDBCompany = ""
'      SSOleDBCompany.RemoveAll
'
'      If Trim$(combo_namespace.Text) = "ALL" Then
'            Call GetALLCampanyName
'      Else
'            Call GetCampanyName
'      End If
'
''    If Len(Trim(combo_namespace)) Then
''        SSOleDBinventory.Enabled = True
''    End If
'
''    Call get_seperator(SSOleDBinventory)
'
''    If Trim$(combo_namespace.Text) = "ALL" Then
'''    str = "select distinct loc_locacode,loc_name from location where loc_npecode in (select npce_code from namespace ) and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
'''    SSOleDBinventory.Text = ""
''        Call GetalllocationName
''    Else
'''    str = "select loc_locacode,loc_name from location where loc_npecode ='" & combo_namespace.Text & "' and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
''        Call GetlocationName
''    End If
'
'
''    SSOleDBinventory.RemoveAll
''    Set rs = get_recordset(str)
''        If get_status(SSOleDBinventory, rs) Then
''        rs.MoveFirst
''        SSOleDBinventory.AddItem (("ALL" & seperator) & "ALL" & "")
''        Do While (Not rs.EOF)
''        SSOleDBinventory.AddItem ((rs!loc_locacode & seperator) & rs!loc_name & "")
''        rs.MoveNext
''        Loop
'
'    Call CleanUp
''      End If
'End Sub



'SQL statement get PO Status list for company combo

Private Sub GetPOstatuName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT sts_code,sts_name  "
        .CommandText = .CommandText & " from status "
        .CommandText = .CommandText & " WHERE sts_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by sts_code"
         Set rst = .Execute
    End With

   
    str = Chr$(1)
    SSOleDBpostatus.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
    SSOleDBpostatus.AddItem "ALL" & str & "ALL"
    Do While ((Not rst.EOF))
       SSOleDBpostatus.AddItem ((rst!sts_code & str) & rst!sts_name & "")
        rst.MoveNext
    Loop
      
         SSOleDBpostatus.AddItem (("CA" & str) & "CANCELLED" & "")
         SSOleDBpostatus.AddItem (("CL" & str) & "CLOSED" & "")
         SSOleDBpostatus.AddItem (("OP" & str) & "OPEN" & "")
         SSOleDBpostatus.AddItem (("OH" & str) & "ON HAND" & "")
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetPOstatuName", Err.Description, Err.number, True)
End Sub

'SQL statement get Buyer list for company combo

Private Sub GetBuyerName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT buy_username "
        .CommandText = .CommandText & " from buyer "
        .CommandText = .CommandText & " WHERE buy_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by buy_username"
         Set rst = .Execute
    End With


    str = Chr$(1)
    SSOleDBbuyer.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
    SSOleDBbuyer.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBbuyer.AddItem rst!buy_username
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetBuyerName", Err.Description, Err.number, True)
End Sub

'SQL statement get Supplier list for company combo

Private Sub GetSupplierName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT sup_code, sup_name "
        .CommandText = .CommandText & " from supplier "
        .CommandText = .CommandText & " WHERE sup_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by sup_code"
         Set rst = .Execute
    End With


    str = Chr$(1)
    SSOleDBsupplier.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
    SSOleDBsupplier.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBsupplier.AddItem ((rst!sup_code & str) & rst!sup_name & "")
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetSupplierName", Err.Description, Err.number, True)
End Sub





'SQL statement get company list for company combo

Private Sub GetCampanyName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT com_compcode, com_name "
        .CommandText = .CommandText & " From Company "
        .CommandText = .CommandText & " WHERE com_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by com_compcode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBCompany.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
     SSOleDBCompany.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCompany.AddItem rst!com_compcode & str & (rst!com_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetCampanyName", Err.Description, Err.number, True)
End Sub

'SQL statement get all company list for company combo

Private Sub GetALLCampanyName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT com_compcode, com_name "
        .CommandText = .CommandText & " From Company "
        .CommandText = .CommandText & " WHERE com_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by com_compcode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBCompany.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
    
    SSOleDBCompany.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCompany.AddItem rst!com_compcode & str & (rst!com_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetALLCampanyName", Err.Description, Err.number, True)
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
        .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
        .CommandText = .CommandText & " and loc_compcode = '" & Company & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBinventory.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
     SSOleDBinventory.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBinventory.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
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
        .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBinventory.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
    SSOleDBinventory.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBinventory.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub





Private Sub DTfrom_GotFocus()
Call HighlightBackground(DTfrom)
End Sub

Private Sub DTfrom_KeyDown(KeyCode As Integer, Shift As Integer)
'If Not DTfrom.DroppedDown Then DTfrom.DroppedDown = True
End Sub

Private Sub DTfrom_KeyPress(KeyAscii As Integer)
'DTfrom.MoveNext
End Sub

Private Sub DTfrom_LostFocus()
Call NormalBackground(DTfrom)
End Sub

Private Sub DTto_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub



'load form populate combo box

Private Sub Form_Load()
Dim str As String
Dim seperator As String
Dim i As Integer

    'Added by Juan (9/13/2000)  for Multilingual
    Call translator.Translate_Forms("frm_order_activity")
    '-------------------------------------------

'Me.Width = 4000 'J hidden
'Me.Height = 4605 'J hidden
seperator = Chr$(1)
    
'    str = "select npce_code from namespace"
'    Set rs = get_recordset(str)
'
'    If get_status(combo_namespace, rs) Then
'        Call PopuLateFromRecordSet(combo_namespace, rs, "npce_code", True)
'    Call combo_namespace.AddItem("ALL", 0)
'    End If

    
      SSOleDBbuyer.Enabled = False
      SSOleDBpostatus.Enabled = True
      SSOleDBinventory.Enabled = False
      SSOleDBsupplier.Enabled = False
      SSOleDBCompany.Enabled = False
      Call CleanUp
    
    Call GetDocType
    Call GetPOstatuName
    Call GetBuyerName
    Call GetSupplierName
    Call GetCampanyName
    
    Caption = Caption + " - " + Tag
    
      SSOleDBbuyer.text = "ALL"
      For i = 0 To SSOleDBbuyer.Rows - 1
        If SSOleDBbuyer.Columns(0).text = SSOleDBbuyer.text Then Exit For
        SSOleDBbuyer.MoveNext
      Next
                  
      SSOleDBinventory.text = "ALL"
      For i = 0 To SSOleDBinventory.Rows - 1
        If SSOleDBinventory.Columns(0).text = SSOleDBinventory.text Then Exit For
        SSOleDBinventory.MoveNext
      Next
      
      SSOleDBsupplier.text = "ALL"
      For i = 0 To SSOleDBsupplier.Rows - 1
        If SSOleDBsupplier.Columns(0).text = SSOleDBsupplier.text Then Exit For
        SSOleDBsupplier.MoveNext
      Next

      SSOleDBCompany.text = "ALL"
      For i = 0 To SSOleDBCompany.Rows - 1
        If SSOleDBCompany.Columns(0).text = SSOleDBCompany.text Then Exit For
        SSOleDBCompany.MoveNext
      Next
    
      If DTfrom.Enabled = False Then
            DTfrom.Enabled = False
            
      Else
        SSOleDBpostatus.RemoveAll
'         SSOleDBpostatus.Enabled = True
         SSOleDBpostatus.AddItem (("ALL" & seperator) & "ALL" & "")
         SSOleDBpostatus.AddItem (("CA" & seperator) & "CANCELLED" & "")
         SSOleDBpostatus.AddItem (("CL" & seperator) & "CLOSED" & "")
         SSOleDBpostatus.AddItem (("OP" & seperator) & "OPEN" & "")
         SSOleDBpostatus.AddItem (("OH" & seperator) & "ON HAND" & "")
      
    End If
    
      SSOleDBpostatus.text = "ALL"
      For i = 0 To SSOleDBpostatus.Rows - 1
        If SSOleDBpostatus.Columns(0).text = SSOleDBpostatus.text Then Exit For
        SSOleDBpostatus.MoveNext
      Next
    
    
    Call GetCurrencylist
    SSOleDBCurrency = "USD"
    DTfrom = FirstOfMonth
    DTto = Now
    
    With frm_order_activity
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'set record set connection

Public Function get_recordset(str As String) As ADODB.Recordset
Set cmd = New ADODB.Command
With cmd
  .ActiveConnection = deIms.cnIms
  .CommandType = adCmdText
  .CommandText = str
  Set get_recordset = .Execute
  End With
   End Function

'check recordset status

Public Function get_status(ctl As Control, rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
      If Not get_status Then
      ctl.Enabled = False
      ctl.text = ""
      Else: ctl.Enabled = True
      End If
End Function

'free memory

Public Sub CleanUp()
Set cmd = Nothing
Set rs = Nothing
End Sub

'resize form

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then
'Me.Width = 4000 'J hidden
'Me.Height = 4605 'J hidden
End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub SSOleDBbuyer_Click()
    If Len(SSOleDBbuyer) <> 0 Then
        SSOleDBsupplier.Enabled = True
'        Call SSOleDBCompany_Click
    End If

End Sub




Private Sub SSOleDBbuyer_GotFocus()
Call HighlightBackground(SSOleDBbuyer)
End Sub

Private Sub SSOleDBbuyer_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBbuyer.DroppedDown Then SSOleDBbuyer.DroppedDown = True
End Sub

Private Sub SSOleDBbuyer_KeyPress(KeyAscii As Integer)
'SSOleDBbuyer.MoveNext
End Sub

Private Sub SSOleDBbuyer_LostFocus()
Call NormalBackground(SSOleDBbuyer)
End Sub

Private Sub SSOleDBbuyer_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBbuyer)) > 0 Then
         If Not SSOleDBbuyer.IsItemInList Then
                SSOleDBbuyer.text = ""
            End If
            If Len(Trim$(SSOleDBbuyer)) = 0 Then
           SSOleDBbuyer.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBCompany_Click()
Dim com As String

    If Len(Trim$(SSOleDBCompany)) <> 0 Then
    
        If Trim$(SSOleDBCompany) = "ALL" Then
            
            SSOleDBinventory = ""
            SSOleDBinventory.RemoveAll
            Call GetalllocationName
        Else
            SSOleDBinventory = ""
            SSOleDBinventory.RemoveAll
            com = Trim$(SSOleDBCompany.Columns(0).text)
            Call GetlocationName(com)
        End If
    End If
        SSOleDBinventory.Enabled = True
End Sub

Private Sub SSOleDBCompany_GotFocus()
Call HighlightBackground(SSOleDBCompany)
End Sub

Private Sub SSOleDBCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCompany.DroppedDown Then SSOleDBCompany.DroppedDown = True
End Sub

Private Sub SSOleDBCompany_KeyPress(KeyAscii As Integer)
'SSOleDBCompany.MoveNext
End Sub

Private Sub SSOleDBCompany_LostFocus()
Call NormalBackground(SSOleDBCompany)
End Sub

Private Sub SSOleDBCompany_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCompany)) > 0 Then
         If Not SSOleDBCompany.IsItemInList Then
                SSOleDBCompany.text = ""
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
                SSOleDBCurrency.text = ""
            End If
            If Len(Trim$(SSOleDBCurrency)) = 0 Then
           SSOleDBCurrency.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBDocType_GotFocus()
Call HighlightBackground(SSOleDBDocType)
End Sub

Private Sub SSOleDBDocType_KeyPress(KeyAscii As Integer)

If SSOleDBDocType.DroppedDown = False Then SSOleDBDocType.DroppedDown = True

End Sub

Private Sub SSOleDBDocType_LostFocus()
Call NormalBackground(SSOleDBDocType)
End Sub

Private Sub SSOleDBDocType_Validate(Cancel As Boolean)

If Len(Trim$(SSOleDBDocType)) > 0 Then
         
         If Not SSOleDBDocType.IsItemInList Then
         
                SSOleDBDocType.text = ""
                
         End If
            
         If Len(Trim$(SSOleDBDocType)) = 0 Then
         
            SSOleDBDocType.SetFocus
            Cancel = True
            
         End If
            
End If

End Sub

'Private Sub SSOleDBCurrency_DropDown()
'    SSOleDBinventory.Columns(0).Caption = "Code"
'    SSOleDBinventory.Columns(0).Width = 800
'    SSOleDBinventory.Columns(1).Width = 3000
'    SSOleDBinventory.Columns(1).Caption = "Description"
'End Sub


'load inventory combo

Private Sub SSOleDBinventory_DropDown()

    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDBinventory.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDBinventory.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBinventory.Columns(0).Width = 800
    SSOleDBinventory.Columns(1).Width = 3000
End Sub

Private Sub SSOleDBinventory_GotFocus()
Call HighlightBackground(SSOleDBinventory)
End Sub

Private Sub SSOleDBinventory_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBinventory.DroppedDown Then SSOleDBinventory.DroppedDown = True
End Sub

Private Sub SSOleDBinventory_KeyPress(KeyAscii As Integer)
'SSOleDBinventory.MoveNext
End Sub

Private Sub SSOleDBinventory_LostFocus()
Call NormalBackground(SSOleDBinventory)
End Sub

Private Sub SSOleDBinventory_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBinventory)) > 0 Then
         If Not SSOleDBinventory.IsItemInList Then
                SSOleDBinventory.text = ""
            End If
            If Len(Trim$(SSOleDBinventory)) = 0 Then
           SSOleDBinventory.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBpostatus_Click()
    
    If Len(SSOleDBpostatus) <> 0 Then
        SSOleDBbuyer.Enabled = True
        Call GetSupplierName
    End If
    
End Sub

'load status combo

Private Sub SSOleDBpostatus_DropDown()

    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDBpostatus.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDBpostatus.Columns(1).Caption = IIf(msg2 = "", "Status", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBpostatus.Columns(0).Width = 600
    SSOleDBpostatus.Columns(1).Width = 3000
    
End Sub

Private Sub SSOleDBpostatus_GotFocus()
Call HighlightBackground(SSOleDBpostatus)
End Sub

Private Sub SSOleDBpostatus_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBpostatus.DroppedDown Then SSOleDBpostatus.DroppedDown = True
End Sub

Private Sub SSOleDBpostatus_KeyPress(KeyAscii As Integer)
'SSOleDBpostatus.MoveNext
End Sub

Private Sub SSOleDBpostatus_LostFocus()
Call NormalBackground(SSOleDBpostatus)
End Sub

Private Sub SSOleDBpostatus_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBpostatus)) > 0 Then
         If Not SSOleDBpostatus.IsItemInList Then
                SSOleDBpostatus.text = ""
            End If
            If Len(Trim$(SSOleDBpostatus)) = 0 Then
           SSOleDBpostatus.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBsupplier_Click()
    If Len(SSOleDBsupplier) <> 0 Then
        SSOleDBCompany.Enabled = True
    End If
End Sub

'load supplier combo

Private Sub SSOleDBsupplier_DropDown()

    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDBsupplier.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDBsupplier.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBsupplier.Columns(0).Width = 1000
    SSOleDBsupplier.Columns(1).Width = 3000
End Sub

Private Sub SSOleDBsupplier_GotFocus()
Call HighlightBackground(SSOleDBsupplier)
End Sub

Private Sub SSOleDBsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBsupplier.DroppedDown Then SSOleDBsupplier.DroppedDown = True
End Sub

Private Sub SSOleDBsupplier_KeyPress(KeyAscii As Integer)
'SSOleDBsupplier.MoveNext
End Sub

Private Sub SSOleDBsupplier_LostFocus()
Call NormalBackground(SSOleDBsupplier)
End Sub

Private Sub SSOleDBsupplier_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBsupplier)) > 0 Then
         If Not SSOleDBsupplier.IsItemInList Then
                SSOleDBsupplier.text = ""
            End If
            If Len(Trim$(SSOleDBsupplier)) = 0 Then
           SSOleDBsupplier.SetFocus
            Cancel = True
            End If
            End If
End Sub

Public Sub GetDocType()


SSOleDBDocType.Columns(0).Caption = "Code"
SSOleDBDocType.Columns(1).Caption = "Description"
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
On Error GoTo Handler
Dim str As String
rs.Source = "select doc_code,doc_desc from doctype where doc_npecode='" & deIms.NameSpace & "' order by doc_code"
rs.ActiveConnection = deIms.cnIms
rs.Open , , adOpenForwardOnly, adLockOptimistic

If rs.RecordCount = 0 Then SSOleDBDocType.Enabled = False: Exit Sub

     
     str = SSOleDBDocType.FieldSeparator
     SSOleDBDocType.text = "ALL"
     SSOleDBDocType.AddItem "ALL" & ";" & "ALL"
     
     
Do While Not rs.EOF
     
     SSOleDBDocType.AddItem rs!doc_code & ";" & rs!doc_desc
     rs.MoveNext
     
Loop

rs.Close
Set rs = Nothing

Exit Sub

Handler:

MsgBox "Errors occurred while trying to populate Document Types." & vbCrLf & "Error Description " & Err.Description, vbCritical, "Imswin"
Err.Clear
SSOleDBDocType.RemoveAll
SSOleDBDocType.Enabled = False

End Sub
