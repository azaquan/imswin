VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_inventoryperstocknu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Per Stock Number"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   4320
   Tag             =   "03030200"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo combo_tostock 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_fromstock 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CheckBox chk_listall 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo_logwar 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   960
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_company 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   240
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_ware 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   600
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_subloc 
      Height          =   315
      Left            =   2400
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDb_category 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
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
   Begin VB.Label lbl_tostock 
      Caption         =   "To Stock number"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2460
      Width           =   2000
   End
   Begin VB.Label lbl_listall 
      Caption         =   "List Detail"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2820
      Width           =   2000
   End
   Begin VB.Label lbl_fromstock 
      Caption         =   "From Stock number"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2100
      Width           =   2000
   End
   Begin VB.Label lbl_catecode 
      Caption         =   "Category Code"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   1740
      Width           =   2000
   End
   Begin VB.Label lbl_subloc 
      Caption         =   "Sublocation"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1380
      Width           =   2000
   End
   Begin VB.Label lbl_logwar 
      Caption         =   "Logical Warehouse"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1020
      Width           =   2000
   End
   Begin VB.Label lbl_locacode 
      Caption         =   "Location Code"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   660
      Width           =   2000
   End
   Begin VB.Label lbl_compcode 
      Caption         =   "Company Code"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   300
      Width           =   2000
   End
End
Attribute VB_Name = "frm_inventoryperstocknu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As Integer

Private Sub chk_listall_GotFocus()
Call HighlightBackground(chk_listall)
End Sub

Private Sub chk_listall_LostFocus()
Call NormalBackground(chk_listall)
End Sub

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get parameter values for crystal report
'and application path

Private Sub cmd_ok_Click()
On Error Resume Next

    With MDI_IMS.CrystalReport1
            .Reset
            
            .ReportFileName = ReportPath + "invtperstocknumb.rpt"
            .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            .ParameterFields(1) = "compcode;" + IIf(Trim$(SSOleDB_company.text) = "ALL", "ALL", SSOleDB_company.Columns(0).text) + ";true"
            .ParameterFields(2) = "locacode;" + IIf(Trim$(SSOleDB_ware.text) = "ALL", "ALL", SSOleDB_ware.Columns(0).text) + ";true"
            .ParameterFields(3) = "logwar ;" + IIf(Trim$(SSOleDBCombo_logwar.text) = "ALL", "ALL", SSOleDBCombo_logwar.text) + ";true"
            .ParameterFields(4) = "subloc ;" + IIf(Trim$(SSOleDB_subloc.text) = "ALL", "ALL", SSOleDB_subloc.Columns(0).text) + ";true"
            .ParameterFields(5) = "catecode ;" + IIf(Trim$(SSOleDb_category.text) = "ALL", "ALL", SSOleDb_category.Columns(0).text) + ";true"
            .ParameterFields(6) = "fromstock;" + IIf(Trim$(Combo_fromstock.text) = "ALL", "ALL", Combo_fromstock.text) + ";true"
            .ParameterFields(7) = "tostock;" + IIf(Trim$(Combo_fromstock.text) = "ALL", "", Combo_tostock.text) + ";true"
            .ParameterFields(8) = "listall;" + IIf(chk_listall.Value = 0, "N", "Y") + ";true"
            
            'Modified by Juan (9/12/2000) for Multilingual
            msg1 = translator.Trans("M00182") 'J added
            .WindowTitle = IIf(msg1 = "", "Inventory per Stock Number", msg1) 'J modified
            Call translator.Translate_Reports("invtperstocknumb.rpt")  ' J added
            Call translator.Translate_SubReports 'J added
            '---------------------------------------------
            
            .Action = 1
            .Reset
    End With
    
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::cmd_ok_Click", Err.Description, Err)
    End If
End Sub

'Search stock recordset, load values, set combo to stock equal to
'combo from stock

Private Sub Combo_fromstock_click()
'Dim RSStock As New adodb.Recordset
If Combo_fromstock.text = "ALL" Then
x = 0
Combo_tostock.text = ""
Combo_tostock.Enabled = False
Else

Combo_tostock.Enabled = True
If x = 1 Then Exit Sub
 x = x + 1
 
rs.MoveFirst
Do While (Not rs.EOF)
  Combo_tostock.AddItem (Trim$(rs!qs1_stcknumb))
  rs.MoveNext
  Loop
End If
    Combo_tostock = Combo_fromstock
End Sub

Private Sub Combo_fromstock_GotFocus()
Call HighlightBackground(Combo_fromstock)
End Sub

Private Sub Combo_fromstock_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_fromstock.DroppedDown Then Combo_fromstock.DroppedDown = True
End Sub

Private Sub Combo_fromstock_KeyPress(KeyAscii As Integer)
'Combo_fromstock.MoveNext
End Sub

Private Sub Combo_fromstock_LostFocus()
Call NormalBackground(Combo_fromstock)
End Sub

Private Sub Combo_fromstock_Validate(Cancel As Boolean)
If Len(Trim$(Combo_fromstock)) > 0 Then
         If Not Combo_fromstock.IsItemInList Then
               Combo_fromstock.text = ""
            End If
            If Len(Trim$(Combo_fromstock)) = 0 Then
            Combo_fromstock.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub combo_tostock_GotFocus()
Call HighlightBackground(Combo_tostock)
End Sub

Private Sub combo_tostock_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_tostock.DroppedDown Then Combo_tostock.DroppedDown = True
End Sub

Private Sub combo_tostock_KeyPress(KeyAscii As Integer)
'combo_tostock.MoveNext
End Sub

Private Sub combo_tostock_LostFocus()
Call NormalBackground(Combo_tostock)

End Sub



Private Sub Combo_tostock_Validate(Cancel As Boolean)
If Len(Trim$(Combo_tostock)) > 0 Then
    If Combo_tostock.Rows > 0 Then
        If Not Combo_tostock.IsItemInList Then
            Combo_tostock.text = ""
        End If
    End If
    If Len(Trim$(Combo_tostock)) = 0 Then
        Combo_tostock.SetFocus
        Cancel = True
    End If
End If
End Sub

'SQL statement get recordsets,load form, populate combo data

Sub Form_Load()

'Added by Juan (9/12/2000) for Multilingual
Call translator.Translate_Forms("frm_inventoryperstocknu")
'------------------------------------------

frm_inventoryperstocknu.Caption = frm_inventoryperstocknu.Caption + " - " + frm_inventoryperstocknu.Tag

'Me.Height = 4605 'J hidden
'Me.Width = 4800 'J hidden
Dim rs1 As ADODB.Recordset
SSOleDB_company.FieldSeparator = Chr$(1)
x = 0
SSOleDB_ware.FieldSeparator = Chr$(1)
SSOleDB_subloc.FieldSeparator = Chr$(1)
SSOleDb_category.FieldSeparator = Chr$(1)
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

    With rs
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "'"
        .Source = .Source & " order by com_compcode"
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
 SSOleDB_company.text = "ALL"
 SSOleDb_category.text = "ALL"
 SSOleDB_subloc.text = "ALL"
 SSOleDB_ware.text = "ALL"
 SSOleDBCombo_logwar.text = "ALL"
 Combo_fromstock.text = "ALL"
 Combo_tostock.text = "ALL"
 
If get_status(rs) Then
SSOleDB_company.AddItem ("ALL" & Chr$(1) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_company.AddItem (rs!com_compcode & Chr$(1) & rs!com_name & "")
rs.MoveNext
Loop
Set rs = Nothing
End If




Set rs = New ADODB.Recordset
rs.Source = "select lw_code ,lw_desc from logwar where lw_npecode ='" & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
If get_status(rs) Then
SSOleDBCombo_logwar.FieldSeparator = Chr$(1)
SSOleDBCombo_logwar.AddItem (("ALL" & Chr$(1)) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDBCombo_logwar.AddItem ((rs!lw_code & Chr$(1)) & rs!lw_desc & "")
rs.MoveNext
Loop
End If
Set rs = Nothing
Set rs = New ADODB.Recordset
rs.Source = "select sb_code,sb_desc from sublocation where sb_npecode='" & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
If get_status(rs) Then
SSOleDB_subloc.AddItem ("ALL" & Chr$(1) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_subloc.AddItem (rs!sb_code & Chr$(1) & rs!sb_desc & "")
rs.MoveNext
Loop
Set rs = Nothing
End If
Set rs = New ADODB.Recordset
rs.Source = "select cate_catecode,cate_catename from category where cate_npecode='" & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
If get_status(rs) Then
SSOleDb_category.AddItem ("ALL" & Chr$(1) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDb_category.AddItem (rs!cate_catecode & Chr$(1) & rs!cate_catename & "")
rs.MoveNext
Loop
Set rs = Nothing
End If


Set rs = New ADODB.Recordset
  
      With rs
        .ActiveConnection = deIms.cnIms
        .Source = "SELECT distinct(qs1_stcknumb)"
        .Source = .Source & " FROM QTYST1 WHERE "
        .Source = .Source & " qs1_npecode = '" & deIms.NameSpace & "'"
        .Source = .Source & " ORDER BY 1"
        .Open
    End With

  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
  
  
            Call Combo_fromstock.AddItem("ALL", 0)
            Do While (Not rs.EOF)
              Combo_fromstock.AddItem (Trim$(rs!qs1_stcknumb))
              rs.MoveNext
            Loop
            
  Else
  
    Exit Sub
    
  End If
  

    
    Caption = Caption + " - " + Tag
  
  
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


'get recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
   End Function

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'load category combo

Private Sub SSOleDb_category_DropDown()

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    SSOleDb_category.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    '---------------------------------------------
    
    SSOleDb_category.Columns(0).Width = 500
    SSOleDb_category.Columns(1).Width = 2000
    
    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    SSOleDb_category.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------
    
End Sub

Private Sub SSOleDb_category_GotFocus()
Call HighlightBackground(SSOleDb_category)
End Sub

Private Sub SSOleDb_category_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDb_category.DroppedDown Then SSOleDb_category.DroppedDown = True
End Sub

Private Sub SSOleDb_category_KeyPress(KeyAscii As Integer)
'SSOleDb_category.MoveNext
End Sub

Private Sub SSOleDb_category_LostFocus()
Call NormalBackground(SSOleDb_category)

End Sub

Private Sub SSOleDb_category_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDb_category)) > 0 Then
         If Not SSOleDb_category.IsItemInList Then
               SSOleDb_category.text = ""
            End If
            If Len(Trim$(SSOleDb_category)) = 0 Then
            SSOleDb_category.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_company_Click()
Dim str As String
    str = Trim$(SSOleDB_company.Columns(0).text)
    If Trim$(SSOleDB_company.Columns(0).text) = "ALL" Then
        SSOleDB_ware = ""
        SSOleDB_ware.RemoveAll
        Call GetalllocationName
    Else
        SSOleDB_ware = ""
        SSOleDB_ware.RemoveAll
        Call GetlocationName(str)
    End If
    
    
End Sub

'load company combo

Private Sub SSOleDB_company_DropDown()

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    SSOleDB_company.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_company.Columns(0).Width = 1500
    SSOleDB_company.Columns(1).Width = 2000
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_company.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------

End Sub

Private Sub SSOleDB_company_GotFocus()
Call HighlightBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_company.DroppedDown Then SSOleDB_company.DroppedDown = True
End Sub

Private Sub SSOleDB_company_KeyPress(KeyAscii As Integer)
'SSOleDB_company.MoveNext
End Sub

Private Sub SSOleDB_company_LostFocus()
Call NormalBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_company)) > 0 Then
         If Not SSOleDB_company.IsItemInList Then
               SSOleDB_company.text = ""
            End If
            If Len(Trim$(SSOleDB_company)) = 0 Then
            SSOleDB_company.SetFocus
            Cancel = True
            End If
            End If
End Sub

'load sub location combo

Private Sub SSOleDB_subloc_DropDown()

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    SSOleDB_subloc.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_subloc.Columns(0).Width = 1000
    SSOleDB_subloc.Columns(1).Width = 3000
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_subloc.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------
    
End Sub

Private Sub SSOleDB_subloc_GotFocus()
Call HighlightBackground(SSOleDB_subloc)
End Sub

Private Sub SSOleDB_subloc_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_subloc.DroppedDown Then SSOleDB_subloc.DroppedDown = True
End Sub

Private Sub SSOleDB_subloc_KeyPress(KeyAscii As Integer)
'SSOleDB_subloc.MoveNext
End Sub

Private Sub SSOleDB_subloc_LostFocus()
Call NormalBackground(SSOleDB_subloc)

End Sub

Private Sub SSOleDB_subloc_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_subloc)) > 0 Then
         If Not SSOleDB_subloc.IsItemInList Then
               SSOleDB_subloc.text = ""
            End If
            If Len(Trim$(SSOleDB_subloc)) = 0 Then
            SSOleDB_subloc.SetFocus
            Cancel = True
            End If
            End If
End Sub

'load warehouse combo

Private Sub SSOleDB_ware_DropDown()

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    SSOleDB_ware.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_ware.Columns(0).Width = 1000
    SSOleDB_ware.Columns(1).Width = 2000
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_ware.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------
    
End Sub

Private Sub SSOleDB_ware_GotFocus()
Call HighlightBackground(SSOleDB_ware)
End Sub

Private Sub SSOleDB_ware_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_ware.DroppedDown Then SSOleDB_ware.DroppedDown = True
End Sub

Private Sub SSOleDB_ware_KeyPress(KeyAscii As Integer)
'SSOleDB_ware.MoveNext
End Sub

Private Sub SSOleDB_ware_LostFocus()
Call NormalBackground(SSOleDB_ware)

End Sub

Private Sub SSOleDB_ware_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_ware)) > 0 Then
    If SSOleDB_ware.Rows > 0 Then
        If Not SSOleDB_ware.IsItemInList Then
            SSOleDB_ware.text = ""
        End If
    End If
    If Len(Trim$(SSOleDB_ware)) = 0 Then
        SSOleDB_ware.SetFocus
        Cancel = True
    End If
End If
End Sub

'load logical warehouse combo

Private Sub SSOleDBCombo_logwar_DropDown()

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    SSOleDBCombo_logwar.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDBCombo_logwar.Columns(0).Width = 1000
    SSOleDBCombo_logwar.Columns(1).Width = 2000
    msg1 = translator.Trans("L00050") 'J added
    SSOleDBCombo_logwar.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------
    
End Sub

Private Sub SSOleDBCombo_logwar_GotFocus()
Call HighlightBackground(SSOleDBCombo_logwar)
End Sub

Private Sub SSOleDBCombo_logwar_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCombo_logwar.DroppedDown Then SSOleDBCombo_logwar.DroppedDown = True
End Sub

Private Sub SSOleDBCombo_logwar_KeyPress(KeyAscii As Integer)
'SSOleDBCombo_logwar.MoveNext
End Sub

Private Sub SSOleDBCombo_logwar_LostFocus()
Call NormalBackground(SSOleDBCombo_logwar)

End Sub

Private Sub SSOleDBCombo_logwar_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCombo_logwar)) > 0 Then
         If Not SSOleDBCombo_logwar.IsItemInList Then
               SSOleDBCombo_logwar.text = ""
            End If
            If Len(Trim$(SSOleDBCombo_logwar)) = 0 Then
            SSOleDBCombo_logwar.SetFocus
            Cancel = True
            End If
            End If
End Sub
