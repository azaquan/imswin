VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_stockhistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock History "
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   3615
   Tag             =   "03020200"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_stocknumb 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      Cols            =   1
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
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
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   765
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.Label lbl_stocknumber 
      Caption         =   "Stock Number"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   400
      Width           =   1700
   End
End
Attribute VB_Name = "frm_stockhistory"
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
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\postkhistory.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stocknumb;" + IIf(UCase(Trim$(Combo_stocknumb.Text)) = "ALL", "ALL", Trim$(Combo_stocknumb.Text)) + ";true"
        '.ParameterFields(2) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(2) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00290") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock History", msg1) 'J modified
        Call translator.Translate_Reports("postkhistory.rpt") 'J added
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

Private Sub Combo_stocknum_DropDown()
Dim RSstkMaster As ADODB.Recordset
 Set RSstkMaster = New ADODB.Recordset

RSstkMaster.Source = "select sap_stcknumb from sap where sap_compcode='" & SSOleDBCompany & "' and sap_loca='" & SSOleDBLocation & "'  and sap_npecode='" & deIms.NameSpace & "'"

RSstkMaster.ActiveConnection = deIms.cnIms
RSstkMaster.Open , , adOpenStatic


combo_stocknum.Text = "ALL"
combo_stocknum.RemoveAll
combo_stocknum.AddItem "All"
Do While Not RSstkMaster.EOF
   combo_stocknum.AddItem RSstkMaster!sap_stcknumb
    RSstkMaster.MoveNext
   
Loop

RSstkMaster.Close
Set RSstkMaster = Nothing




End Sub



Private Sub Combo_stocknumb_Click()
Combo_stocknumb.SelLength = 0
Combo_stocknumb.SelStart = 0
Combo_stocknumb.Tag = Combo_stocknumb.Columns(0).Text
End Sub

Private Sub Combo_stocknumb_DropDown()
    With Combo_stocknumb
        .MoveNext
    End With
End Sub


''Dim RSstkMaster As adodb.Recordset
'' Set RSstkMaster = New adodb.Recordset
''
''RSstkMaster.Source = "select sap_stcknumb from sap where  sap_npecode='" & deIms.NameSpace & "'"
''
''RSstkMaster.ActiveConnection = deIms.cnIms
''RSstkMaster.Open , , adOpenStatic
''
''
''Combo_stocknumb.text = "ALL"
''Combo_stocknumb.RemoveAll
''Combo_stocknumb.AddItem "All"
''Do While Not RSstkMaster.EOF
''   Combo_stocknumb.AddItem RSstkMaster!sap_stcknumb
''    RSstkMaster.MoveNext
''
''Loop
''
''RSstkMaster.Close
''Set RSstkMaster = Nothing






Private Sub Combo_stocknumb_GotFocus()
Call HighlightBackground(Combo_stocknumb)

End Sub

Private Sub Combo_stocknumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_stocknumb.DroppedDown Then Combo_stocknumb.DroppedDown = True
End Sub

Private Sub Combo_stocknumb_KeyPress(KeyAscii As Integer)
'Combo_stocknumb.MoveNext
End Sub

Private Sub Combo_stocknumb_LostFocus()
Call NormalBackground(Combo_stocknumb)
End Sub


Private Sub Combo_stocknumb_Validate(Cancel As Boolean)
If Len(Trim$(Combo_stocknumb)) > 0 Then
         If Not Combo_stocknumb.IsItemInList Then
                Combo_stocknumb.Text = ""
            End If
            If Len(Trim$(Combo_stocknumb)) = 0 Then
            Combo_stocknumb.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Screen.MousePointer = 11
 Set rs = New ADODB.Recordset
 
    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_stockhistory")
    '------------------------------------------
 
 'Me.Height = 2245
 'Me.Width = 3495
 
    
    frm_stockhistory.Caption = frm_stockhistory.Caption + " - " + frm_stockhistory.Tag
    Me.Refresh
    With rs
        .ActiveConnection = deIms.cnIms
'        .Source = "SELECT stk_stcknumb,stk_desc FROM stockmaster"
'        .Source = .Source & " WHERE stk_npecode = '" & deIms.NameSpace & "'"
'        .Source = .Source & " UNION SELECT qs1_stcknumb, qs1_desc"
'        .Source = .Source & " FROM QTYST1 WHERE qs1_stcknumb NOT IN "
'        .Source = .Source & " (SELECT stk_stcknumb FROM stockmaster) "
'        .Source = .Source & " AND qs1_npecode = '" & deIms.NameSpace & "'"
'        .Source = .Source & " ORDER BY 1"
         .Source = "select sap_stcknumb from sap where  sap_npecode='" & deIms.NameSpace & "' order by sap_stcknumb"
        .Open , , adOpenStatic
    End With
    
  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
'   Combo_stocknumb.AddItem "ALL"
'   Combo_stocknumb.text = "ALL"
'  Do While (Not rs.EOF)
'
'  Combo_stocknumb.AddItem (rs!sap_stcknumb)
'  rs.MoveNext
'  Loop
DoEvents
        With Combo_stocknumb
            Set .DataSourceList = rs ' deIms.Commands("getStockOnHandQTYST1").Execute(100, Array(0, deIms.NameSpace))
            .DataFieldToDisplay = "sap_stcknumb"
            .DataFieldList = "sap_stcknumb"
            .Refresh
        End With


   'Call PopuLateFromRecordSet(Combo_stocknumb, rs, "stk_stcknumb", True)
  Set rs = Nothing
  Else
    Exit Sub
  End If
  Combo_stocknumb.Text = "ALL"
    Call GetCurrencylist
    SSOleDBCurrency = "USD"
  
  frm_stockhistory.Caption = frm_stockhistory.Caption + " - " + frm_stockhistory.Tag
  Screen.MousePointer = 0
  
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

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


'resize form

Private Sub Form_Resize()
 If Not (Me.WindowState = vbMinimized) Then
 'Me.Height = 2245
 'Me.Width = 3495
    
 End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub
