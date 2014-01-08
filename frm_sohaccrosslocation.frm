VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_sohaccrosslocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOH Across Location"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   4710
   Tag             =   "03020200"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_tostock 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_fromstock 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CheckBox Summary 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2000
   End
   Begin VB.Label lbl_tostock 
      Caption         =   "To Stock"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   540
      Width           =   2000
   End
   Begin VB.Label lbl_fromstock 
      Caption         =   "From Stock"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   2000
   End
End
Attribute VB_Name = "frm_sohaccrosslocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim str As String
Dim x As Integer

'close recordset free memory

Private Sub cmd_cancel_Click()
x = 0
rs.Close
Set rs = Nothing
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
'On Error GoTo Handled

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\SOHaccrosscomploc.rpt"
        .ParameterFields(0) = "fromstock;" + IIf(UCase(Trim$(Combo_fromstock.Text)) = "ALL", "ALL", Trim$(Combo_fromstock.Text)) + ";true"
        .ParameterFields(1) = "tostock;" + IIf(UCase(Trim$(Combo_fromstock.Text)) = "ALL", "", Trim$(Combo_tostock.Text)) + ";true"
        .ParameterFields(2) = "detail;" + IIf(Summary.value = vbChecked, "Y", "N") + ";true"
        '.ParameterFields(3) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(3) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00540") 'J added
        .WindowTitle = IIf(msg1 = "", "SOH Accross Location-Company", msg1) 'J modified
        Call translator.Translate_Reports("SOHaccrosscomploc.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .Action = 1
        .Reset
    End With
    
    Exit Sub
Handled:
    Call LogErr(Name & "::cmd_ok_Click", Err.Description, Err.number, True)
    If Err Then MsgBox Err.Description: Err.Clear
       
       
End Sub

'load data to data combo

Private Sub Combo_fromstock_click()
If Combo_fromstock.Text = "ALL" Then
x = 0
Combo_tostock.Text = ""
Combo_tostock.Enabled = False
Else
 Combo_tostock.Enabled = True
If x = 1 Then Exit Sub
 x = x + 1
 

  '  Call PopuLateFromRecordSet(Combo_tostock, rs, "stk_stcknumb", False)
 rs.MoveFirst
Do While (Not rs.EOF)
  Combo_tostock.AddItem (rs!stk_stcknumb)
  rs.MoveNext
  
  DoEvents
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
                Combo_fromstock.Text = ""
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
'Combo_tostock.MoveNext
End Sub

Private Sub combo_tostock_LostFocus()
Call NormalBackground(Combo_tostock)
End Sub

Private Sub Combo_tostock_Validate(Cancel As Boolean)
If Len(Trim$(Combo_tostock)) > 0 Then
         If Not Combo_tostock.IsItemInList Then
                Combo_tostock.Text = ""
            End If
            If Len(Trim$(Combo_tostock)) = 0 Then
            Combo_tostock.SetFocus
            Cancel = True
            End If
            End If
End Sub

'SQL statement get stock number and populate combo

Private Sub Form_Load()

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_sohaccrosslocation")
    '------------------------------------------

 'Me.Height = 2800
 'Me.Width = 4300
 Set rs = New ADODB.Recordset
 
 With rs
  .ActiveConnection = deIms.cnIms
  .Source = "SELECT stk_stcknumb FROM stockmaster"
  .Source = .Source & " WHERE stk_npecode = '" & deIms.NameSpace & "'"
  .Source = .Source & " UNION SELECT qs1_stcknumb"
  .Source = .Source & " FROM QTYST1 WHERE qs1_stcknumb NOT IN "
  .Source = .Source & " (SELECT stk_stcknumb FROM stockmaster) "
  .Source = .Source & " AND qs1_npecode = '" & deIms.NameSpace & "'"
  .Source = .Source & " ORDER BY 1"
  .Open
End With
Combo_fromstock.Text = "ALL"
Combo_tostock.Text = "ALL"
SSOleDBCurrency.Text = "ALL"
  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
  Call Combo_fromstock.AddItem("ALL", 0)
  Do While (Not rs.EOF)
  Combo_fromstock.AddItem (rs!stk_stcknumb)
  rs.MoveNext
  Loop
  Else
    Exit Sub
  End If
  
  Call GetCurrencylist
   SSOleDBCurrency = "USD"
  
  frm_sohaccrosslocation.Caption = frm_sohaccrosslocation.Caption + " - " + frm_sohaccrosslocation.Tag
  
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


'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub



Private Sub Summary_GotFocus()
Call HighlightBackground(Summary)
End Sub

Private Sub Summary_LostFocus()
Call NormalBackground(Summary)
End Sub
