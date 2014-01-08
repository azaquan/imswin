VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_transact_order 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Order"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   3870
   Tag             =   "02020400"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   6
      FieldSeparator  =   ";"
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
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
      FieldSeparator  =   ";"
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2540
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3810
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   915
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction Order :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   440
      Width           =   2000
   End
End
Attribute VB_Name = "frm_transact_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\po.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Trim$(SSOleDBCombo1.Text) + ";true"
        '.ParameterFields(2) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'       .ParameterFields(2) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00060") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction Order", msg1) 'J modified
        Call translator.Translate_Reports("po.rpt") 'J added
        Call translator.Translate_SubReports 'J added
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


'SQL statement get po information and populate data grid

Private Sub Form_Load()
 Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    'Add by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_transact_order")
    '----------------------------------------
    
   'Me.Height = 2205
   'Me.Width = 4005
       frm_transact_order.Caption = frm_transact_order.Caption + " - " + frm_transact_order.Tag
    Set cmd = New ADODB.Command
    With cmd
       .ActiveConnection = deIms.cnIms
        .CommandType = adCmdText
        .CommandText = "select po_ponumb,supplier.sup_name,po_stas,po_date,po_priocode, po_revinumb from po,supplier where po.po_suppcode=supplier.sup_code"
        .CommandText = .CommandText & " and po_npecode = '" & deIms.NameSpace & "'and sup_npecode ='" & deIms.NameSpace & "' order by po_ponumb"
        Set rst = .Execute
    End With
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then GoTo CleanUp
     rst.MoveFirst
    Do While (Not rst.EOF)
SSOleDBCombo1.AddItem (rst!PO_PONUMB & ";" & rst!sup_name & ";" & rst!po_stas & ";" & rst!PO_Date & ";" & rst!po_priocode & ";" & rst!po_revinumb & ";")
        rst.MoveNext
    Loop
    
  Call GetCurrencylist
   SSOleDBCurrency = "USD"
   
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

Private Sub Form_Resize()
'Me.Height = 2205
  ' Me.Width = 4005
End Sub

'unload  form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'load combo

Private Sub SSOleDBCombo1_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00543") 'J added
    SSOleDBCombo1.Columns(0).Caption = IIf(msg1 = "", "Order#", msg1) 'J modified
    SSOleDBCombo1.Columns(0).Width = 1500
    msg1 = translator.Trans("L00128") 'J added
    SSOleDBCombo1.Columns(1).Caption = IIf(msg1 = "", "Supplier", msg1) 'J modified
    SSOleDBCombo1.Columns(2).Width = 800
    msg1 = translator.Trans("L00110") 'J added
    SSOleDBCombo1.Columns(2).Caption = IIf(msg1 = "", "Status", msg1) 'J modified
    SSOleDBCombo1.Columns(3).Width = 800
    msg1 = translator.Trans("L00318") 'J added
    SSOleDBCombo1.Columns(3).Caption = IIf(msg1 = "", "Date", msg1) 'J modified
    SSOleDBCombo1.Columns(4).Width = 600
    msg1 = translator.Trans("L00544") 'J added
    SSOleDBCombo1.Columns(4).Caption = IIf(msg1 = "", "Mode", msg1) 'J modified
    msg1 = translator.Trans("L00055") 'J added
    SSOleDBCombo1.Columns(5).Width = 800
    SSOleDBCombo1.Columns(5).Caption = IIf(msg1 = "", "Revision", msg1) 'J modified
    '---------------------------------------------
End Sub
Private Sub SSOleDBCombo1_GotFocus()
Call HighlightBackground(SSOleDBCombo1)
End Sub

Private Sub SSOleDBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCombo1.DroppedDown Then SSOleDBCombo1.DroppedDown = True
End Sub

Private Sub SSOleDBCombo1_KeyPress(KeyAscii As Integer)
'SSOleDBCombo1.MoveNext
End Sub

Private Sub SSOleDBCombo1_LostFocus()
Call NormalBackground(SSOleDBCombo1)
End Sub

Private Sub SSOleDBCombo1_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCombo1)) > 0 Then
         If Not SSOleDBCombo1.IsItemInList Then
                SSOleDBCombo1.Text = ""
            End If
            If Len(Trim$(SSOleDBCombo1)) = 0 Then
            SSOleDBCombo1.SetFocus
            Cancel = True
            End If
            End If
End Sub
