VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_gnrlstatustransac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Status Report (By Transaction)"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   7230
   Tag             =   "02020600"
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3840
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTtodate 
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   60162051
      CurrentDate     =   36525
   End
   Begin MSComCtl2.DTPicker DTfromdate 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   60162051
      CurrentDate     =   36525
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_fromservice 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowSelectionStyle=   1
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_toservice 
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDb_statusitem 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_delivery 
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_statusship 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_statusinventory 
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_DocType 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
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
      BackColorOdd    =   16771818
      RowHeight       =   423
      SplitterPos     =   12
      Columns.Count   =   2
      Columns(0).Width=   1773
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5292
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_comp 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      DividerStyle    =   0
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_loc 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_frombuyer 
      Height          =   315
      Left            =   1920
      TabIndex        =   28
      Top             =   720
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowSelectionStyle=   1
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_tobuyer 
      Height          =   315
      Left            =   5280
      TabIndex        =   29
      Top             =   720
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowSelectionStyle=   1
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_frompo 
      Height          =   315
      Left            =   1920
      TabIndex        =   30
      Top             =   1440
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowSelectionStyle=   1
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_topo 
      Height          =   315
      Left            =   5280
      TabIndex        =   31
      Top             =   1440
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowSelectionStyle=   1
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label lbl_company 
      Caption         =   "Company"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label lbl_ware 
      Caption         =   "Location"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Document Type"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label lbl_statusinventory 
      Caption         =   "Status Inventory"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label lbl_statusship 
      Caption         =   "Status Shipping"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label lbl_topo 
      Caption         =   "To Po"
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   1440
      Width           =   1680
   End
   Begin VB.Label lbl_statusdelivery 
      Caption         =   "Status Delivery"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label lbl_statusitem 
      Caption         =   "Status Item"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label lbl_frompo 
      Caption         =   "From Po"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   1680
   End
   Begin VB.Label lbl_todate 
      Caption         =   "To Date"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label lbl_fromdate 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label lbl_tobuyer 
      Caption         =   "To Buyer"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label lbl_frombuyer 
      Caption         =   "From Buyer"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label lbl_toservice 
      Caption         =   "To Service"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   360
      Width           =   1680
   End
   Begin VB.Label lbl_fromservice 
      Caption         =   "From Service"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   1680
   End
End
Attribute VB_Name = "frm_gnrlstatustransac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_buyer As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs_po As ADODB.Recordset
Dim x As Integer
Dim Y As Integer
Dim z As Integer
Dim a As Integer


'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameters and application path

Private Sub cmd_ok_Click()
On Error Resume Next

If DTtodate.Value >= DTfromdate.Value Then
With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\gnrlstatustransac.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "fromservice ;" + IIf(UCase(Trim$(SSOleDB_fromservice.text)) = "ALL", "ALL", SSOleDB_fromservice.text) + ";true"
        .ParameterFields(2) = "toservice;" + IIf(UCase(Trim$(SSOleDB_fromservice.text)) = "ALL", "", Trim$(SSOleDB_toservice.text)) + ";true"
        .ParameterFields(3) = "frombuyer;" + IIf(UCase(Trim$(Combo_frombuyer.text)) = "ALL", "ALL", Trim$(Combo_frombuyer.text)) + ";true"
        .ParameterFields(4) = "tobuyer;" + IIf(UCase(Trim$(Combo_frombuyer.text)) = "ALL", "", Trim$(Combo_tobuyer.text)) + ";true"
        .ParameterFields(5) = "fromdate;date(" & Year(DTfromdate.Value) & "," & Month(DTfromdate.Value) & "," & Day(DTfromdate.Value) & ");true"
        .ParameterFields(6) = "todate;date(" & Year(DTtodate.Value) & "," & Month(DTtodate.Value) & "," & Day(DTtodate.Value) & ");true"
        .ParameterFields(7) = "frompo;" + IIf(UCase(Trim$(Combo_frompo.text)) = "ALL", "ALL", Trim$(Combo_frompo.text)) + ";true"
        .ParameterFields(8) = "topo;" + IIf(UCase(Trim$(Combo_frombuyer.text)) = "ALL", "", Trim$(Combo_topo.text)) + ";true"
        .ParameterFields(9) = "statusitem;" + IIf(UCase(Trim$(SSOleDb_statusitem.text)) = "ALL", "ALL", Trim$(SSOleDb_statusitem.text)) + ";true"
        .ParameterFields(10) = "statusdel;" + IIf(UCase(Trim$(SSOleDB_delivery.text)) = "ALL", "ALL", Trim$(SSOleDB_delivery.text)) + ";true"
        .ParameterFields(11) = "statusship;" + IIf(UCase(Trim$(SSOleDB_statusship.text)) = "ALL", "ALL", Trim$(SSOleDB_statusship.text)) + ";true"
        .ParameterFields(12) = "statuswh;" + IIf(UCase(Trim$(SSOleDB_statusinventory.text)) = "ALL", "ALL", Trim$(SSOleDB_statusinventory.text)) + ";true"
        .ParameterFields(13) = "doctype;" + IIf(UCase(Trim(SSOleDB_DocType.text)) = "ALL", "ALL", Trim(SSOleDB_DocType.text)) + ";true"
        .ParameterFields(14) = "comp;" + IIf(UCase(Trim(SSOleDB_comp.text)) = "ALL", "ALL", Trim(SSOleDB_comp.text)) + ";true"
        .ParameterFields(15) = "loc;" + IIf(UCase(Trim(SSOleDB_loc.text)) = "ALL", "ALL", Trim(SSOleDB_loc.text)) + ";true"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00537") 'J added
        .WindowTitle = IIf(msg1 = "", "General Status report", msg1) 'J modified
        Call translator.Translate_Reports("gnrlstatustransac.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
     
        .Action = 1
        .Reset
End With
Else

'Modified by Juan (9/11/2000) for Multilingual
msg1 = translator.Trans("M00003") 'J added
msg2 = translator.Trans("L00318") 'J added
MsgBox IIf(msg1 = "", "Make Sure The To Date is greater than the From date", msg1), , IIf(msg2 = "", "DATE", msg2) 'J modified
'---------------------------------------------

DTtodate_Validate ("true")
End If

If Err Then MsgBox Err.Description: Err.Clear
End Sub

'set tobuyer  equal to frombuyer

Private Sub Combo_frombuyer_Click()
If Combo_frombuyer.text = "ALL" Then
Y = 0
Combo_tobuyer.text = ""
Combo_tobuyer.Enabled = False
Else
Combo_tobuyer.Enabled = True
 Y = Y + 1
 End If
 
 Combo_tobuyer = Combo_frombuyer
End Sub

Private Sub Combo_frombuyer_GotFocus()
Call HighlightBackground(Combo_frombuyer)
End Sub

Private Sub Combo_frombuyer_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_frombuyer.DroppedDown Then Combo_frombuyer.DroppedDown = True

End Sub

Private Sub Combo_frombuyer_LostFocus()
Call NormalBackground(Combo_frombuyer)
End Sub

'set tobuyer  equal to frombuyer

Private Sub Combo_frombuyer_Validate(Cancel As Boolean)
     Combo_tobuyer = Combo_frombuyer
     If Len(Trim$(Combo_frombuyer)) > 0 Then
         If Not Combo_frombuyer.IsItemInList Then
                Combo_frombuyer.text = ""
            End If
            If Len(Trim$(Combo_frombuyer)) = 0 Then
           Combo_frombuyer.SetFocus
            Cancel = True
            End If
            End If
End Sub

'set from po number equal to to po number

Private Sub Combo_frompo_Click()
If Combo_frompo.text = "ALL" Then
z = 0
Combo_topo.text = ""
Combo_topo.Enabled = False
Else
Combo_topo.Enabled = True
 z = z + 1
 End If
  Combo_topo = Combo_frompo
End Sub

Private Sub Combo_frompo_GotFocus()
Call HighlightBackground(Combo_frompo)
End Sub

Private Sub Combo_frompo_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_frompo.DroppedDown Then Combo_frompo.DroppedDown = True

End Sub

Private Sub Combo_frompo_LostFocus()
Call NormalBackground(Combo_frompo)
End Sub

'set from po number equal to to po number

Private Sub Combo_frompo_Validate(Cancel As Boolean)
      Combo_topo = Combo_frompo
      If Len(Trim$(Combo_frompo)) > 0 Then
         If Not Combo_frompo.IsItemInList Then
                Combo_frompo.text = ""
            End If
            If Len(Trim$(Combo_frompo)) = 0 Then
           Combo_frompo.SetFocus
            Cancel = True
            End If
            End If
      
End Sub

Private Sub Combo_tobuyer_GotFocus()
Call HighlightBackground(Combo_tobuyer)
End Sub

Private Sub Combo_tobuyer_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_tobuyer.DroppedDown Then Combo_tobuyer.DroppedDown = True
End Sub

Private Sub Combo_tobuyer_LostFocus()
Call NormalBackground(Combo_tobuyer)
End Sub

Private Sub Combo_tobuyer_Validate(Cancel As Boolean)
If Len(Trim$(Combo_tobuyer)) > 0 Then
         If Not Combo_tobuyer.IsItemInList Then
                Combo_tobuyer.text = ""
            End If
            If Len(Trim$(Combo_tobuyer)) = 0 Then
           Combo_tobuyer.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub Combo_topo_GotFocus()
Call HighlightBackground(Combo_topo)
End Sub

Private Sub Combo_topo_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_topo.DroppedDown Then Combo_topo.DroppedDown = True

End Sub

Private Sub Combo_topo_LostFocus()
Call NormalBackground(Combo_topo)
End Sub

Private Sub Combo_topo_Validate(Cancel As Boolean)
If Len(Trim$(Combo_topo)) > 0 Then
         If Not Combo_topo.IsItemInList Then
               Combo_topo.text = ""
            End If
            If Len(Trim$(Combo_topo)) = 0 Then
           Combo_topo.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub DTtodate_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub

Private Sub Form_Activate()
Dim str As String
Dim rs1 As ADODB.Recordset
    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_gnrlstatustransac")
    '------------------------------------------
Screen.MousePointer = 11
Me.Refresh


a = 0
x = 0
Y = 0
z = 0
Set rs = New ADODB.Recordset
str = "select srvc_code, srvc_desc from servcode where srvc_npecode ='"
rs.Source = str & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
SSOleDB_fromservice.FieldSeparator = Chr$(1)
SSOleDB_toservice.FieldSeparator = Chr$(1)

If get_status(rs) Then

'Defalut VAlues
SSOleDB_fromservice.text = "ALL"
SSOleDB_toservice.text = "ALL"

SSOleDB_fromservice.AddItem ("ALL" & Chr$(1) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_fromservice.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
SSOleDB_toservice.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
rs.MoveNext
Loop
End If

'default values added by shakir on 02/13/01

 SSOleDB_comp.text = "ALL"
'SSOleDB_loc.text = "ALL"

SSOleDB_loc.text = ""
SSOleDB_loc.Enabled = False

'SSOleDB_comp.AddItem ("ALL" & vbTab & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_comp.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
SSOleDB_loc.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
rs.MoveNext
Loop
SSOleDB_loc = SSOleDB_comp
 If Err Then Err.Clear
 

'added by shakir on 02/13/01

If SSOleDB_fromservice.text = "ALL" Then
x = 0
SSOleDB_toservice.text = ""
SSOleDB_toservice.Enabled = False
Else
'SSOleDB_toservice.FieldSeparator = Chr$(1)
SSOleDB_toservice.Enabled = True
'If x = 1 Then Exit Sub
 x = x + 1
 End If
   SSOleDB_toservice = SSOleDB_fromservice
    If Err Then Err.Clear
   

'for ssoledb_comp...shah
Set rs = New ADODB.Recordset


SSOleDB_comp.FieldSeparator = Chr$(1)
SSOleDB_loc.FieldSeparator = Chr$(1)

    With rs
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "'"
        .Source = .Source & " order by com_compcode "
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
If get_status(rs) Then
SSOleDB_comp.AddItem ("ALL" & Chr$(1) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_comp.AddItem (rs!com_compcode & Chr$(1) & rs!com_name & " ")
rs.MoveNext
Loop
Set rs = Nothing
End If

'added by shakir on 02/13/01
If Combo_frombuyer.text = "ALL" Then
Y = 0
Combo_tobuyer.text = ""
Combo_tobuyer.Enabled = False
Else
Combo_tobuyer.Enabled = True
 Y = Y + 1
 End If
 Combo_tobuyer = Combo_frombuyer
 If Err Then Err.Clear
 
 'added by shakir on 02/13/01
 If Combo_frompo.text = "ALL" Then
z = 0
Combo_topo.text = ""
Combo_topo.Enabled = False
Else
Combo_topo.Enabled = True
 z = z + 1
 End If
  Combo_topo = Combo_frompo
If Err Then Err.Clear

'BUYER TABLE
Set rs_buyer = New ADODB.Recordset
str = "select buy_username from buyer where buy_npecode ='"
rs_buyer.Source = str & deIms.NameSpace & "'"
rs_buyer.ActiveConnection = deIms.cnIms
rs_buyer.Open
If get_status(rs_buyer) Then

'Defalut VAlues
Combo_frombuyer.text = "ALL"
Combo_tobuyer.text = "ALL"

'Combo_frombuyer.AddItem "ALL"
 Do While (Not rs_buyer.EOF)
 Combo_frombuyer.AddItem rs_buyer!buy_username
 Combo_tobuyer.AddItem rs_buyer!buy_username
 rs_buyer.MoveNext
 Loop
 End If
 
 'PO TABLE
 Set rs_po = New ADODB.Recordset
 str = "select po_ponumb from po where po_npecode ='"
rs_po.Source = str & deIms.NameSpace & "'"
rs_po.ActiveConnection = deIms.cnIms
rs_po.Open
If get_status(rs_po) Then


'Defalut VAlues
Combo_frompo = "ALL"
If rs_po.RecordCount > 0 Then
   rs_po.MoveFirst
   Combo_topo = rs_po!PO_PONUMB
End If

Combo_frompo.AddItem "ALL"
 Do While (Not rs_po.EOF)
 Combo_frompo.AddItem rs_po!PO_PONUMB
 Combo_topo.AddItem rs_po!PO_PONUMB
 rs_po.MoveNext
 Loop
 End If
 SSOleDb_statusitem.FieldSeparator = Chr$(1)
 
 'Modified by Juan (9/11/2000) for Multilingual
 msg1 = translator.Trans("L00520") 'J added
 
'Defalut VAlues
 SSOleDb_statusitem.text = "ALL"
 
 SSOleDb_statusitem.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00521") 'J added
 SSOleDb_statusitem.AddItem ("OP" & Chr$(1) & IIf(msg1 = "", "OPEN", msg1) & "") 'J modified
 msg1 = translator.Trans("L00522") 'J added
 SSOleDb_statusitem.AddItem ("OH" & Chr$(1) & IIf(msg1 = "", "ON HAND", msg1) & "") 'J modified
 msg1 = translator.Trans("L00523") 'J added
 SSOleDb_statusitem.AddItem ("CL" & Chr$(1) & IIf(msg1 = "", "CLOSED", msg1) & "") 'J modified
 msg1 = translator.Trans("L00524") 'J added
 SSOleDb_statusitem.AddItem ("CA" & Chr$(1) & IIf(msg1 = "", "CANCELLED", msg1) & "") 'J modified

 SSOleDB_statusship.FieldSeparator = Chr$(1)
 
 'Defalut VAlues
 SSOleDB_statusship.text = "ALL"

 msg1 = translator.Trans("L00520") 'J added
 SSOleDB_statusship.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00525") 'J added
 SSOleDB_statusship.AddItem ("S2" & Chr$(1) & IIf(msg1 = "", "NOT TOT. SHIPPED", msg1) & "") 'J modified
 msg1 = translator.Trans("L00526") 'J added
 SSOleDB_statusship.AddItem ("NS" & Chr$(1) & IIf(msg1 = "", "NOT SHIPPED", msg1) & "") 'J modified
 msg1 = translator.Trans("L00527") 'J added
 SSOleDB_statusship.AddItem ("SP" & Chr$(1) & IIf(msg1 = "", "SHIPPING,PARTIAL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00528") 'J added
 SSOleDB_statusship.AddItem ("SC" & Chr$(1) & IIf(msg1 = "", "SHIPPING,COMPLETE", msg1) & "") 'J modified
   
 SSOleDB_delivery.FieldSeparator = Chr$(1)
 
'Defalut VAlues
SSOleDB_delivery.text = "ALL"

 msg1 = translator.Trans("L00520") 'J added
 SSOleDB_delivery.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00529") 'J added
 SSOleDB_delivery.AddItem ("D2" & Chr$(1) & IIf(msg1 = "", "NOT TOT. RECEIVED", msg1) & "") 'J modified
 msg1 = translator.Trans("L00530") 'J added
 SSOleDB_delivery.AddItem ("NR" & Chr$(1) & IIf(msg1 = "", "NOT RECEIVED", msg1) & "") 'J modified
 msg1 = translator.Trans("L00531") 'J added
 SSOleDB_delivery.AddItem ("RP" & Chr$(1) & IIf(msg1 = "", "RECEPTION, PARTIAL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00532") 'J added
 SSOleDB_delivery.AddItem ("RC" & Chr$(1) & IIf(msg1 = "", "RECEPTION, COMPLETE", msg1) & "") 'J modified
 
 SSOleDB_statusinventory.FieldSeparator = Chr$(1)
 
 'Defalut VAlues
SSOleDB_statusinventory.text = "ALL"
 
 msg1 = translator.Trans("L00520") 'J added
 SSOleDB_statusinventory.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00533") 'J added
 SSOleDB_statusinventory.AddItem ("W2" & Chr$(1) & IIf(msg1 = "", "NOT TOT. IN INVENTORY", msg1) & "") 'J modified
 msg1 = translator.Trans("L00534") 'J added
 SSOleDB_statusinventory.AddItem ("NI" & Chr$(1) & IIf(msg1 = "", "NOT IN INVENTORY", msg1) & "") 'J modified
 msg1 = translator.Trans("L00535") 'J added
 SSOleDB_statusinventory.AddItem ("IP" & Chr$(1) & IIf(msg1 = "", "INVENTORY, PARTIAL", msg1) & "") 'J modified
 msg1 = translator.Trans("L00536") 'J added
 SSOleDB_statusinventory.AddItem ("IC" & Chr$(1) & IIf(msg1 = "", "INVENTORY, COMPLETE", msg1) & "") 'J modified
    Caption = Caption + " - " + Tag
    DTfromdate = Now
    DTtodate = Now
    
    
    'Added by Juan 15/1/2001
    Dim typesDATA As Recordset
    Dim sql As String
    With SSOleDB_DocType
       'Defalut VAlues
        .text = "ALL"
        
        .AddItem "ALL" + vbTab + "ALL TYPES"
        sql = "SELECT * FROM DOCTYPE WHERE doc_npecode = '" + deIms.NameSpace + "' ORDER BY doc_code"
        Set typesDATA = New ADODB.Recordset
        typesDATA.Open sql, deIms.cnIms, adOpenForwardOnly
        If typesDATA.RecordCount > 0 Then
            Do While Not typesDATA.EOF
                .AddItem typesDATA!doc_code + vbTab + typesDATA!doc_desc
                typesDATA.MoveNext
            Loop
        End If
        .MoveFirst
    End With
    'DTfromdate.Value = Format(Now, "yyyy") + "/" + Format(Now, "mm") + "/1"
    DTfromdate.Value = FirstOfMonth
    DTtodate.Value = Now
    
 
    

Screen.MousePointer = 0
End Sub

'SQL statement get record set data, populate combo box

Private Sub Form_Load()
    With frm_gnrlstatustransac
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'get record set status
Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
 
End Function
'''added by shakir
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
    SSOleDB_loc.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDB_loc.RemoveAll
    
    rst.MoveFirst
      
    SSOleDB_loc.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDB_loc.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub

'''added by shakir

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
    SSOleDB_loc.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    
    SSOleDB_loc.RemoveAll
    
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSOleDB_loc.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetlocationName", Err.Description, Err.number, True)
End Sub

'free memory

Public Sub CleanUp(rs As ADODB.Recordset)
Set rs = Nothing
End Sub
Public Sub fill_SHERIDAN(ctl As Control, STR1 As ADODB.Recordset, STR2 As ADODB.Recordset)
'CTL.FieldSeparator = Chr$(1)
'Do While (Not rs.EOF)
'CTL.AddItem (STR1 & Chr$(1) & STR2 & "")
'STR1.MoveNext
'STR2.MoveNext
'Loop
'rs.Close
'Set rs = Nothing
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
   
End Sub

'on 02/12/2001,,,shah

Private Sub SSOleDB_comp_DropDown()

msg1 = translator.Trans("L00028") 'J added
    msg1 = translator.Trans("L00050") 'J added
    SSOleDB_comp.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_comp.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDB_comp.Columns(0).Width = 1500
    SSOleDB_comp.Columns(1).Width = 2000
End Sub

'on 12/12/2001,,,shah

Private Sub SSOleDB_comp_click()
Dim str As String

    str = Trim$(SSOleDB_comp.Columns(0).text)
    If Trim$(SSOleDB_comp.Columns(0).text) = "ALL" Then
        SSOleDB_loc = ""
     SSOleDB_loc.Enabled = False
        Call GetalllocationName
    Else
        SSOleDB_loc = ""
        SSOleDB_loc.Enabled = True
        Call GetlocationName(str)
    End If
    
End Sub

Private Sub SSOleDB_comp_GotFocus()
Call HighlightBackground(SSOleDB_comp)
End Sub

Private Sub SSOleDB_comp_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_comp.DroppedDown Then SSOleDB_comp.DroppedDown = True

End Sub

Private Sub SSOleDB_comp_LostFocus()
Call NormalBackground(SSOleDB_comp)
End Sub

Private Sub SSOleDB_comp_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_comp)) > 0 Then
         If Not SSOleDB_comp.IsItemInList Then
               SSOleDB_comp.text = ""
            End If
            If Len(Trim$(SSOleDB_comp)) = 0 Then
          SSOleDB_comp.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_delivery_GotFocus()
Call HighlightBackground(SSOleDB_delivery)
End Sub

Private Sub SSOleDB_delivery_LostFocus()
Call NormalBackground(SSOleDB_delivery)
End Sub

Private Sub SSOleDB_delivery_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_delivery)) > 0 Then
         If Not SSOleDB_delivery.IsItemInList Then
               SSOleDB_delivery.text = ""
            End If
            If Len(Trim$(SSOleDB_delivery)) = 0 Then
           SSOleDB_delivery.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_DocType_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_DocType)) > 0 Then
         If Not SSOleDB_DocType.IsItemInList Then
               SSOleDB_DocType.text = ""
            End If
            If Len(Trim$(SSOleDB_DocType)) = 0 Then
          SSOleDB_DocType.SetFocus
            Cancel = True
            End If
            End If
End Sub

'set from service combo data equal to to service combo

Private Sub ssoledb_fromservice_Click()
If SSOleDB_fromservice.text = "ALL" Then
x = 0
SSOleDB_toservice.text = ""
SSOleDB_toservice.Enabled = False
Else
'SSOleDB_toservice.FieldSeparator = Chr$(1)
SSOleDB_toservice.Enabled = True
'If x = 1 Then Exit Sub
 x = x + 1
 ' rs.MoveFirst
'Do While (Not rs.EOF)
'SSOleDB_toservice.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
'  rs.MoveNext
'  Loop
End If
   If Err Then Err.Clear
End Sub

Private Sub SSOleDB_fromservice_GotFocus()
Call HighlightBackground(SSOleDB_fromservice)
End Sub

Private Sub SSOleDB_fromservice_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_fromservice.DroppedDown Then SSOleDB_fromservice.DroppedDown = True
End Sub

Private Sub SSOleDB_fromservice_LostFocus()
Call NormalBackground(SSOleDB_fromservice)
End Sub

'set from service combo data equal to to service combo

Private Sub SSOleDB_fromservice_Validate(Cancel As Boolean)
    SSOleDB_toservice = SSOleDB_fromservice
     
If Len(Trim$(SSOleDB_fromservice)) > 0 Then
         If Not SSOleDB_fromservice.IsItemInList Then
                SSOleDB_fromservice.text = ""
            End If
            If Len(Trim$(SSOleDB_fromservice)) = 0 Then
            SSOleDB_fromservice.SetFocus
            Cancel = True
            End If
            End If
End Sub

' by shah

Private Sub SSOleDB_loc_DropDown()
msg1 = translator.Trans("L00028")
    msg1 = translator.Trans("L00050")
    SSOleDB_loc.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_loc.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDB_loc.Columns(0).Width = 900
    SSOleDB_loc.Columns(1).Width = 2000

End Sub


Private Sub SSOleDB_loc_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_loc)) > 0 Then
         If Not SSOleDB_loc.IsItemInList Then
               SSOleDB_loc.text = ""
            End If
            If Len(Trim$(SSOleDB_loc)) = 0 Then
          SSOleDB_loc.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_statusinventory_GotFocus()
Call HighlightBackground(SSOleDB_statusinventory)
End Sub

Private Sub SSOleDB_statusinventory_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_statusinventory.DroppedDown Then SSOleDB_statusinventory.DroppedDown = True
End Sub

Private Sub SSOleDB_statusinventory_LostFocus()
Call NormalBackground(SSOleDB_statusinventory)
End Sub

Private Sub SSOleDB_statusinventory_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_statusinventory)) > 0 Then
         If Not SSOleDB_statusinventory.IsItemInList Then
               SSOleDB_statusinventory.text = ""
            End If
            If Len(Trim$(SSOleDB_statusinventory)) = 0 Then
           SSOleDB_statusinventory.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDb_statusitem_GotFocus()
Call HighlightBackground(SSOleDb_statusitem)
End Sub



Private Sub SSOleDb_statusitem_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDb_statusitem.DroppedDown Then SSOleDb_statusitem.DroppedDown = True

End Sub

Private Sub SSOleDb_statusitem_LostFocus()
Call NormalBackground(SSOleDb_statusitem)
End Sub

Private Sub SSOleDb_statusitem_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDb_statusitem)) > 0 Then
         If Not SSOleDb_statusitem.IsItemInList Then
               SSOleDb_statusitem.text = ""
            End If
            If Len(Trim$(SSOleDb_statusitem)) = 0 Then
           SSOleDb_statusitem.SetFocus
            Cancel = True
            End If
            End If
End Sub



Private Sub SSOleDB_statusship_GotFocus()
Call HighlightBackground(SSOleDB_statusship)
End Sub

Private Sub SSOleDB_statusship_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_statusship.DroppedDown Then SSOleDB_statusship.DroppedDown = True

End Sub

Private Sub SSOleDB_statusship_LostFocus()
Call NormalBackground(SSOleDB_statusship)
End Sub

Private Sub SSOleDB_statusship_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_statusship)) > 0 Then
         If Not SSOleDB_statusship.IsItemInList Then
               SSOleDB_statusship.text = ""
            End If
            If Len(Trim$(SSOleDB_statusship)) = 0 Then
           SSOleDB_statusship.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_toservice_GotFocus()
Call HighlightBackground(SSOleDB_toservice)
End Sub

Private Sub SSOleDB_toservice_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_toservice.DroppedDown Then SSOleDB_toservice.DroppedDown = True
End Sub

Private Sub SSOleDB_toservice_LostFocus()
Call NormalBackground(SSOleDB_toservice)
End Sub

Private Sub SSOleDB_toservice_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_toservice)) > 0 Then
         If Not SSOleDB_toservice.IsItemInList Then
                SSOleDB_toservice.text = ""
            End If
            If Len(Trim$(SSOleDB_toservice)) = 0 Then
            SSOleDB_toservice.SetFocus
            Cancel = True
            End If
            End If
End Sub

