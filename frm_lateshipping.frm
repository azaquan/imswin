VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lateshipping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Shipping"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   3870
   Tag             =   "03020800"
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox txt_diff 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTtodate 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60030977
      CurrentDate     =   36523
   End
   Begin MSComCtl2.DTPicker DTfromdate 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60030977
      CurrentDate     =   36523
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_FROMSUP 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_tosup 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "days late"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label lbl_todate 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1485
      Width           =   1800
   End
   Begin VB.Label lbl_fromdate 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1125
      Width           =   1800
   End
   Begin VB.Label lbl_tosupplier 
      Caption         =   "To Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   765
      Width           =   1800
   End
   Begin VB.Label lbl_fromsupp 
      Caption         =   "From Supplier"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   405
      Width           =   1800
   End
   Begin VB.Label lbl_diff 
      Caption         =   "More than"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1845
      Width           =   1800
   End
End
Attribute VB_Name = "frm_lateshipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'check enter date data format and from date less than to date
'if wrong type show message

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

If Not IsNumeric(txt_diff.text) Then

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("M00271") 'J added
    MsgBox IIf(msg1 = "", "Please enter a valid Number", msg1), , "Diff" 'J modified
    '---------------------------------------------
  
  Call txt_diff_Validate("true")
  txt_diff.text = ""
 ElseIf DTfromdate.value > DTtodate.value Then
 
    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("M00003") 'J added
    msg2 = translator.Trans("L00318") 'J added
    MsgBox IIf(msg1 = "", "Make Sure The To Date is greater than the From date", msg1), , IIf(msg2 = "", "Date", msg2) 'J modified
    '---------------------------------------------
     DTfromdate_Validate ("true")
  Else

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\lateshipping.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "fromsup;" + IIf(UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL", "ALL", Trim$(SSOleDB_FROMSUP.text)) + ";true"
        .ParameterFields(2) = "tosup;" + IIf(UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL", "", Trim$(SSOleDB_tosup.text)) + ";true"
        .ParameterFields(3) = "fromdate;date(" & Year(DTfromdate.value) & "," & Month(DTfromdate.value) & "," & Day(DTfromdate.value) & ");true"
        .ParameterFields(4) = "todate;date(" & Year(DTtodate.value) & "," & Month(DTtodate.value) & "," & Day(DTtodate.value) & ");true"
        .ParameterFields(5) = "diff;" + txt_diff.text + ";true"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00201") 'J added
        .WindowTitle = IIf(msg1 = "", "Late Shipping", msg1) 'J modified
        Call translator.Translate_Reports("lateshipping.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
     End With
     End If
     Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'set to supplier equal to from supplier

Private Sub ssoledb_fromsup_Click()

If UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL" Then
SSOleDB_tosup.Enabled = False
SSOleDB_tosup.text = ""
Else
SSOleDB_tosup.Enabled = True
End If
    SSOleDB_tosup = SSOleDB_FROMSUP
End Sub

'validate from date

Private Sub DTfromdate_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub

'SQL statement get supplier informatio,populate combo

Private Sub Form_Load()

    'Added by Juan (9/12/2000) for Multilingual
    Call translator.Translate_Forms("frm_lateshipping")
    '------------------------------------------

Screen.MousePointer = 11
Me.Visible = True
Me.Refresh
DoEvents

Set rs = New ADODB.Recordset
'Me.Height = 3600 'J hidden
'Me.Width = 3750 'J hidden
rs.ActiveConnection = deIms.cnIms
rs.Source = "select sup_code,sup_name from supplier where sup_npecode='" & deIms.NameSpace & "' order by sup_code"
rs.Open
SSOleDB_FROMSUP.text = "ALL"
SSOleDB_tosup.text = "ALL"
If get_status(rs) Then
SSOleDB_FROMSUP.ColumnHeaders = True
SSOleDB_tosup.ColumnHeaders = True

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    msg2 = translator.Trans("L00050") 'J added
    SSOleDB_FROMSUP.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_FROMSUP.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    SSOleDB_tosup.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    SSOleDB_tosup.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDB_FROMSUP.FieldSeparator = Chr$(1)
    SSOleDB_tosup.FieldSeparator = Chr$(1)
    
    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00520") 'J added
    SSOleDB_FROMSUP.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
    SSOleDB_tosup.AddItem "ALL" & Chr$(1) & "ALL"
        '---------------------------------------------
    
Do While (Not rs.EOF)
SSOleDB_FROMSUP.AddItem (rs!sup_code & Chr$(1) & rs!sup_name & "")
SSOleDB_tosup.AddItem (rs!sup_code & Chr$(1) & rs!sup_name & "")
rs.MoveNext
Loop
End If
Caption = Caption + " - " + Tag

    DTfromdate = FirstOfMonth
    DTtodate = Now
    
    
    With frm_lateshipping
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
Screen.MousePointer = 0
End Sub

'resize from

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then
'Me.Height = 3600 'J hidden
'Me.Width = 3750 'J hidden
End If
End Sub

Private Sub SSOleDB_FROMSUP_GotFocus()
Call HighlightBackground(SSOleDB_FROMSUP)
End Sub

Private Sub SSOleDB_FROMSUP_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_FROMSUP.DroppedDown Then SSOleDB_FROMSUP.DroppedDown = True
End Sub

Private Sub SSOleDB_FROMSUP_LostFocus()
Call NormalBackground(SSOleDB_FROMSUP)
End Sub

Private Sub SSOleDB_FROMSUP_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_FROMSUP)) > 0 Then
         If Not SSOleDB_FROMSUP.IsItemInList Then
                SSOleDB_FROMSUP.text = ""
            End If
            If Len(Trim$(SSOleDB_FROMSUP)) = 0 Then
            SSOleDB_FROMSUP.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDB_tosup_GotFocus()
Call HighlightBackground(SSOleDB_tosup)
End Sub

Private Sub SSOleDB_tosup_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_tosup.DroppedDown Then SSOleDB_tosup.DroppedDown = True
End Sub

Private Sub SSOleDB_tosup_LostFocus()
Call NormalBackground(SSOleDB_tosup)

End Sub

Private Sub SSOleDB_tosup_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_tosup)) > 0 Then
          If SSOleDB_tosup.Rows > 0 Then
         If Not SSOleDB_tosup.IsItemInList Then
                SSOleDB_tosup.text = ""
            End If
            End If
            If Len(Trim$(SSOleDB_tosup)) = 0 Then
            SSOleDB_tosup.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub txt_diff_GotFocus()
Call HighlightBackground(txt_diff)
End Sub

Private Sub txt_diff_LostFocus()
Call NormalBackground(txt_diff)
End Sub

Private Sub txt_diff_Validate(Cancel As Boolean)
Dim x As Boolean
x = Cancel
If Len(Trim$(Trim$(txt_diff))) > 0 Then
    If CDbl(txt_diff) < 0 Then
       Cancel = True
       MsgBox " You cannot insert Negative Values "
       txt_diff.SetFocus
       txt_diff.text = ""
       Exit Sub
    End If
    End If

End Sub

'check recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
   End Function
