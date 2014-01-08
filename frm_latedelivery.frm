VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_latedelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Delivery"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   4440
   Tag             =   "03020700"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_buyer 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
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
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   " &Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   " &Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt_daysleft 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTenddate 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20447233
      CurrentDate     =   36522
   End
   Begin MSComCtl2.DTPicker DTbegdate 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20447233
      CurrentDate     =   36522
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_FROMSUP 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   480
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_tosup 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   840
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_servcode 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
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
   Begin VB.Label lbl_servcode 
      Caption         =   "Service code"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2320
      Width           =   2000
   End
   Begin VB.Label lbl_dateslate 
      Caption         =   "Days Late"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2680
      Width           =   2000
   End
   Begin VB.Label lbl_buyer 
      Caption         =   "Buyer"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1960
      Width           =   2000
   End
   Begin VB.Label lbl_enddate 
      Caption         =   "End Date"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1600
      Width           =   2000
   End
   Begin VB.Label lbl_begdate 
      Caption         =   "Begining date"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1240
      Width           =   2000
   End
   Begin VB.Label lbl_endsupp 
      Caption         =   "To Supplier"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   880
      Width           =   2000
   End
   Begin VB.Label lbl_beg_supp 
      Caption         =   "From Supplier"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   520
      Width           =   2000
   End
End
Attribute VB_Name = "frm_latedelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset


'close form
Private Sub cmd_cancel_Click()
Unload Me
End Sub

'check enter date format, if wrong data type show message,
'if begin date less than end date show message
'else print crystal report

Private Sub cmd_ok_Click()
Dim Cancel As Boolean
On Error GoTo ErrHandler

If Not IsNumeric(txt_daysleft.text) Then

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("M00271") 'J added
    msg2 = translator.Trans("L00200") 'J added
    MsgBox IIf(msg1 = "", "Please enter a valid Number", msg1), , IIf(msg2 = "", "Days Late", msg2) 'J modified
    '---------------------------------------------
    
  Cancel = True
  Call txt_daysleft_Validate(Cancel)
  txt_daysleft.text = ""
  txt_daysleft.TabIndex = cmd_ok.TabIndex
  ElseIf DTbegdate.value > DTenddate.value Then
  
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00272") 'J added
        msg2 = translator.Trans("L00318") 'J added
        MsgBox IIf(msg1 = "", "Make sure the beginning date is greater than or equal To date", msg1), , IIf(msg2 = "", "Date", msg2) 'J modified
        '---------------------------------------------
        
  DTbegdate_Validate ("true")
  Else
With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\latedelivery.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "begsup;" + IIf(UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL", "ALL", SSOleDB_FROMSUP.text) + ";true"
        .ParameterFields(2) = "endsup;" + IIf(UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL", "", Trim$(SSOleDB_tosup.text)) + ";true"
        .ParameterFields(3) = "begdate;date(" & Year(DTbegdate.value) & "," & Month(DTbegdate.value) & "," & Day(DTbegdate.value) & ");true"
        .ParameterFields(4) = "enddate;date(" & Year(DTenddate.value) & "," & Month(DTenddate.value) & "," & Day(DTenddate.value) & ");true"
        .ParameterFields(5) = "buyer;" + IIf(UCase(Trim$(Combo_buyer.text)) = "ALL", "ALL", Trim$(Combo_buyer.text)) + ";true"
        .ParameterFields(6) = "servcode;" + IIf(UCase(Trim$(SSOleDB_servcode.text)) = "ALL", "ALL", Trim$(SSOleDB_servcode.text)) + ";true"
        .ParameterFields(7) = "dayslate;" + Trim$(txt_daysleft.text) + ";true"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00194") 'J added
        .WindowTitle = IIf(msg1 = "", "Late Delivery", msg1) 'J modified
        Call translator.Translate_Reports("latedelivery.rpt") 'J added
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
'by shah
''''''Private Sub Combo_buyer_Click()
''''''On Error Resume Next
''''''Dim str As String
''''''Dim cmd As ADODB.Command
''''''Dim rst As ADODB.Recordset
''''''
''''''
''''''    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
''''''
''''''    With cmd
''''''        .CommandText = " SELECT buy_username "
''''''        .CommandText = .CommandText & " from buyer "
''''''        .CommandText = .CommandText & " WHERE buy_npecode = '" & deIms.NameSpace & "'"
''''''        .CommandText = .CommandText & " order by buy_username"
''''''         Set rst = .Execute
''''''    End With
''''''
''''''
''''''    str = Chr$(1)
''''''    Combo_buyer.FieldSeparator = str
''''''    If rst.RecordCount = 0 Then GoTo CleanUp
''''''
''''''    rst.MoveFirst
''''''
''''''    Combo_buyer.AddItem (("ALL" & str) & "ALL" & "")
''''''
''''''    Do While ((Not rst.EOF))
''''''        Combo_buyer.AddItem rst!buy_username
''''''
''''''        rst.MoveNext
''''''    Loop
''''''
''''''CleanUp:
''''''    rst.Close
''''''    Set cmd = Nothing
''''''    Set rst = Nothing
''''''If Err Then Call LogErr(Name & "::GetBuyerName", Err.Description, Err.number, True)
''''''
''''''End Sub

Private Sub Combo_buyer_GotFocus()
Call HighlightBackground(Combo_buyer)
End Sub

Private Sub Combo_buyer_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_buyer.DroppedDown Then Combo_buyer.DroppedDown = True
End Sub

Private Sub Combo_buyer_LostFocus()
Call NormalBackground(Combo_buyer)
End Sub

Private Sub Combo_buyer_Validate(Cancel As Boolean)
If Len(Trim$(Combo_buyer)) > 0 Then
         If Not Combo_buyer.IsItemInList Then
                Combo_buyer.text = ""
            End If
            If Len(Trim$(Combo_buyer)) = 0 Then
            Combo_buyer.SetFocus
            Cancel = True
            End If
            End If
End Sub

'resize form

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
'Me.Height = 4470
  'Me.Width = 4000
  End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'set to supplier equal to from supplier

Private Sub ssoledb_fromsup_Click()
Dim str As String
  If UCase(Trim$(SSOleDB_FROMSUP.text)) = "ALL" Then
        
    SSOleDB_tosup.text = ""
    SSOleDB_tosup.Enabled = False
    Else
    SSOleDB_tosup.Enabled = True
   End If
   
   SSOleDB_tosup = SSOleDB_FROMSUP
   'str = "select sup_code from supplier where sup_npecode='" & deIms.NameSpace & " '"
   ' Set rs = get_recordset(str)
   '
   ' If get_status(rs) Then
   '
   ' End If
   '   combo_endsupp.Enabled = True
   ' Call CleanUp
    'Else
    ''  combo_endsupp.Text = ""
    ' combo_endsupp.Enabled = False
    'End If
End Sub

Private Sub DTbegdate_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub

'SQL statement get supplier,buyer,service information and populate combo
 
Sub Form_Load()
Dim str As String

    Screen.MousePointer = 11
    'Added by Juan (9/12/2000) for Multilingual
    Call translator.Translate_Forms("frm_latedelivery")
    '------------------------------------------
    Me.Visible = True
    Me.Refresh
    DoEvents

  'Me.Height = 4470
  'Me.Width = 4000
  SSOleDB_FROMSUP.FieldSeparator = Chr$(1)
  SSOleDB_tosup.FieldSeparator = Chr$(1)
    str = "select sup_code,sup_name from supplier where sup_npecode='" & deIms.NameSpace & " 'order by sup_code"
    Set rs = get_recordset(str)

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
    msg1 = translator.Trans("L00520") 'J added
    SSOleDB_FROMSUP.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
    '---------------------------------------------
    SSOleDB_tosup.AddItem "ALL" & Chr$(1) & "ALL"
    Do While (Not rs.EOF)
        SSOleDB_FROMSUP.AddItem (rs!sup_code & Chr$(1) & rs!sup_name & "")
        SSOleDB_tosup.AddItem (rs!sup_code & Chr$(1) & rs!sup_name & "")
        rs.MoveNext
    Loop
   
    End If
      SSOleDB_FROMSUP.text = "ALL"
      SSOleDB_tosup.text = "ALL"
      SSOleDB_servcode.text = "ALL"
      Combo_buyer.text = "ALL"
    Call CleanUp
   
     str = "select buy_username from buyer where buy_npecode='" & deIms.NameSpace & " 'order by buy_username"
     
    Set rs = get_recordset(str)
    If get_status(rs) Then
    
       'Call PopuLateFromRecordSet(Combo_buyer, rs, "buy_username", True)
        Do While Not rs.EOF
          Combo_buyer.AddItem rs!buy_username
          rs.MoveNext
        Loop
         Call Combo_buyer.AddItem("ALL", 0)
         
    End If
     
  
 
    Call CleanUp
      SSOleDB_servcode.FieldSeparator = Chr$(1)
        str = "select srvc_code,srvc_desc from servcode where srvc_npecode='" & deIms.NameSpace & " '"
    Set rs = get_recordset(str)
    If get_status(rs) Then
    SSOleDB_servcode.ColumnHeaders = True
    
    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00520") 'J added
    SSOleDB_servcode.AddItem ("ALL" & Chr$(1) & IIf(msg1 = "", "ALL", msg1) & "") 'J modified
    msg1 = translator.Trans("L00028") 'J added
    SSOleDB_servcode.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    msg1 = translator.Trans("L00029") 'J added
    SSOleDB_servcode.Columns(1).Caption = IIf(msg1 = "", "Description", msg1) 'J modified
    '---------------------------------------------
    
    Do While (Not rs.EOF)
      SSOleDB_servcode.AddItem (rs!srvc_code & Chr$(1) & rs!srvc_desc & "")
      rs.MoveNext
      Loop
    End If
      
      Caption = Caption + " - " + Tag
      
    Call CleanUp
    
    DTbegdate = FirstOfMonth
    DTenddate = Now
    
    With frm_latedelivery
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
    
    Screen.MousePointer = 0
End Sub

Public Function get_recordset(str As String) As ADODB.Recordset
Set cmd = New ADODB.Command
With cmd
  .ActiveConnection = deIms.cnIms
  .CommandType = adCmdText
  .CommandText = str
  Set get_recordset = .Execute
  End With
   End Function

'set memory free

Public Sub CleanUp()
Set cmd = Nothing
Set rs = Nothing
End Sub

'check recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
 
End Function


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

Private Sub SSOleDB_servcode_GotFocus()
Call HighlightBackground(SSOleDB_servcode)
End Sub

Private Sub SSOleDB_servcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_servcode.DroppedDown Then SSOleDB_servcode.DroppedDown = True
End Sub

Private Sub SSOleDB_servcode_LostFocus()
Call NormalBackground(SSOleDB_servcode)
End Sub

Private Sub SSOleDB_servcode_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_servcode)) > 0 Then
         If Not SSOleDB_servcode.IsItemInList Then
                SSOleDB_servcode.text = ""
            End If
            If Len(Trim$(SSOleDB_servcode)) = 0 Then
            SSOleDB_servcode.SetFocus
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

Private Sub txt_daysleft_GotFocus()
Call HighlightBackground(txt_daysleft)
End Sub

Private Sub txt_daysleft_LostFocus()
Call NormalBackground(txt_daysleft)

End Sub

Private Sub txt_daysleft_Validate(Cancel As Boolean)
Dim x As Boolean
x = Cancel
If Len(Trim$(txt_daysleft)) > 0 Then
If CDbl(txt_daysleft) < 0 Then
       Cancel = True
       MsgBox " You cannot insert Negative Values "
       txt_daysleft.SetFocus
       txt_daysleft.text = ""
       Exit Sub
    End If
    End If

End Sub
