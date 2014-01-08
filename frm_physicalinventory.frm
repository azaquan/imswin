VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_physicalinventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Physical Inventory"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   3885
   Tag             =   "03030500"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_compcode 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      FieldSeparator  =   ";"
      Columns(0).Width=   3200
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_location 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Label showzero 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Zero"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2000
   End
   Begin VB.Label showqty 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Qty"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2000
   End
   Begin VB.Label lbl_locacode 
      Caption         =   "Location Code"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label lbl_compcode 
      Caption         =   "Company Code"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   560
      Width           =   2000
   End
End
Attribute VB_Name = "frm_physicalinventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Check1_GotFocus()
Call HighlightBackground(Check1)
End Sub

Private Sub Check1_LostFocus()
Call NormalBackground(Check1)
End Sub



Private Sub Check2_GotFocus()
Call HighlightBackground(Check2)
End Sub

Private Sub Check2_LostFocus()
Call NormalBackground(Check2)
End Sub

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        Printer.Orientation = 2  'Juan 2012/6/5
        .ReportFileName = FixDir(App.Path) + "CRreports\physinvt.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "company;" + Trim$(Combo_compcode.Text) + ";true"
        .ParameterFields(2) = "ware;" + Trim$(SSOleDB_location.Text) + ";true"
        .ParameterFields(3) = "showqty;" + IIf(Check1.value = 0, "N", "Y") + ";true"
        .ParameterFields(4) = "showzero;" + IIf(Check2.value = 0, "N", "Y") + ";true"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00185") 'J added
        .WindowTitle = IIf(msg1 = "", "Physical Inventory", msg1) 'J modified
        Call translator.Translate_Reports("physinvt.rpt") 'J added
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

Private Sub Combo_compcode_Click()
GetlocationName
End Sub

Private Sub Combo_compcode_DropDown()
msg1 = translator.Trans("L00050") 'J added
    msg2 = translator.Trans("L00028") 'J added
    Combo_compcode.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    Combo_compcode.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
    '---------------------------------------------
    
    Combo_compcode.Columns(0).Width = 2000
    Combo_compcode.Columns(1).Width = 1500
End Sub

Private Sub Combo_compcode_GotFocus()
Call HighlightBackground(Combo_compcode)
End Sub

Private Sub Combo_compcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_compcode.DroppedDown Then Combo_compcode.DroppedDown = True
End Sub

Private Sub Combo_compcode_KeyPress(KeyAscii As Integer)
'Combo_compcode.MoveNext
End Sub

Private Sub Combo_compcode_LostFocus()
Call NormalBackground(Combo_compcode)
End Sub


Private Sub Combo_compcode_Validate(Cancel As Boolean)
If Len(Trim$(Combo_compcode)) > 0 Then
         If Not Combo_compcode.IsItemInList Then
                Combo_compcode.Text = ""
            End If
            If Len(Trim$(Combo_compcode)) = 0 Then
            Combo_compcode.SetFocus
            Cancel = True
            End If
            End If
End Sub


Private Sub Form_Load()
'Me.Height = 3450
'Me.Width = 4000
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim str As String

    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Translate_Forms("frm_physicalinventory")
    '------------------------------------------

 
  With rs
     .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "'"
     .Source = .Source & " order by com_compcode "
     .ActiveConnection = deIms.cnIms
     .Open
 End With
 
 
If get_status(rs) Then

Combo_compcode.Text = "ALL"


   
Combo_compcode.AddItem ("ALL" & ";" & "ALL")
SSOleDB_location.AddItem "ALL" & SSOleDB_location.FieldSeparator & "ALL"
rs.MoveFirst
Do While (Not rs.EOF)
Combo_compcode.AddItem (rs!com_compcode & ";" & rs!com_name & " ")


rs.MoveNext
Loop
Set rs = Nothing
End If

'rs1.Source = "Select loc_locacode,loc_name from location where loc_npecode='" & deIms.NameSpace & "'"
'rs1.ActiveConnection = deIms.cnIms
'rs1.Open
'If get_status(rs1) Then
'SSOleDB_location.FieldSeparator = Chr$(1)
'SSOleDB_location.ColumnHeaders = True
'SSOleDB_location.Columns(0).Caption = "Code"
'SSOleDB_location.Columns(1).Caption = "Name"
'Do While (Not rs1.EOF)
'SSOleDB_location.AddItem (rs1!loc_locacode & Chr$(1) & rs1!loc_name & "")
'rs1.MoveNext
'Loop
'Set rs1 = Nothing
'End If

GetlocationName
 SSOleDB_location.Text = "ALL"
frm_physicalinventory.Caption = frm_physicalinventory.Caption + " - " + frm_physicalinventory.Tag

Me.Left = Round((Screen.Width - Me.Width) / 2)
Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub


'SQL statement get location list for location combo

Private Sub GetlocationName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    If Combo_compcode <> "ALL" Then
        Set cmd = MakeCommand(deIms.cnIms, adCmdText)
        
        With cmd
            .CommandText = " SELECT loc_locacode,loc_name "
            .CommandText = .CommandText & " From location "
            .CommandText = .CommandText & " WHERE loc_npecode = '" & deIms.NameSpace & "'"
            .CommandText = .CommandText & " and loc_compcode = '" & Combo_compcode & "'"
            .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
            .CommandText = .CommandText & " order by loc_locacode"
             Set rst = .Execute
        End With
    
    
        str = Chr$(1)
        SSOleDB_location.FieldSeparator = str
        If rst.RecordCount = 0 Then GoTo CleanUp
           
        rst.MoveFirst
    End If
    SSOleDB_location.RemoveAll
    SSOleDB_location = ""
       
    If UCase(Trim$(Combo_compcode)) = "ALL" Then
        SSOleDB_location.AddItem "ALL" & ";" & "ALL"
    Else
        Do While ((Not rst.EOF))
            SSOleDB_location.AddItem rst!loc_locacode & str & (rst!loc_name & "")
            rst.MoveNext
        Loop
    End If
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetlocationName", Err.Description, Err.number, True)
End Sub


'THIS PROCEDURE WILL NOT BE USED IF LOGICAL WAREHOUSE IS USED INSTEAD OF SUBLOCATION MUZAMMIL  11/06/00
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
    SSOleDB_location.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDB_location.RemoveAll
    
    rst.MoveFirst
      
    SSOleDB_location.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDB_location.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub



'check recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
   End Function

'resize form

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        'Me.Height = 3450
        'Me.Width = 4000
    End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub


Private Sub SSOleDB_location_GotFocus()
Call HighlightBackground(SSOleDB_location)
End Sub

Private Sub SSOleDB_location_InitColumnProps()
SSOleDB_location.Columns(0).Caption = "Code"
SSOleDB_location.Columns(1).Caption = "Name"
End Sub

Private Sub SSOleDB_location_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_location.DroppedDown Then SSOleDB_location.DroppedDown = True
End Sub

Private Sub SSOleDB_location_KeyPress(KeyAscii As Integer)
'SSOleDB_location.MoveNext
End Sub

Private Sub SSOleDB_location_LostFocus()
Call NormalBackground(SSOleDB_location)
End Sub

Private Sub SSOleDB_location_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_location)) > 0 Then
         If Not SSOleDB_location.IsItemInList Then
                SSOleDB_location.Text = ""
            End If
            If Len(Trim$(SSOleDB_location)) = 0 Then
            SSOleDB_location.SetFocus
            Cancel = True
            End If
            End If
End Sub
