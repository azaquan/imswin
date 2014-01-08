VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_sap_analysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sap Analysis"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4650
   Tag             =   "03040300"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo combo_stocknum 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
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
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBLocation 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   840
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCompany 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      FieldSeparator  =   ";"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
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
   Begin VB.Label Label2 
      Caption         =   "Stock Number"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1275
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1635
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lbl_location 
      Caption         =   "Location"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   915
      Width           =   2000
   End
   Begin VB.Label lbl_company 
      Caption         =   "Company"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   555
      Width           =   2000
   End
End
Attribute VB_Name = "frm_sap_analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer

'unload form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameters and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\auditSAP.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "compcode;" + Trim$(UCase(SSOleDBCompany.Text)) + ";true"
        .ParameterFields(2) = "locacode;" + Trim$(UCase(SSOleDBLocation.Text)) + ";true"
        .ParameterFields(3) = "Stocknumb;" + Trim$(UCase(combo_stocknum.Text)) + ";true"
        '.ParameterFields(3) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(3) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00264") 'J added
        .WindowTitle = IIf(msg1 = "", "SAP Analysis", msg1) 'J modified
        Call translator.Translate_Reports("auditSAP.rpt") 'J added
        Call translator.Translate_SubReports
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

RSstkMaster.Source = "select DISTINCT sap_stcknumb from sap where " _
    & "sap_compcode='" & Trim$(SSOleDBCompany) & "' and " _
    & "sap_loca='" & Trim$(SSOleDBLocation) & "'  and " _
    & "sap_npecode='" & deIms.NameSpace & "' ORDER BY sap_stcknumb"

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

Private Sub Combo_stocknum_GotFocus()
Call HighlightBackground(combo_stocknum)
End Sub

Private Sub combo_stocknum_KeyDown(KeyCode As Integer, Shift As Integer)
If Not combo_stocknum.DroppedDown Then combo_stocknum.DroppedDown = True
End Sub

Private Sub combo_stocknum_KeyPress(KeyAscii As Integer)
'combo_stocknum.MoveNext
End Sub

Private Sub Combo_stocknum_LostFocus()
Call NormalBackground(combo_stocknum)
End Sub

Private Sub combo_stocknum_Validate(Cancel As Boolean)
 If Len(Trim$(combo_stocknum)) > 0 Then
         If Not combo_stocknum.IsItemInList Then
                combo_stocknum.Text = ""
            End If
            If Len(Trim$(combo_stocknum)) = 0 Then
            combo_stocknum.SetFocus
            Cancel = True
            End If
            End If
End Sub

'SQL statement company recerdset

Private Sub Form_Load()
Dim str As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim x As Integer

'Modified by Juan (9/17/2000) for Multilingual
Call translator.Translate_Forms("frm_sap_analysis")
'---------------------------------------------

x = 0
'Me.Height = 2715
'Me.Width = 4005
    
    str = Chr$(1)
    Set cmd = New ADODB.Command
        
   frm_sap_analysis.Caption = frm_sap_analysis.Caption + " - " + frm_sap_analysis.Tag
    With cmd
      
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        .CommandText = "select com_compcode,com_name from company "
        .CommandText = .CommandText & " Where com_npecode= '" & deIms.NameSpace & "'"
        Set rs = .Execute
        
    End With
    combo_stocknum.RemoveAll
     combo_stocknum.Text = "ALL"
   If rs Is Nothing Then Exit Sub
   If rs.State And adStateOpen = adStateClosed Then Exit Sub

    If rs.EOF And rs.BOF Then GoTo CleanUp
    If rs.RecordCount = 0 Then GoTo CleanUp
    
    rs.MoveFirst
    SSOleDBCompany.FieldSeparator = str
 
    Do While (Not rs.EOF)

        SSOleDBCompany.AddItem ((rs!com_name & str) & rs!com_compcode & "")
        rs.MoveNext
    Loop
    
    Call GetCurrencylist
     SSOleDBCurrency = "USD"
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

CleanUp:
       rs.Close
       Set cmd = Nothing
       Set rs = Nothing
 
 

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

'resize form size

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then
'Me.Height = 2715
'Me.Width = 4005
End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'call function to get loaction

Private Sub SSOleDBCompany_Click()
Dim str As String
   str = SSOleDBCompany.Columns(1).Text
   If Len(str) Then
        SSOleDBLocation = ""
        SSOleDBLocation.RemoveAll
       Call AddLocation(GetLocation(Trim$(str)))
   End If
End Sub

'SQL statement get location for form

Public Function GetLocation(CompanyCode As String) As ADODB.Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "select loc_name, loc_locacode from location "
        .CommandText = .CommandText & " where loc_npecode ='" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "and loc_compcode='" & CompanyCode & "'"
        
        Set GetLocation = .Execute
    End With
    
    Set cmd = Nothing
End Function

'load recordset to combo

Public Sub AddLocation(rs As ADODB.Recordset)
Dim str As String

   str = Chr$(1)
   If rs Is Nothing Then Exit Sub
   If rs.State And adStateOpen = adStateClosed Then Exit Sub

    If rs.EOF And rs.BOF Then GoTo CleanUp
    If rs.RecordCount = 0 Then GoTo CleanUp
    
    rs.MoveFirst
    SSOleDBLocation.FieldSeparator = str
    SSOleDBLocation.RemoveAll
    Do While (Not rs.EOF)

        SSOleDBLocation.AddItem ((rs!loc_locacode & str) & rs!loc_name & "")
        rs.MoveNext
    Loop


CleanUp:
       rs.Close
       End Sub

'set company combo format

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
'SSOleDBCompany.MoveNext

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
            If Len(Trim$(SSOleDBCompany)) > 20 Then
        MsgBox "Value can not be greater than 20 characters."
        Cancel = True
        SSOleDBCompany.SetFocus
        End If
End If
End Sub

'set location combo format

Private Sub SSOleDBlocation_DropDown()

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    msg2 = translator.Trans("L00028") 'J added
    SSOleDBLocation.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    SSOleDBLocation.Columns(0).Caption = IIf(msg2 = "", "Code", msg2) 'J modified
    '---------------------------------------------
    
    SSOleDBLocation.Columns(1).Width = 2000
    SSOleDBLocation.Columns(0).Width = 800
End Sub

Private Sub SSOleDBlocation_GotFocus()
Call HighlightBackground(SSOleDBLocation)
End Sub

Private Sub SSOleDBlocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBLocation.DroppedDown Then SSOleDBLocation.DroppedDown = True
End Sub

Private Sub SSOleDBlocation_KeyPress(KeyAscii As Integer)
'SSOleDBLocation.MoveNext
End Sub

Private Sub SSOleDBlocation_LostFocus()
Call NormalBackground(SSOleDBLocation)
End Sub

Private Sub SSOleDBlocation_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCompany)) > 0 Then
    If Not SSOleDBLocation.IsItemInList Then
           SSOleDBLocation.Text = ""
       End If
       If Len(Trim$(SSOleDBLocation)) = 0 Then
       SSOleDBLocation.SetFocus
       Cancel = True
       End If
End If
End Sub
