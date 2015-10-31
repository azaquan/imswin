VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_historicalstock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historical Stock Movement"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   4590
   Tag             =   "03030500"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo combo_tostock 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_compcode 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_location 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2175
      DataFieldList   =   "Column 0"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo combo_fromstock 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
      DataFieldList   =   "Column 0"
      _Version        =   196617
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
   Begin VB.Label Label1 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lbl_tostock 
      Caption         =   "To Stock number"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1980
      Width           =   2040
   End
   Begin VB.Label lbl_fromstock 
      Caption         =   "From Stock number"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1620
      Width           =   2040
   End
   Begin VB.Label lbl_locacode 
      Caption         =   "Location Code"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   2040
   End
   Begin VB.Label lbl_compcode 
      Caption         =   "Company Code"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   560
      Width           =   2040
   End
End
Attribute VB_Name = "frm_historicalstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As Integer

'set form close

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter an dapplication path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\histmovtstck.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "company;" + Trim$(UCase(Combo_compcode.Text)) + ";true"
        .ParameterFields(2) = "ware;" + Trim$(UCase(SSOleDB_location.Text)) + ";true"
        .ParameterFields(3) = "fromstock;" + IIf(Trim$(combo_fromstock.Text) = "ALL", "ALL", combo_fromstock.Text) + ";true"
        .ParameterFields(4) = "tostock;" + IIf(Trim$(combo_fromstock.Text) = "ALL", "", combo_tostock.Text) + ";true"
        '.ParameterFields(5) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(5) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00186") 'J added
        .WindowTitle = IIf(msg1 = "", "Historical stock movement", msg1) 'J modified
        Call translator.Translate_Reports("histmovtstck.rpt") 'J added
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
Dim str As String
    str = Trim$(Combo_compcode.Columns(0).Text)
    If Trim$(Combo_compcode.Columns(0).Text) = "ALL" Then
        SSOleDB_location = ""
        SSOleDB_location.RemoveAll
        Call GetalllocationName
    Else
        SSOleDB_location = ""
        SSOleDB_location.RemoveAll
        Call GetlocationName
    End If
    

End Sub

'by shakir

Private Sub Combo_compcode_DropDown()
msg1 = translator.Trans("L00028") 'J added
    Combo_compcode.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    Combo_compcode.Columns(0).Width = 1500
    Combo_compcode.Columns(1).Width = 2000
    msg1 = translator.Trans("L00050") 'J added
    Combo_compcode.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
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

Private Sub combo_fromstock_DropDown()
    With combo_fromstock
        .MoveNext
    End With
End Sub

Private Sub Combo_fromstock_GotFocus()
Call HighlightBackground(combo_fromstock)
End Sub

Private Sub Combo_fromstock_KeyDown(KeyCode As Integer, Shift As Integer)
If Not combo_fromstock.DroppedDown Then combo_fromstock.DroppedDown = True
End Sub

Private Sub Combo_fromstock_KeyPress(KeyAscii As Integer)
'combo_fromstock.MoveNext
End Sub

Private Sub Combo_fromstock_LostFocus()
Call NormalBackground(combo_fromstock)
End Sub

' set to stock number equal to from stock  number

Private Sub Combo_fromstock_Validate(Cancel As Boolean)
    combo_tostock = combo_fromstock
    If Len(Trim$(combo_fromstock)) > 0 Then
         If Not combo_fromstock.IsItemInList Then
               combo_fromstock.Text = ""
            End If
            If Len(Trim$(combo_fromstock)) = 0 Then
            combo_fromstock.SetFocus
            Cancel = True
            End If
            End If
    
End Sub

Private Sub combo_tostock_DropDown()
On Error Resume Next
    With combo_tostock
        .MoveNext
    End With
End Sub

Private Sub combo_tostock_GotFocus()
Call HighlightBackground(combo_tostock)
End Sub

Private Sub combo_tostock_KeyDown(KeyCode As Integer, Shift As Integer)
If Not combo_tostock.DroppedDown Then combo_tostock.DroppedDown = True
End Sub

Private Sub combo_tostock_KeyPress(KeyAscii As Integer)
'combo_tostock.MoveNext
End Sub

Private Sub combo_tostock_LostFocus()
Call NormalBackground(combo_tostock)
End Sub

Private Sub Combo_tostock_Validate(Cancel As Boolean)
If Len(Trim$(combo_tostock)) > 0 Then
         If Not combo_tostock.IsItemInList Then
               combo_tostock.Text = ""
            End If
            If Len(Trim$(combo_tostock)) = 0 Then
            combo_tostock.SetFocus
            Cancel = True
            End If
            End If
End Sub

'SQL statement get record set, and populate combo

Private Sub Form_Load()
'Me.Height = 2605 'J hidden
'Me.Width = 4500 'J hidden


Screen.MousePointer = 11
Me.Visible = True
Me.Refresh
DoEvents

x = 0
'Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

    'Added by Juan (9/11/2000) for Multilingual
    translator.Translate_Forms ("frm_historicalstock")
    '------------------------------------------

    With rs
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "' AND com_actvflag = 1 "
        .Source = .Source & " order by com_compcode "
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
Combo_compcode.Text = "ALL"
combo_fromstock.Text = "ALL"
combo_tostock.Text = "ALL"
SSOleDB_location.Text = "ALL"

If get_status(rs) Then
Combo_compcode.AddItem ("ALL") & vbTab & "ALL"
Do While (Not rs.EOF)
Combo_compcode.AddItem rs!com_compcode & vbTab & rs!com_name
rs.MoveNext
Loop
Set rs = Nothing
End If

    Call GetCurrencylist
    SSOleDBCurrency = "USD"
    
'Set rs = New ADODB.Recordset
'rs.Source = "Select loc_locacode,loc_name from location where loc_npecode='" & deIms.NameSpace & "'"
'rs.ActiveConnection = deIms.cnIms
'rs.Open
'If get_status(rs) Then
'SSOleDB_location.FieldSeparator = Chr$(1)
'SSOleDB_location.AddItem ("ALL" & Chr$(1) & "ALL" & "")
'SSOleDB_location.ColumnHeaders = True
'SSOleDB_location.Columns(0).Caption = "Code"
'SSOleDB_location.Columns(1).Caption = "Name"
'Do While (Not rs.EOF)
'SSOleDB_location.AddItem (rs!loc_locacode & Chr$(1) & rs!loc_name & "")
'rs.MoveNext
'Loop
'Set rs = Nothing
'
'End If


Set rs = New ADODB.Recordset
  
      With rs
        .ActiveConnection = deIms.cnIms
        .Source = "SELECT distinct qs1_stcknumb"
        .Source = .Source & " FROM QTYST1 WHERE qs1_npecode = '" & deIms.NameSpace & "'"
        .Source = .Source & " ORDER BY 1"
        .Open
    End With

  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
  'Call combo_fromstock.AddItem("ALL", 0)
  
  
  
        With combo_fromstock
            Set .DataSourceList = rs
            .DataFieldToDisplay = "qs1_stcknumb"
            .DataFieldList = "qs1_stcknumb"
            .Refresh
        End With

        With combo_tostock
            Set .DataSourceList = rs
            .DataFieldToDisplay = "qs1_stcknumb"
            .DataFieldList = "qs1_stcknumb"
            .Refresh
        End With
  
  
  
'  Do While (Not rs.EOF)
'  combo_fromstock.AddItem (rs!qs1_stcknumb)
'  rs.MoveNext
'  Loop
  Else
    Exit Sub
  End If
  
''''  'by shah
''''  Set rs = New ADODB.Recordset
''''
''''      With rs
''''        .ActiveConnection = deIms.cnIms
''''        .Source = "SELECT distinct(qs1_stcknumb)"
''''        .Source = .Source & " FROM QTYST1 WHERE "
''''        .Source = .Source & " qs1_npecode = '" & deIms.NameSpace & "'"
''''        .Source = .Source & " ORDER BY 1"
''''        .Open
''''    End With
''''
''''  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
''''   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
''''
''''
''''            Call Combo_fromstock.AddItem("ALL", 0)
''''            Do While (Not rs.EOF)
''''              Combo_fromstock.AddItem (Trim$(rs!qs1_stcknumb))
''''              rs.MoveNext
''''            Loop
''''
''''  Else
''''
''''    Exit Sub
''''
''''  End If
 

Caption = Caption + " - " + Tag

Screen.MousePointer = 0

    With frm_historicalstock
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
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

'SQL statement get location list for location combo

Private Sub GetlocationName()
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
        .CommandText = .CommandText & " and loc_compcode = '" & Combo_compcode & "' and loc_actvflag=1 "
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDB_location.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    
    SSOleDB_location.RemoveAll
    
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSOleDB_location.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
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
         .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') and loc_actvflag=1 "
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



'load stock number, and set to stock number equal to from stock number

Private Sub Combo_fromstock_click()

 'Set rs = New ADODB.Recordset

 
If combo_fromstock.Text = "ALL" Then
    x = 0
    combo_tostock.Text = ""
    combo_tostock.Enabled = False
Else

combo_tostock.Enabled = True
'If X = 1 Then Exit Sub
 x = x + 1

    
  
 'rs.MoveFirst
'Do While (Not rs.EOF)
'  combo_tostock.AddItem (rs!qs1_stcknumb)
'  rs.MoveNext
'  Loop
End If
 combo_tostock.Bookmark = combo_fromstock.Bookmark
 combo_tostock.Text = combo_fromstock.Text
 combo_tostock.Enabled = True
End Sub

'get record set status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
   End Function

'resize form

Private Sub Form_Resize()
If Not (Me.WindowState = vbMinimized) Then
'Me.Height = 3500 'J hidden
'Me.Width = 5000 'J hidden
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

