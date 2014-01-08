VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_slowmoving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slow Moving Inventory"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   5010
   Tag             =   "03030600"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_compcode 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      FieldSeparator  =   ";"
      Columns(0).Width=   3200
      _ExtentX        =   4260
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.TextBox txt_difference 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_location 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   2415
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4260
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
      DataFieldList   =   "Column 1"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4260
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Output Currency"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1140
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label lbl_difference 
      Caption         =   "Difference"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1515
      Width           =   2000
   End
   Begin VB.Label lbl_locacode 
      Caption         =   "Location Code"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   780
      Width           =   2000
   End
   Begin VB.Label lbl_compcode 
      Caption         =   "Company Code"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   420
      Width           =   2000
   End
End
Attribute VB_Name = "frm_slowmoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset


'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'call function to get company location

Private Sub Combo_compcode_Click()
    
        SSOleDB_location = ""
        SSOleDB_location.RemoveAll
        
    If Trim$(Combo_compcode) = "ALL" Then
        Call GetalllocationName
    Else
        Call GetlocationName
    End If
    

End Sub

Private Sub Combo_compcode_GotFocus()
Call HighlightBackground(Combo_compcode)
End Sub

Private Sub Combo_compcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_compcode.DroppedDown Then Combo_compcode.DroppedDown = True
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

'SQL statement get company code and populate combo

Private Sub Form_Load()
'Me.Height = 3195
'Me.Width = 4600
Dim rs As ADODB.Recordset
Dim str As String
Dim rs1 As ADODB.Recordset
Dim rst1 As ADODB.Recordset

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_slowmoving")
    '------------------------------------------

Set rs = New ADODB.Recordset

    With rs
    
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "'"
        .Source = .Source & " order by com_compcode "
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
 
 
 
    If get_status(rs) Then
Combo_compcode.AddItem ("ALL" & ";" & "ALL" & "")
Do While (Not rs.EOF)
Combo_compcode.AddItem (rs!com_compcode & ";" & rs!com_name & "")
rs.MoveNext
Loop
Set rs = Nothing
End If

'        Set rs1 = rs.Clone()
'         Do While (Not rs1.EOF)
'            Combo1.AddItem rs1!com_compcode
'            rs1.MoveNext
'        Loop
            
  
            
Set rs = Nothing

      Call GetCurrencylist
       SSOleDBCurrency = "USD"

'Set rs = New ADODB.Recordset
'rs.Source = "Select loc_locacode,loc_name from location where loc_gender = 'BASE' and loc_npecode='" & deIms.NameSpace & "'"
'rs.ActiveConnection = deIms.cnIms
'rs.Open
'If get_status(rs) Then
'SSOleDB_location.FieldSeparator = Chr$(1)
'SSOleDB_location.ColumnHeaders = True
'SSOleDB_location.Columns(0).Caption = "Code"
'SSOleDB_location.Columns(1).Caption = "Name"
'Do While (Not rs.EOF)
'SSOleDB_location.AddItem (rs!loc_locacode & Chr$(1) & rs!loc_name & "")
'rs.MoveNext
'Loop
'Set rs = Nothing
'End If

frm_slowmoving.Caption = frm_slowmoving.Caption + " - " + frm_slowmoving.Tag

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
        .CommandText = .CommandText & " and loc_compcode = '" & Combo_compcode & "'"
        .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
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
If Err Then Call LogErr(Name & "::GetlocationName", Err.Description, Err.number, True)
End Sub

'SQL statement get all location list for location combo

Private Sub GetalllocationName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset

    
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
    SSOleDB_location.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDB_location.RemoveAll
    
    rst.MoveFirst
      
    SSOleDB_location.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDB_location.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
     
'    Set rst1 = rst.Clone()
'    rst1.MoveFirst
'    str = Chr$(1)
'    SSOleDBCombo1.FieldSeparator = str
'    Do While ((Not rst1.EOF))
'        SSOleDBCombo1.AddItem rst1!loc_locacode & str & (rst1!loc_name & "")
'
'        rst1.MoveNext
'    Loop
     
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub

'check recordset stautes

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
   End Function

'resize form

Private Sub Form_Resize()
If Not (Me.WindowState = vbMinimized) Then
'Me.Height = 3195
'Me.Width = 4400
End If
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

If Not IsNumeric(txt_difference.Text) Then

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("M00271") 'J added
    msg2 = translator.Trans("L00191") 'J added
    MsgBox IIf(msg1 = "", "Please enter a valid Number", msg1), , IIf(msg2 = "", "Difference", msg2) 'J modified
    '---------------------------------------------
    
  Call txt_difference_Validate("true")
  txt_difference.Text = ""
  Else
  With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\slowinvt.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "company;" + Trim$(Combo_compcode.Text) + ";true"
        .ParameterFields(2) = "ware;" + Trim$(SSOleDB_location.Text) + ";true"
        .ParameterFields(3) = "Difference;" + Trim$(txt_difference.Text) + ";true"
        '.ParameterFields(4) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(4) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00189") 'J added
        .WindowTitle = IIf(msg1 = "", "Slow Moving Invt", msg1) 'J modified
        Call translator.Translate_Reports("slowinvt.rpt") 'J added
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


Private Sub SSOleDB_location_GotFocus()
Call HighlightBackground(SSOleDB_location)
End Sub

Private Sub SSOleDB_location_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_location.DroppedDown Then SSOleDB_location.DroppedDown = True
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

Private Sub txt_difference_GotFocus()
Call HighlightBackground(txt_difference)
End Sub

Private Sub txt_difference_LostFocus()
Call NormalBackground(txt_difference)
End Sub

Private Sub txt_difference_Validate(Cancel As Boolean)
Dim x As Boolean
If Len(txt_difference) > 0 Then
If CDbl(txt_difference) < 0 Then
       Cancel = True
       MsgBox " You cannot insert Negative Values "
       txt_difference.SetFocus
       txt_difference.Text = ""
       Exit Sub
    End If
    End If

End Sub
