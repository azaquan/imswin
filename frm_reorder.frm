VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_reorder 
   Caption         =   "Re-order Report"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   4335
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBinventory 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   840
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Loca. Code"
      Columns(0).Name =   "Loca. Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Loca. Name"
      Columns(1).Name =   "Loca. Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCompany 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   360
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3016
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3916
      Columns(1).Caption=   "Company Name"
      Columns(1).Name =   "Company Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory Company"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label lbl_Inventorylocation 
      Caption         =   "Inventory Location"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1995
   End
End
Attribute VB_Name = "frm_reorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

'SQL statement get company list for company combo

Private Sub GetCompanyName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT com_compcode, com_name "
        .CommandText = .CommandText & " From Company "
        .CommandText = .CommandText & " WHERE com_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by com_compcode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBCompany.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
     SSOleDBCompany.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCompany.AddItem rst!com_compcode & str & (rst!com_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetCampanyName", Err.Description, Err.number, True)
End Sub



'SQL statement get location list for location combo

Private Sub GetlocationName(Company As String)
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT loc_locacode,loc_name "
        .CommandText = .CommandText & " From location "
        .CommandText = .CommandText & " WHERE "
        .CommandText = .CommandText & "loc_npecode = '" & deIms.NameSpace & "' AND "
        .CommandText = .CommandText & "(UPPER(loc_gender) <> 'OTHER') "
        If RTrim(Company) <> "ALL" Then
            .CommandText = .CommandText & " and loc_compcode = '" & Company & "'"
        End If
        .CommandText = .CommandText & " order by loc_locacode"
        
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBinventory.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
     SSOleDBinventory.AddItem (("ALL" & str) & "ALL" & "")

    Do While ((Not rst.EOF))
        SSOleDBinventory.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetlocationName", Err.Description, Err.number, True)
End Sub

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

        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub




Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\reorder.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        'MsgBox "ParameterFields 0->" + .ParameterFields(0)
        .ParameterFields(1) = "invtloca;" + IIf(Trim$(SSOleDBinventory.Text) = "ALL", "ALL", SSOleDBinventory.Columns(0).Text) + ";true"
        'MsgBox "ParameterFields 6->" + .ParameterFields(1)
        .ParameterFields(2) = "compcode;" + IIf(Trim$(SSOleDBCompany.Text) = "ALL", "ALL", SSOleDBCompany.Columns("code").Text) + ";TRUE"
        'MsgBox "ParameterFields 8->" + .ParameterFields(2)
        
          Call translator.Translate_Reports("reorder.rpt")
          Call translator.Translate_SubReports
        
        .Action = 1: .Reset
End With

    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
    Call GetCompanyName
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

Private Sub SSOleDBCompany_Click()
Dim com As String

    If Len(Trim$(SSOleDBCompany)) <> 0 Then
        SSOleDBinventory = ""
        SSOleDBinventory.RemoveAll
        com = Trim$(SSOleDBCompany.Columns(0).Text)
        Call GetlocationName(com)
    End If
End Sub

Private Sub SSOleDBCompany_InitColumnProps()
Dim com As String
            com = Trim$(SSOleDBCompany.Columns(0).Text)
            Call GetlocationName(com)
End Sub

