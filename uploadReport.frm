VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form uploadReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Report"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   4665
   Tag             =   "03040500"
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   2505
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   4215
      Begin MSComCtl2.DTPicker dtdate2 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   22282243
         CurrentDate     =   36524
      End
      Begin MSComCtl2.DTPicker DTdate1 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   22282243
         CurrentDate     =   36524
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_company 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1935
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         Cols            =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDB_ware 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   1935
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         Cols            =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label lbl_todate 
         Caption         =   "To Date"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lbl_fromdate 
         Caption         =   "From Date"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl_ware 
         Caption         =   "Location"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lbl_company 
         Caption         =   "Company"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboCurrency 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      DataFieldList   =   "Column 0"
      ListAutoPosition=   0   'False
      AllowInput      =   0   'False
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "uploadReport.frx":0000
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "uploadReport.frx":001C
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4180
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2593
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Currency"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "uploadReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'unload form

Private Sub cmd_cancel_Click()
Unload Me
End Sub


Private Sub cmd_ok_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If (DTdate1.value > dtdate2.value) Then
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00003") 'J added
        msg1 = translator.Trans("L00318") 'J added
        MsgBox IIf(msg1 = "", "Make sure the 'From date' is less than the 'To date'", msg1), , IIf(msg2 = "", "Date", msg2) 'J modified
        '---------------------------------------------
        DTdate1_Validate ("true")
    Else
        Dim CompanyCode
        Dim CurrencyCode
        Dim namespaceCode
        Dim LocationCode
        Dim invocation
        Dim Fromdate
        Dim Todate
        If Trim$(SSOleDB_company.Text) = "" Or Trim$(SSOleDB_company.Text) = "ALL" Then
            CompanyCode = "%"
        Else
            CompanyCode = Trim$(SSOleDB_company.Text)
        End If
        CompanyCode = "-company " + Chr(34) + CompanyCode + Chr(34) + " "
        namespaceCode = "-namespace " + Chr(34) + deIms.NameSpace + Chr(34) + " "
        If Trim$(UCase(SSOleDB_ware.Text)) = "" Or Trim$(UCase(SSOleDB_ware.Text)) = "ALL" Then
            LocationCode = "%"
        Else
            LocationCode = Trim$(UCase(SSOleDB_ware.Text))
        End If
        LocationCode = "-location " + Chr(34) + LocationCode + Chr(34) + " "
        Fromdate = "-fromDate " + Chr(34) _
            + Format(Year(DTdate1.value)) + "-" _
            + Format(Month(DTdate1.value), "00") + "-" _
            + Format(Day(DTdate1.value), "00") + Chr(34) + " "
        Todate = "-toDate " + Chr(34) _
            + Format(Year(dtdate2.value)) + "-" _
            + Format(Month(dtdate2.value), "00") + "-" _
            + Format(Day(dtdate2.value), "00") + Chr(34) + " "
        invocation = "cd \imsReportGenerator & java -jar reportGenerator.jar -name uploadReport -xuser " + Chr(34) + CurrentUser + Chr(34) + " "
        Shell "cmd.exe /c " & invocation + CompanyCode + namespaceCode + LocationCode + Fromdate + Todate, vbHide
        MsgBox "An email has been sent to you with the report."
     End If
 
'    If Err Then
'        MsgBox Err.Description
''        Call LogErr(Name & "::cmd_ok_Click", Err.Description, Err)
'   End If
   
Screen.MousePointer = vbArrow
   
End Sub

Private Sub CmdFqa_Click()
Load Frm_FQAReporting
Frm_FQAReporting.Show
End Sub

Private Sub DTdate1_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub


Private Sub EOM_GotFocus()
'Call HighlightBackground(EOM)
SSOleDB_company.Enabled = True
SSOleDB_ware.Enabled = True
End Sub

Private Sub EOM_LostFocus()
'Call NormalBackground(EOM)
End Sub

'SQL statement get company

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

'Added by Juan (9/14/2000) for Multilingual
Call translator.Translate_Forms("frm_tranvaluationreport")
'------------------------------------------

'Me.Height = 3400
'Me.Width = 5000
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
SSOleDB_company.FieldSeparator = Chr$(1)
SSOleDB_ware.FieldSeparator = Chr$(1)

    With rs
        .Source = "select com_compcode,com_name from company where com_npecode='" & deIms.NameSpace & "' AND com_actvflag = 1 "
        .Source = .Source & " order by com_compcode "
        .ActiveConnection = deIms.cnIms
        .Open
    End With
    
If get_status(rs) Then
SSOleDB_company.ColumnHeaders = True

'Modified by Juan (9/14/2000) for Multilingual
msg1 = translator.Trans("L00028") 'J added
msg2 = translator.Trans("L00050") 'J added
SSOleDB_company.Columns(0).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
SSOleDB_company.Columns(1).Caption = IIf(msg2 = "", "Name", msg2) 'J modified
'---------------------------------------------

rs.MoveFirst
SSOleDB_company.Text = "ALL"
SSOleDB_ware.Text = "ALL"
SSOleDB_company.AddItem (("ALL" & Chr$(1)) & "ALL" & "")
Do While (Not rs.EOF)
SSOleDB_company.AddItem (rs!com_compcode & Chr$(1) & rs!com_name & " ")
rs.MoveNext
Loop
Set rs = Nothing
End If


'rs1.Source = "Select loc_locacode,loc_name from location where loc_npecode='" & deIms.NameSpace & "'"
'rs1.ActiveConnection = deIms.cnIms
'rs1.Open
'If get_status(rs1) Then
'SSOleDB_ware.ColumnHeaders = True
'SSOleDB_ware.Columns(0).Caption = "Code"
'SSOleDB_ware.Columns(1).Caption = "Name"
'Do While (Not rs1.EOF)
'SSOleDB_ware.AddItem (rs1!loc_locacode & Chr$(1) & rs1!loc_name & " ")
'rs1.MoveNext
'Loop
'Set rs1 = Nothing
'End If
    Call GetCurrencylist
     SSOleDBCurrency = "USD"
frm_tranvaluationreport.Caption = frm_tranvaluationreport.Caption + " - " + frm_tranvaluationreport.Tag
DTdate1.value = FirstOfMonth
dtdate2.value = Now
'Full.value = True

Me.Left = Round((Screen.Width - Me.Width) / 2)
Me.Top = Round((Screen.Height - Me.Height) / 2)
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
        .CommandText = .CommandText & " WHERE loc_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and loc_compcode = '" & Company & "'"
        .CommandText = .CommandText & " and (UPPER(loc_gender)  <> 'OTHER') and loc_actvflag=1 "
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDB_ware.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    
    SSOleDB_ware.RemoveAll
    
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSOleDB_ware.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
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
         .CommandText = .CommandText & " and (UPPER(loc_gender) ='BASE') and loc_actvflag=1 "
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDB_ware.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDB_ware.RemoveAll
    
    rst.MoveFirst
      
    SSOleDB_ware.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDB_ware.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
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
If Not (Me.WindowState = vbMinimized) Then
'Me.Height = 3400
'Me.Width = 5000
End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'SQL statement get currency list for currency combo

Private Sub GetCurrencylist()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim flagCURR

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT curr_code, curr_desc "
        .CommandText = .CommandText & " From CURRENCY "
        .CommandText = .CommandText & " WHERE curr_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by curr_code"
         Set rst = .Execute
    End With

    str = Chr$(1)
    SSdcboCurrency.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSdcboCurrency.AddItem rst!curr_desc & str & (rst!curr_code & "")
        If rst!curr_code = "USD" Then flagCURR = SSdcboCurrency.Rows - 1
        rst.MoveNext
    Loop
    SSdcboCurrency.Bookmark = flagCURR
    SSdcboCurrency.Text = SSdcboCurrency.Columns(0).Text
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetCurrencylist", Err.Description, Err.number, True)
End Sub



Private Sub Full_GotFocus()
'Call HighlightBackground(Full)
SSOleDB_company.Enabled = True
SSOleDB_ware.Enabled = True

End Sub

Private Sub Full_LostFocus()
'Call NormalBackground(Full)
End Sub

Private Sub OptExcel_Click()
'Call HighlightBackground(OptExcel)
SSOleDB_company.Enabled = False
SSOleDB_ware.Enabled = False
End Sub

Private Sub OptExcel_GotFocus()
'Call HighlightBackground(OptExcel)
SSOleDB_company.Enabled = False
SSOleDB_ware.Enabled = False
End Sub

Private Sub OptExcel_LostFocus()
'Call NormalBackground(OptExcel)
End Sub

Public Sub SSdcboCurrency_Click()
End Sub

'call function get location

Private Sub SSOleDB_company_Click()
Dim str As String

    str = Trim$(SSOleDB_company.Columns(0).Text)
    
    SSOleDB_ware = ""
    SSOleDB_ware.RemoveAll
    
    If Trim$(SSOleDB_company.Columns(0).Text) = "ALL" Then
        Call GetalllocationName
    Else
        Call GetlocationName(str)
    End If

SSOleDB_ware.Text = "ALL"
End Sub


Private Sub SSOleDB_company_GotFocus()
Call HighlightBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_company.DroppedDown Then SSOleDB_company.DroppedDown = True
End Sub

Private Sub SSOleDB_company_LostFocus()
Call NormalBackground(SSOleDB_company)
End Sub

Private Sub SSOleDB_company_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_company)) > 0 Then
         If Not SSOleDB_company.IsItemInList Then
                SSOleDB_company.Text = ""
            End If
            If Len(Trim$(SSOleDB_company)) = 0 Then
            SSOleDB_company.SetFocus
            Cancel = True
            End If
            End If
            
End Sub

Private Sub SSOleDB_ware_GotFocus()
Call HighlightBackground(SSOleDB_ware)
End Sub

Private Sub SSOleDB_ware_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDB_ware.DroppedDown Then SSOleDB_ware.DroppedDown = True
End Sub

Private Sub SSOleDB_ware_LostFocus()
Call NormalBackground(SSOleDB_ware)
End Sub

Private Sub SSOleDB_ware_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDB_ware)) > 0 Then
         If Not SSOleDB_ware.IsItemInList Then
                SSOleDB_ware.Text = ""
            End If
            If Len(Trim$(SSOleDB_ware)) = 0 Then
           SSOleDB_ware.SetFocus
            Cancel = True
            End If
            End If
End Sub

Public Function EomInExcel()
 Dim rs As ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Arr As Variant
    Dim ArrColumns() As String
    Dim Fld As ADODB.Field
    
    Call MDI_IMS.WriteStatus("Exporting EOM data to Excel...", 1)
    
    Call IncrementProgreesBar(1, ProgressBar1)
    cmd.ActiveConnection = deIms.cnIms
    cmd.CommandText = "INVENTORYFORACCOUTING"
    cmd.CommandType = adCmdStoredProc
    cmd.parameters.Append cmd.CreateParameter("@namespace", adVarChar, adParamInput, 15, deIms.NameSpace)
    cmd.parameters.Append cmd.CreateParameter("@fromdate", adDate, adParamInput, , DTdate1.value)
    cmd.parameters.Append cmd.CreateParameter("@todate", adDate, adParamInput, , dtdate2.value)
    Set rs = cmd.Execute
    If rs.RecordCount > 0 Then
            Arr = rs.GetRows
            
            For Each Fld In rs.Fields
            
                    ReDim Preserve ArrColumns(i)
                    ArrColumns(i) = Fld.Name
                    i = i + 1
                    
            Next Fld
                    
            
            Call ExportToExcel(, Arr, ArrColumns, ProgressBar1, "EOMForAccouting")
     End If
    Arr = ""
    
    ReDim ArrColumns(0)
    i = 0
    
    Set rs = Nothing
    
    Call MDI_IMS.WriteStatus("Exporting Purchase data to Excel...", 1)
    Call IncrementProgreesBar(1, ProgressBar1)
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = deIms.cnIms
    cmd.CommandText = "PURCHASINGFORACCOUNTING"
    cmd.CommandType = adCmdStoredProc
    cmd.parameters.Append cmd.CreateParameter("@namespace", adVarChar, adParamInput, 15, deIms.NameSpace)
    cmd.parameters.Append cmd.CreateParameter("@fromdate", adDate, adParamInput, , DTdate1.value)
    cmd.parameters.Append cmd.CreateParameter("@todate", adDate, adParamInput, , dtdate2.value)
    Set rs = cmd.Execute
    If rs.RecordCount > 0 Then
            Arr = rs.GetRows
            
            For Each Fld In rs.Fields
            
                    ReDim Preserve ArrColumns(i)
                    ArrColumns(i) = Fld.Name
                    i = i + 1
                    
            Next Fld
            
            Call ExportToExcel(, Arr, ArrColumns, ProgressBar1, "Purchasing")
     End If
    
    Call MDI_IMS.WriteStatus("", 1)
    
    ProgressBar1.Visible = False
    
End Function
