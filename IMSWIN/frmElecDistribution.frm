VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frmElecDistribution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electronic Distribution"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7785
   Tag             =   "01040500"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      NewEnabled      =   -1  'True
      SaveEnabled     =   0   'False
      AllowDelete     =   0   'False
      DeleteEnabled   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBDDDisCode 
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5318
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2170
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3413
      _ExtentY        =   1720
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGridList 
      Height          =   3255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5910
      _Version        =   196617
      DataMode        =   2
      BorderStyle     =   0
      Col.Count       =   4
      stylesets.count =   2
      stylesets(0).Name=   "colls"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmElecDistribution.frx":0000
      stylesets(1).Name=   "rows"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmElecDistribution.frx":001C
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1535
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HeadStyleSet=   "colls"
      Columns(1).Width=   2249
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).HeadStyleSet=   "colls"
      Columns(2).Width=   3200
      Columns(2).Caption=   "Mail"
      Columns(2).Name =   "Mail"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).HeadStyleSet=   "colls"
      Columns(3).Width=   2646
      Columns(3).Caption=   "Fax"
      Columns(3).Name =   "Fax"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).HeadStyleSet=   "colls"
      _ExtentX        =   10425
      _ExtentY        =   5741
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Visualization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   4920
      TabIndex        =   4
      Top             =   4080
      Width           =   2460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "System Distribution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmElecDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GoodColMove As Boolean
Dim InUnload As Boolean
Dim Modify As String
Dim Create As String
Dim RecSaved As Boolean
Dim Visualize As String
Dim NVBAR_EDIT As Boolean
Dim NVBAR_ADD As Boolean
Dim NVBAR_SAVE As Boolean
Dim CAncelGrid As Boolean
Dim InSave As Boolean

Dim Rstlist As ADODB.Recordset

'SQL statement get document information and populate data combo

Public Sub GetDocumentCode()
On Error Resume Next

Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
       
        .CommandText = " SELECT doc_code, doc_desc "
        .CommandText = .CommandText & " From DOCTYPE "
        .CommandText = .CommandText & " WHERE doc_npecode = '" & deIms.NameSpace & "'"
        
        Set rst = .Execute
    End With

    
    
    str = Chr$(1)
    SSOleDBDDDisCode.FieldSeparator = str

    SSOleDBDDDisCode.RemoveAll
    If rst.BOF And rst.EOF Then Exit Sub
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then GoTo CleanUp

    rst.MoveFirst

    Do While ((Not rst.EOF))
        SSOleDBDDDisCode.AddItem rst!doc_desc & str & (rst!doc_code & "")
         rst.MoveNext
    Loop
    
  
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    
    
End Sub

'SQL statement get transaction type and populate data grid

Public Sub GetTranstypeCode()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        
        .CommandText = " SELECT tty_code, tty_desc "
        .CommandText = .CommandText & " From TRANSACTYPE "
        .CommandText = .CommandText & " WHERE tty_npecode = '" & deIms.NameSpace & "'"

        Set rst = .Execute
    End With

    str = Chr$(1)
   SSOleDBDDDisCode.FieldSeparator = str
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        SSOleDBDDDisCode.AddItem rst!tty_desc & str & (rst!tty_code & "")
         rst.MoveNext
    Loop
    
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
        

End Sub
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT from "
         .CommandText = .CommandText & " UNIT "
        .CommandText = .CommandText & " Where uni_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND uni_code = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function
Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

    msg1 = translator.Trans("M00351") 'J added
    '------------------------------------------
    validate_fields = True
    If Not Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
          '  Call txtmailValidate(True)
            x = txtmailValidate(True)
            If Not x Then
                RecSaved = False
                validate_fields = False
                Exit Function
            End If
       ElseIf Not Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
       
            'Modified by Juan (9/15/2000) for Multilingual
            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------
          RecSaved = False
         validate_fields = False
           
          'txtMail.SetFocus:
          Exit Function
       End If
    Else
        If Not Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
            'Call txtfaxnumber_validate(True)
             x = txtfaxnumber_validate(True)
             If Not x Then
                RecSaved = False
                validate_fields = False
                Exit Function
            End If
         ElseIf Not Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
         
            'Modified by Juan (9/15/2000) for Multilingual
            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------
         RecSaved = False
         validate_fields = False

             'txtfaxNumb.SetFocus:
             Exit Function
        End If
    End If

 If Len(Trim$(SSDBGridList.Columns("MAIL").text)) = 0 And Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
 
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00354") 'J added
    MsgBox IIf(msg1 = "", "You cannot leave Email and Fax empty", msg1) 'J modified
         RecSaved = False
         validate_fields = False
    '---------------------------------------------
 End If
    



'*** validate_fields = True
'***If SSDBGridList.IsAddRow Then
'***   If colnum = 0 Or colnum = 1 Then
'***      x = NotValidLen(SSDBGridList.Columns(colnum).text)
'***      If x = True Then
'***         RecSaved = False
'***         msg1 = translator.Trans("M00702")
'***         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
'***         SSDBGridList.SetFocus
'***         SSDBGridList.Col = colnum
'***         validate_fields = False
'***         Exit Function
'***      End If
'***    End If
'***      If colnum = 0 Then
'***        x = CheckDesCode(SSDBGridList.Columns(0).text)
'***        If x <> False Then
'***             RecSaved = False
'***             msg1 = translator.Trans("M00703")
'***             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
'***             SSDBGridList.SetFocus
'***             SSDBGridList.Col = 0
'***            validate_fields = False
'***         End If
    '***End If
'***   End If

End Function

'call functions get data and populate data grid,set button

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  CAncelGrid = False
   msg1 = translator.Trans("M00126")
   Modify = IIf(msg1 = "", "Modification", msg1)
   msg1 = translator.Trans("M00092")
   Visualize = IIf(msg1 = "", "Visualization", msg1)
   msg1 = translator.Trans("M00125")
   Create = IIf(msg1 = "", "Creation", msg1)
   GoodColMove = True
   RecSaved = True
   InUnload = False
   InSave = False

    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmElecDistribution")
    '------------------------------------------
   
    Call GetDocumentCode
    'Call GetTranstypeCode
    Call GetDistributionCode

    SSDBGridList.DataMode = ssDataModeAddItem
    Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
    Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
    
    NavBar1.NewEnabled = NavBar1.SaveEnabled
    NavBar1.DeleteEnabled = NavBar1.NewEnabled
    
    frmElecDistribution.Caption = frmElecDistribution.Caption + " - " + frmElecDistribution.Tag
     NVBAR_EDIT = NavBar1.EditEnabled
 '  NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_ADD = True
    NVBAR_SAVE = False
    'NVBAR_SAVE = NavBar1.SaveEnabled
    
    NavBar1.NewEnabled = True
    NavBar1.NewVisible = True
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    Call DisableButtons(Me, NavBar1)
     NavBar1.DeleteEnabled = False
       NavBar1.DeleteVisible = False
    SSDBGridList.AllowUpdate = False
End Sub

'assign values to data grid

Public Sub GetDistributionCode()
Dim str As String

    str = Chr(1)
    SSOleDBDDDisCode.FieldSeparator = str

    SSOleDBDDDisCode.AddItem "Update Database" & str & "UD"
    'SSOleDBDDDisCode.AddItem "Delivery" & str & "DL"
    SSOleDBDDDisCode.AddItem "Security" & str & "LO"
    'SSOleDBDDDisCode.AddItem "Shipping" & str & "SH"
    SSOleDBDDDisCode.AddItem "Warehouse Trans." & str & "WH"

End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 InUnload = True
 RecSaved = True
 CAncelGrid = False
 SSDBGridList.Update
 If RecSaved = True Then
    Hide
    If open_forms <= 5 Then ShowNavigator
   If Err Then Err.Clear
    
Else
    Cancel = True
End If
End Sub

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True

End Sub

Private Sub NavBar1_BeforeDeleteClick()
   Dim response As Integer
   
  '   MsgBox ("here")
      msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to delete current record?", msg1)), vbOKCancel, "Imswin")
     If (response = vbOK) Then
            If Len(Trim$(SSDBGridList.Columns("mail").text)) Then
                Call DeleteUserMail(SSDBGridList.Columns("Mail").text)
                Call Clearform
        
            ElseIf Len(Trim$(SSDBGridList.Columns("fax").text)) Then
                Call DeleteUserFax(SSDBGridList.Columns("fax").text)
                Call Clearform
        '        SSDBGridList.MoveLast
            End If
        SSDBGridList.RemoveAll
        Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
        Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
        Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
    '    Call DisableButtons(Me, NavBar1)
        SSDBGridList.AllowUpdate = False
   
        SSDBGridList.SetFocus
        SSDBGridList.Col = 0
    End If
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
     'SSDBGridList.Update

End Sub

Private Sub NavBar1_BeforeNewClick()
  ' SSDBGridList.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.DeleteEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGridList.AllowUpdate = True
    SSDBGridList.Columns(1).locked = True
 '   SSDBGridList.Columns("active").text = 1
  '  SSDBGridList.SetFocus
  '  SSDBGridList.Col = 0

End Sub

Private Sub NavBar1_BeforeSaveClick()
        CAncelGrid = False
        InSave = True
        RecSaved = False
        SSDBGridList.Update
        If RecSaved = True Then
        SSDBGridList.Columns(0).locked = False
        SSDBGridList.Columns(1).locked = False
        SSOleDBDDDisCode.Enabled = True
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.DeleteEnabled = True
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGridList.AllowUpdate = False
       End If
    InSave = False
End Sub

'call function and load data grid

Private Sub NavBar1_OnCancelClick()
Dim response As Integer
   If SSDBGridList.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGridList.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGridList.CancelUpdate
       '   SSDBGridList.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.DeleteEnabled = True
          NavBar1.SaveEnabled = False
          SSDBGridList.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGridList.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGridList.CancelUpdate
    SSDBGridList.Columns(0).locked = False
    SSDBGridList.Columns(1).locked = False
    SSOleDBDDDisCode.Enabled = True
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBGridList.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.DeleteEnabled = True
    NavBar1.SaveEnabled = False
    SSDBGridList.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBGridList.Refresh
End If
 'Dim x As Integer
      'SSDBGridList.MoveLast
       'x = SSDBGridList.Rows
       'If x = SSDBGridList.Row Then Debug.Print "Can Cancel"
     
     'SSDBGridList.AddNew ""
'***    SSDBGridList.RemoveAll
'***    Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
'***    Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
'***    Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
'***    SSDBGridList.MoveLast
    
''    If SSDBGridList.Row = SSDBGridList.Rows Then
''         Call Clearform
''
''         Call NavBar1_OnPreviousClick
''    Else
''         Call NavBar1_OnPreviousClick
''    End If

End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'before detele check data grid

Private Sub NavBar1_OnDeleteClick()
   ' If Len(Trim$(SSDBGridList.Columns("mail").text)) Then
   '     Call DeleteUserMail(SSDBGridList.Columns("Mail").text)
   '     Call Clearform
        
   ' ElseIf Len(Trim$(SSDBGridList.Columns("fax").text)) Then
   '     Call DeleteUserFax(SSDBGridList.Columns("fax").text)
   '     Call Clearform
'   '     SSDBGridList.MoveLast
   ' End If
    '    SSDBGridList.MoveLast

End Sub

Private Sub NavBar1_OnEditClick()
SSDBGridList.AllowUpdate = True
SSDBGridList.Columns(0).locked = True
SSDBGridList.Columns(1).locked = True
NavBar1.CancelEnabled = True
NavBar1.DeleteEnabled = False
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGridList.SetFocus
SSDBGridList.Col = 2
SSOleDBDDDisCode.Enabled = False
SSDBGridList.AllowUpdate = True

End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
On Error Resume Next
    SSDBGridList.MoveFirst
    Call EnableButtons
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
On Error Resume Next
    SSDBGridList.MoveLast
    Call EnableButtons
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
On Error Resume Next
  With SSDBGridList
        If Not SSDBGridList.EOF Then
            SSDBGridList.MoveNext
            Call EnableButtons
        Else
            With SSDBGridList
                If .EOF Then Exit Sub
            End With
        End If
   End With

   
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next

   
        If Not SSDBGridList.BOF Then
            SSDBGridList.MovePrevious
            Call EnableButtons
        Else
            With SSDBGridList
                If .BOF Then Exit Sub
            End With
        End If
   
    
End Sub

'call function reload recordset and move recordset to last position

Private Sub NavBar1_OnNewClick()
   
       
    SSDBGridList.RemoveAll
    Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
    Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
    
    SSDBGridList.AddNew
    SSDBGridList.SetFocus
    SSDBGridList.Col = 0
    
 '   SSDBGridList.AllowUpdate = False
    
    
'     Call NavBar1_OnPreviousClick
'     SSDBGridList.MoveNext
     'Call Clearform

End Sub

'clear data grid

Public Sub Clearform()
    SSDBGridList.Columns(0).text = ""
    SSDBGridList.Columns(1).text = ""
    SSDBGridList.Columns("mail").text = ""
    SSDBGridList.Columns("fax").text = ""
'    txt"mail" = ""
'    txtfaxNumb = ""
    
End Sub

'validate data values

Public Function DataValidate()
    If Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
          'SSDBGridList.Columns("mail").Text.SetFocus:
          Exit Function
       ElseIf Len(Trim$(SSDBGridList.Columns("fax").text)) > 0 Then
       
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00351") 'J added
            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------
            
       End If
    Else
        'txtfaxNumb.SetFocus:
        Exit Function
    End If
    
End Function

'call store procedure insert a record

Private Sub InsertElecDistribution()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "INSERT_DISTRIBUTION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "SYS")
        .Parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 5, SSOleDBDDDisCode.Columns("Code").text)
        .Parameters.Append .CreateParameter("@MAIL", adVarChar, adParamInput, 59, SSDBGridList.Columns("Mail").text)
        .Parameters.Append .CreateParameter("@FAXNUMB", adVarChar, adParamInput, 50, SSDBGridList.Columns("fax").text)
        
        .Execute , , adExecuteNoRecords

    End With

    Set cmd = Nothing
    
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00352") 'J added
    MsgBox IIf(msg1 = "", "Insert into Distribution is completed successfully ", msg1) 'J modified
    '---------------------------------------------
    
  '  SSDBGridList.MovePrevious
    Exit Sub

Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00353") 'J added
        MsgBox IIf(msg1 = "", "Insert into Distribution failed", msg1) 'J modified
        '---------------------------------------------

End Sub

Private Sub UpdateElecDistribution()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "UPDATE_DISTRIBUTION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "SYS")
        .Parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 5, SSOleDBDDDisCode.Columns("Code").text)
        .Parameters.Append .CreateParameter("@MAIL", adVarChar, adParamInput, 59, SSDBGridList.Columns("Mail").text)
        .Parameters.Append .CreateParameter("@FAXNUMB", adVarChar, adParamInput, 50, SSDBGridList.Columns("fax").text)
        
        .Execute , , adExecuteNoRecords

    End With

    Set cmd = Nothing
    
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00352") 'J added
    MsgBox IIf(msg1 = "", " Distribution update successful ", msg1) 'J modified
    '---------------------------------------------
    
   ' SSDBGridList.MovePrevious
    Exit Sub

Noinsert:
        If Err Then Err.Clear
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00353") 'J added
        MsgBox IIf(msg1 = "", "update into Distribution failed", msg1) 'J modified
        '---------------------------------------------

End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\elecdistribution.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "type;SYS;TRUE"
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00445") 'J added
        .WindowTitle = IIf(msg1 = "", "Electronic Distribution", msg1) 'J modified
        Call translator.Translate_Reports("elecdistribution.rpt") 'J added
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

'before save validate data format

Private Sub NavBar1_OnSaveClick()

    'Added by Juan (9/15/2000) for Multilingual
'***    msg1 = translator.Trans("M00351") 'J added
    '------------------------------------------

    '***If Not Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
'***       If Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
'***            Call txtmailValidate(True)

'***       ElseIf Not Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
       
            'Modified by Juan (9/15/2000) for Multilingual
'***            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------
            
          'txtMail.SetFocus:
'***          Exit Sub
'***       End If
'***    Else
'***        If Not Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
'***            Call txtfaxnumber_validate(True)

'***         ElseIf Not Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
         
            'Modified by Juan (9/15/2000) for Multilingual
'***            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------

             'txtfaxNumb.SetFocus:
'***             Exit Sub
'***        End If
'***    End If

'*** If Len(Trim$(SSDBGridList.Columns("MAIL").text)) = 0 And Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
 
    'Modified by Juan (9/15/2000) for Multilingual
'***    msg1 = translator.Trans("M00354") 'J added
'***    MsgBox IIf(msg1 = "", "You cannot leave Email and Fax empty", msg1) 'J modified
    '---------------------------------------------
 '***End If
    


End Sub

'SQL statement get datarecordset and populate data grid

Public Function GetListofdistribution(NameSpace As String, Gender As String, cn As ADODB.Connection) As Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " WHERE dis_npecode = '" & NameSpace & "'"
        '.CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("MAIL").Text & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " ORDER BY dis_id "
        Set Rstlist = .Execute

    End With

    If Rstlist Is Nothing Then Exit Function
        

        
        Set cmd = Nothing
        'Set Rstlist = Nothing
        
End Function

'SQL statement get data recordset and populate data grid

Public Function GetDistributionMail(EmailNunber As String) As Boolean
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").text & "'"
        .CommandText = .CommandText & " AND dis_gender =  'SYS' "
        .CommandText = .CommandText & " AND dis_mail = '" & SSDBGridList.Columns("MAIL").text & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionMail = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
End Function

'SQL statement get data recordset and populate data grid

Public Function GetDistributionFaxnumb(Faxnumber As String) As Boolean

Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").text & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & SSDBGridList.Columns("FAX").text & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionFaxnumb = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing

End Function

'check distribution email exist or not

Public Function txtmailValidate(Cancel As Boolean) As Boolean
        
    Cancel = False
    txtmailValidate = True
    If Len(SSDBGridList.Columns("mail").text) Then
        
        If GetDistributionMail(SSDBGridList.Columns("Mail").text) Then
'            SSDBGridList.Columns("Mail").Text = ""

            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00355") 'J added
            MsgBox IIf(msg1 = "", "This configuration already exist.", msg1) 'J modified
            '---------------------------------------------
            txtmailValidate = False
            'txtMail.SetFocus:
            Exit Function
        Else
            Cancel = False
        '    Call txtfaxnumber_validate(True)
         '   Call InsertElecDistribution
            'txtfaxNumb.SetFocus:
            Exit Function
        End If
    
    End If
End Function

'check distribution fax number exist or not

Public Function txtfaxnumber_validate(Cancel As Boolean) As Boolean
    Cancel = False
    txtfaxnumber_validate = True
    
       If Len(SSDBGridList.Columns("fax").text) Then
        
        If GetDistributionFaxnumb(SSDBGridList.Columns("fax").text) Then
            'txtfaxNumb.Text = ""
            
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00355") 'J added
            MsgBox IIf(msg1 = "", "This configuration already exist.", msg1) 'J modified
            '---------------------------------------------
            txtfaxnumber_validate = False
            'txtfaxNumb.SetFocus: Exit Function
            'SSDBGridList.Columns(fax).Text Exit Function
        Else
            Cancel = False
            txtfaxnumber_validate = True
        '    Call InsertElecDistribution
            'txtMail.SetFocus:
            Exit Function
        End If
    
    End If
    

End Function

'delete exist a email number

Public Function DeleteUserMail(Mail As String) As Boolean
'On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    DeleteUserMail = True
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).text & "'"
        .CommandText = .CommandText & " AND dis_mail ='" & Mail & "'"

        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Call .Execute(0, 0, adExecuteNoRecords)
        
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            DeleteUserMail = False
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Delete from Distribution is failure ", msg1) 'J modified
            '---------------------------------------------
        
        Else
           ' msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Record Successfully Deleted from Distribution", msg1) 'J modified
            SSDBGridList.Columns("Mail").text = ""
            'txtMail.SetFocus
        End If
End Function

'delete exist a fax number

Public Function DeleteUserFax(fax As String) As Boolean
'On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    DeleteUserFax = True
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).text & "'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & fax & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        

        
        Call .Execute(0, 0, adExecuteNoRecords)
       
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Delete from Distribution is failure ", msg1) 'J modified
            '---------------------------------------------
            DeleteUserFax = False
        Else
          '  msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Record Successfully Deleted from Distribution", msg1) 'J modified
            SSDBGridList.Columns("Mail").text = ""
            SSDBGridList.Columns("FAX").text = ""
            
        End If
End Function

'assign data to data grid

Private Sub LoadValues()
On Error Resume Next

    SSDBGridList.Columns("Mail").text = Rstlist!dis_mail & ""
    SSDBGridList.Columns("FAx").text = Rstlist!dis_mail & ""

    If Err Then Err.Clear
End Sub

'add recordset to data grid

Public Sub Addtogrid(Rstlist As ADODB.Recordset)
Dim str As String
    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
    
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!doc_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop
End Sub

'add recordset to data grid

Public Sub Addtogridtran(Rstlist As ADODB.Recordset)
Dim str As String
    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
    
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!tty_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop
End Sub

'assign data to data grid

Public Sub AddtogridCode(Rstlist As ADODB.Recordset)
Dim str As String
Dim desc As String

    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
        If Trim$(Rstlist!dis_id) = "UD" Then desc = "Update Database"
      '  If Trim$(Rstlist!dis_id) = "DL" Then desc = "Delivery"
        If Trim$(Rstlist!dis_id) = "LO" Then desc = "Security"
      '  If Trim$(Rstlist!dis_id) = "SH" Then desc = "Shipping"
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop

End Sub

'SQL statement get data distribution list

Public Function Getlistforgrid(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, doc_desc, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION, doctype "
        .CommandText = .CommandText & " WHERE dis_npecode = doc_npecode "
        .CommandText = .CommandText & " and dis_id = doc_code and "
        .CommandText = .CommandText & " dis_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgrid = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

'SQL statement get data distribution list

Public Function Getlistforgridtran(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, tty_desc, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION,  TRANSACTYPE "
        .CommandText = .CommandText & " WHERE dis_npecode =  tty_npecode "
        .CommandText = .CommandText & " and dis_id = tty_code and "
        .CommandText = .CommandText & " dis_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgridtran = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

'SQL statement get data distribution email list

Public Function GetlistforgridCode(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " where ((dis_id IN ( 'ud', 'sh', 'lo', 'dl')) "
        .CommandText = .CommandText & " AND (dis_gender = 'SYS') "
        .CommandText = .CommandText & " AND (dis_npecode = '" & NameSpace & "'))"
        .CommandText = .CommandText & " order by dis_id "
        

        Set GetlistforgridCode = .Execute

    End With


       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

Private Sub SSDBGridList_AfterUpdate(RtnDispErrMsg As Integer)
'If RecSaved = True Then
'    lblStatus.ForeColor = &HFF00&
'    lblStatus = Visualize
'    NavBar1.SaveEnabled = False
'    NavBar1.CancelEnabled = False
'    NavBar1.EditEnabled = True
'    NavBar1.NewEnabled = NVBAR_ADD
'    SSDBGridList.AllowUpdate = False
'End If

End Sub

Private Sub SSDBGridList_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
'Dim Recchanged As Boolean
'Dim ret  As Integer
'Dim good_field As Boolean
  
 '***         If SSDBGridList.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
'***             If NotValidLen(SSDBGridList.Columns(ColIndex).text) Then
'***                msg1 = translator.Trans("M00702")
'***                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
'***                Cancel = 1
'***                SSDBGridList.SetFocus
'***                SSDBGridList.Columns(ColIndex).text = OldValue
'***                SSDBGridList.Col = 0
'***                RecSaved = False
'***                GoodColMove = False
'***              ElseIf CheckDesCode(SSDBGridList.Columns(ColIndex).text) Then
'***                msg1 = translator.Trans("M00703")
'***                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
'***                Cancel = 1
'***                SSDBGridList.SetFocus
'***                SSDBGridList.Columns(ColIndex).text = OldValue
'***                SSDBGridList.Col = ColIndex
'***                RecSaved = False
'***                GoodColMove = False
'***             End If
        
'***        ElseIf SSDBGridList.IsAddRow And ColIndex = 1 Then
'***              If NotValidLen(SSDBGridList.Columns(ColIndex).text) Then
'***                msg1 = translator.Trans("M00702")
'***                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
'***                Cancel = 1
'***                SSDBGridList.SetFocus
'***                RecSaved = False
'***                SSDBGridList.Col = 1
'***               End If
'***        ElseIf Not SSDBGridList.IsAddRow And ColIndex = 1 Then
'***                If NotValidLen(SSDBGridList.Columns(ColIndex).text) Then
'***               msg1 = translator.Trans("M00702")
'***                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
'***                Cancel = 1
'***                SSDBGridList.SetFocus
'***                'SSDBGridList.Columns(ColIndex).Text =
'***                RecSaved = False
'***                SSDBGridList.Col = ColIndex
'***               End If
'***       End If
 ''    Recchanged = DidFieldChange(Trim(OldValue), Trim(SSDBGridList.Columns(ColIndex).text))
  ''   good_field = validate_fields(ColIndex)
  ''   If Not good_field Then
  ''      RecSaved = False
  ''      Cancel = True
  ''  End If

End Sub

Private Sub SSDBGridList_BeforeRowColChange(Cancel As Integer)
'Dim good_field As Boolean
'    good_field = validate_fields(SSDBGridList.Col)
'    If Not good_field Then
'       Cancel = True
'    End If

End Sub

Private Sub SSDBGridList_BeforeUpdate(Cancel As Integer)
Dim response As Integer
 Dim x, good_field As Boolean
 
  response = -1
If (SSDBGridList.IsAddRow And SSDBGridList.Col = 0 Or _
SSDBGridList.IsAddRow And SSDBGridList.Col = 1) And _
 (Not InSave) Then
   Cancel = True
   Exit Sub
End If
' If CAncelGrid = True Then
'       Cancel = True
'       CAncelGrid = False
'       Exit Sub
'  End If

 RecSaved = True
 If CAncelGrid = True Then
       Cancel = True
       CAncelGrid = False
       Exit Sub
  End If
  If InUnload Then
    msg1 = translator.Trans("M00704") 'J added
    response = MsgBox((IIf(msg1 = "", "Do you wish to save changes before closing?", msg1)), vbYesNo, "Imswin")
  End If
 If response = vbNo Then
    Cancel = True
    Exit Sub
 End If
 If (InUnload = False) Or (response = vbYes) Then
     good_field = validate_fields(SSDBGridList.Col)
     If Not good_field Then
        SSDBGridList.SetFocus
        SSDBGridList.Col = 2
        RecSaved = False
        Cancel = True
        Exit Sub
    End If
 
 '*** If SSDBGridList.IsAddRow Then
'***      x = NotValidLen(SSDBGridList.Columns(1).text)
'***      If x = True Then
'***         RecSaved = False
'***         Cancel = True
'***         msg1 = translator.Trans("M00702")
'***         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
'***                  SSDBGridList.SetFocus
'***         SSDBGridList.Col = 1
'***         Exit Sub
'***      End If
'***      x = CheckDesCode(SSDBGridList.Columns(0).text)
'***      If x <> False Then
'***         RecSaved = False
'***         msg1 = translator.Trans("M00703")
'***         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
'***         SSDBGridList.SetFocus
'***         SSDBGridList.Columns(0).text = ""
'***         SSDBGridList.Col = 0
'***         Exit Sub
'***      End If
   End If
'End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
     '   SSDBGridList.Columns("np").text = deIms.NameSpace
        If SSDBGridList.IsAddRow Then
            InsertElecDistribution
        Else
            UpdateElecDistribution
        End If
     '       SSDBGridList.Columns("create_date").text = Date
     '       SSDBGridList.Columns("create_user").text = CurrentUser
     '   End If
     '   SSDBGridList.Columns("modify_date").text = Date
     '   SSDBGridList.Columns("modify_user").text = CurrentUser
     '   Cancel = 0
     
     Else
       CAncelGrid = True
        RecSaved = False
      ' SSDBGridList.CancelUpdate
     Cancel = True
   End If
  
End Sub

'drop data grid

Private Sub SSDBGridList_InitColumnProps()
    Call EnableButtons
    SSDBGridList.Columns(0).DropDownHwnd = SSOleDBDDDisCode.HWND
   ' SSDBGridList.Columns(0)
    
    
End Sub

Private Sub SSDBGridList_KeyPress(KeyAscii As Integer)
 Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If SSDBGridList.IsAddRow And SSDBGridList.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
  '  If (SSDBGridList.IsAddRow And SSDBGridList.Col = 0) Then
  '     If Len(SSDBGridList.Columns(0).text) > 3 Then
  '        KeyAscii = 0
  '      End If
  '  End If
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGridList.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGridList.Col
        If (cur_col = 2) Then
            If GoodColMove = True Then
                SSDBGridList.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGridList.Col)
                If good_field Then
                    SSDBGridList.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub

'move recordset

Private Sub SSOleDBDDDisCode_Click()
    
            SSDBGridList.MoveLast
            SSDBGridList.MoveNext
            SSDBGridList.Columns("code").text = SSOleDBDDDisCode.Columns("code").text
            SSDBGridList.Columns("description").text = SSOleDBDDDisCode.Columns("description").text
        
End Sub

'set navbar buttom

Private Sub EnableButtons()
Dim i As Integer

    i = SSDBGridList.ROW
    
    If SSDBGridList.Rows = 0 Then
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
        
    ElseIf i = SSDBGridList.Rows - 1 Then
        NavBar1.LastEnabled = False
        NavBar1.NextEnabled = False
        
        NavBar1.FirstEnabled = True
        NavBar1.PreviousEnabled = True
    ElseIf i = 0 Then
        NavBar1.LastEnabled = True
        NavBar1.NextEnabled = True
        
        NavBar1.FirstEnabled = False
        NavBar1.PreviousEnabled = False
    Else
        NavBar1.LastEnabled = True
        NavBar1.NextEnabled = True
        
        NavBar1.FirstEnabled = True
        NavBar1.PreviousEnabled = True
    End If
    
    If Err Then Err.Clear
    'NavBar1.NewEnabled = True
    
        
End Sub
