VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmEUserDistribution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electronic Distribution"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8700
   Tag             =   "01040600"
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Remove Record"
      Top             =   4320
      Width           =   375
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBDDDisCode 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   8295
      _Version        =   196617
      DataMode        =   2
      BorderStyle     =   0
      Col.Count       =   6
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   2
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1508
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3228
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5001
      Columns(2).Caption=   "Mail"
      Columns(2).Name =   "Mail"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2646
      Columns(3).Caption=   "Fax"
      Columns(3).Name =   "Fax"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "recordId"
      Columns(4).Name =   "recordId"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1058
      Columns(5).Caption=   "Active"
      Columns(5).Name =   "Active"
      Columns(5).Alignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   11
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      _ExtentX        =   14631
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
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   4320
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      SaveEnabled     =   0   'False
      AllowCancel     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User  Distribution"
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
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmEUserDistribution"
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
Dim TableLocked As Boolean, currentformname As String   'jawdat
Dim Rstlist As ADODB.Recordset
Const CONSTFREIGHTPACKINGDESC = "FREIGHT/PACKING"
'SQL statement get document type and populate combo

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

'SQL statement get transcation type information and populate
'data grid

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
Private Function validate_fields(colnum As Integer) As Boolean
On Error Resume Next
Dim x As Boolean

    msg1 = translator.Trans("M00351") 'J added
    '------------------------------------------
    validate_fields = True
    If Not Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
          '  Call txtmailValidate(True)
            x = txtmailValidate(True)
            If Not x Then
                RecSaved = False
                validate_fields = False
                Exit Function
            End If
       ElseIf Not Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
       
            'Modified by Juan (9/15/2000) for Multilingual
            MsgBox IIf(msg1 = "", "Please enter either an Email or Fax.", msg1) 'J modified
            '---------------------------------------------
          RecSaved = False
         validate_fields = False
           
          'txtMail.SetFocus:
          Exit Function
       End If
    Else
        If Not Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
            'Call txtfaxnumber_validate(True)
             x = txtfaxnumber_validate(True)
             If Not x Then
                RecSaved = False
                validate_fields = False
                Exit Function
            End If
         ElseIf Not Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
         
            'Modified by Juan (9/15/2000) for Multilingual
            MsgBox IIf(msg1 = "", "You only can select one Email or Fax", msg1) 'J modified
            '---------------------------------------------
         RecSaved = False
         validate_fields = False

             'txtfaxNumb.SetFocus:
             Exit Function
        End If
    End If

 If Len(Trim$(SSDBGridList.Columns("MAIL").Text)) = 0 And Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
 
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00354") 'J added
    MsgBox IIf(msg1 = "", "You cannot leave Email and Fax empty", msg1) 'J modified
         RecSaved = False
         validate_fields = False
    '---------------------------------------------
 End If

End Function



Private Sub Command1_Click()
    Dim response As Integer

  '   MsgBox ("here")
      msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to remove the current record?", msg1)), vbOKCancel, "Imswin")
     If (response = vbOK) Then
            If Len(Trim$(SSDBGridList.Columns("mail").Text)) Then
                Call RemoveUserMail(SSDBGridList.Columns("Mail").Text) 'JCG 2008/1/23
                Call Clearform

            ElseIf Len(Trim$(SSDBGridList.Columns("fax").Text)) Then
                Call RemoveUserFax(SSDBGridList.Columns("fax").Text) 'JCG 200/1/23
                Call Clearform
            End If
        SSDBGridList.RemoveAll
        Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
        Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
        Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
        SSDBGridList.AllowUpdate = False

        SSDBGridList.SetFocus
        SSDBGridList.Col = 0
    End If
End Sub

'call function to get data and populate data grid

Private Sub Form_Load()
On Error Resume Next
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
    Call translator.Translate_Forms("frmEUserDistribution")
    '------------------------------------------
   
    Call GetDocumentCode
    'Call GetTranstypeCode
    Call GetDistributionCode

    SSDBGridList.DataMode = ssDataModeAddItem
    Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
   ' Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))

   NavBar1.NewEnabled = NavBar1.SaveEnabled
    'NavBar1.DeleteEnabled = NavBar1.NewEnabled
    
    frmEUserDistribution.Caption = frmEUserDistribution.Caption + " - " + frmEUserDistribution.Tag
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
    NavBar1.Width = 5055
    Call DisableButtons(Me, NavBar1)
     'NavBar1.DeleteEnabled = True
    ' NavBar1.DeleteVisible = True
    SSDBGridList.AllowUpdate = False
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'add data to data grid

Public Sub GetDistributionCode()
On Error Resume Next
Dim str As String

    str = Chr(1)
    SSOleDBDDDisCode.FieldSeparator = str

'    SSOleDBDDDisCode.AddItem "Update Database" & str & "UD"
'    SSOleDBDDDisCode.AddItem "Delivery" & str & "DL"
'    SSOleDBDDDisCode.AddItem "Login Security" & str & "LO"
'    SSOleDBDDDisCode.AddItem "Shipping" & str & "SH"
'    SSOleDBDDDisCode.AddItem "Warehouse Trans." & str & "WH"  'Commented Out by muzammil. Warehouse is not using autodistribution for now.
    SSOleDBDDDisCode.AddItem CONSTFREIGHTPACKINGDESC & str & "F"  'Added by muzammil. Freight Receipt and Packing List uses this code


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


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If




End Sub

Private Sub NavBar1_BeforeCancelClick()
On Error Resume Next
   CAncelGrid = True

End Sub

Private Sub NavBar1_BeforeDeleteClick()
On Error Resume Next
    Dim response As Integer

  '   MsgBox ("here")
      msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to remove the current record?", msg1)), vbOKCancel, "Imswin")
     If (response = vbOK) Then
            If Len(Trim$(SSDBGridList.Columns("mail").Text)) Then
                'Call DeleteUserMail(SSDBGridList.Columns("Mail").Text) 'JCG 2008/1/23
                Call RemoveUserMail(SSDBGridList.Columns("Mail").Text) 'JCG 2008/1/23
                Call Clearform

            ElseIf Len(Trim$(SSDBGridList.Columns("fax").Text)) Then
                'Call DeleteUserFax(SSDBGridList.Columns("fax").Text) 'JCG 200/1/23
                Call RemoveUserFax(SSDBGridList.Columns("fax").Text) 'JCG 200/1/23
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

Private Sub NavBar1_BeforeNewClick()
On Error Resume Next
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    'NavBar1.DeleteEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGridList.AllowUpdate = True
    SSDBGridList.Columns(1).locked = True

End Sub

Private Sub NavBar1_BeforeSaveClick()
'On Error Resume Next  'JCG 2008-12-12
On Error GoTo ErrHandler  'JCG 2008-12-12

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
            'NavBar1.DeleteEnabled = True
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGridList.AllowUpdate = False
       End If
     Call DisableButtons(Me, NavBar1)
   InSave = False

       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
       
  

            lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
          NavBar1.EditEnabled = True
          Exit Sub
          
          
ErrHandler: 'JCG 2008-12-12
    MsgBox "NabBar1_BeforeSaveClick -->" + Err.Description 'JCG 2008-12-12
    
End Sub

'refresh data grid

Private Sub NavBar1_OnCancelClick()
On Error Resume Next
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
          'NavBar1.DeleteEnabled = True
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
    'NavBar1.DeleteEnabled = True
    NavBar1.SaveEnabled = False
    SSDBGridList.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBGridList.Refresh
End If

If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
 
End Sub



Private Sub NavBar1_OnCloseClick()
On Error Resume Next
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    Unload Me
End Sub


Private Sub NavBar1_OnDeleteClick()
  '  If Len(Trim$(SSDBGridList.Columns("mail").text)) Then
  '      Call DeleteUserMail(SSDBGridList.Columns("Mail").text)
  '      Call Clearform
  '
  '  ElseIf Len(Trim$(SSDBGridList.Columns("fax").text)) Then
  '      Call DeleteUserFax(SSDBGridList.Columns("fax").text)
  '      Call Clearform
' '       SSDBGridList.MoveLast
  '  End If
  '      SSDBGridList.MoveLast
  
End Sub

Private Sub NavBar1_OnEditClick()

'
''copy begin here
'
'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)


   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode

NavBar1.NewEnabled = True
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
TableLocked = True
End If
'NavBar1.SaveEnabled = False
'NavBar1.NewEnabled = False
'NavBar1.CancelEnabled = False
'
'    Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'    Else
'    TableLocked = True
'    End If
'End If

'end copy





SSDBGridList.AllowUpdate = True
SSDBGridList.Columns(0).locked = True
SSDBGridList.Columns(1).locked = True
NavBar1.CancelEnabled = True
'NavBar1.DeleteEnabled = False
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

'move record to first position

Private Sub NavBar1_OnFirstClick()
On Error Resume Next
    SSDBGridList.MoveFirst
End Sub

'move record to last position

Private Sub NavBar1_OnLastClick()
On Error Resume Next
    SSDBGridList.MoveLast
End Sub

'move record to next position

Private Sub NavBar1_OnNextClick()
On Error Resume Next

  '  With SSDBGridList
        If Not SSDBGridList.EOF Then
        SSDBGridList.MoveNext

        Else
            With SSDBGridList
                If .EOF Then Exit Sub
            End With
        End If
  ' End With

   
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next

   
        If Not SSDBGridList.BOF Then
            SSDBGridList.MovePrevious
        Else
            With SSDBGridList
                If .BOF Then Exit Sub
            End With
        End If
   
    
End Sub

'call function get data and populate data grid

Private Sub NavBar1_OnNewClick()
   
       
    SSDBGridList.RemoveAll
    Call Addtogrid(Getlistforgrid(deIms.NameSpace, deIms.cnIms))
   ' Call Addtogridtran(Getlistforgridtran(deIms.NameSpace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.NameSpace, deIms.cnIms))
    
     SSDBGridList.AddNew
   SSDBGridList.SetFocus
    SSDBGridList.Col = 0
    '     Call NavBar1_OnPreviousClick
'     SSDBGridList.MoveNext
     'Call Clearform

End Sub

'clear data grid

Public Sub Clearform()
    SSDBGridList.Columns(0).Text = ""
    SSDBGridList.Columns(1).Text = ""
    SSDBGridList.Columns("mail").Text = ""
    SSDBGridList.Columns("fax").Text = ""
'    txt"mail" = ""
'    txtfaxNumb = ""
    
End Sub

'validate data format

Public Function DataValidate()
    If Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
          'SSDBGridList.Columns("mail").Text.SetFocus:
          Exit Function
       ElseIf Len(Trim$(SSDBGridList.Columns("fax").Text)) > 0 Then
       
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

'call store procedure to insert a record to database

Private Sub InsertElecDistribution()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "INSERT_DISTRIBUTION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "USER")
        .parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 5, SSOleDBDDDisCode.Columns("Code").Text)
        .parameters.Append .CreateParameter("@MAIL", adVarChar, adParamInput, 59, SSDBGridList.Columns("Mail").Text)
        .parameters.Append .CreateParameter("@FAXNUMB", adVarChar, adParamInput, 50, SSDBGridList.Columns("fax").Text)
        
        'JCG 2008-12-13
        Dim activeVal As String
        activeVal = Abs(val(SSDBGridList.Columns("Active").Text))
        .parameters.Append .CreateParameter("@ACTIVE", adVarChar, adParamInput, 1, activeVal)
        '---------------------------
        
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
        MsgBox IIf(msg1 = "", "Insert into Distribution failed.", msg1) 'J modified
        '---------------------------------------------

End Sub
Private Sub UpdateElecDistribution()
On Error GoTo Noinsert
Dim cmd As ADODB.Command
Dim activeVal   As String

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "UPDATE_DISTRIBUTION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "USER")
        .parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 5, SSDBGridList.Columns("Code").Text)
        .parameters.Append .CreateParameter("@MAIL", adVarChar, adParamInput, 59, SSDBGridList.Columns("Mail").Text)
        .parameters.Append .CreateParameter("@FAXNUMB", adVarChar, adParamInput, 50, SSDBGridList.Columns("fax").Text)
        
        'JCG 2008-12-13
        .parameters.Append .CreateParameter("@RECORDID", adVarChar, adParamInput, 4, SSDBGridList.Columns("recordId").Text)
        activeVal = Abs(val(SSDBGridList.Columns("Active").Text))
        .parameters.Append .CreateParameter("@ACTIVE", adVarChar, adParamInput, 1, activeVal)
        '---------------------------
        
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
MsgBox Err.Description

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
        .ParameterFields(1) = "type;USER;TRUE"
        
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

'before save record validate data format

Private Sub NavBar1_OnSaveClick()

    'Added by Juan (9/15/2000) for Multilingual
'***    msg1 = translator.Trans("M00351")
    '------------------------------------------


'***    If Not Len(Trim$(SSDBGridList.Columns("mail").text)) = 0 Then
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

 '***If Len(Trim$(SSDBGridList.Columns("MAIL").text)) = 0 And Len(Trim$(SSDBGridList.Columns("fax").text)) = 0 Then
 
    'Modified by Juan (9/15/2000) for Multilingual
'***    msg1 = translator.Trans("M00354") 'J added
'***    MsgBox IIf(msg1 = "", "You cannot leave Email and Fax empty", msg1) 'J modified
    '---------------------------------------------

 '***End If
    


End Sub

'SQL statement get distribution mail information and populate data grid

Public Function GetListofdistribution(NameSpace As String, Gender As String, cn As ADODB.Connection) As Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        'JCG 2008-12-13
        '.CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb, dis_recordId, dis_active "
        '-----------------------
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " WHERE dis_npecode = '" & NameSpace & "'"
        '.CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("MAIL").Text & "'"
        '.CommandText = .CommandText & " AND dis_gender = 'USER' " 'JCG 2008/1/23
        .CommandText = .CommandText & " AND dis_gender = 'USER' AND dis_id<>'-1'" 'JCG 2008/1/23
        .CommandText = .CommandText & " ORDER BY dis_id "
        Set Rstlist = .Execute

    End With

    If Rstlist Is Nothing Then Exit Function
        

        
        Set cmd = Nothing
        'Set Rstlist = Nothing
        
End Function

'SQL statement check distribution mailnumber exist or not

Public Function GetDistributionMail(EmailNunber As String) As Boolean
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").Text & "'"
        '.CommandText = .CommandText & " AND dis_gender =  'USER' " 'JCG 2008/1/23
        .CommandText = .CommandText & " AND dis_gender =  'USER' AND dis_id<>'-1'" 'JCG 2008/1/23
        .CommandText = .CommandText & " AND dis_mail = '" & SSDBGridList.Columns("MAIL").Text & "'"
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionMail = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
End Function

'SQL statement check distribution fax number exist or not

Public Function GetDistributionFaxnumb(Faxnumber As String) As Boolean

Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").Text & "'"
        '.CommandText = .CommandText & " AND dis_gender = 'USER'" 'JCG 2008/1/23
        .CommandText = .CommandText & " AND dis_gender = 'USER' and dis_id<>'-1' " 'JCG 2008/1/23
        .CommandText = .CommandText & " AND dis_faxnumb ='" & SSDBGridList.Columns("FAX").Text & "'"
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionFaxnumb = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing

End Function

'call function check mail exist or not

Public Function txtmailValidate(Cancel As Boolean) As Boolean
        
    Cancel = False
    txtmailValidate = True
    
    If Len(SSDBGridList.Columns("mail").Text) Then
            If GetDistributionMail(SSDBGridList.Columns("Mail").Text) Then
    '            SSDBGridList.Columns("Mail").Text = ""
                            
                'msg1 = translator.Trans("M00355") 'J added  'JCG 2008-12-13
                'MsgBox IIf(msg1 = "", "This configuration already exist.", msg1) 'J modified
                '---------------------------------------------
                'following line added
                'txtmailValidate = False 'JCG 2008-12-13
               
                'txtMail.SetFocus:
                'Exit Function 'JCG 2008-12-13
            Else
                'Cancel = False 'JCG 2008-12-13
        '***     Call txtfaxnumber_validate(True)
        '***     Call InsertElecDistribution
                'txtfaxNumb.SetFocus:
                'Exit Function 'JCG 2008-12-13
            End If
    End If
End Function

'call function check fax number exist or not

Public Function txtfaxnumber_validate(Cancel As Boolean) As Boolean
    Cancel = False
    txtfaxnumber_validate = True
    
       If Len(SSDBGridList.Columns("fax").Text) Then
            If GetDistributionFaxnumb(SSDBGridList.Columns("fax").Text) Then
                'txtfaxNumb.Text = ""
    
                'Modified by Juan (9/15/2000) for Multilingual
                msg1 = translator.Trans("M00355") 'J added
                MsgBox IIf(msg1 = "", "This configuration already exist.", msg1) 'J modified
                '---------------------------------------------
                'following line added
                txtfaxnumber_validate = False
    
                'txtfaxNumb.SetFocus: Exit Function
                'SSDBGridList.Columns(fax).Text Exit Function
            Else
                Cancel = False
                txtfaxnumber_validate = True
    '         Call InsertElecDistribution
                'txtMail.SetFocus:
                Exit Function
            End If
    End If
    

End Function

Public Function RemoveUserMail(Mail As String) As Boolean 'JCG 2008/1/23
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
RemoveUserMail = True
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "UPDATE DISTRIBUTION "
        .CommandText = .CommandText & "SET dis_id='-1' "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_mail ='" & Mail & "'"

        '.Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Call .Execute(0, 0, adExecuteNoRecords)
        
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            
            'Modified by Juan (9/15/2000) for Multilingual
            RemoveUserMail = False
            msg1 = translator.Trans("M00356") 'J
            MsgBox IIf(msg1 = "", "Remove from Distribution is failure ", msg1) 'J modified
            '---------------------------------------------
            
        Else
           ' msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Record Successfully Removed from Distribution", msg1) 'J modified
            SSDBGridList.Columns("Mail").Text = ""
            'txtMail.SetFocus
        End If
End Function


'SQL statement to detele  a mail records

Public Function DeleteUserMail(Mail As String) As Boolean
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
DeleteUserMail = True
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'USER'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_mail ='" & Mail & "'"

        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Call .Execute(0, 0, adExecuteNoRecords)
        
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            
            'Modified by Juan (9/15/2000) for Multilingual
            DeleteUserMail = False
            msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Delete from Distribution is failure ", msg1) 'J modified
            '---------------------------------------------
            
        Else
           ' msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Record Successfully Deleted from Distribution", msg1) 'J modified
            SSDBGridList.Columns("Mail").Text = ""
            'txtMail.SetFocus
        End If
End Function

Public Function RemoveUserFax(fax As String) As Boolean 'JCG 2008/1/23
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    RemoveUserFax = True
    With cmd
        .CommandText = "UPDATE DISTRIBUTION "
        .CommandText = .CommandText & "SET dis_id='-1' "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'USER' "
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & fax & "'"
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        

        
        Call .Execute(0, 0, adExecuteNoRecords)
       
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00356") 'J added
            MsgBox IIf(msg1 = "", "Remove from Distribution is failure ", msg1) 'J modified
            '---------------------------------------------
            RemoveUserFax = False
        Else
            SSDBGridList.Columns("FAX").Text = ""
            
        End If
End Function


'SQL statement to detele a fax number

Public Function DeleteUserFax(fax As String) As Boolean
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    DeleteUserFax = True
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'USER' "
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & fax & "'"
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        

        
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
            SSDBGridList.Columns("FAX").Text = ""
            
        End If
End Function

'assige value to data grid

Private Sub LoadValues()
On Error Resume Next

    SSDBGridList.Columns("Mail").Text = Rstlist!dis_mail & ""
    SSDBGridList.Columns("FAx").Text = Rstlist!dis_mail & ""

    If Err Then Err.Clear
End Sub

'populate data grid

Public Sub Addtogrid(Rstlist As ADODB.Recordset)
On Error GoTo errorHandler

Dim str As String
    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
    
        'JCG 2008-12-13
        'SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!doc_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!doc_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & "" & str & Rstlist!dis_recordId & "" & str & Rstlist!dis_active & ""
        '----------------------
    
        Rstlist.MoveNext
    Loop
    
    Exit Sub
    
errorHandler:
    MsgBox "Addtogrid-->" + Err.Description
End Sub

'populate transcation data grid

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

'add data to data grid

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
        If Trim$(Rstlist!dis_id) = "DL" Then desc = "Delivery"
        If Trim$(Rstlist!dis_id) = "LO" Then desc = "Login Security"
        If Trim$(Rstlist!dis_id) = "SH" Then desc = "Shipping"
        If Trim$(Rstlist!dis_id) = "F" Then desc = CONSTFREIGHTPACKINGDESC
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop

End Sub

'SQL statement get distribution list for form

Public Function Getlistforgrid(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        'JCG 2008-12-13
        '.CommandText = " SELECT  dis_id, doc_desc, dis_mail, dis_faxnumb "
        .CommandText = " SELECT  dis_id, doc_desc, dis_mail, dis_faxnumb,dis_recordId, dis_active "
        '----------------------
        .CommandText = .CommandText & " From DISTRIBUTION, doctype "
        .CommandText = .CommandText & " WHERE dis_npecode = doc_npecode "
        '.CommandText = .CommandText & " and dis_id = doc_code and "  'JCG 2008/1/23
        .CommandText = .CommandText & " and dis_id = doc_code and dis_id<>'-1' and " 'JCG 2008/1/23
        .CommandText = .CommandText & " dis_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'USER' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgrid = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function


'SQL statement get transcation for form

Public Function Getlistforgridtran(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, tty_desc, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION,  TRANSACTYPE "
        .CommandText = .CommandText & " WHERE dis_npecode =  tty_npecode "
        '.CommandText = .CommandText & " and dis_id = tty_code and " ''JCG 2008/1/23
        .CommandText = .CommandText & " and dis_id = tty_code and dis_id<>'-1' and "
        .CommandText = .CommandText & " dis_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'USER' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgridtran = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

'SQL statement get distribution recordset for form

Public Function GetlistforgridCode(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " where ((dis_id IN ( 'ud', 'sh', 'lo', 'dl','F')) "
        .CommandText = .CommandText & " AND (dis_gender = 'USER') "
        .CommandText = .CommandText & " AND (dis_npecode = '" & NameSpace & "'))"
        .CommandText = .CommandText & " order by dis_id "
        

        Set GetlistforgridCode = .Execute

    End With


       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

Private Sub SSDBGridList_BeforeUpdate(Cancel As Integer)
On Error Resume Next
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
    End If
'End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
     
        If SSDBGridList.IsAddRow Then
            InsertElecDistribution
        Else
            UpdateElecDistribution
        End If
        
        SSDBGridList.Columns(0).locked = False
        SSDBGridList.Columns(1).locked = False
        SSOleDBDDDisCode.Enabled = True
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        'NavBar1.DeleteEnabled = True
        lblStatus.ForeColor = &HFF00&
        lblStatus = Visualize
        NavBar1.EditEnabled = True
        NavBar1.NewEnabled = NVBAR_ADD
        SSDBGridList.AllowUpdate = False
        
    
        Call DisableButtons(Me, NavBar1)


        
     Else
        CAncelGrid = True
        RecSaved = False
     Cancel = True
   End If
End Sub

'drop down data grid

Private Sub SSDBGridList_InitColumnProps()
On Error Resume Next
    SSDBGridList.Columns(0).DropDownHwnd = SSOleDBDDDisCode.HWND

End Sub

Private Sub SSDBGridList_KeyPress(KeyAscii As Integer)
On Error Resume Next
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

'assign data to data grid

Private Sub SSOleDBDDDisCode_Click()
On Error Resume Next
            SSDBGridList.MoveLast
            SSDBGridList.MoveNext
            SSDBGridList.Columns("code").Text = SSOleDBDDDisCode.Columns("code").Text
            SSDBGridList.Columns("description").Text = SSOleDBDDDisCode.Columns("description").Text
   
End Sub

