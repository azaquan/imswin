VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Document 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document type"
   ClientHeight    =   3870
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   9780
   Tag             =   "01010700"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowUpdate     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLine 
      Bindings        =   "frm_Document.frx":0000
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   8950
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      FieldSeparator  =   ";"
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
      stylesets(0).Picture=   "frm_Document.frx":0014
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
      stylesets(1).Picture=   "frm_Document.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      CheckBox3D      =   0   'False
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   12
      Columns(0).Width=   1138
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "code"
      Columns(0).DataField=   "doc_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HeadStyleSet=   "colls"
      Columns(1).Width=   5292
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "doc_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1640
      Columns(2).Caption=   "Revision Flag"
      Columns(2).Name =   "Revision Flag"
      Columns(2).DataField=   "doc_reviflag"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(2).HeadStyleSet=   "colls"
      Columns(3).Width=   1535
      Columns(3).Caption=   "Ideas"
      Columns(3).Name =   "Ideas"
      Columns(3).DataField=   "doc_ideaflag"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "colls"
      Columns(4).Width=   1402
      Columns(4).Caption=   "Active"
      Columns(4).Name =   "Active"
      Columns(4).DataField=   "doc_actvflag"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   1
      Columns(4).Style=   2
      Columns(4).HeadStyleSet=   "colls"
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "np"
      Columns(5).Name =   "np"
      Columns(5).DataField=   "doc_npecode"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "create_date"
      Columns(6).Name =   "create_date"
      Columns(6).DataField=   "doc_creadate"
      Columns(6).DataType=   135
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create_user"
      Columns(7).Name =   "create_user"
      Columns(7).DataField=   "doc_creauser"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   5292
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "modify_date"
      Columns(8).Name =   "modify_date"
      Columns(8).DataField=   "doc_modidate"
      Columns(8).DataType=   135
      Columns(8).FieldLen=   256
      Columns(9).Width=   5292
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "modify_user"
      Columns(9).Name =   "modify_user"
      Columns(9).DataField=   "doc_modiuser"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1402
      Columns(10).Caption=   "Invoice"
      Columns(10).Name=   "Invoice"
      Columns(10).DataField=   "doc_invcreqd"
      Columns(10).DataType=   11
      Columns(10).FieldLen=   256
      Columns(10).Style=   2
      Columns(11).Width=   2566
      Columns(11).Caption=   "Autodistribution"
      Columns(11).Name=   "Autodistribution"
      Columns(11).DataField=   "doc_autodist"
      Columns(11).FieldLen=   256
      Columns(11).Style=   2
      UseDefaults     =   0   'False
      _ExtentX        =   15787
      _ExtentY        =   4471
      _StockProps     =   79
      DataMember      =   "DOCTYPE"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
      Left            =   5400
      TabIndex        =   3
      Top             =   3240
      Width           =   2460
   End
   Begin VB.Label lbl_Document 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_Document"
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
Dim GridEnabled As Boolean
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBLine.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBLine.Columns(colnum).Text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         If SSDBLine.Enabled = True Then SSDBLine.SetFocus
         SSDBLine.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckTypeexist(SSDBLine.Columns(0).Text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBLine.SetFocus
             SSDBLine.Columns(0).Text = ""
             SSDBLine.Col = 0
            validate_fields = False
         End If
    End If
   End If

End Function
'Load form populate data to combe,set button

Private Sub Form_Load()



Dim ctl As Control
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
 
    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Document")
    '------------------------------------------
 
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.DocType(deIms.NameSpace)
    
    Visible = True
 '   SSDBLine.DataMember = "DOCTYPE"
    Set NavBar1.Recordset = deIms.rsDOCTYPE
    Set SSDBLine.DataSource = deIms
    
    Caption = Caption + " - " + Tag
     
    
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
   ' NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    Call DisableButtons(Me, NavBar1)
    
     GridEnabled = SSDBLine.Enabled
     SSDBLine.Enabled = True
     
    NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
    
    SSDBLine.AllowUpdate = False
    SSDBLine.Columns(0).FieldLen = 2
    SSDBLine.Columns(1).FieldLen = 30

    With frm_Document
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)



If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If



On Error Resume Next
Dim response As String
On Error Resume Next
 InUnload = True
 RecSaved = True
 CAncelGrid = False
 SSDBLine.Update
 If RecSaved = True Then
    Hide
    deIms.rsDOCTYPE.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
 Else
    Cancel = True
End If
End Sub

'cancel recordset update

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
'    SSDBLine.CancelUpdate
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBLine.Update
 
End Sub

'set combo name space equal to current name space

Private Sub NavBar1_BeforeNewClick()
     SSDBLine.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBLine.AllowUpdate = True
    SSDBLine.Columns("active").Text = 1
    SSDBLine.Columns("Revision Flag").Text = 1
    SSDBLine.Columns("Ideas").Text = 1
    SSDBLine.Columns("autodistribution").Text = 1
    If SSDBLine.Enabled = True Then SSDBLine.SetFocus
    SSDBLine.Col = 0 '  SSDBLine.AddNew
    
   ' SSDBLine.Columns(2).Text = 0
   ' SSDBLine.Columns(3).Text = 0
    'SSDBLine.Columns("np").Text = deIms.NameSpace
End Sub

'before save check code exist or not, if code exist, give message
'cancel recordset update

Private Sub NavBar1_BeforeSaveClick()
On Error Resume Next
    CAncelGrid = False
     SSDBLine.Update
       If RecSaved = True Then
            SSDBLine.Columns(0).locked = False
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBLine.AllowUpdate = False
        End If
'Dim Numb As Integer
'Dim number As Integer
'Dim numbe As Integer

  
    
 '   Numb = SSDBLine.Rows
 '   number = SSDBLine.GetBookmark(-1)
 '   numbe = SSDBLine.Bookmark
    
'    If (Numb - number) = 1 And (Numb > numbe) Then
 '           If CheckTypeexist(SSDBLine.Columns(0).text) Then
            
  '              'Modified by Juan (9/11/2000) for Multilingual
   '             msg1 = translator.Trans("M00013") 'J added
    '            MsgBox IIf(msg1 = "", "Code exist, Please make new one", msg1) 'J modified
    '            '---------------------------------------------
    '
    '            SSDBLine.CancelUpdate
    '            Exit Sub
    '        End If
            
   ' End If
            
    '        SSDBLine.Update
    '    Call SSDBLine.MoveRecords(0)
'        Call SSDBLine.MoveRecords(-1)
       
       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
       
       
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As Integer
   If SSDBLine.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBLine.CancelUpdate
        ' Cancel = -1
          CAncelGrid = False
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBLine.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBLine.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBLine.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
  '  SSDBLine.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBLine.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
    SSDBLine.Columns(0).locked = False
'    SSDBLine.Refresh
End If


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
 
'  FormMode = ChangeModeOfForm(lblStatus, mdVisualization)
' If FormMode = mdVisualization Then MakeReadOnly (False)
'
'    If sstSup.Tab = 3 Then
'        ssdbgContacts.CancelUpdate
'    ElseIf sstSup.Tab = 0 Then
'        i = deIms.rsINtSupplier.editmode
'        rs.CancelUpdate
'        deIms.rsINtSupplier.CancelUpdate
'
'        If Not i = adEditAdd Then GetOriginalValues
'
'        'Call deIms.rsIntSupplier.CancelBatch(adAffectCurrent)
'        'dcboSuppCode.locked = True
'    End If

      
End Sub

'close form

Private Sub NavBar1_OnCloseClick()



If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    Unload Me
End Sub

Public Function chk_LI() As String
'    Dim Missing_field As String
'    Dim ls_code As String
'    Dim Focus_Flag As Boolean
'
'
'    If Trim(txt_Description.Text) = "" Then
'        Missing_field = Missing_field & "Description~"
'        If Focus_Flag = False Then
'            txt_Description.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'
'    If Trim(cbo_Code.Text) = "" Then
'        Missing_field = Missing_field & "Code~"
'        If Focus_Flag = False Then
'            cbo_Code.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'
'    ls_code = piece1(Missing_field, "~")
'    chk_LI = ls_code
'
'    Do While Missing_field <> ""
'        ls_code = piece1(Missing_field, "~")
'        chk_LI = chk_LI & "," & ls_code
'    Loop
End Function

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
                                                     'Exit Edit sub because theres nothing the user can do

NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False
Exit Sub
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

Else



SSDBLine.AllowUpdate = True

SSDBLine.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
If SSDBLine.Enabled = True Then SSDBLine.SetFocus
SSDBLine.Col = 1
SSDBLine.AllowUpdate = True


TableLocked = True
End If


End Sub

Private Sub NavBar1_OnFirstClick()
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
End Sub

Private Sub NavBar1_OnLastClick()
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
End Sub

Private Sub NavBar1_OnNewClick()
    
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else
    
    SSDBLine.AllowUpdate = False

End Sub

Private Sub NavBar1_OnNextClick()
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
End Sub

Private Sub NavBar1_OnPreviousClick()
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
End Sub

'get crystal report parameter, and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Doctype.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00054") 'J added
        .WindowTitle = IIf(msg1 = "", "Document type", msg1) 'J modified
        Call translator.Translate_Reports("Doctype.rpt") 'J added
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

'commit save tramsaction

Private Sub NavBar1_OnSaveClick()



On Error Resume Next
    Call deIms.rsDOCTYPE.Move(0)
    If Err Then Err.Clear
'    Call CommitTransaction(deIms.cnIms)
'    MsgBox "Insert into Document type was completed"
End Sub

Private Sub SSDBLine_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBLine.AllowUpdate = False
End If
End Sub

'before colume update check code field, if code field value have been changed
'show message and set back values

Private Sub SSDBLine_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
      
'   If TransCancelled = False Then
    
    
          If SSDBLine.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBLine.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                SSDBLine.Columns(ColIndex).Text = oldVALUE
                RecSaved = False
                SSDBLine.Col = 0
                GoodColMove = False
              ElseIf CheckTypeexist(SSDBLine.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                SSDBLine.Columns(ColIndex).Text = oldVALUE
                RecSaved = False
                SSDBLine.Col = 0
                GoodColMove = False
             End If
        
        ElseIf SSDBLine.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBLine.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                RecSaved = False
                'SSDBLine.Columns(ColIndex).Text =
                SSDBLine.Col = 1
               End If
        ElseIf Not SSDBLine.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBLine.Columns(ColIndex).Text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                'SSDBLine.Columns(ColIndex).Text =
                RecSaved = False
                SSDBLine.Col = 1
               End If
       End If
            Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBLine.Columns(ColIndex).Text))
       'Dim oldstr As String
'Dim newstr As String
'Dim Numb As Integer
'Dim number As Integer
'Dim numbe As Integer
'
    'oldstr = SSDBLine.Columns(0).CellText(SSDBLine.Bookmark)
    'newstr = SSDBLine.Columns(0).text
'
    'Numb = SSDBLine.Rows
    'number = SSDBLine.GetBookmark(-1)
    'numbe = SSDBLine.Bookmark
'
    'If (Numb - number) = 1 And (Numb > numbe) Then
            'Exit Sub
        'Else
        'If ColIndex = 0 Then
            'If oldstr <> newstr Then
                'Cancel = True
'
                ''Modified by Juan (9/11/2000) for Multilingual
                'msg1 = translator.Trans("M00015") 'J added
                'MsgBox IIf(msg1 = "", " Code can not changed once it is saved, Please make new one.", msg1) 'J modified
                ''---------------------------------------------
'
            'End If
        'End If
    'End If
'
End Sub

Private Sub SSDBLine_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBLine.Col)
    If Not good_field Then
       Cancel = True
    End If
End Sub

'before save validate data fields

Private Sub SSDBLine_BeforeUpdate(Cancel As Integer)
  Dim response As Integer
  Dim x As Boolean
 response = 0
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
  
  If SSDBLine.IsAddRow Then
      x = NotValidLen(SSDBLine.Columns(1).Text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBLine.SetFocus
         SSDBLine.Col = 0
         Exit Sub
      End If
      x = CheckTypeexist(SSDBLine.Columns(0).Text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBLine.Columns(0).Text = ""
        SSDBLine.SetFocus
         SSDBLine.Col = 0
         Exit Sub
      End If
   End If
End If
      
    
  '  Cancel = 0
  'Else
'    If InUnload Then
     '     msg1 = translator.Trans("M00704") 'J added
      '    Response = MsgBox((IIf(msg1 = "", "Do you wish to save changes before closing?", msg1)), vbOKCancel, "Imswin")
'    Else
    If InUnload = False Then
          msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
    End If
 '  End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBLine.Columns("np").Text = deIms.NameSpace
        If SSDBLine.IsAddRow Then
            SSDBLine.Columns("create_date").Text = Date
            SSDBLine.Columns("create_user").Text = CurrentUser
        End If
        SSDBLine.Columns("modify_date").Text = Date
        SSDBLine.Columns("modify_user").Text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
 '      SSDBLine.CancelUpdate
     Cancel = True
  '   SSDBLine.Refresh
   End If
     'Cancel = True
'
    ''Modified by Juan (9/11/2000) for Multilingual
    'msg1 = translator.Trans("M00016") 'J added
    'If SSDBLine.Columns(0).text = "" Then
        'MsgBox SSDBLine.Columns(0).Caption & IIf(msg1 = "", " Cannot be left empty", " " + msg1): Exit Sub 'J modified
    'ElseIf SSDBLine.Columns(1).text = "" Then
        'MsgBox SSDBLine.Columns(1).Caption & IIf(msg1 = "", " Cannot be left empty", " " + msg1): Exit Sub 'J modified
    'Else
        'Cancel = False
    'End If
    ''--------------------------------------------
    
End Sub

'SQL statement check code exist or not

Private Function CheckTypeexist(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DOCTYPE"
        .CommandText = .CommandText & " Where doc_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND doc_code = '" & Code & "'"
        
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckTypeexist = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckTypeexist", Err.Description, Err.number, True)

End Function

Private Sub SSDBLine_KeyPress(KeyAscii As Integer)
 Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If Not SSDBLine.IsAddRow And SSDBLine.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBLine.Col = 4)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBLine.Col
        If (cur_col = 4) Then
            If GoodColMove = True Then
                SSDBLine.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBLine.Col)
                If good_field Then
                    SSDBLine.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub
