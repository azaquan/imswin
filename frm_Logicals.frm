VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Logicals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logical Warehouse"
   ClientHeight    =   4305
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   7275
   Tag             =   "01030500"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLogical 
      Bindings        =   "frm_Logicals.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6315
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
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
      stylesets(0).Picture=   "frm_Logicals.frx":0014
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frm_Logicals.frx":0030
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   1693
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "lw_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   10
      Columns(1).Width=   7117
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "lw_desc"
      Columns(1).DataType=   8
      Columns(1).Case =   2
      Columns(1).FieldLen=   40
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NP"
      Columns(2).Name =   "NP"
      Columns(2).DataField=   "lw_npecode"
      Columns(2).FieldLen=   256
      Columns(3).Width=   1588
      Columns(3).Caption=   "Active"
      Columns(3).Name =   "Active"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "lw_actvflag"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "ColHeader"
      Columns(3).StyleSet=   "ColHeader"
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "modify_date"
      Columns(4).Name =   "modify_date"
      Columns(4).DataField=   "lw_modidate"
      Columns(4).DataType=   135
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_user"
      Columns(5).Name =   "modify_user"
      Columns(5).DataField=   "lw_modiuser"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "create_date"
      Columns(6).Name =   "create_date"
      Columns(6).DataField=   "lw_creadate"
      Columns(6).DataType=   135
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create_user"
      Columns(7).Name =   "create_user"
      Columns(7).DataField=   "lw_creauser"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   11139
      _ExtentY        =   4895
      _StockProps     =   79
      DataMember      =   "LOGWAR"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   2460
   End
   Begin VB.Label lbl_Logicals 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logical Warehouse"
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
      Left            =   1980
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frm_Logicals"
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
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBLogical.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBLogical.Columns(colnum).text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBLogical.SetFocus
         SSDBLogical.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(SSDBLogical.Columns(0).text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBLogical.SetFocus
             SSDBLogical.Col = 0
             SSDBLogical.Columns(0).text = ""
            validate_fields = False
         End If
    End If
   End If

End Function
'get data and load form,set navbar button

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
    Call translator.Translate_Forms("frm_Logicals")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    'Call deIms.Destination(deIms.NameSpace)
    deIms.rsLOGWAR.Close
     Call deIms.LOGWAR(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsLOGWAR
    Set SSDBLogical.DataSource = deIms
    
    
    Visible = True
    Screen.MousePointer = vbDefault
    'SSDBLogical.BatchUpdate = True
    
    Caption = Caption + " - " + Tag
    
     NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    Call DisableButtons(Me, NavBar1)
    SSDBLogical.AllowUpdate = False
    
    'Call DisableButtons(Me, NavBar1)
    
    With frm_Logicals
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

 
 
 
 
 Dim response As String
On Error Resume Next
 RecSaved = True
 InUnload = True
 CAncelGrid = False
 SSDBLogical.Update
 If RecSaved = True Then
    Hide
    deIms.rsLOGWAR.Close
    
    If open_forms <= 5 Then ShowNavigator
    
    If Err Then Err.Clear
    
Else
    Cancel = True
End If
End Sub

'cancel record update

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBLogical.Update

End Sub

'set create user equal to current user and name space

Private Sub NavBar1_BeforeNewClick()
      SSDBLogical.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBLogical.AllowUpdate = True
    SSDBLogical.Columns("active").text = 1
    SSDBLogical.SetFocus
    NavBar1.Width = 5050
    SSDBLogical.Col = 0
End Sub

'before save check logical warehouse exist or not
'show message

Private Sub NavBar1_BeforeSaveClick()
    CAncelGrid = False
    SSDBLogical.Update
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBLogical.AllowUpdate = False
       End If
       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
 
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As Integer

If TableLocked = True Then


   SSDBLogical.Columns("code").locked = True
   SSDBLogical.Columns("description").locked = True
   SSDBLogical.Columns("active").locked = True

Else

   If SSDBLogical.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBLogical.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBLogical.CancelUpdate
       '   SSDBLogical.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBLogical.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBLogical.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBLogical.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBLogical.CancelUpdate
    SSDBLogical.Columns(0).locked = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBLogical.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBLogical.Refresh
End If
      

If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
 End If
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
'    Set frm_Logicals = Nothing
End Sub

Private Sub NavBar1_OnEditClick()

''copy begin here
'
'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)


   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode


   SSDBLogical.Columns("code").locked = True
   SSDBLogical.Columns("description").locked = True
   SSDBLogical.Columns("active").locked = True
NavBar1.NewEnabled = False
NavBar1.SaveEnabled = False

Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
SSDBLogical.AllowUpdate = True
SSDBLogical.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBLogical.SetFocus
SSDBLogical.Col = 1
SSDBLogical.AllowUpdate = True


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
End Sub

Private Sub NavBar1_OnNewClick()



    SSDBLogical.AllowUpdate = False

If TableLocked = True Then
   SSDBLogical.Columns("code").locked = False
   SSDBLogical.Columns("description").locked = False
   SSDBLogical.Columns("active").locked = False
End If
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Logwar.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00116") 'J added
        .WindowTitle = IIf(msg1 = "", "Logical Warehouse", msg1) 'J modified
        Call translator.Translate_Reports("Logwar.rpt") 'J added
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

'before save logical location exist or not, show message

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsLOGWAR.Move(0)
    If Err Then Err.Clear

End Sub

Private Sub SSDBLogical_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBLogical.AllowUpdate = False
End If

End Sub

'check logical code exist or not show message

Private Sub SSDBLogical_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSDBLogical.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBLogical.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBLogical.SetFocus
                SSDBLogical.Columns(ColIndex).text = oldVALUE
                SSDBLogical.Col = 0
                RecSaved = False
                GoodColMove = False
              ElseIf CheckDesCode(SSDBLogical.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBLogical.SetFocus
                SSDBLogical.Columns(ColIndex).text = oldVALUE
                SSDBLogical.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf SSDBLogical.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBLogical.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBLogical.SetFocus
                'SSDBLogical.Columns(ColIndex).Text =
                RecSaved = False
                SSDBLogical.Col = 1
               End If
        ElseIf Not SSDBLogical.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBLogical.Columns(ColIndex).text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBLogical.SetFocus
                'SSDBLogical.Columns(ColIndex).Text =
                RecSaved = False
                SSDBLogical.Col = 1
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBLogical.Columns(ColIndex).text))
     
        
End Sub

'SQL statement check logical code exist or not

Private Function CheckLogwarexist(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT ? = count(*)"
        .CommandText = .CommandText & " From LOGWAR"
        .CommandText = .CommandText & " Where lw_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND lw_code = '" & Code & "'"
        
        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute(Options:=adExecuteNoRecords)
        
        CheckLogwarexist = cmd.parameters(0)
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckLogwarexist", Err.Description, Err.number, True)

End Function

Private Sub SSDBLogical_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBLogical.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

'before add a new record check code exist or not
'show message

Private Sub SSDBLogical_BeforeUpdate(Cancel As Integer)

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
 
  If SSDBLogical.IsAddRow Then
      x = NotValidLen(SSDBLogical.Columns(1).text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBLogical.SetFocus
         SSDBLogical.Col = 1
         Exit Sub
      End If
      x = CheckDesCode(SSDBLogical.Columns(0).text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBLogical.SetFocus
         SSDBLogical.Col = 0
         SSDBLogical.Columns(0).text = ""
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBLogical.Columns("np").text = deIms.NameSpace
        If SSDBLogical.IsAddRow Then
            SSDBLogical.Columns("create_date").text = Date
            SSDBLogical.Columns("create_user").text = CurrentUser
        End If
        SSDBLogical.Columns("modify_date").text = Date
        SSDBLogical.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
       'SSDBLogical.CancelUpdate
     Cancel = True
   End If
  
  
End Sub
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From LOGWAR "
        .CommandText = .CommandText & " Where lw_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND lw_code = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function

Private Sub SSDBLogical_KeyPress(KeyAscii As Integer)
Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If Not SSDBLogical.IsAddRow And SSDBLogical.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBLogical.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBLogical.Col
        If (cur_col = 2) Then
            If GoodColMove = True Then
                SSDBLogical.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBLogical.Col)
                If good_field Then
                    SSDBLogical.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub
