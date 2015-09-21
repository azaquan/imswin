VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmTermCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terms of Condition"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   8295
   Tag             =   "01011100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3960
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
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGTermCondition 
      Bindings        =   "frmTermCondition.frx":0000
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
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
      stylesets(0).Name=   "Colls"
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
      stylesets(0).Picture=   "frmTermCondition.frx":0014
      stylesets(1).Name=   "Rows"
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
      stylesets(1).Picture=   "frmTermCondition.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorFrame =   -2147483630
      BevelColorHighlight=   14737632
      BevelColorShadow=   -2147483633
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   2117
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "tac_taccode"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   10
      Columns(0).BackColor=   16777215
      Columns(0).HeadStyleSet=   "Colls"
      Columns(1).Width=   3572
      Columns(1).Caption=   "File Name"
      Columns(1).Name =   "File Name"
      Columns(1).DataField=   "tac_filename"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(1).HeadStyleSet=   "Colls"
      Columns(2).Width=   5424
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "tac_desc"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   80
      Columns(2).HeadStyleSet=   "Colls"
      Columns(3).Width=   5292
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "np"
      Columns(3).Name =   "np"
      Columns(3).DataField=   "tac_npecode"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   5
      Columns(4).Width=   1455
      Columns(4).Caption=   "Active"
      Columns(4).Name =   "Active"
      Columns(4).DataField=   "tac_actvflag"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(4).HeadStyleSet=   "Colls"
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_date"
      Columns(5).Name =   "modify_date"
      Columns(5).DataField=   "tac_modidate"
      Columns(5).DataType=   135
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "modify_user"
      Columns(6).Name =   "modify_user"
      Columns(6).DataField=   "tac_modiuser"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create_date"
      Columns(7).Name =   "create_date"
      Columns(7).DataField=   "tac_creadate"
      Columns(7).DataType=   135
      Columns(7).FieldLen=   256
      Columns(8).Width=   5292
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "create_user"
      Columns(8).Name =   "create_user"
      Columns(8).DataField=   "tac_creauser"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   13361
      _ExtentY        =   5636
      _StockProps     =   79
      DataMember      =   "TermCondition"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   3840
      Width           =   3300
   End
   Begin VB.Label lbl_ServiceCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terms of Condition"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmTermCondition"
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
Dim X As Boolean

validate_fields = True
If SSDBGTermCondition.IsAddRow Then
   If colnum = 0 Or colnum = 2 Then
      X = NotValidLen(SSDBGTermCondition.Columns(colnum).Text)
      If X = True Then
         RecSaved = False
         msg1 = translator.Trans("M00921")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGTermCondition.SetFocus
         SSDBGTermCondition.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        X = CheckDesCode(SSDBGTermCondition.Columns(0).Text)
        If X <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBGTermCondition.SetFocus
             SSDBGTermCondition.Col = 0
             SSDBGTermCondition.Columns(0).Text = ""
            validate_fields = False
         End If
    End If
   End If

End Function

'load form and set back ground color

Private Sub Form_Load()






Dim ctl As Control
Dim rs As ADODB.Recordset

    Screen.MousePointer = vbHourglass
   CAncelGrid = False
   msg1 = translator.Trans("L00126")
   Modify = IIf(msg1 = "", "Modification", msg1)
   msg1 = translator.Trans("L00684")
   Visualize = IIf(msg1 = "", "Visualization", msg1)
   msg1 = translator.Trans("L00125")
   Create = IIf(msg1 = "", "Creation", msg1)
   GoodColMove = True
   RecSaved = True
   InUnload = False
    lblStatus.Caption = Visualize
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmTermCondition")
    '------------------------------------------
    
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    If deIms.rsTermCondition.State = adStateClosed Then
        Call deIms.TermCondition(deIms.NameSpace)
    End If
    Set NavBar1.Recordset = deIms.rsTermCondition
    
    Set SSDBGTermCondition.DataSource = deIms
    
    Set rs = Nothing
    Screen.MousePointer = vbDefault
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
    SSDBGTermCondition.AllowUpdate = False
End Sub

'unload form set recordset to close

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
On Error Resume Next
 InUnload = True
  RecSaved = True
  CAncelGrid = False
SSDBGTermCondition.Update
 If RecSaved = True Then

    Hide
    deIms.rsTermCondition.Close
    If open_forms <= 5 Then ShowNavigator
 Else
    Cancel = True
End If
       
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

End Sub

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
'''''If SSDBGTermCondition.IsAddRow Then SSDBGTermCondition.CancelUpdate
'''''    SSDBGTermCondition.Refresh
End Sub

'set record sset update

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGTermCondition.Update
    
End Sub

'set recordset add new

Private Sub NavBar1_BeforeNewClick()
    SSDBGTermCondition.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGTermCondition.AllowUpdate = True
    SSDBGTermCondition.Columns("active").Text = 1
    SSDBGTermCondition.SetFocus
    SSDBGTermCondition.Col = 0
End Sub

'before save records set record update

Private Sub NavBar1_BeforeSaveClick()
    CAncelGrid = False
     SSDBGTermCondition.Update
     ' SSDBGTermCondition.Refresh
    'Call SSDBGTermCondition.MoveRecords(0)
'''    NavBar1.NewEnabled = NVBAR_ADD
'''    NavBar1.EditEnabled = NVBAR_EDIT
'''    NavBar1.SaveEnabled = False
'''    lblStatus.Caption = Visualize
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            SSDBGTermCondition.Columns(0).locked = False
            SSDBGTermCondition.Columns(1).locked = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGTermCondition.AllowUpdate = False
       End If
       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
              
       
       
       
End Sub

'cancel recordset update

Private Sub NavBar1_OnCancelClick()
 Dim response As Integer
   If SSDBGTermCondition.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGTermCondition.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGTermCondition.CancelUpdate
       '   SSDBGTermCondition.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBGTermCondition.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGTermCondition.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGTermCondition.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBGTermCondition.CancelUpdate
    SSDBGTermCondition.Columns(0).locked = False
    SSDBGTermCondition.Columns(1).locked = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBGTermCondition.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBGTermCondition.Refresh
End If


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If


   ' Else
   ' SSDBGTermCondition.Refresh
   ' End If
'    SSDBGTermCondition.CancelUpdate
    'NavBar1.SaveEnabled =
    'SSDBGTermCondition.Columns(1).text = SSDBGTermCondition.Columns(1).CellText(SSDBGTermCondition.Bookmark)
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

Private Sub NavBar1_OnEditClick()

'copy begin here

'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
   
   
   SSDBGTermCondition.Columns("code").locked = True
   SSDBGTermCondition.Columns("description").locked = True
   SSDBGTermCondition.Columns("active").locked = True
   SSDBGTermCondition.Columns("file name").locked = True
   
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
    Else
 
'End If

'end copy


'SSDBGTermCondition.Refresh
SSDBGTermCondition.AllowUpdate = True
SSDBGTermCondition.Columns(0).locked = True
SSDBGTermCondition.Columns(1).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGTermCondition.SetFocus
SSDBGTermCondition.Col = 2
SSDBGTermCondition.AllowUpdate = True
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

'set name space equal to current name space

Private Sub NavBar1_OnNewClick()
  
    SSDBGTermCondition.AllowUpdate = False
    'deIms.rsDestination!des_npecode = deIms.NameSpace
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

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\termcond.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00483") 'J added
        .WindowTitle = IIf(msg1 = "", "Terms of Condition", msg1) 'J modified
        Call translator.Translate_Reports("termcond.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With

Handler:
    If Err Then MsgBox Err.Description: Err.Clear
End Sub


Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsTermCondition.Move(0)
    If Err Then Err.Clear

End Sub

Private Sub SSDBGTermCondition_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBGTermCondition.AllowUpdate = False
End If
End Sub

Private Sub SSDBGTermCondition_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSDBGTermCondition.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGTermCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00921")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBGTermCondition.SetFocus
                SSDBGTermCondition.Columns(ColIndex).Text = oldVALUE
                SSDBGTermCondition.Col = 0
                RecSaved = False
                GoodColMove = False
              ElseIf CheckDesCode(SSDBGTermCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                RecSaved = False
                SSDBGTermCondition.SetFocus
                SSDBGTermCondition.Columns(ColIndex).Text = oldVALUE
                SSDBGTermCondition.Col = 0
                GoodColMove = False
             End If
        
        ElseIf SSDBGTermCondition.IsAddRow And ColIndex = 2 Then
              If NotValidLen(SSDBGTermCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00921")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                RecSaved = False
                SSDBGTermCondition.SetFocus
                'SSDBGTermCondition.Columns(ColIndex).Text =
                SSDBGTermCondition.Col = 2
               End If
        ElseIf Not SSDBGTermCondition.IsAddRow And ColIndex = 2 Then
                If NotValidLen(SSDBGTermCondition.Columns(ColIndex).Text) Then
               msg1 = translator.Trans("M00921")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGTermCondition.SetFocus
                'SSDBGTermCondition.Columns(ColIndex).Text =
                RecSaved = False
                SSDBGTermCondition.Col = 2
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBGTermCondition.Columns(ColIndex).Text))

End Sub

Private Sub SSDBGTermCondition_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBGTermCondition.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

Private Sub SSDBGTermCondition_BeforeUpdate(Cancel As Integer)
 Dim response As Integer
 Dim X, Y As Boolean
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
 
  If SSDBGTermCondition.IsAddRow Then
      X = NotValidLen(SSDBGTermCondition.Columns(2).Text)
      If (X = True) Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00921")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBGTermCondition.SetFocus
         SSDBGTermCondition.Col = 2
         Exit Sub
      End If
      X = CheckDesCode(SSDBGTermCondition.Columns(0).Text)
      If X <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBGTermCondition.SetFocus
         SSDBGTermCondition.Columns(0).Text = ""
         SSDBGTermCondition.Col = 0
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBGTermCondition.Columns("np").Text = deIms.NameSpace
        If SSDBGTermCondition.IsAddRow Then
            SSDBGTermCondition.Columns("create_date").Text = Date
            SSDBGTermCondition.Columns("create_user").Text = CurrentUser
        End If
        SSDBGTermCondition.Columns("modify_date").Text = Date
        SSDBGTermCondition.Columns("modify_user").Text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
 '      SSDBGTermCondition.CancelUpdate
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
         .CommandText = .CommandText & " From TERMSANDCONDITION "
        .CommandText = .CommandText & " Where tac_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND tac_taccode = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function


Private Sub SSDBGTermCondition_KeyPress(KeyAscii As Integer)
 Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If (Not SSDBGTermCondition.IsAddRow And SSDBGTermCondition.Col = 0 And KeyAscii <> 13) Or _
  (Not SSDBGTermCondition.IsAddRow And SSDBGTermCondition.Col = 1 And KeyAscii <> 13) Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGTermCondition.Col = 3)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGTermCondition.Col
        If (cur_col = 3) Then
            If GoodColMove = True Then
                SSDBGTermCondition.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGTermCondition.Col)
                If good_field Then
                    SSDBGTermCondition.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub
