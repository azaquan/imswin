VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Condition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Condition Code"
   ClientHeight    =   4845
   ClientLeft      =   1890
   ClientTop       =   1500
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   8295
   Tag             =   "01030700"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4320
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGCondition 
      Bindings        =   "frm_Condition.frx":0000
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7935
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
      stylesets(0).Picture=   "frm_Condition.frx":0014
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
      stylesets(1).Picture=   "frm_Condition.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorHighlight=   16777215
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
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   1191
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "cond_condcode"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   2
      Columns(0).HeadStyleSet=   "Colls"
      Columns(0).StyleSet=   "Rows"
      Columns(1).Width=   8440
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "cond_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   40
      Columns(1).HeadStyleSet=   "Colls"
      Columns(1).StyleSet=   "Rows"
      Columns(2).Width=   1693
      Columns(2).Caption=   "Percent"
      Columns(2).Name =   "Percent"
      Columns(2).DataField=   "cond_perc"
      Columns(2).DataType=   131
      Columns(2).Case =   2
      Columns(2).FieldLen=   256
      Columns(2).HeadStyleSet=   "Colls"
      Columns(2).StyleSet=   "Rows"
      Columns(3).Width=   1323
      Columns(3).Caption=   "Active"
      Columns(3).Name =   "Active"
      Columns(3).DataField=   "cond_actvflag"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "Colls"
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "np"
      Columns(4).Name =   "np"
      Columns(4).DataField=   "cond_npecode"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_date"
      Columns(5).Name =   "modify_date"
      Columns(5).DataField=   "cond_modidate"
      Columns(5).DataType=   135
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "modify_user"
      Columns(6).Name =   "modify_user"
      Columns(6).DataField=   "cond_modiuser"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create-date"
      Columns(7).Name =   "create_date"
      Columns(7).DataField=   "cond_creadate"
      Columns(7).DataType=   135
      Columns(7).FieldLen=   256
      Columns(8).Width=   5292
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "create_user"
      Columns(8).Name =   "create_user"
      Columns(8).DataField=   "cond_creauser"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   13996
      _ExtentY        =   6165
      _StockProps     =   79
      BackColor       =   -2147483638
      DataMember      =   "CONDITION"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label lbl_Condition 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Condition Code"
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
      Left            =   3015
      TabIndex        =   1
      Top             =   60
      Width           =   2070
   End
End
Attribute VB_Name = "frm_Condition"
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
Dim FormMode As FormMode

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBGCondition.IsAddRow Then
   If colnum = 0 Or colnum = 1 Or colnum = 2 Then
    x = NotValidLen(SSDBGCondition.Columns(colnum).Text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGCondition.SetFocus
         SSDBGCondition.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(SSDBGCondition.Columns(0).Text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBGCondition.SetFocus
             SSDBGCondition.Col = 0
             SSDBGCondition.Columns(0).Text = ""
            validate_fields = False
         End If
    End If
         If colnum = 2 Then
        x = CheckPercent(SSDBGCondition.Columns(2).Text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00717")
             MsgBox IIf(msg1 = "", "Value must be between 0 and 100 percent with 2 decimal places.", msg1)
             SSDBGCondition.SetFocus
             SSDBGCondition.Col = 2
            validate_fields = False
         End If
    End If
End If

End Function
Private Sub Form_Load()


'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
FormMode = ChangeModeOfForm(lblStatus, mdVisualization)
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
  

'end copy




   TableLocked = True
    End If
End If

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
    Call translator.Translate_Forms("frm_Condition")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.Condition(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsCONDITION
    
    Set SSDBGCondition.DataSource = deIms
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    
    Caption = Caption + " - " + Tag
       NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
'    NavBar1.EditEnabled = True
'    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
'    NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    SSDBGCondition.AllowUpdate = False
    
    With frm_Condition
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
  


On Error Resume Next
 RecSaved = True
 InUnload = True
 CAncelGrid = False
 SSDBGCondition.Update
 If RecSaved = True Then
    Hide
  '  deIms.rsCONDITION.Update
   ' deIms.rsCONDITION.CancelUpdate
    
    deIms.rsCONDITION.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
Else
    Cancel = True
End If
    
End Sub

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGCondition.Update
End Sub

Private Sub NavBar1_BeforeNewClick()
    SSDBGCondition.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGCondition.AllowUpdate = True
    SSDBGCondition.Columns("active").value = 1
    SSDBGCondition.Columns("active").Text = 1
    SSDBGCondition.SetFocus
    SSDBGCondition.Col = 0
End Sub

Private Sub NavBar1_BeforeSaveClick()
     CAncelGrid = False
     SSDBGCondition.Update
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGCondition.AllowUpdate = False
       End If
End Sub

Private Sub NavBar1_OnCancelClick()
 Dim response As Integer
   If SSDBGCondition.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGCondition.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGCondition.CancelUpdate
       '   SSDBGCondition.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBGCondition.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGCondition.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGCondition.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBGCondition.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBGCondition.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
    SSDBGCondition.Columns("Active").locked = False
'    SSDBGCondition.Refresh
End If
      

End Sub

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
SSDBGCondition.AllowUpdate = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGCondition.Columns("Active").locked = True
SSDBGCondition.SetFocus
SSDBGCondition.Col = 1
SSDBGCondition.AllowUpdate = True

End Sub

Private Sub NavBar1_OnNewClick()
    SSDBGCondition.AllowUpdate = False

End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Condition.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00118") 'J added
        .WindowTitle = IIf(msg1 = "", "Condition", msg1) 'J modified
        Call translator.Translate_Reports("Condition.rpt") 'J added
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

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsCONDITION.Move(0)
    If Err Then Err.Clear

End Sub

Private Sub SSDBGCondition_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBGCondition.AllowUpdate = False
    SSDBGCondition.Columns("Active").locked = False

End If
End Sub

Private Sub SSDBGCondition_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSDBGCondition.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBGCondition.SetFocus
                SSDBGCondition.Columns(ColIndex).Text = oldVALUE
                RecSaved = False
                SSDBGCondition.Col = 0
                GoodColMove = False
              ElseIf CheckDesCode(SSDBGCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBGCondition.SetFocus
                SSDBGCondition.Columns(ColIndex).Text = oldVALUE
                SSDBGCondition.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf SSDBGCondition.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBGCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGCondition.SetFocus
                'SSDBGCondition.Columns(ColIndex).Value =
                RecSaved = False
                SSDBGCondition.Col = 1
               End If
        ElseIf SSDBGCondition.IsAddRow And ColIndex = 2 Then
              If CheckPercent(SSDBGCondition.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00717")
                MsgBox IIf(msg1 = "", "Value must be between 0 and 100 percent with 2 decimal places.", msg1)
                Cancel = 1
                SSDBGCondition.SetFocus
                'SSDBGCondition.Columns(ColIndex).Value =
                RecSaved = False
                SSDBGCondition.Col = 2
               End If
        ElseIf Not SSDBGCondition.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBGCondition.Columns(ColIndex).Text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGCondition.SetFocus
                'SSDBGCondition.Columns(ColIndex).Value =
                RecSaved = False
                SSDBGCondition.Col = 1
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBGCondition.Columns(ColIndex).Text))
     

End Sub

Private Sub SSDBGCondition_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBGCondition.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

Private Sub SSDBGCondition_BeforeUpdate(Cancel As Integer)
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
 
  If SSDBGCondition.IsAddRow Then
      x = NotValidLen(SSDBGCondition.Columns(1).Text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBGCondition.SetFocus
         SSDBGCondition.Col = 0
         Exit Sub
      End If
      x = CheckDesCode(SSDBGCondition.Columns(0).Text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBGCondition.SetFocus
         SSDBGCondition.Col = 0
         SSDBGCondition.Columns(0).Text = ""
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBGCondition.Columns("np").Text = deIms.NameSpace
        If SSDBGCondition.IsAddRow Then
            SSDBGCondition.Columns("create_date").Text = Date
            SSDBGCondition.Columns("create_user").Text = CurrentUser
        End If
        SSDBGCondition.Columns("modify_date").Text = Date
        SSDBGCondition.Columns("modify_user").Text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
    '   SSDBGCondition.CancelUpdate
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
         .CommandText = .CommandText & " From CONDITION "
        .CommandText = .CommandText & " Where cond_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND cond_condcode = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function
Private Function CheckPercent(Code As Variant) As Boolean
On Error Resume Next

If (Code < 0) Or (Code > 100) Then
    CheckPercent = True
Else
    CheckPercent = False
End If
End Function

Private Sub SSDBGCondition_KeyPress(KeyAscii As Integer)
  Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If (Not SSDBGCondition.IsAddRow And SSDBGCondition.Col = 0 And KeyAscii <> 13) Or _
   (Not SSDBGCondition.IsAddRow And SSDBGCondition.Col = 2 And KeyAscii <> 13) Or _
   (Not SSDBGCondition.IsAddRow And SSDBGCondition.Col = 3 And KeyAscii <> 13) Then
        KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGCondition.Col = 3)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGCondition.Col
        If (cur_col = 3) Then
            If GoodColMove = True Then
                SSDBGCondition.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGCondition.Col)
                If good_field Then
                    SSDBGCondition.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub
