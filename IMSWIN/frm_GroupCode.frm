VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frm_GroupCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group"
   ClientHeight    =   4755
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   8115
   Tag             =   "01011800"
   Visible         =   0   'False
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelEnabled   =   0   'False
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgGroup 
      Bindings        =   "frm_GroupCode.frx":0000
      Height          =   3435
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   7095
      _Version        =   196617
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
      stylesets(0).Picture=   "frm_GroupCode.frx":0014
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
      stylesets(1).Picture=   "frm_GroupCode.frx":0030
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      HeadStyleSet    =   "ColHeader"
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   5292
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "grp_npecode"
      Columns(0).Name =   "np"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "grp_npecode"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1614
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "grp_code"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   4
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   6906
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "grp_desc"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   40
      Columns(2).HeadStyleSet=   "ColHeader"
      Columns(2).StyleSet=   "RowFont"
      Columns(3).Width=   1773
      Columns(3).Caption=   "Reorder"
      Columns(3).Name =   "Reorder"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "grp_reorproc"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "ColHeader"
      Columns(3).StyleSet=   "RowFont"
      Columns(4).Width=   1508
      Columns(4).Caption=   "Active"
      Columns(4).Name =   "Active"
      Columns(4).DataField=   "grp_actvflag"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(4).HeadStyleSet=   "ColHeader"
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_date"
      Columns(5).Name =   "modify_date"
      Columns(5).DataField=   "grp_modidate"
      Columns(5).DataType=   135
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "modify_user"
      Columns(6).Name =   "modify_user"
      Columns(6).DataField=   "grp_modiuser"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create_date"
      Columns(7).Name =   "create_date"
      Columns(7).DataField=   "grp_creadate"
      Columns(7).DataType=   135
      Columns(7).FieldLen=   256
      Columns(8).Width=   5292
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "create_user"
      Columns(8).Name =   "create_user"
      Columns(8).DataField=   "grp_creauser"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   12515
      _ExtentY        =   6059
      _StockProps     =   79
      BackColor       =   -2147483643
      DataMember      =   "GROUPE"
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
   Begin VB.PictureBox VisM1 
      Height          =   480
      Left            =   480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   5040
      Width           =   1200
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
      Left            =   4080
      TabIndex        =   4
      Top             =   4080
      Width           =   2460
   End
   Begin VB.Label lbl_GroupCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code"
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
      Width           =   5055
   End
End
Attribute VB_Name = "frm_GroupCode"
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

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If ssdbgGroup.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(ssdbgGroup.Columns(colnum).Value)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         ssdbgGroup.SetFocus
         ssdbgGroup.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(ssdbgGroup.Columns(0).Value)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             ssdbgGroup.SetFocus
             ssdbgGroup.Col = 0
            validate_fields = False
         End If
    End If
   End If

End Function
' unload form free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 RecSaved = True
 InUnload = True
 ssdbgGroup.Update
 If RecSaved = True Then
   
    Hide
    deIms.rsGROUPE.Close
    
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
    ssdbgGroup.Update

End Sub

Private Sub NavBar1_BeforeNewClick()
   ssdbgGroup.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    ssdbgGroup.AllowUpdate = True
    ssdbgGroup.Columns("active").Value = 1
    ssdbgGroup.SetFocus
    ssdbgGroup.Col = 0

End Sub

Private Sub NavBar1_BeforeSaveClick()
     ssdbgGroup.Update
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            ssdbgGroup.AllowUpdate = False
       End If

End Sub

'cancel record set update

Private Sub NavBar1_OnCancelClick()
 Dim response As Integer
   If ssdbgGroup.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           ssdbgGroup.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          ssdbgGroup.CancelUpdate
       '   ssdbgGroup.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          ssdbgGroup.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     ssdbgGroup.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    ssdbgGroup.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    ssdbgGroup.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    ssdbgGroup.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    ssdbgGroup.Refresh
End If
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'move record set to first position

Private Sub NavBar1_OnFirstClick()
    ssdbgGroup.MoveFirst
    
    Call EnableNav(True, False)

End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    ssdbgGroup.MoveLast
    Call EnableNav(False, True)
End Sub

'set data grid name space equal to current name space

Private Sub NavBar1_OnNewClick()
    ssdbgGroup.AllowUpdate = False
End Sub


'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Groupe.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00057")
        .WindowTitle = IIf(msg1 = "", "Group Code", msg1)
        Call translator.Translate_Reports("Groupe.rpt") 'J added
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

'load form,populate combo data, set controls

Private Sub Form_Load()
On Error Resume Next

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
    Call translator.Translate_Forms("frm_GroupCode")
    '------------------------------------------

    Dim li_x As Integer
    For li_x = 0 To (Controls.Count - 1)
        Call gsb_fade_to_black(Controls(li_x))
    Next li_x
 
    Visible = True
    Screen.MousePointer = vbDefault
 
    Call deIms.GROUPE(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsGROUPE
    Set ssdbgGroup.DataSource = deIms
    Call DisableButtons(Me, NavBar1)
    
    Caption = Caption + " - " + Tag
     NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.CloseEnabled = True
    ssdbgGroup.AllowUpdate = False
    
    
End Sub

'save record set data

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsDestination.Move(0)
    If Err Then Err.Clear
End Sub

'set navbar button

Public Sub EnableNav(bNext As Boolean, bPrior As Boolean)
    NavBar1.LastEnabled = bNext
    NavBar1.LastEnabled = bNext
    NavBar1.FirstEnabled = bPrior
    NavBar1.PreviousEnabled = bPrior
End Sub

Private Sub ssdbgGroup_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    ssdbgGroup.AllowUpdate = False
End If
End Sub

Private Sub ssdbgGroup_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If ssdbgGroup.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(ssdbgGroup.Columns(ColIndex).Value) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                ssdbgGroup.SetFocus
                ssdbgGroup.Columns(ColIndex).Value = oldVALUE
                ssdbgGroup.Col = 0
                GoodColMove = False
              ElseIf CheckDesCode(ssdbgGroup.Columns(ColIndex).Value) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                ssdbgGroup.SetFocus
                ssdbgGroup.Columns(ColIndex).Value = oldVALUE
                ssdbgGroup.Col = 0
                GoodColMove = False
             End If
        
        ElseIf ssdbgGroup.IsAddRow And ColIndex = 1 Then
              If NotValidLen(ssdbgGroup.Columns(ColIndex).Value) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                ssdbgGroup.SetFocus
                'ssdbgGroup.Columns(ColIndex).Value =
                ssdbgGroup.Col = 0
               End If
        ElseIf Not ssdbgGroup.IsAddRow And ColIndex = 1 Then
                If NotValidLen(ssdbgGroup.Columns(ColIndex).Value) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                ssdbgGroup.SetFocus
                'ssdbgGroup.Columns(ColIndex).Value =
                ssdbgGroup.Col = 0
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(ssdbgGroup.Columns(ColIndex).Value))
     
End Sub

Private Sub ssdbgGroup_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(ssdbgGroup.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

Private Sub ssdbgGroup_BeforeUpdate(Cancel As Integer)
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
 If (InUnload = False) Or (response = vbOK) Then
 
  If ssdbgGroup.IsAddRow Then
      x = NotValidLen(ssdbgGroup.Columns(1).Value)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  ssdbgGroup.SetFocus
         ssdbgGroup.Col = 0
         Exit Sub
      End If
      x = CheckDesCode(ssdbgGroup.Columns(0).Value)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         ssdbgGroup.SetFocus
         ssdbgGroup.Col = 0
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        ssdbgGroup.Columns("np").text = deIms.NameSpace
        If ssdbgGroup.IsAddRow Then
            ssdbgGroup.Columns("create_date").text = Date
            ssdbgGroup.Columns("create_user").text = CurrentUser
        End If
        ssdbgGroup.Columns("modify_date").Value = Date
        ssdbgGroup.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
       ssdbgGroup.CancelUpdate
     Cancel = True
   End If
  
End Sub

Private Sub ssdbgGroup_BtnClick()

End Sub
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From GROUPE "
        .CommandText = .CommandText & " Where grp_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND grp_code = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function
Private Sub ssdbgGroup_InitColumnProps()

End Sub

Private Sub ssdbgGroup_KeyPress(KeyAscii As Integer)
Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If Not ssdbgGroup.IsAddRow And ssdbgGroup.Col = 0 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (ssdbgGroup.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = ssdbgGroup.Col
        If (cur_col = 3) Then
            If GoodColMove = True Then
                ssdbgGroup.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(ssdbgGroup.Col)
                If good_field Then
                    ssdbgGroup.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub

Private Sub VisM1_Click()

End Sub
