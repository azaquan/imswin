VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~2.OCX"
Begin VB.Form frm_SubLocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Location"
   ClientHeight    =   4515
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   7365
   Tag             =   "01030600"
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGSubLocation 
      Bindings        =   "frm_Sublocation.frx":0000
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6675
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
      stylesets(0).Picture=   "frm_Sublocation.frx":0014
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
      stylesets(1).Picture=   "frm_Sublocation.frx":0030
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
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
      Columns.Count   =   8
      Columns(0).Width=   2461
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "sb_code"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   10
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(1).Width=   6826
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "sb_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   30
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "sb_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   5
      Columns(3).Width=   1720
      Columns(3).Caption=   "Active"
      Columns(3).Name =   "Active"
      Columns(3).DataField=   "sb_actvflag"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "ColHeader"
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "modify_date"
      Columns(4).Name =   "modify_date"
      Columns(4).DataField=   "sb_modidate"
      Columns(4).DataType=   135
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_user"
      Columns(5).Name =   "modify_user"
      Columns(5).DataField=   "sb_modiuser"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "create_date"
      Columns(6).Name =   "create_date"
      Columns(6).DataField=   "sb_creadate"
      Columns(6).DataType=   135
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "create_user"
      Columns(7).Name =   "create_user"
      Columns(7).DataField=   "sb_creauser"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   11774
      _ExtentY        =   5741
      _StockProps     =   79
      DataMember      =   "SUBLOCATION"
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
      Left            =   240
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
      DeleteToolTipText=   ""
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
      Top             =   3840
      Width           =   2460
   End
   Begin VB.Label lbl_SubLocation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Location"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   255
      TabIndex        =   0
      Top             =   60
      Width           =   6525
   End
End
Attribute VB_Name = "frm_SubLocation"
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
If SSDBGSubLocation.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBGSubLocation.Columns(colnum).text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGSubLocation.SetFocus
         SSDBGSubLocation.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(SSDBGSubLocation.Columns(0).text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBGSubLocation.SetFocus
             SSDBGSubLocation.Col = 0
            SSDBGSubLocation.Columns(0).text = ""
           validate_fields = False
         End If
    End If
   End If

End Function
'load form get data for data grid

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

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_SubLocation")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.SubLocation(deIms.NameSpace)
     Set NavBar1.Recordset = deIms.rsSUBLOCATION
   Set SSDBGSubLocation.DataSource = deIms
    
    Screen.MousePointer = vbDefault
    
    Caption = Caption + " - " + Tag
    
    
       NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
  '  NavBar1.CloseEnabled = True
    Call DisableButtons(Me, NavBar1)
    NavBar1.Width = 5050
    SSDBGSubLocation.AllowUpdate = False
   
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
Dim response As String
On Error Resume Next
 RecSaved = True
 InUnload = True
 CAncelGrid = False
 SSDBGSubLocation.Update
 If RecSaved = True Then

    Hide
'    deIms.rsSUBLOCATION.Update
  '  deIms.rsSUBLOCATION.CancelUpdate
    
    deIms.rsSUBLOCATION.Close
    Set frm_SubLocation = Nothing
    If open_forms <= 5 Then ShowNavigator
   If Err Then Err.Clear
    
Else
    Cancel = True
End If


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If




End Sub

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True

End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGSubLocation.Update

End Sub

Private Sub NavBar1_BeforeNewClick()
   SSDBGSubLocation.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGSubLocation.AllowUpdate = True
    SSDBGSubLocation.Columns("active").text = 1
    SSDBGSubLocation.SetFocus
    SSDBGSubLocation.Col = 0

End Sub

'before save check location code exist or not
'show message

Private Sub NavBar1_BeforeSaveClick()
    CAncelGrid = False
     SSDBGSubLocation.Update
        If RecSaved = True Then
            SSDBGSubLocation.Columns(0).locked = False
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGSubLocation.AllowUpdate = False
       End If
       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
       
  
End Sub

'cancel update

Private Sub NavBar1_OnCancelClick()
 Dim response As Integer
   If SSDBGSubLocation.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGSubLocation.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGSubLocation.CancelUpdate
       '   SSDBGSubLocation.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBGSubLocation.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGSubLocation.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGSubLocation.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBGSubLocation.CancelUpdate
    SSDBGSubLocation.Columns(0).locked = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBGSubLocation.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBGSubLocation.Refresh
End If


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
 


End Sub

'close form

Private Sub NavBar1_OnCloseClick()

If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
        
    
    
    
    Unload Me
End Sub


Private Sub NavBar1_OnEditClick()

'
''copy begin here
'
'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)


   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode



   SSDBGSubLocation.Columns("code").locked = True
   SSDBGSubLocation.Columns("description").locked = True
   SSDBGSubLocation.Columns("active").locked = True
NavBar1.CancelEnabled = False
NavBar1.NewEnabled = False
NavBar1.SaveEnabled = False

Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

SSDBGSubLocation.AllowUpdate = True
SSDBGSubLocation.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGSubLocation.SetFocus
SSDBGSubLocation.Col = 1
SSDBGSubLocation.AllowUpdate = True


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

'set data grid user equal to current user and name space equal to
'current name space

Private Sub NavBar1_OnNewClick()
    SSDBGSubLocation.AllowUpdate = False
End Sub




'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Subloc.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00117") 'J added
        .WindowTitle = IIf(msg1 = "", "Sub Location", msg1) 'J modified
        Call translator.Translate_Reports("Subloc.rpt") 'J added
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
    Call deIms.rsSUBLOCATION.Move(0)
    If Err Then Err.Clear
    
End Sub

Private Sub SSDBGSubLocation_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBGSubLocation.AllowUpdate = False
End If
End Sub

'before save data check code changed or not if code changed resize
'data grid and show message

Private Sub SSDBGSubLocation_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSDBGSubLocation.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGSubLocation.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBGSubLocation.SetFocus
                SSDBGSubLocation.Columns(ColIndex).text = oldVALUE
                SSDBGSubLocation.Col = 0
                RecSaved = False
                GoodColMove = False
              ElseIf CheckDesCode(SSDBGSubLocation.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBGSubLocation.SetFocus
                SSDBGSubLocation.Columns(ColIndex).text = oldVALUE
                SSDBGSubLocation.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf SSDBGSubLocation.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBGSubLocation.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGSubLocation.SetFocus
                'SSDBGSubLocation.Columns(ColIndex).Value =
                RecSaved = False
                SSDBGSubLocation.Col = 1
               End If
        ElseIf Not SSDBGSubLocation.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBGSubLocation.Columns(ColIndex).text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGSubLocation.SetFocus
                'SSDBGSubLocation.Columns(ColIndex).Value =
                RecSaved = False
                SSDBGSubLocation.Col = 1
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBGSubLocation.Columns(ColIndex).text))
     

End Sub

Private Sub SSDBGSubLocation_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBGSubLocation.Col)
    If Not good_field Then
       Cancel = True
    End If
End Sub

'set data grid  name space equal to current name space

Private Sub SSDBGSubLocation_BeforeUpdate(Cancel As Integer)
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
 
  If SSDBGSubLocation.IsAddRow Then
      x = NotValidLen(SSDBGSubLocation.Columns(1).text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBGSubLocation.SetFocus
         SSDBGSubLocation.Col = 1
         Exit Sub
      End If
      x = CheckDesCode(SSDBGSubLocation.Columns(0).text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBGSubLocation.SetFocus
         SSDBGSubLocation.Col = 0
         SSDBGSubLocation.Columns(0).text = ""
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBGSubLocation.Columns("np").text = deIms.NameSpace
        If SSDBGSubLocation.IsAddRow Then
            SSDBGSubLocation.Columns("create_date").text = Date
            SSDBGSubLocation.Columns("create_user").text = CurrentUser
        End If
        SSDBGSubLocation.Columns("modify_date").text = Date
        SSDBGSubLocation.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
     '  SSDBGSubLocation.CancelUpdate
     Cancel = True
   End If
  End Sub

'SQL statement check exist sub location
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From sublocation "
        .CommandText = .CommandText & " Where sb_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND sb_code = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function


Private Function CheckLocationexist(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT ? = count(*) "
        .CommandText = .CommandText & " From sublocation"
        .CommandText = .CommandText & " Where sb_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sb_code = '" & Code & "'"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckLocationexist = .Parameters(0)
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckLocationexist", Err.Description, Err.number, True)

End Function

Private Sub SSDBGSubLocation_KeyPress(KeyAscii As Integer)
Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If (Not SSDBGSubLocation.IsAddRow And SSDBGSubLocation.Col = 0 And KeyAscii <> 13) Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGSubLocation.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGSubLocation.Col
        If (cur_col = 2) Then
            If GoodColMove = True Then
                SSDBGSubLocation.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGSubLocation.Col)
                If good_field Then
                    SSDBGSubLocation.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub
