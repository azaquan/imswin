VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~2.OCX"
Begin VB.Form frm_Originator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Originator"
   ClientHeight    =   4635
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   7755
   Tag             =   "01010400"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowUpdate     =   0   'False
      AllowCancel     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGOrig 
      Bindings        =   "frm_Originator.frx":0000
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6855
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
      stylesets(0).Picture=   "frm_Originator.frx":0014
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
      stylesets(1).Picture=   "frm_Originator.frx":0030
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
      SelectTypeRow   =   0
      MaxSelectedRows =   0
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   4313
      Columns(0).Caption=   "Name"
      Columns(0).Name =   "Name"
      Columns(0).DataField=   "ori_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   20
      Columns(0).HeadStyleSet=   "colls"
      Columns(1).Width=   3625
      Columns(1).Caption=   "Phone Number"
      Columns(1).Name =   "Phone Number"
      Columns(1).DataField=   "ori_phonnumb"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   20
      Columns(1).HeadStyleSet=   "colls"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NP"
      Columns(2).Name =   "NP"
      Columns(2).DataField=   "ori_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3175
      Columns(3).Caption=   "Active Flag"
      Columns(3).Name =   "Active Flag"
      Columns(3).DataField=   "ori_actvflag"
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(3).HeadStyleSet=   "colls"
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "create_user"
      Columns(4).Name =   "create_user"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_user"
      Columns(5).Name =   "modify_user"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   12091
      _ExtentY        =   5741
      _StockProps     =   79
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
      Left            =   5040
      TabIndex        =   3
      Top             =   3960
      Width           =   2460
   End
   Begin VB.Label lbl_Originator 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Originator"
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
      Left            =   210
      TabIndex        =   0
      Top             =   60
      Width           =   6720
   End
End
Attribute VB_Name = "frm_Originator"
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
If SSDBGOrig.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBGOrig.Columns(colnum).text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGOrig.SetFocus
         SSDBGOrig.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckCodeexist(SSDBGOrig.Columns(0).text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBGOrig.SetFocus
             SSDBGOrig.Col = 0
            validate_fields = False
         End If
    End If
   End If

End Function

Private Function NotValidLen(Code As String) As Boolean

On Error Resume Next
If Len(Trim(Code)) > 0 Then
    NotValidLen = False
Else
    NotValidLen = True
End If
End Function


'set back ground color, get recordset value and populate combo

Private Sub Form_Load()
On Error Resume Next
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

    'Added by Juan (9/13/2000) for Mutilingual
    Call translator.Translate_Forms("frm_Originator")
    '-----------------------------------------

    Screen.MousePointer = vbHourglass
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    SSDBGOrig.DataMember = "ORIGINATOR"
    Call deIms.Originator(deIms.NameSpace)
    
    Set SSDBGOrig.DataSource = deIms
    Set NavBar1.Recordset = deIms.rsORIGINATOR
    Screen.MousePointer = vbDefault
    NavBar1.CloseEnabled = True
    frm_Originator.Caption = frm_Originator.Caption + " - " + frm_Originator.Tag
     NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.Width = 5050
    Call DisableButtons(Me, NavBar1)
    SSDBGOrig.AllowUpdate = False
End Sub

'unload form cancel  recordset uodate

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 InUnload = True
 CAncelGrid = False
  RecSaved = True
 SSDBGOrig.Update
  If RecSaved = True Then
   Hide
   ' SSDBGOrig.Update
   ' deIms.rsORIGINATOR.CancelUpdate
   ' deIms.rsORIGINATOR.UpdateBatch
    deIms.rsORIGINATOR.Close
    If Err Then Err.Clear
     If open_forms <= 5 Then ShowNavigator
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

'cancel record update

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
  '  SSDBGOrig.CancelUpdate
   ' deIms.rsORIGINATOR.CancelUpdate
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGOrig.Update

End Sub

'before save record set create user and name space equal to current user
'and current name space


Private Sub NavBar1_BeforeNewClick()
   ' SSDBGOrig.Update
   ' SSDBGOrig.AddNew
    
   ' deIms.rsORIGINATOR!ori_creauser = CurrentUser
   ' SSDBGOrig.Columns("NP").text = deIms.NameSpace
    SSDBGOrig.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGOrig.AllowUpdate = True
    SSDBGOrig.Columns("Active Flag").text = 1
    SSDBGOrig.SetFocus
    SSDBGOrig.Col = 0
End Sub

'before save check code exist or not, show message

Private Sub NavBar1_BeforeSaveClick()
On Error Resume Next
Dim Numb As Integer
Dim number As Integer
Dim numbe As Integer
    CAncelGrid = False
      SSDBGOrig.Update
      If RecSaved = True Then
            SSDBGOrig.Columns(0).locked = False
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGOrig.AllowUpdate = False
     End If
       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
        
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As Integer
   If SSDBGOrig.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGOrig.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGOrig.CancelUpdate
       '   SSDBGOrig.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBGOrig.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGOrig.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGOrig.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True

    SSDBGOrig.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBGOrig.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
    SSDBGOrig.Columns(0).locked = False
'    SSDBGOrig.Refresh
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


Dim currentformname
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False


   SSDBGOrig.Columns("name").locked = True
   SSDBGOrig.Columns("phone number").locked = True
'   SSDBGOrig.Columns("transaction type").locked = True
   SSDBGOrig.Columns("active flag").locked = True


FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
    Else



SSDBGOrig.AllowUpdate = True
SSDBGOrig.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGOrig.SetFocus
SSDBGOrig.Col = 1
SSDBGOrig.AllowUpdate = True


TableLocked = True
End If

End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
  '  SSDBGOrig.MoveFirst
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
 '   SSDBGOrig.MoveLast
End Sub

'set data grid name space equal to current name space

Private Sub NavBar1_OnNewClick()
    SSDBGOrig.AllowUpdate = False
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
    SSDBGOrig.MoveNext
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
 '   SSDBGOrig.MovePrevious
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Orig.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00070") 'J added
        .WindowTitle = IIf(msg1 = "", "Originator", msg1) 'J modified
        Call translator.Translate_Reports("Orig.rpt") 'J added
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
    Call deIms.rsDestination.Move(0)
    If Err Then Err.Clear
  
'    deIms.rsORIGINATOR!ori_npecode = deIms.NameSpace
'    SSDBGOrig.Update
End Sub

Private Sub SSDBGOrig_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBGOrig.AllowUpdate = False
End If

End Sub

'before save check code exist or not, show message

Private Sub SSDBGOrig_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
      
'   If TransCancelled = False Then
    
    
          If SSDBGOrig.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGOrig.Columns(ColIndex).text) Then
                 msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBGOrig.SetFocus
                SSDBGOrig.Columns(ColIndex).text = oldVALUE
                RecSaved = False
                SSDBGOrig.Col = 0
                GoodColMove = False
              ElseIf CheckCodeexist(SSDBGOrig.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                 Cancel = 1
                SSDBGOrig.SetFocus
                SSDBGOrig.Columns(ColIndex).text = oldVALUE
                RecSaved = False
                SSDBGOrig.Col = 0
                GoodColMove = False
             End If
        
        ElseIf SSDBGOrig.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBGOrig.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGOrig.SetFocus
                'SSDBGOrig.Columns(ColIndex).Text = oldValue
                RecSaved = False
                SSDBGOrig.Col = 1
                GoodColMove = False
               End If
        ElseIf Not SSDBGOrig.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBGOrig.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGOrig.SetFocus
                'SSDBGOrig.Columns(ColIndex).Text = oldValue
                RecSaved = False
                SSDBGOrig.Col = 1
                GoodColMove = False
                End If
         End If
            Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBGOrig.Columns(ColIndex).text))
 '       End If
    
   'Dim oldstr As String
'Dim newstr As String
'Dim Numb As Integer
'Dim number As Integer
'Dim numbe As Integer

'    oldstr = SSDBGOrig.Columns(0).CellText(SSDBGOrig.Bookmark)
'    newstr = SSDBGOrig.Columns(0).text
'
'    Numb = SSDBGOrig.Rows
'    number = SSDBGOrig.GetBookmark(-1)
'    numbe = SSDBGOrig.Bookmark
'
'    If (Numb - number) = 1 And (Numb > numbe) Then
'            Exit Sub
'        Else
 '       If ColIndex = 0 Then
 '           If oldstr <> newstr Then
 '               Cancel = True
 '
 '               'Modified by Juan (9/13/2000) for Multilingual
 '               msg1 = translator.Trans("M00015") 'J added
 '               MsgBox IIf(msg1 = "", " Code cannot be changed once it is saved, Please make new one.", msg1)
 '               '---------------------------------------------
 '
  '          End If
 '       End If
 '   End If
 '
End Sub

Private Sub SSDBGOrig_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean

If SSDBGOrig.Col = 1 And (SSDBGOrig.IsAddRow Or SSDBGOrig.AllowUpdate = True) Then
   
   If Len(Trim(SSDBGOrig.Columns(1).text)) = 0 Then SSDBGOrig.Columns(1).text = "N/A"

End If

    good_field = validate_fields(SSDBGOrig.Col)
    If Not good_field Then
       Cancel = True
    End If
End Sub

Private Sub SSDBGOrig_BeforeUpdate(Cancel As Integer)
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
  If SSDBGOrig.IsAddRow Then
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGOrig.SetFocus
         SSDBGOrig.Col = 0
          Exit Sub
      End If
      x = CheckCodeexist(SSDBGOrig.Columns(0).text)
      If x <> False Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBGOrig.SetFocus
         SSDBGOrig.Columns(0).text = ""
         SSDBGOrig.Col = 0
         RecSaved = False
         Exit Sub
      End If
    '  x = NotValidLen(SSDBGOrig.Columns(1).Text)
      If NotValidLen(SSDBGOrig.Columns(1).text) Then
       msg1 = translator.Trans("M00702")
       MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGOrig.SetFocus
                'SSDBGOrig.Columns(ColIndex).Text = oldValue
                SSDBGOrig.Col = 1
                RecSaved = False
                GoodColMove = False
        Exit Sub
    End If
  End If
End If

   'End If
      
    
  '  Cancel = 0
  'Else
    If InUnload = False Then
         msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
    End If
    
      If (response = vbOK) Or (response = vbYes) Then
       
        SSDBGOrig.Columns("np").text = deIms.NameSpace
        If SSDBGOrig.IsAddRow Then
          '  SSDBGOrig.Columns("create_date").text = Date
            SSDBGOrig.Columns("create_user").text = CurrentUser
        End If
        'SSDBGOrig.Columns("modify_date").Text = Date
        SSDBGOrig.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
   '    SSDBGOrig.CancelUpdate
     Cancel = True
  '   SSDBGOrig.Refresh
   End If
  
'    Cancel = True
'
'    If SSDBGOrig.Columns(0).Text = "" Then
'        MsgBox SSDBGOrig.Columns(0).Caption & " Cannot be left empty": Exit Sub
'
'    ElseIf SSDBGOrig.Columns(1).Text = "" Then
'        MsgBox SSDBGOrig.Columns(1).Caption & " Cannot be left empty": Exit Sub
'    Else
'        Cancel = False
'    End If
End Sub

'SQL statement check code exist

Private Function CheckCodeexist(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From ORIGINATOR"
        .CommandText = .CommandText & " Where ori_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND ori_code = '" & Code & "'"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckCodeexist = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckCodeexist", Err.Description, Err.number, True)
End Function

Private Sub SSDBGOrig_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim good_field As Boolean
'Dim tempcol As Integer

'If KeyCode = 9 Then
'   good_field = validate_fields(SSDBGOrig.Col)
'   tempcol = SSDBGOrig.Col
'   If Not good_field Then
'    SSDBGOrig.Col = tempcol - 1
'   End If
'End If

End Sub

Private Sub SSDBGOrig_KeyPress(KeyAscii As Integer)
  Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean

If Not SSDBGOrig.IsAddRow And SSDBGOrig.Col = 0 And KeyAscii <> 0 Then
             KeyAscii = 0
Else
''    Char = Chr(KeyAscii)
''    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGOrig.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGOrig.Col
        If (cur_col = 2) Then
            If GoodColMove = True Then
                SSDBGOrig.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGOrig.Col)
                If good_field Then
                    SSDBGOrig.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub

