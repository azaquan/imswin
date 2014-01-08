VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Destination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destination"
   ClientHeight    =   4485
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7365
   Tag             =   "01020400"
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGDestination 
      Bindings        =   "frm_Destination.frx":0000
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5055
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
      stylesets(0).Picture=   "frm_Destination.frx":0014
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
      stylesets(1).Picture=   "frm_Destination.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorHighlight=   16777215
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
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
      Columns.Count   =   8
      Columns(0).Width=   1799
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "des_destcode"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   3
      Columns(0).HeadStyleSet=   "Colls"
      Columns(0).StyleSet=   "Rows"
      Columns(1).Width=   4419
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "des_destname"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   20
      Columns(1).HeadStyleSet=   "Colls"
      Columns(1).StyleSet=   "Rows"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "des_npecode"
      Columns(2).FieldLen=   256
      Columns(3).Width=   5292
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "create_date"
      Columns(3).Name =   "create_date"
      Columns(3).DataField=   "des_creadate"
      Columns(3).DataType=   135
      Columns(3).FieldLen=   256
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "create_user"
      Columns(4).Name =   "create_user"
      Columns(4).DataField=   "des_creauser"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "modify_date"
      Columns(5).Name =   "modify_date"
      Columns(5).DataField=   "des_modidate"
      Columns(5).DataType=   135
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "modify_user"
      Columns(6).Name =   "modify_user"
      Columns(6).DataField=   "des_modiuser"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1773
      Columns(7).Caption=   "Active"
      Columns(7).Name =   "active"
      Columns(7).DataField=   "des_active"
      Columns(7).DataType=   11
      Columns(7).FieldLen=   256
      Columns(7).Style=   2
      Columns(7).HeadStyleSet=   "Colls"
      Columns(7).StyleSet=   "Rows"
      _ExtentX        =   8916
      _ExtentY        =   5318
      _StockProps     =   79
      BackColor       =   -2147483638
      DataMember      =   "DESTINATION"
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
      Left            =   360
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
      Left            =   4560
      TabIndex        =   3
      Top             =   3840
      Width           =   2460
   End
   Begin VB.Label lbl_ServiceCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frm_Destination"
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

'load form,populate combo data,set navbar button
Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBGDestination.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBGDestination.Columns(colnum).Text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBGDestination.SetFocus
         SSDBGDestination.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(SSDBGDestination.Columns(0).Text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBGDestination.SetFocus
             SSDBGDestination.Col = 0
             SSDBGDestination.Columns(0).Text = ""
            validate_fields = False
         End If
    End If
   End If

End Function

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
    Call translator.Translate_Forms("frm_Destination")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.Destination(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsDestination
    
    
    Visible = True
    Screen.MousePointer = vbDefault
    'SSDBGDestination.BatchUpdate = True
    Set SSDBGDestination.DataSource = deIms
    
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
    SSDBGDestination.AllowUpdate = False
    Call DisableButtons(Me, NavBar1)
    
    'Call DisableButtons(Me, NavBar1)
    
    With frm_Destination
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

' unload form,free memory

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
 SSDBGDestination.Update
 If RecSaved = True Then
    Hide
    deIms.rsDestination.Close
    
    If open_forms <= 5 Then ShowNavigator
    
    If Err Then Err.Clear
    
Else
    Cancel = True
End If
    
    
 'smm   If Not SSDBGDestination.IsAddRow And DidFieldChange(SSDBGDestination.Columns(1).text, SSDBGDestination.Columns(1).CellText(SSDBGDestination.Bookmark)) Then 'Or SSDBGDestination.IsAddRow Then
          'response = MsgBox("Do you want to Save the Changes", vbOKCancel, "Imswin")

          'If response = vbOK Then
  'smm            SSDBGDestination.Update
          'Else
          '    CAncelGrid = True
          '    SSDBGDestination.CancelUpdate
          'End If
          
  'smm   ElseIf SSDBGDestination.IsAddRow Then
    'smm         Call SSDBGDestination_BeforeUpdate(0)
      'smm     End If
            
       
    
    
   'smm  deIms.rsDestination.CancelBatch
    
End Sub

'cancel recordset update

Private Sub NavBar1_BeforeCancelClick()
   CAncelGrid = True
'''''If SSDBGDestination.IsAddRow Then SSDBGDestination.CancelUpdate
'''''    SSDBGDestination.Refresh
End Sub

'set record sset update

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGDestination.Update
    
End Sub

'set recordset add new

Private Sub NavBar1_BeforeNewClick()
    SSDBGDestination.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBGDestination.AllowUpdate = True
    SSDBGDestination.Columns("active").Text = 1
    SSDBGDestination.SetFocus
    SSDBGDestination.Col = 0
End Sub

'before save records set record update

Private Sub NavBar1_BeforeSaveClick()
    CAncelGrid = False
     SSDBGDestination.Update
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            SSDBGDestination.Columns(0).locked = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBGDestination.AllowUpdate = False
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
   If SSDBGDestination.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           SSDBGDestination.CancelUpdate
        ' Cancel = -1
          CAncelGrid = True
          SSDBGDestination.CancelUpdate
       '   SSDBGDestination.Refresh
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          SSDBGDestination.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
     '     SSDBGDestination.Refresh
    Else
        CAncelGrid = False
    End If
Else
  '  CAncelGrid = True
    SSDBGDestination.CancelUpdate
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBGDestination.CancelUpdate
    SSDBGDestination.Columns(0).locked = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBGDestination.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBGDestination.Refresh
End If
'
'If TableLocked = True Then    'jawdat
'Dim imsLock As imsLock.lock
'Set imsLock = New imsLock.lock
'currentformname = Forms(3).Name
'Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
'End If

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

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)


   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode

NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

SSDBGDestination.AllowUpdate = False

FormMode = ChangeModeOfForm(lblStatus, mdvisualization)

Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

'
''copy begin here
'
'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

'SSDBGDestination.Refresh
SSDBGDestination.AllowUpdate = True
SSDBGDestination.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBGDestination.SetFocus
SSDBGDestination.Col = 1
SSDBGDestination.AllowUpdate = True




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

'set name space equal to current name space

Private Sub NavBar1_OnNewClick()
  
    SSDBGDestination.AllowUpdate = False
    'deIms.rsDestination!des_npecode = deIms.NameSpace
End Sub

'get crystal report paramenter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Destination.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00053") 'J added
        .WindowTitle = IIf(msg1 = "", "Destination", msg1) 'J modified
        Call translator.Translate_Reports("Destination.rpt") 'J added
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

'save record set

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsDestination.Move(0)
    If Err Then Err.Clear
    
    
End Sub

Private Sub SSDBGDestination_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    SSDBGDestination.AllowUpdate = False
End If
'If CAncelGrid = False Then MsgBox "Changes Saved"
 ' SSDBGDestination.Move (0)
'  SSDBGDestination.Refresh
End Sub

Private Sub SSDBGDestination_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)

Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSDBGDestination.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGDestination.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBGDestination.SetFocus
                SSDBGDestination.Columns(ColIndex).Text = oldVALUE
                SSDBGDestination.Col = 0
                RecSaved = False
                GoodColMove = False
              ElseIf CheckDesCode(SSDBGDestination.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBGDestination.SetFocus
                SSDBGDestination.Columns(ColIndex).Text = oldVALUE
                SSDBGDestination.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf SSDBGDestination.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBGDestination.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGDestination.SetFocus
                'SSDBGDestination.Columns(ColIndex).Text =
                RecSaved = False
                SSDBGDestination.Col = 1
               End If
        ElseIf Not SSDBGDestination.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBGDestination.Columns(ColIndex).Text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBGDestination.SetFocus
                'SSDBGDestination.Columns(ColIndex).Text =
                RecSaved = False
                SSDBGDestination.Col = 1
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBGDestination.Columns(ColIndex).Text))
     
        
End Sub


Private Sub SSDBGDestination_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBGDestination.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

'set data grip value to current name space

Private Sub SSDBGDestination_BeforeUpdate(Cancel As Integer)
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
 
  If SSDBGDestination.IsAddRow Then
      x = NotValidLen(SSDBGDestination.Columns(1).Text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSDBGDestination.SetFocus
         SSDBGDestination.Col = 1
         Exit Sub
      End If
      x = CheckDesCode(SSDBGDestination.Columns(0).Text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSDBGDestination.SetFocus
         SSDBGDestination.Col = 0
         SSDBGDestination.Columns(0).Text = ""
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        
        SSDBGDestination.Columns("np").Text = deIms.NameSpace
        If SSDBGDestination.IsAddRow Then
            SSDBGDestination.Columns("create_date").Text = Date
            SSDBGDestination.Columns("create_user").Text = CurrentUser
        End If
        SSDBGDestination.Columns("modify_date").Text = Date
        SSDBGDestination.Columns("modify_user").Text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
       'SSDBGDestination.CancelUpdate
     Cancel = True
   End If
  
End Sub


Private Function NotValidLen(Code As String) As Boolean

On Error Resume Next
If Len(Trim(Code)) > 0 Then
    NotValidLen = False
Else
    NotValidLen = True
End If
End Function


'Added 11/20/00 by S. McMorrow to check for duplicate key valuePrivate Function CheckDesCode(Code As String) As Boolean
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From destination "
        .CommandText = .CommandText & " Where des_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND des_destcode = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function

Private Function DidFieldChange(strOldValue As String, strNewValue As String)
Dim ret
    ret = StrComp(Trim(strOldValue), Trim(strNewValue), vbTextCompare)
            If ret <> 0 Then
                DidFieldChange = True
            Else
                DidFieldChange = False
            End If

End Function

Private Sub SSDBGDestination_KeyPress(KeyAscii As Integer)
  Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If Not SSDBGDestination.IsAddRow And SSDBGDestination.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBGDestination.Col = 2)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBGDestination.Col
        If (cur_col = 2) Then
            If GoodColMove = True Then
                SSDBGDestination.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(SSDBGDestination.Col)
                If good_field Then
                    SSDBGDestination.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub

