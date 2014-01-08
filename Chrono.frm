VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmChrono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Numbering"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   8640
   Tag             =   "01040400"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   600
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
      AllowDelete     =   0   'False
      Mode            =   0
      CommandType     =   0
      CursorLocation  =   0
      CommandType     =   0
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgKeys 
      Bindings        =   "Chrono.frx":0000
      Height          =   3135
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   8070
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "Chrono.frx":0014
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "Chrono.frx":0030
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
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
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   12
      Columns(0).Width=   1270
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "chr_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "Field1"
      Columns(1).Name =   "Field1"
      Columns(1).DataField=   "chr_fld1"
      Columns(1).DataType=   8
      Columns(1).Case =   2
      Columns(1).FieldLen=   256
      Columns(2).Width=   2037
      Columns(2).Caption=   "Field2"
      Columns(2).Name =   "Field2"
      Columns(2).DataField=   "chr_fld2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1879
      Columns(3).Caption=   "Field3"
      Columns(3).Name =   "Field3"
      Columns(3).DataField=   "chr_fld3"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   256
      Columns(4).Width=   2011
      Columns(4).Caption=   "Field4"
      Columns(4).Name =   "Field4"
      Columns(4).DataField=   "chr_fld4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1905
      Columns(5).Caption=   "Field5"
      Columns(5).Name =   "Field5"
      Columns(5).DataField=   "chr_fld5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2170
      Columns(6).Caption=   "Field6"
      Columns(6).Name =   "Field6"
      Columns(6).DataField=   "chr_fld6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "modify_date"
      Columns(7).Name =   "modify_date"
      Columns(7).DataField=   "chr_modidate"
      Columns(7).DataType=   135
      Columns(7).FieldLen=   256
      Columns(8).Width=   5292
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "modify_user"
      Columns(8).Name =   "modify_user"
      Columns(8).DataField=   "chr_modiuser"
      Columns(8).DataType=   8
      Columns(8).Case =   2
      Columns(8).FieldLen=   256
      Columns(9).Width=   5292
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "create_date"
      Columns(9).Name =   "create_date"
      Columns(9).DataField=   "chr_creadate"
      Columns(9).DataType=   135
      Columns(9).FieldLen=   256
      Columns(9).Nullable=   2
      Columns(10).Width=   5292
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "create_user"
      Columns(10).Name=   "create_user"
      Columns(10).DataField=   "chr_creauser"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   5292
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "np"
      Columns(11).Name=   "np"
      Columns(11).DataField=   "chr_npecode"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      _ExtentX        =   14235
      _ExtentY        =   5530
      _StockProps     =   79
      DataMember      =   "CHRONO"
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
   Begin LRNavigators.LROleDBNavBar LROleDBNavBar11 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      FirstEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowDelete     =   0   'False
      DeleteVisible   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Auto Numbering"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7815
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
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   2460
   End
End
Attribute VB_Name = "frmChrono"
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
Private Function RemoveNulls(ByRef val As Variant) As String
Dim strVarType As String
Dim intMyComp As Integer

strVarType = TypeName(val)
intMyComp = StrComp(val, "", 1)
If intMyComp = 0 Then
    val = "     "
End If
RemoveNulls = val
End Function

Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
      '  .CommandText = "SELECT count(*) RT"
      '   .CommandText = .CommandText & " From CHRONO "
      '  .CommandText = .CommandText & " Where chr_npecode = '" & deIms.NameSpace & "'"
      ' .CommandText = .CommandText & " AND chr_code = '" & Code & "'"
      .CommandText = " select count(*) RT from chrono where chr_npecode = '" & deIms.NameSpace & "'  AND chr_code = '" & Code & "' and chr_fld1 ='" & Trim(ssdbgKeys.Columns(2).value) & "'"
        Set rst = .Execute
        CheckDesCode = rst!rt
End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function
Private Function CheckInt(Code As Variant) As Boolean
On Error Resume Next
Dim strVarType As String
Dim intMyComp As Integer

strVarType = TypeName(Code)
intMyComp = StrComp(strVarType, "Integer", 1)
If intMyComp = 0 Then
    CheckInt = True
Else
    CheckInt = False
End If
End Function
Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If ssdbgKeys.IsAddRow Then
   If colnum = 0 Or colnum = 3 Then
   'Line Commented out by Muzammil,Since the Project would not run.
      x = NotValidLen(ssdbgKeys.Columns(colnum).Text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
              MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         ssdbgKeys.SetFocus
         ssdbgKeys.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
        x = CheckDesCode(ssdbgKeys.Columns(0).Text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             ssdbgKeys.SetFocus
             ssdbgKeys.Columns(0).Text = ""
             ssdbgKeys.Col = 0
            validate_fields = False
         End If
    End If
    
   If colnum = 1 Then
        x = DoesCombinationExists
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             ssdbgKeys.SetFocus
            ' ssdbgKeys.Columns(1).text = ""
             ssdbgKeys.Col = 1
            validate_fields = False
         End If
    End If
    
    'If colnum = 0 Or colnum = 1 Or colnum = 2 Or colnum = 4 Or colnum = 5 Or colnum = 6 Then
    
   '   Select Case conumb
    '
    '  Case 0
      
            
      
     ' Case 1
    '
    '  Case 2
      
    '  Case 3
      
    '  Case 4
      
    '  Case 5
      
    '  Case 6
      
     ' End Select
    
  '    if len(ssdbgKeys.Columns(colnumb).text) =
    '
      '  If colnum = 3 Then
      '  x = CheckInt(ssdbgKeys.Columns(3).text)
      '  If x = False Then
      '       RecSaved = False
      '       msg1 = translator.Trans("M00718")
      '       MsgBox IIf(msg1 = "", "Value must be and Integer", msg1)
      '       ssdbgKeys.SetFocus
      '       ssdbgKeys.Col = 3
      '       ssdbgKeys.Columns(3).text = ""
      '      validate_fields = False
      '   End If
      ' End If
    End If

End Function

'get data and set navbar buttom

Private Sub Form_Load()
Dim ctl As Control

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
    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmChrono")
    '------------------------------------------
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor

    For Each ctl In Controls
        gsb_fade_to_black (ctl)
    Next ctl
    
   deIms.rsCHRONO.Close
    If Err Then Err.Clear
    
    Call deIms.CHRONO(deIms.NameSpace)
    
    ssdbgKeys.DataMember = "CHRONO"
    Set NavBar1.Recordset = deIms.rsCHRONO
    Set ssdbgKeys.DataSource = deIms
    Screen.MousePointer = vbDefault
    
    
    frmChrono.Caption = frmChrono.Caption + " - " + frmChrono.Tag
    NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
  '  NavBar1.FirstEnabled = True
  '  NavBar1.FirstVisible = True
  '  NavBar1.LastEnabled = True
  '  NavBar1.LastVisible = True
    NavBar1.NextEnabled = True
    NavBar1.NextVisible = True
    NavBar1.PreviousEnabled = True
    NavBar1.PreviousVisible = True
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    ssdbgKeys.AllowUpdate = False
    
    
    Call DisableButtons(Me, NavBar1)
    
    ssdbgKeys.Columns(0).FieldLen = 3
    ssdbgKeys.Columns(1).FieldLen = 10
    ssdbgKeys.Columns(2).FieldLen = 10
    ssdbgKeys.Columns(3).FieldLen = 4
    ssdbgKeys.Columns(4).FieldLen = 10
    ssdbgKeys.Columns(5).FieldLen = 10
    ssdbgKeys.Columns(6).FieldLen = 10
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
InUnload = True
  RecSaved = True
  CAncelGrid = False
ssdbgKeys.Update
 If RecSaved = True Then
    Hide
    deIms.rsCHRONO.Close
    
    If Err Then Err.Clear
    Set frmChrono = Nothing
    If open_forms <= 5 Then ShowNavigator
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
   CAncelGrid = True

End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    ssdbgKeys.Update

End Sub

Private Sub NavBar1_BeforeNewClick()
   ssdbgKeys.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    ssdbgKeys.AllowUpdate = True
' there is no active field in this table
'    ssdbgKeys.Columns("active").text = 1
    ssdbgKeys.SetFocus
    ssdbgKeys.Col = 0

End Sub

Private Sub NavBar1_BeforeSaveClick()
   CAncelGrid = False
    ssdbgKeys.Update
      If RecSaved = True Then
            NavBar1.SaveEnabled = False
            ssdbgKeys.Columns(0).locked = False
            ssdbgKeys.Columns(3).locked = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            ssdbgKeys.AllowUpdate = False
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
   If ssdbgKeys.IsAddRow Then
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbOKCancel, "Imswin")
      If response = vbOK Then
          CAncelGrid = True
           ssdbgKeys.CancelUpdate
          CAncelGrid = True
          ssdbgKeys.CancelUpdate
          NavBar1.EditEnabled = True
          NavBar1.NewEnabled = True
          NavBar1.CancelEnabled = False
          NavBar1.SaveEnabled = False
          ssdbgKeys.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
          lblStatus.Caption = Visualize
    Else
        CAncelGrid = False
    End If
Else
    ssdbgKeys.CancelUpdate
    ssdbgKeys.Columns(0).locked = False
    ssdbgKeys.Columns(3).locked = False
    ssdbgKeys.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    ssdbgKeys.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize

End If


If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
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
End Sub

Private Sub NavBar1_OnEditClick()


'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False


ssdbgKeys.Columns("code").locked = True
ssdbgKeys.Columns("field1").locked = True
ssdbgKeys.Columns("field2").locked = True
ssdbgKeys.Columns("field3").locked = True
ssdbgKeys.Columns("field4").locked = True
ssdbgKeys.Columns("field5").locked = True
ssdbgKeys.Columns("field6").locked = True



FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
  Exit Sub
    Else




ssdbgKeys.AllowUpdate = True
ssdbgKeys.Columns(0).locked = True
ssdbgKeys.Columns(3).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
ssdbgKeys.SetFocus
ssdbgKeys.Col = 1
ssdbgKeys.AllowUpdate = True

TableLocked = True
End If

End Sub





'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Chrono.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00332") 'J added
        .WindowTitle = IIf(msg1 = "", "Auto Numbering", msg1) 'J modified
        Call translator.Translate_Reports("Chrono.rpt") 'J added
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
    Call deIms.rsCHRONO.Move(0)
    If Err Then Err.Clear
End Sub

Private Sub ssdbgKeys_AfterUpdate(RtnDispErrMsg As Integer)
If RecSaved = True Then
    lblStatus.ForeColor = &HFF00&
    lblStatus = Visualize
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = NVBAR_ADD
    ssdbgKeys.AllowUpdate = False
End If
End Sub

Private Sub ssdbgKeys_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret, x As Integer
  
          If ssdbgKeys.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
          'Added the next line by Muzammil.
          'Line Commented out by Muzammil,Since the Project would not run as directed by Francois.
         ' If x = 1 Then
             If NotValidLen(ssdbgKeys.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                ssdbgKeys.SetFocus
                ssdbgKeys.Columns(ColIndex).Text = oldVALUE
                ssdbgKeys.Col = 0
                RecSaved = False
                GoodColMove = False
              ElseIf CheckDesCode(ssdbgKeys.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                ssdbgKeys.SetFocus
                ssdbgKeys.Columns(ColIndex).Text = oldVALUE
                ssdbgKeys.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf ssdbgKeys.IsAddRow And ColIndex = 3 Then
           ' If x = 1 Then
            'Added the next line by Muzammil.
          'Line Commented out by Muzammil,Since the Project would not run as directed by Francois.
              If NotValidLen(ssdbgKeys.Columns(ColIndex).Text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                RecSaved = False
                ssdbgKeys.SetFocus
                'ssdbgKeys.Columns(ColIndex).Text =
                ssdbgKeys.Col = 3
            '   Else
            '      x = CheckInt(ssdbgKeys.Columns(2).text)
            '      If x = False Then
            '        RecSaved = False
            '        msg1 = translator.Trans("M00718")
            '        MsgBox IIf(msg1 = "", "Value must be and Integer", msg1)
            '        ssdbgKeys.SetFocus
            '        ssdbgKeys.Col = 3
            '     End If
               End If
        ElseIf Not ssdbgKeys.IsAddRow And ColIndex = 3 Then
            ' If x = 1 Then
             'Added the next line by Muzammil.
          'Line Commented out by Muzammil,Since the Project would not run as directed by Francois.
              If NotValidLen(ssdbgKeys.Columns(ColIndex).Text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                RecSaved = False
                ssdbgKeys.SetFocus
                'ssdbgKeys.Columns(ColIndex).Text =
                ssdbgKeys.Col = 3
               End If
       End If
     Recchanged = DidFieldChange(Trim(oldVALUE), Trim(ssdbgKeys.Columns(ColIndex).Text))

End Sub

Private Sub ssdbgKeys_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(ssdbgKeys.Col)
    If Not good_field Then
       Cancel = True
    End If
End Sub

Private Sub ssdbgKeys_BeforeUpdate(Cancel As Integer)
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
  If ssdbgKeys.IsAddRow Then
      x = NotValidLen(ssdbgKeys.Columns(3).Text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  ssdbgKeys.SetFocus
         ssdbgKeys.Col = 3
         Exit Sub
      End If
      x = CheckDesCode(ssdbgKeys.Columns(0).Text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         ssdbgKeys.SetFocus
         ssdbgKeys.Col = 0
         Exit Sub
      End If
   End If
End If
    If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbOKCancel, "Imswin")
   End If
     If (response = vbOK) Or (response = vbYes) Then
        ssdbgKeys.Columns(1).Text = RemoveNulls(Trim(ssdbgKeys.Columns(1).Text))
        ssdbgKeys.Columns(2).Text = RemoveNulls(Trim(ssdbgKeys.Columns(2).Text))
        ssdbgKeys.Columns(4).Text = RemoveNulls(Trim(ssdbgKeys.Columns(4).Text))
        ssdbgKeys.Columns(5).Text = RemoveNulls(Trim(ssdbgKeys.Columns(5).Text))
        ssdbgKeys.Columns(6).Text = RemoveNulls(Trim(ssdbgKeys.Columns(6).Text))
        ssdbgKeys.Columns("np").Text = deIms.NameSpace
        If ssdbgKeys.IsAddRow Then
            ssdbgKeys.Columns("create_date").Text = Date
            ssdbgKeys.Columns("create_user").Text = CurrentUser
        End If
        ssdbgKeys.Columns("modify_date").Text = Date
   '    ssdbgKeys.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
       'ssdbgKeys.CancelUpdate
     Cancel = True
  '   ssdbgKeys.Refresh
   End If
End Sub


Private Sub ssdbgKeys_KeyPress(KeyAscii As Integer)
 Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean
    
If Not ssdbgKeys.IsAddRow And ssdbgKeys.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (ssdbgKeys.Col = 6)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = ssdbgKeys.Col
        If (cur_col = 6) Then
            If GoodColMove = True Then
                ssdbgKeys.Col = 0
            Else
                GoodColMove = True
            End If
        Else
            If GoodColMove = True Then
                good_field = validate_fields(ssdbgKeys.Col)
                If good_field Then
                    ssdbgKeys.Col = cur_col + 1
                End If
            Else
                GoodColMove = True
            End If
        End If
    End If
End If
End Sub

'if print error cause, disable it

Private Sub ssdbgKeys_PrintError(ByVal PrintError As Long, response As Integer)
    If PrintError = 30457 Then response = 0
End Sub

Public Function DoesCombinationExists() As Boolean

On Error GoTo ErrHandler

    Dim rsDOCTYPE As New ADODB.Recordset
    
DoesCombinationExists = True
    
    rsDOCTYPE.Source = "select count(*) RT from chrono where chr_code ='" & Trim(ssdbgKeys.Columns(0).Text) & "' and chr_fld1='" & Trim(ssdbgKeys.Columns(1).Text) & "' and chr_npecode ='" & deIms.NameSpace & "'"
    
    rsDOCTYPE.ActiveConnection = deIms.cnIms
    
    rsDOCTYPE.Open

If rsDOCTYPE("rt") = 0 Then DoesCombinationExists = False

Exit Function

ErrHandler:

MsgBox "Errors occurred while trying to verify if that Record alredy exists." & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function
