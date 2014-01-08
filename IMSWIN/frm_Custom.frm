VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frm_Custom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Category"
   ClientHeight    =   4500
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   7035
   Tag             =   "01010900"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLine 
      Bindings        =   "frm_Custom.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
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
      stylesets(0).Name=   "COLLS"
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
      stylesets(0).Picture=   "frm_Custom.frx":0014
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
      stylesets(1).Picture=   "frm_Custom.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowDelete     =   -1  'True
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
      HeadStyleSet    =   "COLLS"
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   5027
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "cust_cate"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   20
      Columns(1).Width=   5292
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "np"
      Columns(1).Name =   "np"
      Columns(1).DataField=   "cust_npecode"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1402
      Columns(2).Caption=   "Active"
      Columns(2).Name =   "active"
      Columns(2).DataField=   "cust_actvflag"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(3).Width=   5292
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "modify_date"
      Columns(3).Name =   "modify_date"
      Columns(3).DataField=   "cust_modidate"
      Columns(3).DataType=   135
      Columns(3).FieldLen=   256
      Columns(4).Width=   5292
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "modify_user"
      Columns(4).Name =   "modify_user"
      Columns(4).DataField=   "cust_modiuser"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "create_date"
      Columns(5).Name =   "create_date"
      Columns(5).DataField=   "cust_creadate"
      Columns(5).DataType=   135
      Columns(5).FieldLen=   256
      Columns(6).Width=   5292
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "create_user"
      Columns(6).Name =   "create_user"
      Columns(6).DataField=   "cust_creauser"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   7435
      _ExtentY        =   5106
      _StockProps     =   79
      DataMember      =   "CUSTOM"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   2
      Top             =   3720
      Width           =   2460
   End
   Begin VB.Label lbl_Custom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Category"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frm_Custom"
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
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From custom "
        .CommandText = .CommandText & " Where cust_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND cust_cate = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBLine.IsAddRow Then
   If colnum = 0 Or colnum = 1 Then
      x = NotValidLen(SSDBLine.Columns(colnum).text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSDBLine.SetFocus
         SSDBLine.Col = colnum
         validate_fields = False
         Exit Function
      End If
    End If
      If colnum = 0 Then
       x = CheckDesCode(SSDBLine.Columns(0).text)
        If x <> False Then
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
             SSDBLine.Columns(0).text = ""
             SSDBLine.SetFocus
             SSDBLine.Col = 0
            validate_fields = False
         End If
    End If
   End If

End Function
'load form get data for combo

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
    
    Call deIms.Custom(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsCUSTOM
    
    
    Visible = True
    Screen.MousePointer = vbDefault
    'SSDBLinen.BatchUpdate = True
    Set SSDBLine.DataSource = deIms
    
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
    SSDBLine.AllowUpdate = False
    
    'Call DisableButtons(Me, NavBar1)

End Sub

' unload form

Private Sub Form_Unload(Cancel As Integer)
Dim response As String
On Error Resume Next
 InUnload = True
  RecSaved = True
  CAncelGrid = False
SSDBLine.Update
 If RecSaved = True Then
     Hide
      deIms.rsCUSTOM.Close
  If open_forms <= 5 Then ShowNavigator
    Else
    Cancel = True
End If

End Sub

'cancel update recordset

Private Sub NavBar1_BeforeCancelClick()
    CAncelGrid = True
'   SSDBLine.CancelUpdate
End Sub

'set recordset update

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBLine.Update
End Sub

'click new set recordset update

Private Sub NavBar1_BeforeNewClick()
    SSDBLine.AddNew
    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    SSDBLine.AllowUpdate = True
    SSDBLine.Columns("active").text = 1
    SSDBLine.SetFocus
    SSDBLine.Col = 0

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
          CAncelGrid = True
          SSDBLine.CancelUpdate
       '   SSDBLine.Refresh
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
    SSDBLine.Columns(0).locked = False
   ' Cancel = -1
   ' CAncelGrid = True
    SSDBLine.CancelUpdate
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    SSDBLine.AllowUpdate = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
'    SSDBLine.Refresh
End If
      
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnEditClick()
SSDBLine.AllowUpdate = True
SSDBLine.Columns(0).locked = True
NavBar1.CancelEnabled = True
NavBar1.EditEnabled = False
NavBar1.SaveEnabled = True
NavBar1.NewEnabled = False
lblStatus.ForeColor = &HFF0000
lblStatus.Caption = Modify
SSDBLine.SetFocus
SSDBLine.Col = 1
SSDBLine.AllowUpdate = True

End Sub

Private Sub NavBar1_OnNewClick()
    SSDBLine.AllowUpdate = False

End Sub

'get parameter for crystal report and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Custom.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00146") 'J added
        .WindowTitle = IIf(msg1 = "", "Custom", msg1) 'J modified
        Call translator.Translate_Reports("Custom.rpt") 'J added
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

'save recordset

Private Sub NavBar1_BeforeSaveClick()
    CAncelGrid = False
    SSDBLine.Update
        If RecSaved = True Then
            NavBar1.SaveEnabled = False
            SSDBLine.Columns(0).locked = False
            NavBar1.CancelEnabled = False
            lblStatus.ForeColor = &HFF00&
            lblStatus = Visualize
            NavBar1.SaveEnabled = False
            NavBar1.CancelEnabled = False
            NavBar1.EditEnabled = True
            NavBar1.NewEnabled = NVBAR_ADD
            SSDBLine.AllowUpdate = False
       End If
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

Private Sub SSDBLine_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
      
'   If TransCancelled = False Then
    
    
          If SSDBLine.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBLine.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value.", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                SSDBLine.Columns(ColIndex).text = oldVALUE
                RecSaved = False
                SSDBLine.Col = 0
                GoodColMove = False
              ElseIf CheckDesCode(SSDBLine.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                SSDBLine.Columns(ColIndex).text = oldVALUE
                SSDBLine.Col = 0
                RecSaved = False
                GoodColMove = False
             End If
        
        ElseIf SSDBLine.IsAddRow And ColIndex = 1 Then
              If NotValidLen(SSDBLine.Columns(ColIndex).text) Then
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                SSDBLine.SetFocus
                'SSDBLine.Columns(ColIndex).Text =
                RecSaved = False
                SSDBLine.Col = 1
               End If
        ElseIf Not SSDBLine.IsAddRow And ColIndex = 1 Then
                If NotValidLen(SSDBLine.Columns(ColIndex).text) Then
               msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                Cancel = 1
                RecSaved = False
                SSDBLine.SetFocus
                'SSDBLine.Columns(ColIndex).Text =
                SSDBLine.Col = 1
               End If
       End If
                   Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSDBLine.Columns(ColIndex).text))

End Sub

Private Sub SSDBLine_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBLine.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

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
     ' x = NotValidLen(SSDBLine.Columns(1).text)
     ' If x = True Then
     '    RecSaved = False
     '    Cancel = True
     '    msg1 = translator.Trans("M00702")
     '    MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
     '             SSDBLine.SetFocus
     '    SSDBLine.Col = 0
     '    Exit Sub
     ' End If
      x = CheckDesCode(SSDBLine.Columns(0).text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
          SSDBLine.Columns(0).text = ""
         SSDBLine.SetFocus
         SSDBLine.Col = 0
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
        
        SSDBLine.Columns("np").text = deIms.NameSpace
        If SSDBLine.IsAddRow Then
            SSDBLine.Columns("create_date").text = Date
            SSDBLine.Columns("create_user").text = CurrentUser
        End If
        SSDBLine.Columns("modify_date").text = Date
        SSDBLine.Columns("modify_user").text = CurrentUser
        Cancel = 0
     Else
       CAncelGrid = True
        RecSaved = False
       'SSDBLine.CancelUpdate
     Cancel = True
  '   SSDBLine.Refresh
   End If
End Sub

Private Sub SSDBLine_KeyPress(KeyAscii As Integer)
  Dim Char
  Dim cur_col As Integer
  Dim good_field As Boolean


    
    
If Not SSDBLine.IsAddRow And SSDBLine.Col = 0 And KeyAscii <> 13 Then
    KeyAscii = 0
Else
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSDBLine.Col = 1)) Then
        GoodColMove = True
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        cur_col = SSDBLine.Col
        If (cur_col = 1) Then
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
