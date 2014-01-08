VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmEccn 
   Caption         =   "ECCN"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   9615
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   6360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      NewEnabled      =   -1  'True
      AllowDelete     =   0   'False
      DeleteVisible   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSoleEccn 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9570
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      ExtraHeight     =   291
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "id"
      Columns(0).Name =   "id"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Eccnno"
      Columns(1).Name =   "Eccnno"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   10425
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2090
      Columns(3).Caption=   "Active"
      Columns(3).Name =   "Active"
      Columns(3).Alignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      _ExtentX        =   16880
      _ExtentY        =   10821
      _StockProps     =   79
      Caption         =   "ECCN"
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
      Left            =   6840
      TabIndex        =   1
      Top             =   6240
      Width           =   2460
   End
End
Attribute VB_Name = "frmEccn"
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
Dim NVBAR_DELETE As Boolean
Dim CAncelGrid As Boolean
Dim TableLocked As Boolean, currentformname As String
Dim FormMode As FormMode

Private Function validate_fields(colnum As Integer) As Boolean

Dim x As Boolean

validate_fields = True
If SSoleEccn.IsAddRow Then

   If colnum = 1 Then
      
      x = NotValidLen(SSoleEccn.Columns(colnum).Text)
      If x = True Then
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
         SSoleEccn.SetFocus
         SSoleEccn.Col = colnum
         validate_fields = False
         Exit Function
      End If
    
    End If
    
    If colnum = 1 Then
      
     
      If Len(SSoleEccn.Columns(colnum).Text) > ConnInfo.EccnLength Then
      
         RecSaved = False
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Eccn# field can not be greater than " & ConnInfo.EccnLength, msg1)
         SSoleEccn.SetFocus
         SSoleEccn.Col = colnum
         validate_fields = False
         Exit Function
         
      End If
    
    End If
    
      If colnum = 1 Then
        
        x = DoesEccnNoExist(SSoleEccn.Columns(1).Text)
        
        If x = True Then
        
             RecSaved = False
             msg1 = translator.Trans("M00703")
             MsgBox IIf(msg1 = "", "The Eccn# already exists, please type in a unique Eccn#.", msg1)
             SSoleEccn.SetFocus
             SSoleEccn.Col = 1
            validate_fields = False
            
         End If
     
     End If
     
End If

End Function
'load form, get data for data grid set button

Private Sub Form_Load()

Dim ctl As Control
Dim textboxes As Control
Dim cmd As New ADODB.Command
Dim Rseccn As ADODB.Recordset
On Error GoTo ErrHand

Me.Tag = "01010106"

 FormMode = ChangeModeOfForm(lblStatus, mdvisualization)

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes

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

    Call translator.Translate_Forms("frmeccn")
    
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    

With cmd

    .CommandText = "eccnselect"
    .CommandType = adCmdStoredProc
    .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
    .ActiveConnection = deIms.cnIms
    
    Set Rseccn = .Execute

End With

'SSoleEccn.Cols = 4
SSoleEccn.ColumnHeaders = True
SSoleEccn.Columns(0).Caption = "id"
SSoleEccn.Columns(0).Visible = False
SSoleEccn.Columns(1).Caption = "#"
SSoleEccn.Columns(2).Caption = "Description"
SSoleEccn.Columns(3).Caption = "Active"

SSoleEccn.Columns(1).FieldLen = 10
SSoleEccn.Columns(2).FieldLen = 100

Do While Not Rseccn.EOF

    SSoleEccn.AddItem Rseccn("eccnid") & vbTab & Rseccn("eccn_no") & vbTab & Rseccn("eccn_desc") & vbTab & Rseccn("eccn_active")
    Rseccn.MoveNext
    
Loop
    
    Screen.MousePointer = vbDefault
    
    Me.Caption = Me.Caption + " - " + Me.Tag
    
    NavBar1.EditVisible = True
    NavBar1.EditEnabled = True
    NavBar1.DeleteVisible = True
    NavBar1.NextVisible = False
    NavBar1.PreviousVisible = False
    NavBar1.FirstVisible = False
    NavBar1.LastVisible = False
    NavBar1.EMailVisible = False
    NavBar1.PrintVisible = False
    NavBar1.LastPrintSepVisible = False
    NavBar1.PrintSaveSepVisible = False

    Call DisableButtons(Me, NavBar1)
    
    NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    NVBAR_DELETE = NavBar1.NewEnabled
    'navbar1.EditEnabled = True
    'navbar1.EditVisible = True
    'navbar1.CancelEnabled = False
    'navbar1.SaveEnabled = False
    'navbar1.CloseEnabled = True
    NavBar1.Width = 5050

    

    
    FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
    
    Call ToggleNavButtons(FormMode)
    
    Me.Width = 9735
    Me.Height = 7170
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
    Exit Sub
    
ErrHand:
    Err.Clear
End Sub

'unload form
Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 InUnload = True
 RecSaved = True
 CAncelGrid = False
 SSoleEccn.Update
 If FormMode <> mdvisualization Then
    
    If MsgBox("Are you sure you want to close?", vbInformation + vbYesNo, "Imswin") = vbNo Then
       
       Cancel = 1
       Exit Sub
       
    End If
    
End If



   If open_forms <= 5 Then ShowNavigator
   If Err Then Err.Clear
    

If TableLocked = True Then
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If


End Sub

Private Sub NavBar1_BeforeDeleteClick()
If MsgBox("Are you sure you want to delete?", vbInformation + vbYesNo) = vbYes Then
If SaveRecord(True) = True Then
    SSoleEccn.DeleteSelected
    MsgBox "Delete successfully"
End If
End If

If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

End Sub



'before save a record set modify user equal to current user

''Private Sub NavBar1_BeforeMove(bCancel As Boolean)
''
''     SSoleEccn.Update
''End Sub

'set name space equal to current name space

Private Sub NavBar1_BeforeNewClick()
   SSoleEccn.AddNew

    NavBar1.CancelEnabled = True
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create

    SSoleEccn.Columns("active").Text = 1
    SSoleEccn.SetFocus
    SSoleEccn.Col = 1

FormMode = ChangeModeOfForm(lblStatus, mdCreation)
Call ToggleNavButtons(FormMode)
End Sub

'before save check unit code exist or not shoaw message

Private Sub NavBar1_BeforeSaveClick()
        Dim cmd As New Command
        Dim rs As ADODB.Recordset
        On Error GoTo ErrHand
        
        SSoleEccn.Update
   'If validate_fields( = False Then Exit Function
   If SaveRecord(False) = False Then Exit Sub
                
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        lblStatus.ForeColor = &HFF00&
        lblStatus = Visualize
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        NavBar1.EditEnabled = True
        NavBar1.NewEnabled = NVBAR_ADD
        SSoleEccn.AllowUpdate = False
        

       
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Me.Name  'Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)

End If
FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
Exit Sub
ErrHand:
MsgBox Err.Description
        Err.Clear
       
End Sub

Private Sub NavBar1_OnCancelClick()
 Dim response As Integer
   
      msg1 = translator.Trans("M00706")
      response = MsgBox((IIf(msg1 = "", " Are you sure you want to cancel changes?", msg1)), vbYesNo, "Imswin")
   
      If response = vbYes Then
          
         SSoleEccn.CancelUpdate
         NavBar1.EditEnabled = True
         NavBar1.NewEnabled = True
         NavBar1.CancelEnabled = False
         NavBar1.SaveEnabled = False
         SSoleEccn.AllowUpdate = False
         lblStatus.ForeColor = &HFF00&
         lblStatus.Caption = Visualize
    SSoleEccn.CancelUpdate


    FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
    Call ToggleNavButtons(mdvisualization)
    
  ' If RsEccn.RecordCount > 0 Then SSoleEccn.Columns(1).locked = False
  


'End If


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
End Sub

Private Sub NavBar1_OnEditClick()

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Me.Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

If TableLocked = True Then
   
        NavBar1.SaveEnabled = False
        NavBar1.NewEnabled = False
        NavBar1.CancelEnabled = False
        
        SSoleEccn.AllowUpdate = False
        FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
        Exit Sub
Else

        SSoleEccn.AllowUpdate = True
        'SSoleEccn.Columns(1).locked = True
        NavBar1.CancelEnabled = True
        NavBar1.EditEnabled = False
        NavBar1.SaveEnabled = True
        NavBar1.NewEnabled = False
        lblStatus.ForeColor = &HFF0000
        lblStatus.Caption = Modify
        SSoleEccn.SetFocus
        SSoleEccn.Col = 1
        SSoleEccn.AllowUpdate = True
        FormMode = ChangeModeOfForm(lblStatus, mdModification)
        Call ToggleNavButtons(FormMode)
        
        TableLocked = True

End If

End Sub



'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Unit.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00144") 'J added
        .WindowTitle = IIf(msg1 = "", "Unit", msg1) 'J modified
        Call translator.Translate_Reports("Unit.rpt") 'J added
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



Private Sub LROleDBNavBar1_Click()

End Sub

''Private Sub ssoleEccn_AfterUpdate(RtnDispErrMsg As Integer)
''If RecSaved = True Then
''    lblStatus.ForeColor = &HFF00&
''    lblStatus = Visualize
''    navbar1.SaveEnabled = False
''    navbar1.CancelEnabled = False
''    navbar1.EditEnabled = True
''    navbar1.NewEnabled = NVBAR_ADD
''    SSoleEccn.AllowUpdate = False
''End If
''
''End Sub



Private Sub ssoleEccn_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
  
          If SSoleEccn.IsAddRow And ColIndex = 1 Then
             
             If NotValidLen(SSoleEccn.Columns(ColIndex).Text) Then
                
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Please enter a value for Eccn #.", msg1)
                Cancel = 1
                SSoleEccn.SetFocus
                SSoleEccn.Columns(ColIndex).Text = oldVALUE
                SSoleEccn.Col = 0
                RecSaved = False
                GoodColMove = False
              
              ElseIf DoesEccnNoExist(SSoleEccn.Columns(ColIndex).Text) Then
                
                msg1 = translator.Trans("M00703")
                MsgBox IIf(msg1 = "", "Eccn # already exists. Please create a unique one.", msg1)
                Cancel = 1
                SSoleEccn.SetFocus
                SSoleEccn.Columns(ColIndex).Text = oldVALUE
                SSoleEccn.Col = ColIndex
                RecSaved = False
                GoodColMove = False
                
             End If
        
        ElseIf ColIndex = 2 Then
              
              If NotValidLen(SSoleEccn.Columns(ColIndex).Text) Then
                
                msg1 = translator.Trans("M00702")
                MsgBox IIf(msg1 = "", "Please enter a description for the Eccn Description.", msg1)
                Cancel = 1
                SSoleEccn.SetFocus
                RecSaved = False
                SSoleEccn.Col = 1
               
               End If
               
       End If
     'Recchanged = DidFieldChange(Trim(oldVALUE), Trim(SSoleEccn.Columns(ColIndex).text))
     

End Sub



Private Sub SSoleEccn_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
DispPromptMsg = 0
End Sub

Private Sub ssoleEccn_BeforeRowColChange(Cancel As Integer)

Dim good_field As Boolean
    good_field = validate_fields(SSoleEccn.Col)
    If Not good_field Then
       Cancel = 1
    End If


End Sub
''''
Private Sub ssoleEccn_BeforeUpdate(Cancel As Integer)
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

  If SSoleEccn.IsAddRow Then
      x = NotValidLen(SSoleEccn.Columns(1).Text)
      If x = True Then
         RecSaved = False
         Cancel = True
         msg1 = translator.Trans("M00702")
         MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
                  SSoleEccn.SetFocus
         SSoleEccn.Col = 1
         Exit Sub
      End If
      x = DoesEccnNoExist(SSoleEccn.Columns(1).Text)
      If x <> False Then
         RecSaved = False
         msg1 = translator.Trans("M00703")
         MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
         SSoleEccn.SetFocus
         SSoleEccn.Columns(1).Text = ""
         SSoleEccn.Col = 1
         Exit Sub
      End If
   End If
End If
    
   If InUnload = False Then
           msg1 = translator.Trans("M00705") 'J added
          response = MsgBox((IIf(msg1 = "", "Are you sure you want to save the changes?", msg1)), vbYesNo, "Imswin")
   End If
   
     If (response = vbOK) Or (response = vbYes) Then
        
        If SaveRecord(False) = False Then
            
            Cancel = 1
            
        Else
        
            FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
            ToggleNavButtons (FormMode)
        End If
        
     Else
       CAncelGrid = True
        RecSaved = False
      ' ssoleEccn.CancelUpdate
     Cancel = True
   End If

End Sub

Private Function DoesEccnNoExist(Eccnno As String) As Boolean

On Error Resume Next
Dim cmd As New ADODB.Command
Dim rst As ADODB.Recordset

On Error GoTo ErrHand
    
    With cmd
        
        .CommandText = "doeseccnnoexist"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        .parameters.Append .CreateParameter("@eccnno", adVarChar, adParamInput, 25, Trim(Eccnno))
        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        Set rst = .Execute
    
        DoesEccnNoExist = rst!countit
        
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    
Exit Function
ErrHand:

    
    If Err Then Call LogErr(Name & "::DoesEccnNoExist", Err.Description, Err.number, True)
End Function

Private Sub ssoleEccn_KeyPress(KeyAscii As Integer)
 
 Dim Char
 Dim cur_col As Integer
 Dim good_field As Boolean

    
If Not SSoleEccn.IsAddRow And SSoleEccn.Col = 1 And KeyAscii <> 13 Then

    KeyAscii = 0
    
Else

    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
  '  If (ssoleEccn.IsAddRow And ssoleEccn.Col = 0) Then
  '     If Len(ssoleEccn.Columns(0).text) > 3 Then
  '        KeyAscii = 0
  '      End If
  '  End If
    If KeyAscii = 13 Or ((KeyAscii = 9) And (SSoleEccn.Col = 2)) Then
    
        GoodColMove = True
        
    End If
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        
        cur_col = SSoleEccn.Col
        
        If (cur_col = 2) Then
            
            If GoodColMove = True Then
                SSoleEccn.Col = 0
            Else
                GoodColMove = True
            End If
        
        Else
            
            If GoodColMove = True Then
                
                good_field = validate_fields(SSoleEccn.Col)
                If good_field Then
                    SSoleEccn.Col = cur_col + 1
                End If
            
            Else
                
                GoodColMove = True
            
            End If
            
        End If
        
    End If
    
End If
End Sub


Public Function SaveRecord(Delete As Boolean) As Boolean

        Dim cmd As New Command
        Dim rs As ADODB.Recordset
        Dim candelete As Integer
        On Error GoTo ErrHand
        
        'SSoleEccn.Update
        
        With cmd
        
            .ActiveConnection = deIms.cnIms
            .CommandText = "ECCNUpdate"
            .CommandType = adCmdStoredProc
            .parameters.Append .CreateParameter("@eccnid", adBigInt, adParamInputOutput, , IIf(FormMode = mdCreation, 0, SSoleEccn.Columns(0).value))
            .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
            .parameters.Append .CreateParameter("@eccn_no", adVarChar, adParamInput, 25, SSoleEccn.Columns(1).value)
            .parameters.Append .CreateParameter("@eccn_desc", adVarChar, adParamInput, 500, SSoleEccn.Columns(2).value)
            .parameters.Append .CreateParameter("@actiondelete", adBoolean, adParamInput, , Delete)
            .parameters.Append .CreateParameter("@eccn_active", adBoolean, adParamInput, , SSoleEccn.Columns(3).value)
            .parameters.Append .CreateParameter("@CanDelete", adInteger, adParamOutput, , Null)
        
            Set rs = .Execute
        
        End With
        
        If FormMode = mdCreation Then SSoleEccn.Columns(0).value = cmd.parameters("@eccnid").value
        If Delete = True Then candelete = cmd.parameters("@CanDelete").value
        
        If candelete = 1 Then
                MsgBox "This Eccn# has records associated with Stockmaster\ Purchase Order and can not be deleted. De-activating it will stop users from including it in any new records.", vbInformation
                Exit Function
        End If
        
        SaveRecord = True
        
    Exit Function
ErrHand:
    
    MsgBox "Errors Occurred while trying to save this Record. Error Description :" & Err.Description
    Err.Clear
    
End Function

Public Function ToggleNavButtons(FMode As FormMode) As Boolean
 
        If FormMode = mdvisualization Then
                    
                    NavBar1.EditEnabled = NVBAR_EDIT  'True
                    NavBar1.NewEnabled = NVBAR_ADD 'True
                    NavBar1.CancelEnabled = False
                    NavBar1.SaveEnabled = False
                    
                    
                    SSoleEccn.AllowUpdate = False

        ElseIf FormMode = mdCreation Then
                 
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    
                    SSoleEccn.AllowUpdate = True
                    
        ElseIf FormMode = mdModification Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    SSoleEccn.AllowUpdate = True
                    
        End If
        
        NavBar1.DeleteEnabled = NavBar1.NewEnabled
        
End Function

Private Sub SSoleEccn_Validate(Cancel As Boolean)
Dim good_field As Boolean

If FormMode <> mdvisualization Then

    good_field = validate_fields(SSoleEccn.Col)
    If Not good_field Then
       Cancel = True
    End If
    
End If

End Sub
