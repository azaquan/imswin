VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmBuyerRight 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buyer"
   ClientHeight    =   4050
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6585
   Icon            =   "BuyerRight.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "04010100"
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      Height          =   375
      Left            =   1620
      TabIndex        =   12
      Top             =   3660
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DisableSaveOnSave=   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2880
      Index           =   1
      Left            =   240
      ScaleHeight     =   2880
      ScaleWidth      =   6105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   6105
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgRights 
         Height          =   2775
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   6015
         _Version        =   196617
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
         stylesets(0).Picture=   "BuyerRight.frx":000C
         stylesets(0).AlignmentText=   0
         stylesets(1).Name=   "ColHeader"
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "BuyerRight.frx":0028
         stylesets(1).AlignmentText=   0
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   5027
         Columns(0).Caption=   "Document Type"
         Columns(0).Name =   "doctype"
         Columns(0).DataField=   "buyr_docutype"
         Columns(0).FieldLen=   256
         Columns(0).HeadStyleSet=   "ColHeader"
         Columns(0).StyleSet=   "RowFont"
         Columns(1).Width=   4948
         Columns(1).Caption=   "Maximum Amount"
         Columns(1).Name =   "amount"
         Columns(1).DataField=   "buyr_maxiamou"
         Columns(1).FieldLen=   256
         Columns(1).HeadStyleSet=   "ColHeader"
         Columns(1).StyleSet=   "RowFont"
         Columns(2).Width=   5292
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "user"
         Columns(2).Name =   "user"
         Columns(2).DataField=   "buyr_username"
         Columns(2).FieldLen=   256
         Columns(3).Width=   5292
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "namespace"
         Columns(3).Name =   "namespace"
         Columns(3).DataField=   "buyr_npecode"
         Columns(3).FieldLen=   256
         _ExtentX        =   10610
         _ExtentY        =   4895
         _StockProps     =   79
         BackColor       =   -2147483643
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
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6006
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Buyer"
            Key             =   "buyer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rights"
            Key             =   "rights"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown ssdbgDocType 
      Height          =   1275
      Left            =   1020
      TabIndex        =   13
      Top             =   1200
      Width           =   3555
      DataFieldList   =   "doc_code"
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnHeaders   =   0   'False
      HeadFont3D      =   3
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4868
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "doc_desc"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "doc_code"
      Columns(1).FieldLen=   256
      _ExtentX        =   6271
      _ExtentY        =   2249
      _StockProps     =   77
      DataFieldToDisplay=   "doc_desc"
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   240
      ScaleHeight     =   960
      ScaleWidth      =   6105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   6105
      Begin VB.TextBox txt_Building 
         DataField       =   "buy_builroom"
         DataMember      =   "BUYER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   4335
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Width           =   1668
      End
      Begin VB.TextBox txt_Phone 
         DataField       =   "buy_phonnumb"
         DataMember      =   "BUYER"
         DataSource      =   "deIms"
         Height          =   288
         Left            =   1185
         MaxLength       =   14
         TabIndex        =   3
         Top             =   480
         Width           =   1668
      End
      Begin VB.TextBox txt_Extension 
         DataField       =   "buy_extenumb"
         DataMember      =   "BUYER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   4
         Top             =   0
         Width           =   1668
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboUserName 
         DataField       =   "buy_username"
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1635
         DataFieldList   =   "usr_userid"
         _Version        =   196617
         ColumnHeaders   =   0   'False
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   2884
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "usr_username"
      End
      Begin VB.Label lbl_Name 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lbl_Building 
         BackStyle       =   0  'Transparent
         Caption         =   "Building room"
         Height          =   285
         Left            =   3150
         TabIndex        =   9
         Top             =   120
         Width           =   1110
      End
      Begin VB.Label lbl_Phone 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label lbl_Extension 
         BackStyle       =   0  'Transparent
         Caption         =   "Extension"
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Top             =   495
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmBuyerRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim cn As ADODB.Connection

Dim AddingRecord As Boolean
Dim rsBuyer As ADODB.Recordset
Dim rsRight As ADODB.Recordset
Dim ObjXevents As ImsXevents
Dim TRANSACTIONNUBMER As Integer
'set tab controls

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


Dim i As Integer
On Error Resume Next

    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    
        i = tbsOptions.SelectedItem.Index
        
        If i = tbsOptions.Tabs.count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
    
    
End Sub

'load form and set caption

Private Sub Form_Load()
On Error Resume Next
    NavBar.SaveEnabled = Getmenuuser(np, CurrentUser, Me.Tag, cn)
    
    Set ObjXevents = New ImsXevents 'Shakir
    ObjXevents.ConnectionObject = cn
    'TRANSACTIONNUBMER =
    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmBuyerRight")
    '--------------------------------------------------
    Caption = Caption & " - " & Tag
    tbsOptions.Tabs(1).Selected = True
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
   
    'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
End Sub

Private Sub Form_Paint()
    If Not NavBar.SaveEnabled Then
        ssdbgRights.AllowUpdate = False
        ssdcboUserName.AllowInput = False
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
'    Call CrystalReport1.LogOffServer(1, True)
   'Shakir 12-17-00
   'The Object is Updating the XEVENTS table with all the Events the User has Generated.
   
End Sub

Private Sub NavBar_BeforeCancelClick()
    If tbsOptions.SelectedItem.Index = 2 Then ssdbgRights.CancelUpdate
End Sub

'get buyer recordset and assign name space values

Private Sub NavBar_BeforeNewClick()
On Error Resume Next

    'NavBar.Update
    AddingRecord = True
    
    If tbsOptions.SelectedItem.Index = 1 Then
        Call GetBuyerNames(False)
        NavBar.Recordset.AddNew
        NavBar.Recordset!buy_npecode = np
    Else
        ssdbgRights.update
        ssdbgRights.AddNew
        NavBar.AllowAddNew = False
        
        With ssdbgRights
            .SetFocus
            .Columns("namespace").Text = np
            .Columns("user").Text = ssdcboUserName.Columns(1).Text
        End With

    End If
    If Err Then Call LogErr("frmBuyerRight::NavBar_BeforeNewClick", Err.Description, Err.Number, True): Err.Clear
End Sub

'before save assign name space equal to current name space

Private Sub NavBar_BeforeSaveClick()
On Error Resume Next
    
    If tbsOptions.SelectedItem.Index = 1 And rsBuyer.EditMode = 2 Then
    
       If checkFIELDS = False Then
          NavBar.AllowUpdate = False
          Exit Sub
       End If
       
    End If
    
    rsBuyer!buy_npecode = np
      If tbsOptions.SelectedItem.Index = 2 Then ssdbgRights.update
      
      If Err Then Call LogErr("frmBuyerRight::NavBar_BeforeSaveClick", Err.Description, Err, True): Err.Clear
End Sub

Private Sub NavBar_MoveComplete()
    If Not NavBar.SaveEnabled Then
        ssdbgRights.AllowUpdate = False
        ssdcboUserName.AllowInput = False
    End If
End Sub

'cancel recordset update

Private Sub NavBar_OnCancelClick()
On Error Resume Next

    If AddingRecord Then
    
        AddingRecord = Err <> 0
        
        If tbsOptions.SelectedItem.Index = 1 Then
            
            rsBuyer.CancelUpdate
            rsBuyer.CancelBatch adAffectCurrent
            
            GetBuyerNames
            ssdcboUserName.MoveLast
            ssdcboUserName.MoveFirst
            GetBuyers (ssdcboUserName.Columns(1).Text)
            Set NavBar.Recordset = rsBuyer
        Else
            ssdbgRights.CancelUpdate
            ssdbgRights.CancelUpdate
        End If
        
    End If
    
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnCancelClick", Err.Description, Err, True): Err.Clear
End Sub

'close form cancel recordset update

Private Sub NavBar_OnCloseClick()
On Error Resume Next

    Unload Me
  
    ssdbgRights.CancelUpdate
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnCloseClick", Err.Description, Err, True): Err.Clear
End Sub

'move recordset to first position

Private Sub NavBar_OnFirstClick()
On Error Resume Next
    If tbsOptions.SelectedItem.Index = 1 Then
        ssdcboUserName.MoveFirst
        ssdcboUserName.Text = ssdcboUserName.Columns(0).Text
        ssdcboUserName_Click
    End If
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnFirstClick", Err.Description, Err, True): Err.Clear
End Sub

'move recordset to last position

Private Sub NavBar_OnLastClick()
On Error Resume Next
    If tbsOptions.SelectedItem.Index = 1 Then
        ssdcboUserName.MoveLast
        ssdcboUserName.Text = ssdcboUserName.Columns(0).Text
        ssdcboUserName_Click
    End If
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnLastClick", Err.Description, Err, True): Err.Clear
End Sub

'set name space to current name space and recordset buyer
'name space to current name space

Private Sub NavBar_OnNewClick()
On Error Resume Next
    If tbsOptions.SelectedItem.Index = 2 Then
    
        NavBar.CancelUpdate
        
        With ssdbgRights
            .SetFocus
            .Columns("namespace").Text = np
            .Columns("user").Text = ssdcboUserName.Columns(1).Text
        End With
    Else
        rsBuyer!buy_npecode = np
    End If
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnNewClick", Err.Description, Err, True): Err.Clear
End Sub

'move recordset to next position

Private Sub NavBar_OnNextClick()
On Error Resume Next
    If tbsOptions.SelectedItem.Index = 1 Then
        ssdcboUserName.MoveNext
        
        ssdcboUserName.Text = ssdcboUserName.Columns(0).Text
        ssdcboUserName_Click
    End If
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnNextClick", Err.Description, Err, True): Err.Clear
End Sub

'move recordset to previous position

Private Sub NavBar_OnPreviousClick()
On Error Resume Next
    If tbsOptions.SelectedItem.Index = 1 Then
        ssdcboUserName.MovePrevious
        ssdcboUserName.Text = ssdcboUserName.Columns(0).Text
        ssdcboUserName_Click
    End If
End Sub

'call function to print crystal report

Private Sub NavBar_OnPrintClick()
'On Error Resume Next
'Dim retval As PrintOpts

    Load frmPrintDialog
    With frmPrintDialog
        .Show 1

        DoEvents: DoEvents
        If .Result = poPrintCurrent Then
            Call BeforePrint(poPrintCurrent)

        ElseIf .Result = poPrintAll Then
            Call BeforePrint(poPrintAll)

        Else
            Exit Sub

        End If

    End With
        
    Unload frmPrintDialog
    Set frmPrintDialog = Nothing
    CrystalReport1.Action = 1
    If Err Then Call LogErr("frmBuyerRight::NavBar_OnPrintClick", Err.Description, Err, True): Err.Clear
End Sub

'get crystal report parameters and application path

Public Sub BeforePrint(iOption As PrintOpts)
On Error Resume Next
Dim Path As String

    Path = ReportPath
    
    With CrystalReport1
        .ReportFileName = Path & "buyright.rpt"
        If iOption = poPrintCurrent Then
            .ParameterFields(1) = "username;" & ssdcboUserName.Columns(1).Text & ";TRUE"
        Else
            .ParameterFields(1) = "username;ALL;TRUE"
        End If
        .ParameterFields(0) = "namespace;" + np + ";TRUE"
        
        'Modified by Juan (10/23/00) for Multilingual 'J added
        Call translate_reports(Me.Name, "buyright.rpt", True, cn, CrystalReport1) 'J added
        msg1 = Trans("M00689") 'J added
        .WindowTitle = IIf(msg1 = "", "Buyer Right", msg1) 'J modified
        '--------------------------------------------------
    End With
    
    If Err Then Call LogErr("frmBuyerRight::BeforePrint", Err.Description, Err, True): Err.Clear
End Sub

'before save record to get name space and current user

Private Sub NavBar_OnSaveClick()
On Error Resume Next
Dim i As Integer, X As Integer

    If AddingRecord Then
    
        
        
        If tbsOptions.SelectedItem.Index = 2 Then
            
            With ssdbgRights
                X = .Rows
                .MoveFirst
                
                For i = 0 To X
                    .Columns("namespace").Text = np
                    .Columns("user").Text = ssdcboUserName.Columns(1).Text
                    
                    .update
                    .MoveNext
                Next i
            End With
        
        Else
            rsBuyer.update
            rsBuyer.UpdateBatch
            ssdbgRights.update
            Call rsRight.Move(0)
                    
            
            ssdcboUserName.Text = ssdcboUserName.Columns(0).Text
            ssdcboUserName_Click
        End If
        
    Else
        Call rsBuyer.Move(0)
        rsBuyer.UpdateBatch
        ssdbgRights.update
    End If
        
    If tbsOptions.SelectedItem.Index = 1 And AddingRecord Then
        GetBuyerNames
        Call FindBuyer
    End If
    If AddingRecord Then AddingRecord = Err <> 0

    If Err Then Call LogErr("frmBuyerRight::NavBar_OnSaveClick", Err.Description, Err, True): Err.Clear
End Sub


Private Sub ssdbgDocType_Click()
 Dim X  As String
''  x = ssdbgRights.Columns(0).Text
''   ssdbgRights.Columns(0).Text = ""
''    If ssdbgRights.IsItemInList = True Then
''    ssdbgRights.Columns(0).Text = ""
''      MsgBox "User already had rights for the specified document. Please select a different one.", vbInformation, "Imswin"
''    Else
''     ssdbgRights.Columns(0).Text = x
''
''    End If

End Sub

Private Sub ssdbgRights_AfterUpdate(RtnDispErrMsg As Integer)
  If Not ObjXevents Is Nothing Then
      ObjXevents.update
    Set ObjXevents = Nothing
   End If  'RtnDispErrMsg = 0
End Sub

Private Sub ssdbgRights_BeforeRowColChange(Cancel As Integer)
Dim X As String
Dim count As Integer
Dim rs As ADODB.Recordset

If ssdbgRights.IsAddRow And Len(Trim$(ssdbgRights.Columns(0).Text)) > 0 Then

     Set rs = New ADODB.Recordset
  
     rs.Source = "select count(*) countIt from buyer_right where buyr_npecode='" & np & "' and buyr_username = '" & Trim$(ssdbgRights.Columns(2).Text) & "' and buyr_docutype='" & Trim$(ssdbgDocType.Columns(1).Text) & "'"
     rs.ActiveConnection = cn
     rs.Open
  
     If rs!countit > 0 Then
      ssdbgRights.Columns(0).Text = ""
      MsgBox "User already had rights for the specified document. Please select a different one.", vbInformation, "Imswin"
      Cancel = -1
     End If
     
 End If

'x = Trim$(ssdbgRights.Columns(0).Text)
'ssdbgRights.Columns(0).Text = ""
'If ssdbgRights.IsItemInList Then
'  MsgBox ""

'''    x = ssdbgRights.Columns(0).Text
'''    ssdbgRights.Columns(0).Text = ""
'''    If ssdbgRights.IsItemInList = True Then
'''    ssdbgRights.Columns(0).Text = ""
'''      MsgBox "User already had rights for the specified document. Please select a different one.", vbInformation, "Imswin"
'''      Cancel = -1
'''    Else
'''     ssdbgRights.Columns(0).Text = x
'''
'''    End If


'ssdbgRights.MoveFirst
 
'count = ssdbgRights.Rows
 
'For i = 0 To x - 1
'If x = Trim$(ssdbgRights.Columns(0).Text) Then
   
  

End Sub

'set data grid name space to current name space and
' user to current user

Private Sub ssdbgRights_BeforeUpdate(Cancel As Integer)
  Dim OldVAlue As String
  Dim NewVAlue As String
    ssdbgRights.Columns("namespace").Text = np
    ssdbgRights.Columns("user").Text = ssdcboUserName.Columns(1).Text
    'ssdbgrights.columns("namespace").text
    'SHAKIR
    
If ssdbgRights.IsAddRow = True Then
   OldVAlue = 0
   Else
   OldVAlue = ssdbgRights.Columns(1).CellText(ssdbgRights.Bookmark)
   
End If

 NewVAlue = ssdbgRights.Columns(1).Text
 
 'Shakir 12-17-00
  'The Object is Being Fed over here with the Data to be fed in the Xevents Table.
  'Update method on this Object is Exceuted when the User Exist the Form.
  Set ObjXevents = New ImsXevents
  ObjXevents.ConnectionObject = cn
  ObjXevents.AddNew
  ObjXevents.Namespace = np
  ObjXevents.MyLoginId = CurrentUser
  ObjXevents.HisLoginId = Trim$(ssdcboUserName.Columns(1).Text)
If ssdbgRights.IsAddRow = True Then
  ObjXevents.EventDetail = "New right on " & Trim$(ssdbgRights.Columns(0).Text) & " has been granted with " & NewVAlue & " approval amount for the User " & Trim$(ssdcboUserName.Columns(1).Text) & "."
  Else
  'ObjXevents.EventDetail = "The approval amount on " & Trim$(ssdbgRights.Columns(0).Text) & " has been changed from " & OldVAlue & " to " & NewVAlue & " for the User " & Trim$(ssdcboUserName.Columns(1).Text) & "."
   ObjXevents.EventDetail = "The record with doctype " & ssdbgRights.Columns(0).CellText(ssdbgRights.Bookmark) & " / $" & ssdbgRights.Columns(1).CellText(ssdbgRights.Bookmark) & " has been changed to doctype " & Trim$(ssdbgRights.Columns(0).Text) & " / $" & Trim$(ssdbgRights.Columns(1).Text) & " for the User " & Trim$(ssdcboUserName.Columns(1).Text) & "."
  End If

ObjXevents.OldVAlue = IIf(Len(OldVAlue) = 0 Or IsNull(OldVAlue), 0, OldVAlue)
ObjXevents.NewVAlue = IIf(Len(NewVAlue) = 0 Or IsNull(NewVAlue), 0, NewVAlue)
  ObjXevents.STAs = "A"
  'ObjXevents.EventDetail = " the record " & Trim$(ssdbgRights.Columns(0).Text) & " has been changed to " & NewVAlue & "  for the User " & Trim$(ssdcboUserName.Columns(1).Text) & "."
  
End Sub

Private Sub ssdbgRights_GotFocus()
    If Not NavBar.SaveEnabled Then
        ssdbgRights.AllowUpdate = False
    End If
End Sub

'drop down document type data combo

Private Sub ssdbgRights_InitColumnProps()


'    With ssdbgRights
'         .Columns(0).Width = 3000
'         .Columns(0).Caption = "Document Type"
'         .Columns(0).Name = "doctype"
'         .Columns(0).DataField = "buyr_docutype"
'
'         .Columns(0).HeadStyleSet = "ColHeader"
'         .Columns(0).StyleSet = "RowFont"
'         .Columns(1).Width = 2650
'         .Columns(1).Caption = "Maximum Amount"
'         .Columns(1).Name = "amount"
'         .Columns(1).DataField = "buyr_maxiamou"
'
'         .Columns(1).HeadStyleSet = "ColHeader"
'         .Columns(1).StyleSet = "RowFont"
'         .Columns(2).Width = 5292
'         .Columns(2).Visible = False
'         .Columns(2).Caption = "user"
'         .Columns(2).Name = "user"
'         .Columns(2).DataField = "buyr_username"
'
'         .Columns(3).Width = 5292
'         .Columns(3).Visible = False
'         .Columns(3).Caption = "namespace"
'         .Columns(3).Name = "namespace"
'         .Columns(3).DataField = "buyr_npecode"
'    End With
    
    ssdbgRights.AllowUpdate = True
    ssdbgRights.Columns(0).DropDownHwnd = ssdbgDocType.Hwnd
End Sub

'assign value to recordset and call function to get buyers

Private Sub ssdcboUserName_Click()

    If AddingRecord Then
        NavBar.Recordset!buy_username = ssdcboUserName.Columns(1).Text
    Else
        'Call LockWindowUpdate(Hwnd)
        Call GetBuyers(ssdcboUserName.Columns(1).Text)
        'Call LockWindowUpdateOff
    End If
End Sub

Private Sub ssdcboUserName_GotFocus()
    If Not NavBar.SaveEnabled Then
        ssdcboUserName.AllowInput = False
    End If
End Sub


 'show and enable the selected tab's controls
'and hide and disable all others

Private Sub tbsOptions_Click()
    
    Dim i As Integer
   
    For i = 0 To tbsOptions.Tabs.count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).ZOrder
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = True
            'picOptions(i).Enabled = False
        End If
    Next
    
    picOptions(i - 1).Refresh
    If tbsOptions.SelectedItem.Index = 1 Then
        Set NavBar.Recordset = rsBuyer
    Else
        
        If ssdbgRights.Rows Then
        
            If Trim$(ssdbgRights.Columns("user").Text) <> _
                Trim$(ssdcboUserName.Columns(1).Text) Then _
                Call GetRights(ssdcboUserName.Columns(1).Text)
        Else
            Call GetRights(ssdcboUserName.Columns(1).Text)
        End If
    
    End If
    
End Sub

'call function to get datas for data grids

Public Sub SetConnection(con As ADODB.Connection)
 
    Set cn = con
    
    If IsNothing(cn) Then Exit Sub
    If Not IsConnectionOpen(con) Then Exit Sub
    
    Set rsBuyer = New ADODB.Recordset
    Set rsBuyer.ActiveConnection = con
    
   
    
    Call GetBuyerNames
  
    Call AddDocumentType
 
    ssdcboUserName.MoveFirst
  
    Call GetBuyers(ssdcboUserName.Columns(1).Text)
  

    Call GetRights(ssdcboUserName.Columns(1).Text)

    Set NavBar.Recordset = rsBuyer
End Sub

'set name space to current name space

Public Sub SetNameSpace(Namespace As String)
    np = Namespace
End Sub

'SQL statement to get buyer user name

Private Sub GetBuyers(UserId As String)
    
    With rsBuyer
        .CancelUpdate
        If .State And adStateOpen = adStateOpen Then .Close
        
        Set .ActiveConnection = cn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        
        .Source = "SELECT buy_username, buy_npecode, buy_builroom, "
        .Source = .Source & " buy_phonnumb , buy_extenumb FROM BUYER"
        .Source = .Source & " WHERE ( buy_npecode = '" & np & "')"
        .Source = .Source & " AND ( buy_username = '" & UserId & "')"
        .Source = .Source & " order by buy_username "
        .Open
        
    End With
    
    Set NavBar.Recordset = rsBuyer
    
    
    Set txt_Phone.DataSource = NavBar
    Set txt_Building.DataSource = NavBar
    Set txt_Extension.DataSource = NavBar
    Set ssdcboUserName.DataSource = NavBar
End Sub

'SQL statement to get buyer document type
'and maximum approval amount

Private Sub GetRights(UserId As String)
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    Set rsRight = New ADODB.Recordset
    Set rsRight.ActiveConnection = cn

    With cmd
        .Prepared = True
        Set .ActiveConnection = cn
        .CommandType = adCmdText
        
        .CommandText = "SELECT buyr_username, buyr_npecode, buyr_docutype, "
        .CommandText = .CommandText & " buyr_maxiamou  FROM BUYER_RIGHT"
        
        .CommandText = .CommandText & " WHERE ( buyr_npecode = ?)"
        .CommandText = .CommandText & " AND ( buyr_username = ?)"
        .CommandText = .CommandText & " order by buyr_username "
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, np)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, UserId)
        
    End With
    
    
    Set rsRight = cmd.Execute
    
    rsRight.Close
    
    With rsRight
        '.ActiveConnection = cn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
    End With
    
    Call rsRight.Open(, , , adLockBatchOptimistic)
    
    Set NavBar.Recordset = rsRight
    
    Set ssdbgRights.DataSource = rsRight
    ssdbgRights.Columns(0).DropDownHwnd = ssdbgDocType.Hwnd
End Sub

'SQL statement to get document type

Public Sub AddDocumentType()
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        .CommandType = adCmdText
        Set .ActiveConnection = cn
        
        .CommandText = "SELECT doc_desc, doc_code FROM DOCTYPE"
        .CommandText = .CommandText & " WHERE doc_npecode = '" & np & "'"
        .CommandText = .CommandText & " order by doc_code "
        
        Set ssdbgDocType.DataSource = .Execute
    End With
    
    Set cmd = Nothing
        
End Sub

'call function to get buyer names

Public Sub GetBuyerNames(Optional InList As Boolean = True)
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "GETBUYERNAMES"
        .CommandType = adCmdStoredProc
        
        Set ssdcboUserName.DataSourceList = .Execute(, Array(np, InList))
    End With
    Call DisableButtons(Me, NavBar, np, CurrentUser, cn)
    tbsOptions.Enabled = True
    picOptions(0).Enabled = True
    'ssdcboUserName.Enabled = True
    
End Sub

'look for buyer from buyer recordset

Private Sub FindBuyer()
On Error Resume Next
Dim str As String, i As Integer, X As Integer
    
    str = NavBar.Recordset(0)
    X = ssdcboUserName.Rows - 1
    
    Do While i < X
        i = i + 1
        If ssdcboUserName.Columns(1).Text = str Then
           'ssdcboUserName.Columns(0).Text = str
           Exit Do
        End If
        ssdcboUserName.MoveNext
    Loop
End Sub

Public Function checkFIELDS() As Boolean
checkFIELDS = False
   If Len(Trim$(txt_Phone)) = 0 Then
       MsgBox "The phone number can not be left empty.", vbInformation, "Imswin"
       txt_Phone.SetFocus
       'txt_Building.BackColor =
       Exit Function
   End If
   If Len(Trim$(txt_Extension)) = 0 Then
       MsgBox "The Extension field can not be left empty.", vbInformation, "Imswin"
       txt_Extension.SetFocus
       'txt_Building.BackColor =
       Exit Function
   End If
   checkFIELDS = True
End Function



