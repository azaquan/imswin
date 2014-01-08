VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_ShiptermsEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ship Terms and Conditions"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   7335
   Tag             =   "01010300"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin TabDlg.SSTab sst_Tab 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ship Terms"
      TabPicture(0)   =   "frm_ShipTermsEdit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Description"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Notes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSDBLine"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_Description"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dcboNotes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Text"
      TabPicture(1)   =   "frm_ShipTermsEdit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRemarks"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtRemarks 
         DataField       =   "stc_clau"
         DataMember      =   "SHIPTERM"
         DataSource      =   "deIms"
         Height          =   3255
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   420
         Width           =   6615
      End
      Begin MSDataListLib.DataCombo dcboNotes 
         Bindings        =   "frm_ShipTermsEdit.frx":0038
         DataField       =   "stc_note"
         DataMember      =   "SHIPTERM"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1830
         TabIndex        =   7
         Top             =   510
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "stc_note"
         Text            =   ""
         Object.DataMember      =   "SHIPTERM"
      End
      Begin VB.TextBox txt_Description 
         DataField       =   "stc_desc"
         DataMember      =   "SHIPTERM"
         DataSource      =   "deIms"
         Height          =   288
         Left            =   1830
         TabIndex        =   4
         Top             =   855
         Width           =   4035
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLine 
         Height          =   2445
         Left            =   240
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1260
         Width           =   7515
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
         FieldSeparator  =   ""
         DefColWidth     =   5292
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
         SelectByCell    =   -1  'True
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   1402
         Columns(0).Caption=   "Notes"
         Columns(0).Name =   "Notes"
         Columns(0).DataField=   "stc_note"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   7488
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "stc_desc"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Active"
         Columns(2).Name =   "Active"
         Columns(2).DataField=   "stc_actvflag"
         Columns(2).DataType=   11
         Columns(2).FieldLen=   256
         Columns(2).Style=   2
         TabNavigation   =   1
         _ExtentX        =   13256
         _ExtentY        =   4313
         _StockProps     =   79
         BackColor       =   -2147483638
         DataMember      =   "SHIPTERM"
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
      Begin VB.Label lbl_Notes 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   510
         Width           =   1600
      End
      Begin VB.Label lbl_Description 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   855
         Width           =   1600
      End
   End
End
Attribute VB_Name = "frm_ShiptermsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat
Private Sub PrintCurrent()
Dim Path As String
On Error GoTo ErrHandler

    Path = FixDir(App.Path) + "CRreports\"

      With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Shipterm.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "shiptermcode;" & dcboNotes.Text & ";TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00539") 'J added
        .WindowTitle = IIf(msg1 = "", "Shipment Terms", msg1) 'J modified
        Call translator.Translate_Reports("Shipterm.rpt") 'J added
        '---------------------------------------------

        .Action = 1
        .Reset
    End With
        Exit Sub

ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'get crystal report parameter and application path

Private Sub PrintAll()
      With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Shipterm.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "shiptermcode;ALL;TRUE"

        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00539") 'J added
        .WindowTitle = IIf(msg1 = "", "Shipment Terms", msg1) 'J modified
        Call translator.Translate_Reports("Shipterm.rpt") 'J added
        '---------------------------------------------

        .Action = 1
        .Reset
    End With
End Sub

'seacrh shipper term note

Private Sub dcboNotes_Click(Area As Integer)
Dim str As String

    If Area = 2 Then
        str = dcboNotes
        NavBar1.CancelUpdate
        deIms.rsSHIPTERM.CancelUpdate
        Call RecordsetFind(deIms.rsSHIPTERM, "stc_note = '" & str & "'")
    End If
End Sub

'set back ground color

Private Sub dcboNotes_GotFocus()
    Call HighlightBackground(dcboNotes)
End Sub

'set back ground color

Private Sub dcboNotes_LostFocus()
    Call NormalBackground(dcboNotes)
End Sub

'unlock note combo

Private Sub dcboNotes_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    dcboNotes.locked = False
End Sub

'if navbar is not new position lock note combo

Private Sub dcboNotes_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  dcboNotes.locked = Not NavBar1.NewEnabled
End Sub

'load form set recordset to text box

Private Sub Form_Load()



If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False


   SSDBLine.Columns("notes").locked = True
   SSDBLine.Columns("description").locked = True
''   SSDBGOrig.Columns("transaction type").locked = True
'   SSDBLine.Columns("active flag").locked = True




    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
 

   
    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_ShiptermsEdit")
    '------------------------------------------
    
    deIms.shipterm (deIms.NameSpace)
    DoEvents: DoEvents: DoEvents: DoEvents
    
    Call BindAll(Me, deIms)
    Set dcboNotes.RowSource = deIms
    Set NavBar1.Recordset = deIms.rsSHIPTERM
        Call DisableButtons(Me, NavBar1)
        
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
    NavBar1.EditVisible = True 'Juan Gonzalez 12/29/2006
    
    frm_ShiptermsEdit.Caption = frm_ShiptermsEdit.Caption + " - " + frm_ShiptermsEdit.Tag
Else

  Call translator.Translate_Forms("frm_ShiptermsEdit")
    '------------------------------------------
    
    deIms.shipterm (deIms.NameSpace)
    DoEvents: DoEvents: DoEvents: DoEvents
    
    Call BindAll(Me, deIms)
    Set dcboNotes.RowSource = deIms
    Set NavBar1.Recordset = deIms.rsSHIPTERM
        Call DisableButtons(Me, NavBar1)
        
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
    NavBar1.EditVisible = True 'Juan Gonzalez 12/29/2006
        
    frm_ShiptermsEdit.Caption = frm_ShiptermsEdit.Caption + " - " + frm_ShiptermsEdit.Tag




    TableLocked = True

End If
End If

Me.Left = Round((Screen.Width - Me.Width) / 2)
Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form, free memory,close recordset

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Hide
    deIms.rsSHIPTERM.Update
    deIms.rsSHIPTERM.UpdateBatch
    deIms.rsSHIPTERM.CancelBatch
    deIms.rsSHIPTERM.Close
    Set NavBar1.Recordset = Nothing
    Set frm_ShiptermsEdit = Nothing
    
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator

    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If



End Sub

Private Sub NavBar1_BeforeNewClick()
    NavBar1.EditEnabled = False 'Juan Gonzalez 12/29/2006
End Sub

'set recordset to update

Private Sub NavBar1_BeforeSaveClick()
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
    deIms.rsSHIPTERM.Update
End Sub

Private Sub NavBar1_OnCancelClick()
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
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
    NavBar1.CancelEnabled = True 'Juan Gonzalez 12/29/2006
    NavBar1.EditEnabled = False 'Juan Gonzalez 12/29/2006
    NavBar1.SaveEnabled = True 'Juan Gonzalez 12/29/2006
End Sub

'set name space equal to current name space

Private Sub NavBar1_OnNewClick()
    deIms.rsSHIPTERM!stc_npecode = deIms.NameSpace
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handled

Dim retval As PrintOpts

    Load frmPrintDialog
    'frmPrintDialog.optprintSel.Visible = False
    With frmPrintDialog

        .Show 1
        retval = .Result

        DoEvents: DoEvents
        If retval = poPrintCurrent Then

            PrintCurrent

        ElseIf retval = poPrintAll Then
            PrintAll

        Else
            Exit Sub

        End If

    End With

'    MDI_IMS.CrystalReport1 = "Supplier"
    ' MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset

    Unload frmPrintDialog
    Set frmPrintDialog = Nothing

Handled:
    If Err Then MsgBox Err.Description
'On Error GoTo ErrHandler
'    With MDI_IMS.CrystalReport1
'        .Reset
'        .ReportFileName = FixDir(App.Path) + "CRreports\Shipterm.rpt"
'        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'
'        'Modified by Juan (9/14/2000) for Multilingual
'        msg1 = translator.Trans("L00539") 'J added
'        .WindowTitle = IIf(msg1 = "", "Shipment Terms", msg1) 'J modified
'        Call translator.Translate_Reports("Shipterm.rpt") 'J added
'        '---------------------------------------------
'
'        .Action = 1
'        .Reset
'    End With
'        Exit Sub
'
'ErrHandler:
'    If Err Then
'        MsgBox Err.Description
'        Err.Clear
'    End If

End Sub

'set recordset to upadate

Private Sub NavBar1_OnSaveClick()
On Error Resume Next

    Call deIms.rsSHIPTERM.Move(0)
    'Call deIms.rsSHIPTERM.UpdateBatch 'Juan Gonzalez 12/29/2006
End Sub

'set back ground color

Private Sub txt_Description_GotFocus()
    Call HighlightBackground(txt_Description)
End Sub

'set back ground color

Private Sub txt_Description_LostFocus()
        Call NormalBackground(txt_Description)
End Sub
