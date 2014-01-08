VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_ShipTerms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipment Term and condition"
   ClientHeight    =   6165
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtNote 
      CausesValidation=   0   'False
      DataField       =   "stc_note"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   1860
      TabIndex        =   8
      Top             =   630
      Width           =   4000
   End
   Begin VB.TextBox txtRemarks 
      CausesValidation=   0   'False
      DataField       =   "stc_clau"
      DataSource      =   "deIms"
      Height          =   2010
      Left            =   1860
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   4000
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   5640
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      PrintEnabled    =   0   'False
      PrintVisible    =   0   'False
      AllowAddNew     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLine 
      CausesValidation=   0   'False
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   6975
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
      RecordSelectors =   0   'False
      FieldSeparator  =   ""
      DefColWidth     =   5292
      AllowUpdate     =   0   'False
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
      Columns(0).Width=   2011
      Columns(0).Caption=   "Notes"
      Columns(0).Name =   "Notes"
      Columns(0).DataField=   "stc_note"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   9075
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "stc_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Active"
      Columns(2).Name =   "Active"
      Columns(2).DataField=   "stc_actvflag"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      TabNavigation   =   1
      _ExtentX        =   12303
      _ExtentY        =   3836
      _StockProps     =   79
      BackColor       =   -2147483638
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
   Begin VB.TextBox txt_Description 
      CausesValidation=   0   'False
      DataField       =   "stc_desc"
      DataSource      =   "deIms"
      Height          =   288
      Left            =   1860
      TabIndex        =   3
      Top             =   948
      Width           =   4000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shipment Terms && Conditions"
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
      Top             =   105
      Width           =   7410
   End
   Begin VB.Label lbl_Remarks 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   396
      Left            =   120
      TabIndex        =   4
      Top             =   1332
      Width           =   2000
   End
   Begin VB.Label lbl_Description 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   228
      Left            =   120
      TabIndex        =   2
      Top             =   948
      Width           =   2000
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
      Height          =   228
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2000
   End
End
Attribute VB_Name = "frm_ShipTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cancelled As Boolean
Event Completed(Cancelled As Boolean, Terms As String)

'unload form

Private Sub cmdSelect_Click()
    Cancelled = False
    Unload Me
End Sub

Private Sub Form_Paint()
txt_Description.SetFocus
End Sub

'set back ground color

Private Sub txtNote_GotFocus()
    Call HighlightBackground(txtNote)
End Sub

'set back ground color

Private Sub txtNote_LostFocus()
    Call NormalBackground(txtNote)
End Sub

'unload form if error cause raise event

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu Then Cancelled = True
    
    If Not Cancelled Then _
        RaiseEvent Completed(Cancelled, txtRemarks.Text)
End Sub

'load recordset

Private Sub Form_Load()
On Error Resume Next

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_ShipTerms")
    '------------------------------------------
    
    If deIms.rsSHIPTERM2.State And adStateOpen Then
       Set NavBar1.Recordset = deIms.rsSHIPTERM2.Clone(adLockReadOnly)
        
    Else
        Call deIms.shipterm2(deIms.NameSpace)
        
        Set NavBar1.Recordset = deIms.rsSHIPTERM2.Clone(adLockReadOnly)
        
        deIms.rsSHIPTERM.Close
    End If
        

      
    NavBar1.Width = 0
    Call BindAll(Me, NavBar1)
    
    frm_ShipTerms.Caption = frm_ShipTerms.Caption + " - " + frm_ShipTerms.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Cancelled = True
    Unload Me
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) & "Reports\tandc1.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";True "
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00232") 'J added
        .WindowTitle = IIf(msg1 = "", "Ship Terms", msg1) 'J modified
        Call translator.Translate_Reports("tandc1.rpt") 'J added
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

Private Sub SSDBLine_DblClick()
    Cancelled = False
    Unload Me
End Sub

'set back ground color

Private Sub txt_Description_GotFocus()
    Call HighlightBackground(txt_Description)
End Sub

'set back ground color

Private Sub txt_Description_LostFocus()
    Call NormalBackground(txt_Description)
End Sub
