VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#8.0#0"; "LRNavigators.ocx"
Begin VB.Form frm_Currency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Currency"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   Tag             =   "01010500"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frm_Currency.frx":0000
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1650
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2910
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info"
      TabPicture(0)   =   "frm_Currency.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Conversion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Name"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Code"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_Value"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt_Name"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Rates"
      TabPicture(1)   =   "frm_Currency.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtpEndDate"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConversion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dtpStartDate"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.ComboBox cboCode 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         DataField       =   " "
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   7
         Top             =   825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   68812803
         CurrentDate     =   36549.4640972222
      End
      Begin VB.TextBox txtConversion 
         Alignment       =   1  'Right Justify
         DataField       =   " "
         Height          =   315
         Left            =   -73080
         TabIndex        =   0
         Top             =   1170
         Width           =   1575
      End
      Begin VB.TextBox txt_Name 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   830
         Width           =   2775
      End
      Begin VB.TextBox txt_Value 
         DataField       =   "curr_convration"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   1170
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         DataField       =   " "
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   68812803
         CurrentDate     =   36549.4635532407
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   195
         Left            =   -74880
         TabIndex        =   13
         Top             =   830
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   1170
         Width           =   1800
      End
      Begin VB.Label lbl_Code 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label lbl_Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   825
         Width           =   1800
      End
      Begin VB.Label Conversion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convertion Ratio"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1170
         Width           =   1800
      End
   End
   Begin VB.Label lbl_Currency 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   135
      TabIndex        =   8
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frm_Currency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Nav As imsNav
Attribute Nav.VB_VarHelpID = -1
Dim WithEvents cr As imsCurrency
Attribute cr.VB_VarHelpID = -1
Dim WithEvents crd As imsCurrencyDetl
Attribute crd.VB_VarHelpID = -1

Private Sub cboCode_Click()
    If cr.Navigator.Find("Code = '" & cboCode & "'") = False Then
    LoadCurrency
    End If
End Sub

Private Sub cboCode_Validate(Cancel As Boolean)
    If Len(cboCode) < 4 Then
      cr.Code = cboCode
    Else
    
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00391") 'J added
        MsgBox IIf(msg1 = "", "The field 'Code' can not have more than 3 Characters", msg1) 'J modified
        '---------------------------------------------

      Cancel = True
    End If
    
    If cr.IsCodeRepeated(deIms.NameSpace) = True Then
    Cancel = True
    
    'Modified by Juan (9/27/2000) for Multilingual
    msg1 = translator.Trans("M00013") 'J added
    MsgBox IIf(msg1 = "", "The code Already Exist, Please Enter a different One", msg1) 'J modified
    '---------------------------------------------

    End If
End Sub


Private Sub cr_OnError(Description As String)
    If Len(Description) Then MsgBox Description
End Sub

Private Sub cr_OnMoveComplete()
    LoadCurrency
    If Not cr.Inserting Then
        Call crd.GetValues(deIms.cnIms, cr.NameSpace, cr.Code)
    Else
        Call crd.GetValues(deIms.cnIms, cr.NameSpace, "")
        crd.Navigator.AddNew
    End If
        
End Sub

Private Sub crd_OnError(Description As String)
    If Len(Description) Then MsgBox Description
End Sub

Private Sub crd_OnMoveComplete()
    LoadCurrencyDetl
End Sub

Private Sub Form_Load()

    'Added by Juan (9/27/2000) for Multilingual
    Call translator.Translate_Forms("frm_Currency")
    '------------------------------------------

    Set cr = New imsCurrency
    Set crd = New imsCurrencyDetl
    
    'deIms.cnIms.Open
    Set Nav = cr
    Set crd.imsCurrency = cr
    
    Call cr.GetValues(deIms.cnIms, deIms.NameSpace)
    
    Call PopuLateFromRecordSet(cboCode, cr.CodeList _
        (deIms.cnIms, deIms.NameSpace), "code", True)
    
    LoadCurrency
    Call DisableButtons(Me, NavBar1)
    Caption = Caption + " - " + Tag
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set cr = Nothing
    Set crd = Nothing
    Set Nav = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub NavBar1_Click()
    UpdateFields
End Sub

Private Sub NavBar1_OnCancelClick()
    Nav.CancelUpdate
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnFirstClick()
    If Nav.Validate Then Nav.MoveFirst
End Sub

Private Sub NavBar1_OnLastClick()
    If Nav.Validate Then Nav.MoveLast
End Sub

Private Sub NavBar1_OnNewClick()
  On Error Resume Next  'M
   txtConversion.Enabled = True
    If TypeName(Nav) = TypeName(crd) And Nav.AbsolutePosition = -1 Then
    Nav.AddNew
    Exit Sub
    End If
    If Nav.Validate Then Nav.AddNew
    
End Sub

Private Sub NavBar1_OnNextClick()
    If Nav.Validate Then Nav.MoveNext
End Sub

Private Sub NavBar1_OnPreviousClick()
    If Nav.Validate Then Nav.MovePrevious
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Currency.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("L00047") 'J added
        .WindowTitle = IIf(msg1 = "", "Currency", msg1) 'J modified
        Call translator.Translate_Reports("Currency.rpt") 'J added
        '---------------------------------------------

        .Action = 1
        .Reset
        
    End With
'    MDI_IMS.CrystalReport1.Reset
'    MDI_IMS.CrystalReport1.ReportFileName = FixDir(App.Path) + "CRreports\Currency.rpt"
'    MDI_IMS.CrystalReport1.ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'    MDI_IMS.CrystalReport1.Action = 1
'    MDI_IMS.CrystalReport1.WindowTitle = "Currency"
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If

End Sub

Private Sub LoadCurrency()
    cboCode = cr.Code
    txt_Name = cr.Description
    txt_Value = cr.ConversionRatio & ""
End Sub

Private Sub NavBar1_OnSaveClick()

    'If SSTab1.Tab = 0 Then  'M commented this out.
    
    If cr.Navigator.Editting = True Then   'M
    
    'If SSTab1.Tab = 0 Or (SSTab1.Tab = 1 And cr.Navigator.Editting = True) Then  'M
        If cr.Navigator.Validate Then
        
            If crd.Navigator.Validate Then
                cr.Navigator.Update
                crd.Navigator.UpdateBatch
            End If
        End If
        'ELSE                                   'M Commented out
        'If Nav.Validate Then Nav.Update        'M COMMENTED OUT
    ElseIf crd.Navigator.Editting = True Then   'M
        If crd.Navigator.Validate Then crd.Navigator.Update
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    If SSTab1.Tab = 0 Then
        Set Nav = cr
        NavBar1.SaveEnabled = True
    Else
        
        Set Nav = crd
        
      '  If Not Nav.Editting = True Then  'M
      '     txtConversion.Enabled = False 'M
      '  Else                             'M
         '  txtConversion.Enabled = True  'M
      '  End If                           'M
        
      NavBar1.SaveEnabled = False      'M
    End If
    
End Sub

Private Sub txt_Name_Validate(Cancel As Boolean)
    cr.Description = txt_Name
End Sub

Private Sub txt_Value_Validate(Cancel As Boolean)
    
    If Len(Trim$(txt_Value)) = 0 Then Exit Sub
    
    If Not IsNumeric(txt_Value) Then
        Cancel = True
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00012") 'J added
        MsgBox IIf(msg1 = "", "Conversion Ratio has to be a number", msg1) 'J modified
        '---------------------------------------------
        
        Exit Sub
    End If
    
    cr.ConversionRatio = txt_Value
End Sub

Private Sub UpdateFields()

    If SSTab1.Tab = 0 Then
        Call UpdateCurrency
        Call UpdateCurrencyDetl
    ElseIf SSTab1.Tab = 1 Then
        Call UpdateCurrencyDetl
    End If
End Sub

Private Sub UpdateCurrency()
    cr.Code = cboCode
    cr.Description = txt_Name
    cr.NameSpace = deIms.NameSpace
    cr.ConversionRatio = txt_Value
End Sub

Private Sub UpdateCurrencyDetl()
On Error Resume Next

    crd.Value = txtConversion
    crd.EndDate = dtpEndDate.Value
    crd.Startdate = dtpStartDate.Value
    crd.imsCurrency.NameSpace = deIms.NameSpace
    
    If Err Then Err.Clear
End Sub

Public Sub LoadCurrencyDetl()
    
  'REASON - if there was no detl record in the Second tab
  'then it would simply display a 0,which the user might think
  'try to modify resulting in the Value not Save.The If statement
  'basically disable the text box in such a Scenario.
      
    
    txtConversion = crd.Value
    dtpEndDate.Value = crd.EndDate
    dtpStartDate.Value = crd.Startdate
    
    If IsNull(dtpEndDate.Value) And IsNull(dtpStartDate.Value) And cr.Navigator.Editting = False Then 'M
      txtConversion.Enabled = False    'M
      Else 'M
      txtConversion.Enabled = True  'M
      End If  'M
    
End Sub

