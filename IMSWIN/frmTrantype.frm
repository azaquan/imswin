VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmTrantype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Type "
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   5745
   Tag             =   "01030100"
   Begin LRNavigators.NavBar NavBar1 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   960
      TabIndex        =   2
      Top             =   3360
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frmTrantype.frx":0000
      CloseToolTipText=   ""
      PrintToolTipText=   ""
      EmailToolTipText=   ""
      NewToolTipText  =   ""
      SaveToolTipText =   ""
      CancelToolTipText=   ""
      NextToolTipText =   ""
      LastToolTipText =   ""
      FirstToolTipText=   ""
      PreviousToolTipText=   ""
      DeleteToolTipText=   ""
      EditToolTipText =   ""
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgTrantype 
      Height          =   2295
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   5355
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
      stylesets(0).Picture=   "frmTrantype.frx":001C
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
      stylesets(1).Picture=   "frmTrantype.frx":0038
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowGroupSizing=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1296
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "tty_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   2
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   5609
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "tty_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   25
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   1614
      Columns(2).Caption=   "Sign 1=+"
      Columns(2).Name =   "Sign"
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "tty_sign"
      Columns(2).DataType=   11
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(2).HeadStyleSet=   "ColHeader"
      Columns(2).StyleSet=   "RowFont"
      Columns(3).Width=   5292
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "tty_npecode"
      Columns(3).Name =   "NameSpace"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "tty_npecode"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   9446
      _ExtentY        =   4048
      _StockProps     =   79
      BackColor       =   -2147483643
      DataMember      =   "Get_Transaction_Type"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Transaction Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   5430
   End
End
Attribute VB_Name = "frmTrantype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'load form and get transaction type recordset and set navbar buttom

Private Sub Form_Load()
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim l As Long

    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmTrantype")
    '------------------------------------------

    Set cmd = deIms.Commands("Get_Transaction_Type")
    Set rs = deIms.rsGet_Transaction_Type
    
    If ((rs.State And adStateOpen) = adStateOpen) Then rs.Close
    cmd.Parameters("@NAMESPACE").Value = deIms.NameSpace
    cmd.Execute
    
    l = cmd.Parameters("Return_Value")
    Set ssdbgTrantype.DataSource = deIms
    Call DisableButtons(Me, NavBar1)
    frmTrantype.Caption = frmTrantype.Caption + " - " + frmTrantype.Tag
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'cancel recordset update

Private Sub NavBar1_OnCancelClick()
    ssdbgTrantype.CancelUpdate
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    ssdbgTrantype.MoveFirst
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    ssdbgTrantype.MoveLast
End Sub

'move recordset to add new position and
'set name space to currrent name space

Private Sub NavBar1_OnNewClick()
    ssdbgTrantype.AddNew
    ssdbgTrantype.Columns("Sign").Value = 0
    ssdbgTrantype.Columns("NameSpace").Value = deIms.NameSpace
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
    ssdbgTrantype.MoveNext
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
    ssdbgTrantype.MovePrevious
End Sub

'get parameters to print crystal report

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Transtype.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("M00112") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction Type", msg1) 'J modified
        Call translator.Translate_Reports("Transtype.rpt") 'J added
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

Private Sub NavBar1_OnSaveClick()
    ssdbgTrantype.Update
End Sub

'validate data format and show messege

Private Sub NavBar1_Validate(Cancel As Boolean)
    
    Cancel = True
    With ssdbgTrantype
        If Trim$(.Columns("Code").Value) = "" Then
        
            'Modified by Juan (9/25/2000) for Multilingual
            msg1 = translator.Trans("M00014") 'J added
            MsgBox IIf(msg1 = "", "Code cannot be left empty", msg1) 'J modified
            '---------------------------------------------
            
            Exit Sub
        End If
        
        If Trim$(.Columns("Description").Value = "") Then
        
            'Modified by Juan (9/25/2000) for Multilingual
            msg1 = translator.Trans("M00372") 'J added
            MsgBox IIf(msg1 = "", "Description canot be left empty", msg1) 'J modified
            '---------------------------------------------

            Exit Sub
        End If
    End With
    
    Cancel = False
End Sub

