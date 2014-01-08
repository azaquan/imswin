VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frm_StockType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   7500
   Tag             =   "01012100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   1830
      TabIndex        =   2
      Top             =   3870
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   767
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowUpdate     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      DeleteToolTipText=   ""
      Mode            =   0
      CommandType     =   0
      CursorLocation  =   0
      CommandType     =   0
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGStockType 
      Height          =   3195
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   7260
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
      stylesets(0).Picture=   "frm_StockType.frx":0000
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frm_StockType.frx":001C
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1111
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "sty_stcktype"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5345
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "sty_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2910
      Columns(2).Caption=   "Capital/Expense"
      Columns(2).Name =   "Capital"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "sty_cenc"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   900
      Columns(3).Caption=   "Own"
      Columns(3).Name =   "Own"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "sty_owle"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(4).Width=   1826
      Columns(4).Caption=   "Idea Flag"
      Columns(4).Name =   "Idea"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "sty_idaeflag"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(5).Width=   5292
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "sty_npecode"
      Columns(5).Name =   "NameSpace"
      Columns(5).CaptionAlignment=   0
      Columns(5).DataField=   "sty_npecode"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   12806
      _ExtentY        =   5636
      _StockProps     =   79
      DataMember      =   "StockType"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\My Documents\IMSWin\CRreports\stocktype.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      Connect         =   """DSN = RABBITFOOT; UID = sa; PWD = ;"""
      UserName        =   "sa"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_StockType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Type"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3439
      TabIndex        =   0
      Top             =   60
      Width           =   1776
   End
End
Attribute VB_Name = "frm_StockType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Unload(Cancel As Integer)
If open_forms <= 5 Then frmNavigator.Visible = True
End Sub

Private Sub NavBar1_BeforeNewClick()
    SSDBGStockType.AddNew
    SSDBGStockType.Columns("Own").Value = 0
    SSDBGStockType.Columns("Idea").Value = 0
    SSDBGStockType.Columns("NameSpace").Value = deIms.Namespace
   
End Sub

Private Sub NavBar1_OnEMailClick()
    'Send Via Email
    Dim Email_Error As Integer
    Dim Message As String

    Message = Space(34) & "Stock Type" & vbCrLf
    Message = Message & "" & vbCrLf
    Message = Message & "        " & Format(Date, "MM/DD/YYYY") & vbCrLf
    Message = Message & "        " & Format(Time, "HH:MM:SS") & vbCrLf
    Message = Message & "" & vbCrLf
    Message = Message & "        Code        : " & SSDBGStockType.Columns("Code").CellValue(SSDBGStockType.RowBookmark(SSDBGStockType.Row)) & vbCrLf
    Message = Message & "        Description : " & SSDBGStockType.Columns("Description").CellValue(SSDBGStockType.RowBookmark(SSDBGStockType.Row)) & vbCrLf
    Message = Message & "" & vbCrLf
'    Email_Error = send_mail("c:\attmsg\out\test.msq", lst_Destination, "Stock Type", message)
'    Debug.Print Email_Error
    If Email_Error = 0 Then
        MsgBox "Message Queued"
    Else
        Err.Number = Email_Error
        MsgBox Err.Number
        MsgBox Err.Description
    End If
End Sub

Private Sub Form_Load()
Dim ctl As Control
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim l As Long

'    Set cmd = deIms.Commands("STOCKTYPE")
'    Set rs = deIms.rsSTOCKTYPE
'
'    If ((rs.State And adStateOpen) = adStateOpen) Then rs.Close
'    cmd.Parameters("NAMESPACE").Value = deIms.NameSpace
'    cmd.Execute
'
'    l = cmd.Parameters("Return_Value")
'    Set SSDBGStockType.DataSource = deIms
'
'    Screen.MousePointer = vbHourglass
'    'color the controls and form backcolor
'    'Me.BackColor = frm_Color.txt_WBackground.BackColor
'
'    For Each ctl In Controls
'        Call gsb_fade_to_black(ctl)
'    Next ctl
'
'
'    Set NavBar1.Recordset = deIms.rsSTOCKTYPE
'    Screen.MousePointer = vbDefault
'
'    Show
End Sub

Private Sub NavBar1_BeforeCancelClick()
    SSDBGStockType.CancelUpdate
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\Stocktype.rpt"
        .ParameterFields(0) = "namespace;" + deIms.Namespace + ";TRUE"
        .Action = 1: .Reset
    End With
     Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

Private Sub NavBar1_BeforeSaveClick()
    SSDBGStockType.Update
End Sub


Private Sub SSDBGStockType_Validate(Cancel As Boolean)

    Cancel = True
    
    If Len(Trim$(SSDBGStockType.Columns("Code").Value)) = 0 Then
        MsgBox "Code Cannot be left empty"
        Exit Sub
    
    End If
    
    If Len(Trim$(SSDBGStockType.Columns("Description").Value)) = 0 Then
        MsgBox "Description Cannot be left empty"
        Exit Sub
    End If
    
    If Len(Trim$(SSDBGStockType.Columns("Capital").Value)) = 0 Then
        MsgBox "Capital/Expense Cannot be left empty"
        Exit Sub
    End If
    
    Cancel = False
End Sub

