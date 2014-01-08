VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form FrmCopyPORemarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Po Remarks"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6375
   Begin VB.TextBox TxtTotalLI 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1000
      Width           =   615
   End
   Begin VB.TextBox TxtLineNumb 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1000
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1040
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1040
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   2175
      Left            =   120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   6135
   End
   Begin LRNavigators.LROleDBNavBar Navbar1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelVisible   =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      NewVisible      =   0   'False
      PrintVisible    =   0   'False
      SaveVisible     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBPO 
      Height          =   320
      Left            =   3600
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      Rows            =   1
      Columns(0).Width=   3200
      _ExtentX        =   3625
      _ExtentY        =   564
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Of"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label LblRemarks 
      Alignment       =   2  'Center
      Caption         =   "Po Remarks"
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
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lbl_PO2 
      Caption         =   "Purchase Order#"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   2145
   End
End
Attribute VB_Name = "FrmCopyPORemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAdd_Click()
Dim strremk As String
strremk = frm_NewPurchase.txtRemarks

If Len(txtRemarks) > 0 Then
 'frm_NewPurchase.txtRemarks = IIf(strremk = "", Null, strremk & vbCrLf) & txtRemarks & vbCrLf
 
 If InStr(strremk, "~") > 0 Then
    strremk = Replace(strremk, "~", vbCrLf & txtRemarks & vbCrLf)
 Else
    strremk = IIf(strremk = "", Null, strremk & vbCrLf) & txtRemarks & vbCrLf
 End If
 frm_NewPurchase.txtRemarks = strremk
 frm_NewPurchase.txtRemarks.SelStart = Len(strremk)
End If

Unload Me
End Sub

Private Sub cmdReplace_Click()

frm_NewPurchase.txtRemarks = txtRemarks & vbCrLf
frm_NewPurchase.txtRemarks.SelStart = Len(frm_NewPurchase.txtRemarks)
Unload Me
End Sub


Private Sub Form_Load()

Dim ObjPo As New imsPO
Dim RsPONumbs As New ADODB.Recordset
Dim RsRemarks As New ADODB.Recordset

ObjPo.NameSpace = deIms.NameSpace
Set RsPONumbs = ObjPo.GetAllPOnumb("POREM")

txtRemarks.locked = True

Do While Not RsPONumbs.EOF
   SSOleDBPO.AddItem RsPONumbs!POR_PONUMB
   RsPONumbs.MoveNext
Loop
   
Set NavBar1.ActiveConnection = deIms.cnIms
   
   Set NavBar1.Recordset = ObjPo.GetRemarks
   
   Set ObjPo = Nothing
   'SSOleDBPO.SetFocus
   
NavBar1.CancelLastSepVisible = False
NavBar1.LastPrintSepVisible = False
NavBar1.PrintSaveSepVisible = False
   
End Sub

Private Sub Form_Paint()
SSOleDBPO.SetFocus
Call HighlightBackground(SSOleDBPO)
End Sub

Private Sub Form_Unload(Cancel As Integer)
CleanUp
End Sub

Private Sub NavBar1_OnCloseClick()
Unload Me
End Sub

Private Sub NavBar1_OnFirstClick()

If Not NavBar1.Recordset.BOF Then
'  NavBar1.Recordset.MoveFirst
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnLastClick()

If Not NavBar1.Recordset.EOF Then
 ' NavBar1.Recordset.MoveLast
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnNextClick()

If Not NavBar1.Recordset.EOF Then
'  NavBar1.Recordset.MoveNext
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnPreviousClick()

If Not NavBar1.Recordset.BOF Then
  'NavBar1.Recordset.MovePrevious
  ShowRecords
End If

End Sub

Private Sub SSOleDBPO_Click()
NavBar1.Recordset.Filter = ""
NavBar1.Recordset.Filter = "por_ponumb='" & Trim$(SSOleDBPO.text) & "'"
 
 Dim Count As Integer
 Count = NavBar1.Recordset.RecordCount
 
If Count > 0 Then
   NavBar1.Recordset.MoveFirst
   
   If Count = 1 Then
      NavBar1.PreviousEnabled = False
      NavBar1.FirstEnabled = False
      NavBar1.NextEnabled = False
      NavBar1.LastEnabled = False
   Else
      NavBar1.PreviousEnabled = False
      NavBar1.FirstEnabled = False
      NavBar1.NextEnabled = True
      NavBar1.LastEnabled = True
   End If
   
   ShowRecords
   
Else
   txtRemarks.text = ""
   TxtLineNumb = ""
   NavBar1.PreviousEnabled = False
      NavBar1.FirstEnabled = False
      NavBar1.NextEnabled = False
      NavBar1.LastEnabled = False
   'NavBar1.PreviousEnabled = NavBar1.FirstEnabled = _


   
End If


End Sub


Public Sub ShowRecords()

If Not (NavBar1.Recordset.EOF = True Or NavBar1.Recordset.BOF = True) Then
txtRemarks.text = IIf(IsNull(NavBar1.Recordset!por_remk) = True, "", NavBar1.Recordset!por_remk)
TxtLineNumb = NavBar1.Recordset!por_linenumb
TxtTotalLI = NavBar1.Recordset.RecordCount
End If
   
End Sub

Public Sub CleanUp()
On Error Resume Next
NavBar1.Recordset.Close
If Err.number > 0 Then Err.Clear
Set NavBar1.Recordset = Nothing
End Sub

Private Sub SSOleDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBPO.DroppedDown Then SSOleDBPO.DroppedDown = True
End Sub
Private Sub SSOleDBPO_GotFocus()
Call HighlightBackground(SSOleDBPO)
End Sub


Private Sub SSOleDBPO_LostFocus()
Call NormalBackground(SSOleDBPO)
End Sub
