VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form FrmCopyPOItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Order Items"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7800
   Begin VB.TextBox TxtTotalLI 
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1275
      Width           =   615
   End
   Begin VB.TextBox TxtStknumb 
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1755
      Width           =   2175
   End
   Begin VB.TextBox TxtLineNumb 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1275
      Width           =   615
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   6375
      TabIndex        =   2
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1755
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBPO 
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      Rows            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3625
      _ExtentY        =   564
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin LRNavigators.LROleDBNavBar Navbar1 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
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
   Begin VB.TextBox txt_Desc1 
      Height          =   675
      Left            =   1200
      MaxLength       =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   6420
   End
   Begin VB.TextBox txt_Remk1 
      Height          =   675
      Left            =   1200
      MaxLength       =   256
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2880
      Width           =   6420
   End
   Begin VB.Label Lbl_StkNUmb 
      Caption         =   "StockNumber"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblLineNumb 
      Caption         =   "Of"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label LblPoitems 
      Alignment       =   2  'Center
      Caption         =   "Transaction Order Items"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lbl_Description 
      Caption         =   "Description"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   975
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
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   2025
   End
End
Attribute VB_Name = "FrmCopyPOItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim non_stock As Boolean
Private Sub cmdAdd_Click()

Dim StrDesc As String
Dim strremk As String
  Call IsItNon_Stock(non_stock)
  
   StrDesc = frm_NewPurchase.txt_Descript
   strremk = frm_NewPurchase.txt_remk
If non_stock = True Then
    If Len(Me.txt_Desc1) > 0 Then
     frm_NewPurchase.txt_Descript.text = IIf(StrDesc = "", Null, StrDesc & vbCrLf) & Me.txt_Desc1.text
    End If
End If
If Len(txt_Remk1.text) > 0 Then
 frm_NewPurchase.txt_remk.text = IIf(strremk = "", Null, strremk & vbCrLf) & Me.txt_Remk1.text
End If
Unload Me

End Sub

Private Sub cmdReplace_Click()

Call IsItNon_Stock(non_stock)

If non_stock = True Then
  If Len(Me.txt_Desc1) > 0 Then
  frm_NewPurchase.txt_Descript = txt_Desc1.text
  End If
End If

If Len(txt_Remk1.text) > 0 Then
frm_NewPurchase.txt_remk = txt_Remk1.text
End If

Unload Me
End Sub

Private Sub Form_Load()
Dim ObjPo As New imsPO
Dim RsPONumbs As New ADODB.Recordset


non_stock = False
txt_Remk1.locked = True
txt_Desc1.locked = True


ObjPo.NameSpace = deIms.NameSpace

Set RsPONumbs = ObjPo.GetAllPOnumb

Do While Not RsPONumbs.EOF
   SSOleDBPO.AddItem RsPONumbs!PO_PONUMB
   RsPONumbs.MoveNext
Loop

 Set Me.NavBar1.ActiveConnection = deIms.cnIms
   
 Set Me.NavBar1.Recordset = ObjPo.getLINEitems
 Set ObjPo = Nothing
 
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
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnLastClick()

If Not NavBar1.Recordset.EOF Then
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnNextClick()

If Not NavBar1.Recordset.EOF Then
  ShowRecords
End If

End Sub

Private Sub NavBar1_OnPreviousClick()

If Not NavBar1.Recordset.BOF Then
  ShowRecords
End If

End Sub

Private Sub SSOleDBPO_Click()
NavBar1.Recordset.Filter = ""
NavBar1.Recordset.Filter = "poi_ponumb='" & Trim$(SSOleDBPO.text) & "'"
 
If NavBar1.Recordset.RecordCount > 0 Then
   NavBar1.Recordset.MoveFirst
   
   If NavBar1.Recordset.RecordCount = 1 Then
      
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

   NavBar1.PreviousEnabled = False
   NavBar1.FirstEnabled = False
   NavBar1.NextEnabled = False
   NavBar1.LastEnabled = False
   
   Me.txt_Desc1 = ""
   Me.txt_Remk1 = ""
   
End If


End Sub


Public Sub ShowRecords()
If Not (NavBar1.Recordset.EOF = True Or NavBar1.Recordset.BOF = True) Then

txt_Desc1.text = IIf(IsNull(NavBar1.Recordset!poi_desc) = True, "", NavBar1.Recordset!poi_desc)
txt_Remk1.text = IIf(IsNull(NavBar1.Recordset!poi_remk) = True, "", NavBar1.Recordset!poi_remk)
TxtLineNumb = NavBar1.Recordset!poi_liitnumb
TxtStknumb = NavBar1.Recordset!poi_comm
TxtTotalLI = NavBar1.Recordset.RecordCount
End If
   
End Sub

Public Sub CleanUp()
On Error Resume Next
NavBar1.Recordset.Close
If Err.number > 0 Then Err.Clear
Set NavBar1.Recordset = Nothing
End Sub


Public Sub IsItNon_Stock(non_stock As Boolean)
  
Dim ObjPo As New imsPO
Dim rsSTOCKMASTER As ADODB.Recordset

'ObjPo.NameSpace = deIms.NameSpace
'Set rsSTOCKMASTER = ObjPo.GetStocKsFromStockmaster

'Here this filter should be done with the records from Stockmaster.
 
    'If Len(Trim$(frm_NewPurchase.ssdcboCommoditty.text)) = 0 Then
    If frm_NewPurchase.chk_FrmStkMst.Value = 0 Then
      non_stock = True
      
    Else
    
       non_stock = False
    End If
    
    '  rsSTOCKMASTER.Find "stk_stcknumb='" & Trim$(frm_NewPurchase.ssdcboCommoditty.text) & "'"
    
   '   If rsSTOCKMASTER.EOF = True Then
   '     non_stock = True
   '   Else
   '     non_stock = False
   '   End If
    
   '    rsSTOCKMASTER.Close
   '    Set rsSTOCKMASTER = Nothing
    
'    End If
    
    Set ObjPo = Nothing
    
    
'     txt_Desc1.Visible = Not non_stock
'     lbl_Description.Visible = Not non_stock
'     txt_Remk1.Top = txt_Desc1.Top
'     txt_Remk1.Height = txt_Remk1.Height + txt_Desc1.Height
     
End Sub

Private Sub SSOleDBPO_GotFocus()
Call HighlightBackground(SSOleDBPO)
End Sub

Private Sub SSOleDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBPO.DroppedDown Then SSOleDBPO.DroppedDown = True
End Sub

Private Sub SSOleDBPO_LostFocus()
Call NormalBackground(SSOleDBPO)
End Sub

