VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_ConvertToPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert a transaction"
   ClientHeight    =   1695
   ClientLeft      =   645
   ClientTop       =   930
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   4675
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssOleDbPO 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton CmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Requisition Number"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   255
      Width           =   2040
   End
End
Attribute VB_Name = "frm_ConvertToPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConvert_Click()
''Dim CmdConvertToRequisition As ADODB.Command
''Set CmdConvertToRequisition = New ADODB.Command

If Len(Trim(ssOleDbPO)) = 0 Or ssOleDbPO.IsItemInList = False Then
    
    MsgBox " Please fill in a Valid Transaction Number.", vbInformation, "Ims"
    
    ssOleDbPO.SetFocus
    
    Exit Sub
    
End If

Call frm_NewPurchase.ConvertRequisition(Trim(ssOleDbPO.Text))


End Sub

Private Sub Form_Load()

Me.Height = 2100
Me.Width = 4800

    'Added by Juan (2015-02-13) for Multilingual
    Call translator.Translate_Forms("frm_ConvertToPO")
    '------------------------------------------

Call GetPosForDoctype
ssOleDbPO.Columns(0).Width = ssOleDbPO.Width

    With frm_ConvertToPO
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub


Public Function GetPosForDoctype()

Dim rs As New ADODB.Recordset

On Error GoTo ErrHandler

'rs.Source = "select po_ponumb, po_buyr from po where po_docutype = '" & DocumentType & "' and po_npecode='" & deIms.NameSpace & "' AND LEN(RTRIM(ISNULL(po_apprby,''))) > 0 AND po_buyr ='" & CurrentUser & "'"
'rs.Source = "select po_ponumb, po_buyr from po where po_npecode='" & deIms.NameSpace & "' AND LEN(RTRIM(ISNULL(po_apprby,''))) > 0 AND po_buyr ='" & CurrentUser & "'"
rs.Source = "select po_ponumb, po_buyr from po where po_npecode='" & deIms.NameSpace & "' order by po_ponumb"

rs.ActiveConnection = deIms.cnIms

rs.Open
        
        Do While Not rs.EOF
           
           ssOleDbPO.AddItem rs!PO_PONUMB
           rs.MoveNext
           
        Loop

Exit Function

ErrHandler:

MsgBox "Errors Occurred while tying to get the Requisitions. Error Description : " & Err.Description

Err.Clear
End Function

Private Sub SSOleDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If Not ssOleDbPO.DroppedDown Then ssOleDbPO.DroppedDown = True
 
End Sub

Private Sub ssOleDbPO_KeyPress(KeyAscii As Integer)
 ssOleDbPO.MoveNext
End Sub
