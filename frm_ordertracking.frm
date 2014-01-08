VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_ordertracking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Tracking"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   4230
   Tag             =   "03020500"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo combo_ponumb 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      DataFieldList   =   "Column 0"
      _Version        =   196617
      Cols            =   1
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label lbl_ponumb 
      Caption         =   "Po Number"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   645
      Width           =   1695
   End
End
Attribute VB_Name = "frm_ordertracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and applition path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\ordertracking.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Trim$(combo_ponumb.Text) + ";true"
        
        'Modified by Juan (9/13/2000) for Mutilingual
        msg1 = translator.Trans("frm_ordertracking") 'J added
        .WindowTitle = IIf(msg1 = "", "Order Tracking", msg1) 'J modified
        Call translator.Translate_Reports("ordertracking.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '--------------------------------------------
        
        .Action = 1: .Reset
       End With
           Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

Private Sub combo_ponumb_DropDown()
    With combo_ponumb
        .MoveNext
    End With
End Sub

Private Sub combo_ponumb_GotFocus()
Call HighlightBackground(combo_ponumb)
End Sub

Private Sub combo_ponumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not combo_ponumb.DroppedDown Then combo_ponumb.DroppedDown = True
End Sub

Private Sub combo_ponumb_KeyPress(KeyAscii As Integer)
'combo_ponumb.MoveNext
End Sub

Private Sub combo_ponumb_LostFocus()
Call NormalBackground(combo_ponumb)
End Sub

Private Sub combo_ponumb_Validate(Cancel As Boolean)
If Len(Trim$(combo_ponumb)) > 0 Then
         If Not combo_ponumb.IsItemInList Then
                combo_ponumb.Text = ""
            End If
            If Len(Trim$(combo_ponumb)) = 0 Then
           combo_ponumb.SetFocus
            Cancel = True
            End If
            End If
End Sub

'SQL statement get po number and populate recordset

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
 
 'Added by Juan (9/13/2000) for Mutilingual
 Call translator.Translate_Forms("frm_ordertracking")
 '-----------------------------------------
 
 'Me.Height = 2520
 'Me.Width = 3800
 
  With rs
        .ActiveConnection = deIms.cnIms
  .Source = "select po_ponumb from po where po_npecode  ='" & deIms.NameSpace & "' order by po_ponumb "
  .Open , , adOpenStatic
  End With
  If Not ((rs Is Nothing) Or (rs.State And adStateOpen = adStateClosed) _
   Or (rs.EOF And rs.BOF) Or (rs.RecordCount = 0)) Then
   'Call PopuLateFromRecordSet(combo_ponumb, rs, "po_ponumb", True)
'   combo_ponumb.AddItem "ALL"
'  Do While (Not rs.EOF)
'
'  combo_ponumb.AddItem (rs!PO_PONUMB)
'  rs.MoveNext
'  Loop
   
        With combo_ponumb
            Set .DataSourceList = rs ' deIms.Commands("getStockOnHandQTYST1").Execute(100, Array(0, deIms.NameSpace))
            .DataFieldToDisplay = "po_ponumb"
            .DataFieldList = "po_ponumb"
            .Refresh
        End With
   
   
   
  Set rs = Nothing
  Else
    Exit Sub
  End If
  
  Caption = Caption + " - " & Tag
  
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub


'resize form

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then
 'Me.Height = 2520
 'Me.Width = 3800
 End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub
