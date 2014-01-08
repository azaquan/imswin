VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAllActivityReport 
   Caption         =   "All Activity Report"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton getReport 
      Caption         =   "&Get Report"
      Height          =   320
      Left            =   7800
      TabIndex        =   12
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   4560
      Width           =   1092
   End
   Begin VB.TextBox resultsText 
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   720
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4560
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4560
      Width           =   1092
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_userid 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      Columns(0).Width=   3200
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComCtl2.DTPicker DTfromdate 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61014017
      CurrentDate     =   36523
   End
   Begin MSComCtl2.DTPicker DTtodate 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61014017
      CurrentDate     =   36523
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Send Email to:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Id"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   645
   End
   Begin VB.Label lbl_fromdate 
      Alignment       =   1  'Right Justify
      Caption         =   "From Date"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label lbl_todate 
      Alignment       =   1  'Right Justify
      Caption         =   "To Date"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmAllActivityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Source = "select usr_USERID  from xuserprofile where usr_npecode='" & deIms.NameSpace & "'"
    rs.ActiveConnection = deIms.cnIms
    rs.Open
    Combo_userid.AddItem "ALL"
    Do While (Not rs.EOF)
        Combo_userid.AddItem (rs!usr_userid)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    DTfromdate = FirstOfMonth
    DTtodate = Now

    With frmAllActivityReport
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub


Private Sub getReport_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim result As Variant
    Sql = "execute getAllTables '" + Format(DTfromdate, "yyyy/mm/dd") + "', '" + Format(DTtodate, "yyyy/mm/dd") + "', " + Combo_userid.value
    rs.Open Sql, deIms.cnIms, adOpenForwardOnly
    If rs.State > 0 Then
        resultsText = rs.GetString
    End If
End Sub
