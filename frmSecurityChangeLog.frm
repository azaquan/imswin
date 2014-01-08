VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSecurityChangeLog 
   Caption         =   "Security Change Log"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Tag             =   "03050400"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_userid 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1092
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   1092
   End
   Begin MSComCtl2.DTPicker DTfromdate 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60030977
      CurrentDate     =   36523
   End
   Begin MSComCtl2.DTPicker DTtodate 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60030977
      CurrentDate     =   36523
   End
   Begin VB.Label Label1 
      Caption         =   "User Id"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lbl_fromdate 
      Caption         =   "From Date"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label lbl_todate 
      Caption         =   "To Date"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1080
   End
End
Attribute VB_Name = "frmSecurityChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'validate from date

Private Sub Combo_userid_GotFocus()
Call HighlightBackground(Combo_userid)
End Sub

Private Sub Combo_userid_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_userid.DroppedDown Then Combo_userid.DroppedDown = True
End Sub

Private Sub Combo_userid_KeyPress(KeyAscii As Integer)
'Combo_userid.MoveNext
End Sub

Private Sub Combo_userid_LostFocus()
Call NormalBackground(Combo_userid)
End Sub

Private Sub Combo_userid_Validate(Cancel As Boolean)
If Len(Trim$(Combo_userid)) > 0 Then
         If Not Combo_userid.IsItemInList Then
                Combo_userid.Text = ""
            End If
            If Len(Trim$(Combo_userid)) = 0 Then
           Combo_userid.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub DTfromdate_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub
'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler
If DTfromdate.value > DTtodate.value Then
 
    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("M00003") 'J added
    msg2 = translator.Trans("L00318") 'J added
    MsgBox IIf(msg1 = "", "Make Sure The To Date is greater than the From date", msg1), , IIf(msg2 = "", "Date", msg2) 'J modified
    '---------------------------------------------
     DTfromdate_Validate ("true")
  Else
With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\securitchange.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "userid;" + IIf(Trim$(Combo_userid.Text) = "ALL", "ALL", Trim$(Combo_userid.Text)) + ";true"
        'added fromdate and todate parameters
        .ParameterFields(2) = "fromdate;date(" & Year(DTfromdate.value) & "," & Month(DTfromdate.value) & "," & Day(DTfromdate.value) & ");true"
        .ParameterFields(3) = "todate;date(" & Year(DTtodate.value) & "," & Month(DTtodate.value) & "," & Day(DTtodate.value) & ");true"
       
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00388") 'J added
        .WindowTitle = IIf(msg1 = "", "Security change", msg1) 'J modified
        Call translator.Translate_Reports("securitchange.rpt") 'J added
        '---------------------------------------------
'        MsgBox (Combo_userid.text)
'        MsgBox (DTfromdate.Value)
'        MsgBox (DTtodate.Value)
        .Action = 1: .Reset
End With
End If
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    
End Sub

'SQL statement get userid and populate combo

Private Sub Form_Load()
Dim rs As ADODB.Recordset

'Added by Juan (9/13/2000) for Multilingual
Call translator.Translate_Forms("frm_SecurityChangeLog")
'------------------------------------------

'Me.Height = 2500 'J hidden
'Me.Width = 4000 'J hidden
Set rs = New ADODB.Recordset
rs.Source = "select usr_USERID  from xuserprofile where usr_npecode='" & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
If get_status(rs) Then
Combo_userid.AddItem "ALL"
Do While (Not rs.EOF)
    Combo_userid.AddItem (rs!usr_userid)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Else
End If
Combo_userid.MoveFirst
Combo_userid.Text = Combo_userid.Columns(0).Text

Caption = Caption + " - " + Tag
DTfromdate = FirstOfMonth
DTtodate = Now

Me.Left = Round((Screen.Width - Me.Width) / 2)
Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'check recordset status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
End Function

'resize form

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then
'Me.Height = 2500 'J hidden
'Me.Width = 4000 'J hidden
End If
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

