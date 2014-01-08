VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_individualuserprofile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Individual Profile"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1890
   ScaleWidth      =   3510
   Tag             =   "03050500"
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo Combo_userid 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   480
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
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Id"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "frm_individualuserprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

'get crystal report parameter and application path

Private Sub cmd_ok_Click()
On Error GoTo ErrHandler

With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\indiuserprof.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "UserId;" + IIf(Trim$(Combo_userid.text) = "ALL", "ALL", Trim$(Combo_userid.text)) + ";true"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00198") 'J added
        .WindowTitle = IIf(msg1 = "", "Individual User Profile", msg1) 'J modified
        Call translator.Translate_Reports("indiuserprof.rpt") 'J added
        Call translator.Translate_SubReports 'J added
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

Private Sub Combo_userid_GotFocus()
Call HighlightBackground(Combo_userid)
End Sub

Private Sub Combo_userid_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Combo_userid.DroppedDown Then Combo_userid.DroppedDown = True
End Sub

Private Sub Combo_userid_LostFocus()
Call NormalBackground(Combo_userid)
End Sub

Private Sub Combo_userid_Validate(Cancel As Boolean)
If Len(Trim$(Combo_userid)) > 0 Then
         If Not Combo_userid.IsItemInList Then
                Combo_userid.text = ""
            End If
            If Len(Trim$(Combo_userid)) = 0 Then
            Combo_userid.SetFocus
            Cancel = True
            End If
            End If
End Sub

'SQL statement,get values for combo box

Private Sub Form_Load()
Dim rs As ADODB.Recordset

'Added by Juan (9/12/2000) for Multilingual
Call translator.Translate_Forms("frm_individualuserprofile")
'---------------------------------------------

'Me.Height = 2500 'J hidden
'Me.Width = 3630 'J hidden

Set rs = New ADODB.Recordset
rs.Source = "select usr_USERID  from xuserprofile where usr_npecode='" & deIms.NameSpace & "'"
rs.ActiveConnection = deIms.cnIms
rs.Open
If get_status(rs) Then
Do While (Not rs.EOF)
Combo_userid.AddItem (rs!usr_userid)
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Else
End If

Caption = Caption + " - " + Tag
    With frm_individualuserprofile
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'get record set status

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
End Function

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub
