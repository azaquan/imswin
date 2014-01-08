VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_newSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier by Date created"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   3315
   Tag             =   "03020104"
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1092
   End
   Begin MSComCtl2.DTPicker DTfromdate 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60096513
      CurrentDate     =   36523
   End
   Begin MSComCtl2.DTPicker DTtodate 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60096513
      CurrentDate     =   36523
   End
   Begin VB.Label lbl_todate 
      Caption         =   "To Date"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label lbl_fromdate 
      Caption         =   "From Date"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "frm_newSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo_userid_GotFocus()

End Sub

Private Sub DTfromdate_Validate(Cancel As Boolean)
Dim x As Boolean
End Sub
'close form

Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_ok_Click()
On Error GoTo Errhandler
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
        
        .ReportFileName = FixDir(App.Path) + "CRreports\newSupplier.rpt"
        'MsgBox "ReportFileName->" + .ReportFileName
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        'MsgBox "ParameterFields 0->" + .ParameterFields(0)
        .ParameterFields(1) = "from;date(" & Year(DTfromdate.value) & "," & Month(DTfromdate.value) & "," & Day(DTfromdate.value) & ");true"
        'MsgBox "ParameterFields 1->" + .ParameterFields(1)
        .ParameterFields(2) = "to;date(" & Year(DTtodate.value) & "," & Month(DTtodate.value) & "," & Day(DTtodate.value) & ");true"
        'MsgBox "ParameterFields2->" + .ParameterFields(2)
        
        Call translator.Translate_Reports("newSupplier.rpt")
        .Action = 1: .Reset
End With
End If
    Exit Sub
    
Errhandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    
End Sub

'SQL statement get userid and populate combo

Private Sub Form_Load()
Dim Rs As ADODB.Recordset

'Added by Juan (9/13/2000) for Multilingual
Call translator.Translate_Forms("frm_newSupplier")
'------------------------------------------

Set Rs = New ADODB.Recordset
Rs.Source = "select usr_USERID  from xuserprofile where usr_npecode='" & deIms.NameSpace & "'"
Rs.ActiveConnection = deIms.cnIms
Rs.Open

Caption = Caption + " - " + Tag
DTfromdate = FirstOfMonth
DTtodate = Now

    With frm_newSupplier
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Public Function get_status(rst As Recordset) As Boolean
  get_status = IIf(rst Is Nothing, (False), (True))
   If rst.State And adStateOpen = adStateClosed Then get_status = False
   If rst.EOF And rst.BOF Then get_status = False
   If rst.RecordCount = 0 Then get_status = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub
