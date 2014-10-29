VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form Frm_TrackingPONew 
   Caption         =   "Tracking Message for PO"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   8925
   Tag             =   "02020200"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   11086
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Message"
      TabPicture(0)   =   "Frm_TrackingPONew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblStatu"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblOperator"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "SSOleDBPO"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "NavBar1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "SScmbMessage"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtRemark"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkYesorNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtMessage"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtSubject"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmd_Addterms"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Recepients"
      TabPicture(1)   =   "Frm_TrackingPONew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdSupEmail"
      Tab(1).Control(1)=   "CmdSupFax"
      Tab(1).Control(2)=   "fra_FaxSelect"
      Tab(1).Control(3)=   "cmd_Add"
      Tab(1).Control(4)=   "txt_Recipient"
      Tab(1).Control(5)=   "cmdRemove"
      Tab(1).Control(6)=   "Txt_search"
      Tab(1).Control(7)=   "OptFax"
      Tab(1).Control(8)=   "OptEmail"
      Tab(1).Control(9)=   "dgRecipientList"
      Tab(1).Control(10)=   "SSOLEDBEmail"
      Tab(1).Control(11)=   "SSOLEDBFax"
      Tab(1).Control(12)=   "lbl_Recipients"
      Tab(1).Control(13)=   "Lbl_search"
      Tab(1).ControlCount=   14
      Begin VB.CommandButton cmd_Addterms 
         Caption         =   "Add Clause"
         Height          =   288
         Left            =   1320
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton CmdSupEmail 
         Caption         =   "Supplier Email"
         Height          =   288
         Left            =   -74640
         TabIndex        =   15
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton CmdSupFax 
         Caption         =   "Supplier Fax"
         Height          =   288
         Left            =   -74640
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox TxtSubject 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   8295
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   1170
         Left            =   -74745
         TabIndex        =   27
         Top             =   3915
         Width           =   1635
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   795
         End
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74685
         TabIndex        =   9
         Top             =   3270
         Width           =   1335
      End
      Begin VB.TextBox txt_Recipient 
         Height          =   288
         Left            =   -72885
         MaxLength       =   60
         TabIndex        =   10
         Top             =   3270
         Width           =   6150
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74685
         TabIndex        =   6
         Top             =   1110
         Width           =   1335
      End
      Begin VB.TextBox Txt_search 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   -72885
         MaxLength       =   60
         TabIndex        =   11
         Top             =   3630
         Width           =   3855
      End
      Begin VB.OptionButton OptFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   -72885
         TabIndex        =   7
         Top             =   2910
         Width           =   615
      End
      Begin VB.OptionButton OptEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -72045
         TabIndex        =   8
         Top             =   2910
         Width           =   735
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkYesorNo 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7680
         TabIndex        =   3
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtRemark 
         Height          =   3135
         Left            =   240
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2520
         Width           =   8295
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScmbMessage 
         Height          =   315
         Left            =   6120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "MessageNumber"
         Columns(0).Name =   "MessageNumber"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "MessageDate"
         Columns(1).Name =   "MessageDate"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin LRNavigators.NavBar NavBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   5880
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   661
         ButtonHeight    =   329.953
         ButtonWidth     =   345.26
         Style           =   1
         MouseIcon       =   "Frm_TrackingPONew.frx":0038
         PreviousVisible =   0   'False
         LastVisible     =   0   'False
         NextVisible     =   0   'False
         FirstVisible    =   0   'False
         EMailVisible    =   -1  'True
         PrintEnabled    =   0   'False
         EmailEnabled    =   -1  'True
         SaveEnabled     =   0   'False
         CancelEnabled   =   0   'False
         NextEnabled     =   0   'False
         DeleteEnabled   =   -1  'True
         EditEnabled     =   -1  'True
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBPO 
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3201
         _ExtentY        =   573
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
         Height          =   2085
         Left            =   -72765
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   570
         Width           =   6015
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
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
         stylesets(0).Picture=   "Frm_TrackingPONew.frx":0054
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
         stylesets(1).Picture=   "Frm_TrackingPONew.frx":0070
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowAddNew     =   -1  'True
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   5292
         Columns(0).Caption=   "Column 0"
         Columns(0).Name =   "Column 0"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10610
         _ExtentY        =   3678
         _StockProps     =   79
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBEmail 
         Height          =   2055
         Left            =   -72840
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4080
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Caption=   "Email"
         Columns(1).Name =   "Email"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         _ExtentX        =   10927
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Email"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBFax 
         Height          =   2055
         Left            =   -72840
         TabIndex        =   32
         Top             =   4080
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Caption=   "Fax"
         Columns(1).Name =   "Fax"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         _ExtentX        =   10927
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Fax"
      End
      Begin VB.Label Label7 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Subject"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74760
         TabIndex        =   30
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label Lbl_search 
         Caption         =   "Search by name"
         Height          =   255
         Left            =   -74565
         TabIndex        =   29
         Top             =   3630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Message Date"
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Message Date"
         Height          =   315
         Left            =   -2040
         TabIndex        =   24
         Top             =   -1560
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Operator"
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label20 
         Caption         =   "Include Original Message"
         Height          =   315
         Left            =   5040
         TabIndex        =   22
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label LblOperator 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6120
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblStatu 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Visualization"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   5160
         TabIndex        =   19
         Top             =   5760
         Width           =   3420
      End
      Begin VB.Label Label2 
         Caption         =   "Transaction #"
         Height          =   315
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Message #"
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tracking Message For PO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "Frm_TrackingPONew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GGridFilledWithEmails As Boolean
Dim GGridFilledWithFax As Boolean
Dim FormMode As FormMode
Dim WithEvents st As frm_ShipTerms
Attribute st.VB_VarHelpID = -1

Private Sub cmd_Add_Click()
On Error Resume Next
If (OptEmail.value = True Or OptFax.value = True) Then
    
        If Len(Trim$(txt_Recipient)) > 0 Then
               txt_Recipient = UCase(txt_Recipient)
               
               If OptEmail.value = True Then txt_Recipient = (txt_Recipient)
               If OptFax.value = True Then txt_Recipient = (txt_Recipient)
           
              'dgRecipientList.AddItem txt_Recipient
              
              AddRecepient txt_Recipient
            
              txt_Recipient = ""
      
        End If
 Else
    MsgBox "Please check Email or Fax.", vbInformation, "Imswin"
    
 End If
End Sub

Private Sub cmd_Addterms_Click()

On Error Resume Next
 Me.MousePointer = vbHourglass

    If st Is Nothing Then Set st = New frm_ShipTerms
    st.Show
    st.txt_Description.SetFocus

    If Err Then Err.Clear
   Me.MousePointer = vbArrow
   
End Sub

Private Sub cmdRemove_Click()
Dim x As Integer

If Len(dgRecipientList.SelBookmarks(0)) = 0 Then
    
    MsgBox "Please make a selection first.", vbInformation, "Imswin"
    
    Exit Sub
 
 End If
 
If FormMode = mdCreation Then

    dgRecipientList.DeleteSelected
        
End If

End Sub

Private Sub CmdSupEmail_Click()

Dim RsSupEmail As New ADODB.Recordset

RsSupEmail.Source = "select sup_mail from po, supplier where po_npecode = '" & deIms.NameSpace & "' and  sup_code=  po_suppcode and po_ponumb ='" & ssOleDbPO & "'"

RsSupEmail.ActiveConnection = deIms.cnIms

RsSupEmail.Open

If Len(RsSupEmail("sup_mail") & "") > 0 Then

    AddRecepient RsSupEmail("sup_mail")
    
 Else
 
   MsgBox "No Email exists for the supplier.", vbInformation, "Imswin"
    
End If

End Sub

Private Sub CmdSupFax_Click()

Dim RsSupFax As New ADODB.Recordset

RsSupFax.Source = "select sup_faxnumb from po, supplier where po_npecode = '" & deIms.NameSpace & "' and  sup_code=  po_suppcode and po_ponumb ='" & ssOleDbPO & "'"

RsSupFax.ActiveConnection = deIms.cnIms

RsSupFax.Open

If Len(RsSupFax("sup_faxnumb") & "") > 0 Then

    AddRecepient RsSupFax("sup_faxnumb")
    
 Else
 
   MsgBox "No Fax exists for the supplier.", vbInformation, "Imswin"
    
End If

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11

    Me.Width = 9045
    Me.Height = 7600
    dgRecipientList.Columns(0).locked = True
    LblOperator = CurrentUser

    Screen.MousePointer = 11
    Me.Refresh
    DoEvents
    GetPOnumber
    GetSupplierPhoneDirEmails
    GetSupplierPhoneDirFAX

    Call ChangeMode(mdvisualization)
    Call EnableControls(False)
    Caption = Caption + " - " + Tag
    Call DisableButtons(Frm_TrackingPONew, NavBar1)
    Screen.MousePointer = 0
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

Public Sub GetPOnumber()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT po_ponumb From PO "
        .CommandText = .CommandText & " WHERE po_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by po_ponumb "

        Set rst = .Execute
    End With

    If rst.RecordCount = 0 Then GoTo clearup

    rst.MoveFirst
    Do While ((Not rst.EOF))
        ssOleDbPO.AddItem rst!PO_PONUMB & ""
        rst.MoveNext
    Loop

clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing



End Sub


Private Sub Form_Unload(Cancel As Integer)
    Hide

     GGridFilledWithEmails = False
     GGridFilledWithFax = False

    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub NavBar1_OnCancelClick()
       
            LblOperator = ""
            dgRecipientList.RemoveAll
            TxtSubject = ""
            chkYesorNo.value = 0
            txtRemark = ""
            txtMessage = ""
            SScmbMessage = ""
            ChangeMode (mdvisualization)
            Call EnableControls(False)
            NavBar1.NewEnabled = True
            NavBar1.CancelEnabled = False

End Sub

Private Sub NavBar1_OnCloseClick()
Unload Me
End Sub

Private Sub NavBar1_OnEMailClick()
Dim IFile As IMSFile
Dim Filename(1) As String
Dim Recepients() As String
Dim rsr As ADODB.Recordset
Dim rptinfo As RPTIFileInfo
Dim subject As String
Dim attention As String
Dim ParamsForCrystalReports() As String
Dim ParamsForRPTI() As String
Dim FieldName As String
Dim Message As String

ReDim ParamsForCrystalReports(2)
ReDim ParamsForRPTI(2)

    Set rsr = GetObsRecipients(deIms.NameSpace, ssOleDbPO, SScmbMessage.Text)
    
    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    
    ParamsForCrystalReports(1) = "mesgnumb;" + SScmbMessage + ";TRUE"
    
    ParamsForCrystalReports(2) = "ponumb;" + ssOleDbPO + ";true"
    
    ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
    
    ParamsForRPTI(1) = "mesgnumb=" + SScmbMessage
    
    ParamsForRPTI(2) = "ponumb=" + ssOleDbPO
    
    FieldName = "Recipient"
    
    subject = "Tracking Message for PO -" & ssOleDbPO
    
    If ConnInfo.EmailClient = Outlook Then
    
        Call sendOutlookEmailandFax(Report_EmailFax_TrackingPo, "Tracking Message Header", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, subject, attention)
    
    ElseIf ConnInfo.EmailClient = ATT Then
    
        Call SendAttFaxAndEmail("obs.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, subject, Message, FieldName)

    ElseIf ConnInfo.EmailClient = Unknown Then
    
        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

    End If


    If chkYesorNo.value = 1 Then
    
        ReDim ParamsForCrystalReports(1)
        
        ReDim ParamsForRPTI(1)
    
    
            ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            
            ParamsForCrystalReports(1) = "ponumb;" + ssOleDbPO + ";true"
            
            ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
            
            ParamsForRPTI(1) = "ponumb=" + ssOleDbPO
            
            FieldName = "Recipient"
            
            If ConnInfo.EmailClient = Outlook Then
            
                'Call sendOutlookEmailandFax("PO.rpt", "Tracking Message", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, subject, attention)  MM  / using CR 11 report for emails now
                Call sendOutlookEmailandFax(Report_EmailFax_PO_name, "Tracking Message PO", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, subject, attention)
            
            ElseIf ConnInfo.EmailClient = ATT Then
            
                Call SendAttFaxAndEmail("PO.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, subject, Message, FieldName)
        
            ElseIf ConnInfo.EmailClient = Unknown Then
            
                MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"
        
            End If

    End If
    
End Sub


Private Sub NavBar1_OnNewClick()
Dim str As String
       
   If Len(Trim(ssOleDbPO)) > 0 Then

        Call Clearform
        Call ChangeMode(mdCreation)
        SScmbMessage = GetMessageNumber
        txtMessage = Format$(Now(), "mm/dd/yyyy")
        LblOperator = CurrentUser
        
        
        NavBar1.SaveEnabled = True
        NavBar1.CancelEnabled = True
        Call EnableControls(True)
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
        NavBar1.NewEnabled = False
        
   Else
   
       MsgBox "Please select a Transaction number before creating a new Tracking message.", vbInformation, "Imswin"
     
   End If

End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    SetOBsReportParam

    'Modified by Juan (9/25/2000) for Multilingual
    msg1 = translator.Trans("M00370") 'J added
    MDI_IMS.CrystalReport1.WindowTitle = IIf(msg1 = "", "Tracking Message", msg1) 'J modified
    '------------------------------------------

    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset

    If chkYesorNo.value = 1 Then
        SetPOReportParam

        'Modified by Juan (9/25/2000) for Multilingual
        MDI_IMS.CrystalReport1.WindowTitle = IIf(msg1 = "", "Tracking Message", msg1) 'J modified
        '---------------------------------------------

        MDI_IMS.CrystalReport1.Action = 1
        MDI_IMS.CrystalReport1.Reset

    End If


    Exit Sub

ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub
Private Sub SetOBsReportParam()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\obs.rpt"

        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("obs.rpt") 'J added
        '---------------------------------------------

        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "mesgnumb;" + SScmbMessage + ";TRUE"
        .ParameterFields(2) = "ponumb;" + ssOleDbPO + ";true"
    End With
End Sub
'get po report parameters to print po report

Private Sub SetPOReportParam()

    If chkYesorNo.value = 1 Then
        With MDI_IMS.CrystalReport1
            .Reset
            .ReportFileName = FixDir(App.Path) + "CRreports\po.rpt"

            'Modified by Juan (8/28/2000) for Multilingual
            Call translator.Translate_Reports("po.rpt") 'J added
            Call translator.Translate_SubReports 'J added
            '---------------------------------------------

            .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            .ParameterFields(1) = "ponumb;" + ssOleDbPO + ";true"
        End With
    End If

End Sub
'SQL  statement check message number exist or not

Private Function CheckMessageNumber(Numb As String) As Integer
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From OBS "
        .CommandText = .CommandText & " Where ob_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND ob_ponumb = '" & Numb & "'"

        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)

        Set rst = .Execute
        CheckMessageNumber = rst!rt
    End With


    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckMessageNumber", Err.Description, Err.number, True)
End Function

'before save a record check data fields
'and set store procedure parameters

Private Sub NavBar1_OnSaveClick()

Dim rec(7) As String
     If dgRecipientList.Rows = 0 Then

            'Modified by Juan (9/25/2000) for Multilingual
            msg1 = translator.Trans("M00369") 'J added
            MsgBox IIf(msg1 = "", "Please make sure you have selected atleast one recepient.", msg1), vbCritical, "ImsWin" 'J modified
            '---------------------------------------------
            NavBar1.SaveEnabled = True
            
        Exit Sub
    End If

    dgRecipientList.MoveFirst

  Do While Not dgRecipientList.row > dgRecipientList.Rows - 1
  
    rec(dgRecipientList.row) = dgRecipientList.Columns(0).value
  
    dgRecipientList.MoveNext
  
  Loop


Call InsertandUpdateTable(ssOleDbPO, SScmbMessage, txtMessage, 1, rec, TxtSubject, txtRemark, chkYesorNo)


    If Len(Trim(ssOleDbPO)) <> 0 And Len(Trim(SScmbMessage)) <> 0 Then
        NavBar1.CancelEnabled = False
        NavBar1.SaveEnabled = False
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
        NavBar1.NewEnabled = True
    End If
    
   Call ChangeMode(mdvisualization)
   
   Call EnableControls(False)
   
   Call GetOBSMessagelist
   
End Sub

'set store procedure parameters

Public Function GetMessageNumber() As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GetMessageNumber"
        Set .ActiveConnection = deIms.cnIms


        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)

        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, ssOleDbPO)
'        .Parameters.Append .CreateParameter("@MESSAGE", adVarChar, adParamInput, 15, SScmbMessage.Columns(0).Text)
        .parameters.Append .CreateParameter("@STRING", adVarChar, adParamOutput, 15, GetMessageNumber)

        .Execute , , adExecuteNoRecords

        GetMessageNumber = .parameters("@STRING").value & ""
    End With

    Set cmd = Nothing

End Function


Private Sub opt_Email_Click()

SSOLEDBEmail.Visible = True
SSOLEDBFax.Visible = False

If GGridFilledWithEmails = False Then

   Call GetSupplierPhoneDirEmails

   GGridFilledWithEmails = True

End If

End Sub


Public Sub GetSupplierPhoneDirEmails()

Dim str As String
Dim cmd As Command
Dim rst As New Recordset
Dim Sql As String

Sql = " select sup_name Names,  upper( sup_mail) Emails  from supplier where sup_npecode='" & deIms.NameSpace & "' and sup_mail is not null and len(sup_mail) > 0   union "

Sql = Sql & " select phd_name Names, upper(phd_mail) Emails from phonedir  where phd_npecode='" & deIms.NameSpace & "'and phd_mail is not null and len(phd_mail)>0 order by names "

rst.Source = Sql

rst.ActiveConnection = deIms.cnIms

rst.Open

    If rst.RecordCount = 0 Then GoTo clearup

    rst.MoveFirst
    
    Do While Not rst.EOF
    
        SSOLEDBEmail.AddItem rst("Names") & Chr(9) & rst("Emails")
        
        rst.MoveNext
    
    Loop
    
    GGridFilledWithEmails = True
    
clearup:

    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub


Public Sub GetSupplierPhoneDirFAX()
Dim str As String
Dim cmd As Command
Dim rst As New Recordset
Dim Sql As String

Sql = "select sup_name Names,  upper( sup_faxnumb) Fax  from supplier where sup_npecode='" & deIms.NameSpace & "' and sup_faxnumb is not null and len(sup_faxnumb) > 0   union"

Sql = Sql & " select phd_name Names, upper(phd_faxnumb) Fax from phonedir  where phd_npecode='" & deIms.NameSpace & "'and phd_faxnumb is not null and len(phd_faxnumb)>0 order by names"

rst.Source = Sql

rst.ActiveConnection = deIms.cnIms

rst.Open

    If rst.RecordCount = 0 Then GoTo clearup

    rst.MoveFirst
    
    Do While Not rst.EOF
    
        SSOLEDBFax.AddItem rst("Names") & Chr(9) & rst("FAX")
        
        rst.MoveNext
    
    Loop
    
    GGridFilledWithFax = True
    
clearup:

    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub opt_FaxNum_Click()

SSOLEDBFax.Visible = True
SSOLEDBEmail.Visible = False

If GGridFilledWithFax = False Then

    Call GetSupplierPhoneDirFAX
    
    GGridFilledWithFax = True
    
 End If


End Sub

Private Sub SScmbMessage_Click()
    If Not Len(SScmbMessage.Columns(0).Text) = 0 And Not Len(SScmbMessage.Columns(1).Text) = 0 Then
        dgRecipientList.RemoveAll 'JCG 2008/10/10
        
        Call GetOBSList
         NavBar1.PrintEnabled = True
         NavBar1.EMailEnabled = True
    Else
        If Not Len(SScmbMessage.Columns(0).Text) = 0 Then
           ' NavBar1.SaveEnabled = SaveEnabled
            NavBar1.CancelEnabled = True
        End If

    End If
End Sub

Private Sub SScmbMessage_GotFocus()
 Call HighlightBackground(SScmbMessage)
End Sub

Private Sub SScmbMessage_LostFocus()
Call NormalBackground(SScmbMessage)
End Sub

Private Sub SSOLEDBEmail_DblClick()
On Error Resume Next


 If FormMode = mdCreation Then AddRecepient SSOLEDBEmail.Columns(1).value
    

End Sub

Private Sub SSOLEDBFax_DblClick()
On Error Resume Next
    
    'dgRecipientList.AddItem SSOLEDBFax.Columns(1).Value
    
    AddRecepient SSOLEDBFax.Columns(1).value
    
    If Err Then Err.Clear
End Sub

Private Sub SSOleDBPO_Click()

Dim exist As Integer
Dim rsPO As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim query As String
    
    Call Clearform
    
    If Len(ssOleDbPO) Then Call GetOBSMessagelist
         
'-----------------------------------
'Commented out by Muzammil 07/11. Added the last line instead of all
'this to remove the contact in the list.
         
''     query = "SELECT  po_suppcode, sup_mail,sup_faxnumb"
''     query = query & " From PO, supplier "
''     query = query & " WHERE po_ponumb = '" & ssOleDbPO & "' AND"
''     query = query & " po_npecode = '" & deIms.NameSpace & "' AND po_suppcode = sup_code AND"
''     query = query & "  po_npecode = sup_npecode"
''
''      rsPO.ActiveConnection = deIms.cnIms
''      rsPO.Open query
''
''    If rsPO.RecordCount > 0 Then
''
''    If Len((rsPO!sup_mail) & "") > 0 Then
''
''        dgRecipientList.RemoveAll
''
''        'dgRecipientList.AddItem rsPO!sup_mail
''
''
''    End If
''
''    If Len((rsPO!sup_faxnumb) & "") > 0 Then
''
''        dgRecipientList.RemoveAll
''        'dgRecipientList.AddItem rsPO!sup_faxnumb
''
''    End If
''
''    End If
''
''    Set rsPO = Nothing

    
          dgRecipientList.RemoveAll
    
 '--------------------------------------------------
 
          exist = CheckMessageNumber(ssOleDbPO)
        
        If exist > 0 Then
           ' EnableControls (False)
            SScmbMessage.Enabled = True
        Else
           ' EnableControls (False)
            NavBar1.PrintEnabled = False
            NavBar1.EMailEnabled = False
        End If
     '   If SSOleDBPO <> "" Then NavBar1.NewEnabled = SaveEnabled  'J added

End Sub

'SQL statement get list and populate it

Public Sub GetOBSMessagelist()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT OBS.ob_mesgnumb, OBS.ob_mesgdate"
        .CommandText = .CommandText & " FROM OBS INNER JOIN PO "
        .CommandText = .CommandText & " ON OBS.ob_ponumb = PO.po_ponumb AND "
        .CommandText = .CommandText & " OBS.ob_npecode = PO.po_npecode "
        .CommandText = .CommandText & " WHERE (OBS.ob_ponumb = '" & ssOleDbPO & "') AND "
        .CommandText = .CommandText & " OBS.ob_npecode  = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND OBS.ob_flag  = 1 "
         Set rst = .Execute
    End With
    
    str = Chr$(1)
    
'    EnableControls True

    
    SScmbMessage.Enabled = True
    
    SScmbMessage.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    
    rst.MoveFirst
    SScmbMessage.RemoveAll
    
    Do While ((Not rst.EOF))
        SScmbMessage.AddItem rst!ob_mesgnumb & "" & str & (rst!ob_mesgdate & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    

End Sub
'call store procedure to get OBS list an d populate data grid

Public Sub GetOBSList()
Dim cmd As ADODB.Command
Dim rst As Recordset

    Set cmd = New ADODB.Command
    
    With cmd
        
        
        .CommandType = adCmdStoredProc
        .CommandText = "GetOBSList"
        Set .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, ssOleDbPO)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@mesgnumb", adVarChar, adParamInput, 15, SScmbMessage)
        
         Set rst = .Execute

    End With
        

        txtMessage = rst!ob_mesgdate & ""
        TxtSubject = rst!ob_subj & ""
        
            
        If Len(rst!ob_forwreci & "") > 0 Then dgRecipientList.AddItem rst!ob_forwreci & ""
        If Len(rst!ob_suppreci & "") > 0 Then dgRecipientList.AddItem rst!ob_suppreci & ""
        If Len(rst!ob_rec1 & "") > 0 Then dgRecipientList.AddItem rst!ob_rec1 & ""
        If Len(rst!ob_rec2 & "") > 0 Then dgRecipientList.AddItem rst!ob_rec2 & ""
        If Len(rst!ob_rec3 & "") > 0 Then dgRecipientList.AddItem rst!ob_rec3 & ""
        If Len(rst!ob_rec4 & "") > 0 Then dgRecipientList.AddItem rst!ob_rec4 & ""
        If Len(rst!ob_rec5 & "") > 0 Then dgRecipientList.AddItem rst!ob_rec5 & ""
        
        LblOperator = rst!ob_oper & ""
        
        chkYesorNo.value = IIf((rst!ob_inclmesg), 1, 0)
        txtRemark = rst!ob_remk & ""
          
    Set cmd = Nothing
     
   
End Sub


Public Sub Clearform()

SScmbMessage.RemoveAll

SScmbMessage = ""

dgRecipientList.RemoveAll

txtMessage = ""

LblOperator = ""

txtRemark = ""

txt_Recipient = ""

Txt_search = ""

chkYesorNo.value = 0

TxtSubject = ""

End Sub

Public Function InsertandUpdateTable(PO As String, Message As String, MessageDate As Date, Flag As Boolean, _
                                     rec() As String, subject As String, remark As String, Inflag As Boolean) _
                                    As Boolean

On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = deIms.cnIms
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATE_INSERT_OBS"
    
        .parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, PO)
        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@mesgnumb", adVarChar, adParamInput, 12, Message)
        .parameters.Append .CreateParameter("@subj", adVarChar, adParamInput, 60, subject)
        .parameters.Append .CreateParameter("@flag", adBoolean, adParamInput, 1, Flag)
        .parameters.Append .CreateParameter("@mesgdate", adVarChar, adParamInput, 15, MessageDate)
        .parameters.Append .CreateParameter("@suppreci", adVarChar, adParamInput, 60, rec(0))
        .parameters.Append .CreateParameter("@forwreci", adVarChar, adParamInput, 60, rec(1))
        .parameters.Append .CreateParameter("@rec1", adVarChar, adParamInput, 60, rec(2))
        .parameters.Append .CreateParameter("@rec2", adVarChar, adParamInput, 60, rec(3))
        .parameters.Append .CreateParameter("@rec3", adVarChar, adParamInput, 60, rec(4))
        .parameters.Append .CreateParameter("@rec4", adVarChar, adParamInput, 60, rec(5))
        .parameters.Append .CreateParameter("@rec5", adVarChar, adParamInput, 60, rec(6))
        .parameters.Append .CreateParameter("@oper", adVarChar, adParamInput, 30, CurrentUser)
        .parameters.Append .CreateParameter("@newdelvdate", adVarChar, adParamInput, 12, Null)
        .parameters.Append .CreateParameter("@etd", adVarChar, adParamInput, 12, Null)
        .parameters.Append .CreateParameter("@eta", adVarChar, adParamInput, 12, Null)
        .parameters.Append .CreateParameter("@shipvia", adVarChar, adParamInput, 2, Null)
        .parameters.Append .CreateParameter("@inclmesg", adBoolean, adParamInput, 1, Inflag)
        .parameters.Append .CreateParameter("@REMARKS", adVarChar, adParamInput, 1000, remark)
        .parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, CurrentUser)
        
        
        .Execute , , adExecuteNoRecords
    End With
    
    InsertandUpdateTable = True
    Set cmd = Nothing
    MsgBox "Record saved successfully."
    Exit Function
    
Noinsert:
    InsertandUpdateTable = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Sub SSOleDBPO_GotFocus()
 Call HighlightBackground(ssOleDbPO)
End Sub

Private Sub ssOleDbPO_KeyPress(KeyAscii As Integer)
  If Not ssOleDbPO.DroppedDown Then ssOleDbPO.DroppedDown = True
End Sub

Private Sub SSOleDBPO_LostFocus()
Call NormalBackground(ssOleDbPO)
End Sub

Private Sub SSOleDBPO_Validate(Cancel As Boolean)

Dim Count As Integer

Count = 1

Do While Not ssOleDbPO.Rows = Count

If ssOleDbPO = ssOleDbPO.Columns(0).value Then

    Exit Sub

End If

Count = Count + 1
Loop

MsgBox "Please make sure that you have entered a valid Transaction Number.", vbInformation, "Imswin"
Cancel = True

ssOleDbPO.SetFocus
ssOleDbPO.SelLength = 0
ssOleDbPO.SelStart = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab

Case 0


Case 1

    opt_Email.value = 1
    
End Select

End Sub

Public Function AddRecepient(RecepientAddress As String)

Dim Count As Integer

'If dgRecipientList.Rows = 0 Then dgRecipientList.AddItem RecepientAddress: Exit Sub

If dgRecipientList.Rows = 7 Then

   MsgBox "Can not Add more than 7 recepients.", vbInformation + vbOKOnly, "Imswin"

ElseIf dgRecipientList.Rows < 7 Then

    'count = 1

    dgRecipientList.MoveFirst
    
    Do While Not dgRecipientList.Rows = Count

        If dgRecipientList.Columns(0).value = RecepientAddress Then
        
            MsgBox "Recepient already exists, Please choose a different one.", vbInformation + vbOKOnly, "Imswin"
            
            Exit Function
        
        End If
            
        dgRecipientList.MoveNext
            
        Count = Count + 1
            
    Loop

    dgRecipientList.AddItem RecepientAddress

End If

End Function

Private Sub Txt_search_Change()

Dim Grid As SSOleDBGrid
    
Dim x As Integer

Dim Count As Integer

Dim i As Integer

If SSOLEDBEmail.Visible = True Then Set Grid = SSOLEDBEmail

If SSOLEDBFax.Visible = True Then Set Grid = SSOLEDBFax

i = Len(Txt_search)

Count = 1

    Grid.MoveFirst

    Do While Not Grid.Rows = Count

        If UCase(Txt_search) = UCase(Mid(Grid.Columns(0).value, 1, i)) Then
    
           Grid.Scroll 0, Grid.row
           
           Exit Sub
          
        End If
        
        Grid.MoveNext
    
        Count = Count + 1

    Loop

End Sub


Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next

    
    If FMode = mdCreation Then
        lblStatu.ForeColor = vbRed
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00125") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        
                
        
        ChangeMode = True
  
    ElseIf FMode = mdvisualization Then
        lblStatu.ForeColor = vbGreen
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00092") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        
    End If
    
       
    FormMode = FMode

End Function

Public Function EnableControls(Enabled As Boolean)

Dim value As Boolean
    value = Not (Enabled)
    
    txtMessage.Enabled = value
    
    'LblOperator
    
    txtRemark.locked = value
    
    txt_Recipient.locked = value
    
    Txt_search.locked = value
    
    chkYesorNo.Enabled = Enabled
    
    TxtSubject.locked = value
    
    Txt_search.locked = value
    
    txt_Recipient.locked = value
    
    fra_FaxSelect.Enabled = Enabled

    CmdSupEmail.Enabled = Enabled
    
    CmdSupFax.Enabled = Enabled
    
    cmd_Add.Enabled = Enabled
    
    cmdRemove.Enabled = Enabled
    
    cmd_Addterms.Enabled = Enabled
    
End Function
Private Sub st_Completed(Cancelled As Boolean, Terms As String)
On Error Resume Next

    If Not Cancelled Then
        'txtClause.SelText = Terms
        txtRemark.Text = txtRemark.Text & Terms
        txtRemark.SelStart = Len(txtRemark)
        
    End If
    
    Set st = Nothing
End Sub

Private Sub Txt_search_GotFocus()
 Call HighlightBackground(Txt_search)
End Sub

Private Sub Txt_search_LostFocus()
Call NormalBackground(Txt_search)
End Sub

Private Sub txtRemark_GotFocus()
 Call HighlightBackground(txtRemark)
End Sub

Private Sub txtRemark_LostFocus()
Call NormalBackground(txtRemark)
End Sub

Private Sub TxtSubject_GotFocus()
 Call HighlightBackground(TxtSubject)
End Sub

Private Sub TxtSubject_LostFocus()
Call NormalBackground(TxtSubject)
End Sub
