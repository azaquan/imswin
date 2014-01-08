VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.UserControl UserControl1 
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ScaleHeight     =   5610
   ScaleWidth      =   8190
   Begin VB.Frame fra_FaxSelect 
      Height          =   1170
      Left            =   15
      TabIndex        =   6
      Top             =   3360
      Width           =   1635
      Begin VB.OptionButton opt_Email 
         Caption         =   "Email"
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   795
      End
      Begin VB.OptionButton opt_FaxNum 
         Caption         =   "Fax Numbers"
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmd_Add 
      Caption         =   "Add"
      Height          =   288
      Left            =   75
      TabIndex        =   5
      Top             =   2670
      Width           =   1335
   End
   Begin VB.TextBox txt_Recipient 
      Height          =   288
      Left            =   1875
      MaxLength       =   60
      TabIndex        =   4
      Top             =   2670
      Width           =   6150
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   288
      Left            =   75
      TabIndex        =   3
      Top             =   630
      Width           =   1335
   End
   Begin VB.TextBox Txt_search 
      BackColor       =   &H00C0E0FF&
      Height          =   288
      Left            =   1875
      MaxLength       =   60
      TabIndex        =   2
      Top             =   3030
      Width           =   3855
   End
   Begin VB.OptionButton OptFax 
      Caption         =   "Fax"
      Height          =   255
      Left            =   1875
      TabIndex        =   1
      Top             =   2310
      Width           =   615
   End
   Begin VB.OptionButton OptEmail 
      Caption         =   "Email"
      Height          =   255
      Left            =   2715
      TabIndex        =   0
      Top             =   2310
      Width           =   735
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
      Height          =   2085
      Left            =   1920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   6195
      _Version        =   196617
      DataMode        =   2
      ColumnHeaders   =   0   'False
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "ImsRecepeint.ctx":0000
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "ImsRecepeint.ctx":001C
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
      Columns(0).Width=   10319
      Columns(0).Caption=   "Column 0"
      Columns(0).Name =   "Column 0"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10927
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
      Left            =   1920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   6195
      _Version        =   196617
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   10927
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Email"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBFax 
      Height          =   2055
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   6195
      _Version        =   196617
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   10927
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Fax"
   End
   Begin VB.Label lbl_Recipients 
      Caption         =   "Recipient(s)"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Lbl_search 
      Caption         =   "Search by name"
      Height          =   255
      Left            =   195
      TabIndex        =   11
      Top             =   3030
      Width           =   1215
   End
   Begin VB.Menu MnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mNUDummy 
         Caption         =   "Dummy"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim GGridFilledWithEmails As Boolean
Dim GGridFilledWithFax As Boolean

Public Property Get EmailGrid() As Variant

Set EmailGrid = SSOLEDBEmail

End Property

Public Property Get FaxGrid() As Variant

Set FaxGrid = SSOLEDBFax

End Property

Public Property Get RecepientList() As Variant

Set RecepientList = dgRecipientList

End Property


Private Sub dgRecipientList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then

    'PopupMenu  MnuMenu, , X + 2000, Y
    PopupMenu MnuMenu, vbPopupMenuLeftAlign
    
End If
    
End Sub





Private Sub MnuDelete_Click()
Call cmdRemove_Click
End Sub

Private Sub txt_Recipient_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then Call AddRecepient(txt_Recipient): txt_Recipient = ""
    
End Sub

Public Function AddRecepient(RecepientAddress As String)

Dim count As Integer

    dgRecipientList.MoveFirst

    Do While Not dgRecipientList.Rows = count

        If UCase(Trim(dgRecipientList.Columns(0).Value)) = UCase(Trim(RecepientAddress)) Then

            MsgBox "Recepient already exists, Please choose a different one.", vbInformation + vbOKOnly, "Imswin"

            Exit Function

        End If

        dgRecipientList.MoveNext

        count = count + 1

    Loop

    dgRecipientList.AddItem RecepientAddress

'End If

End Function


Private Sub Txt_search_Change()

Dim rs As ADODB.Recordset

Dim X As Integer

Dim count As Integer

Dim i As Integer

If Len(Trim(Txt_search)) = 0 Then Exit Sub

If SSOLEDBEmail.Visible = True Then Set rs = SSOLEDBEmail.DataSource

If SSOLEDBFax.Visible = True Then Set rs = SSOLEDBFax.DataSource

rs.MoveFirst

rs.Find rs.Fields(0).Name & " like '" & Trim(Txt_search) & "%'"



End Sub

Private Sub SSOLEDBEmail_DblClick()

On Error Resume Next

 AddRecepient SSOLEDBEmail.Columns(1).Value


End Sub

Private Sub opt_FaxNum_Click()

SSOLEDBFax.Visible = True
SSOLEDBEmail.Visible = False



'    Call GetSupplierPhoneDirFAX

'    GGridFilledWithFax = True

' End If


End Sub

Private Sub opt_Email_Click()

SSOLEDBEmail.Visible = True
SSOLEDBFax.Visible = False

''If GGridFilledWithEmails = False Then
''
''   Call GetSupplierPhoneDirEmails
''
''   GGridFilledWithEmails = True
''   GGridFilledWithFax
''End If

End Sub

Private Sub cmd_Add_Click()
On Error Resume Next
If (OptEmail.Value = True Or OptFax.Value = True) Then

        If Len(Trim$(txt_Recipient)) > 0 Then
               txt_Recipient = UCase(txt_Recipient)

               If OptEmail.Value = True Then txt_Recipient = (txt_Recipient)
               If OptFax.Value = True Then txt_Recipient = (txt_Recipient)

              'dgRecipientList.AddItem txt_Recipient

              AddRecepient txt_Recipient

              txt_Recipient = ""

        End If
 Else
    MsgBox "Please check Email or Fax.", vbInformation, "Imswin"

 End If
End Sub

Private Sub cmdRemove_Click()

Dim X As Integer

If Len(dgRecipientList.SelBookmarks(0)) = 0 Then

    MsgBox "Please make a selection first.", vbInformation, "Imswin"

    Exit Sub

 End If



    dgRecipientList.DeleteSelected

' dgRecipientList.SelBookmarks.RemoveAll

End Sub
Private Sub SSOLEDBFax_DblClick()
On Error Resume Next

    AddRecepient SSOLEDBFax.Columns(1).Value

    If Err Then Err.Clear
End Sub

Private Sub Txt_search_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

Dim rs As ADODB.Recordset

If SSOLEDBEmail.Visible = True Then Set rs = SSOLEDBEmail.DataSource

If SSOLEDBFax.Visible = True Then Set rs = SSOLEDBFax.DataSource

    If rs.AbsolutePosition <> adPosBOF And rs.AbsolutePosition <> adPosEOF And rs.AbsolutePosition <> adPosUnknown Then
    
            Call AddRecepient(rs.Fields(1).Value)    ': txt_Recipient = ""
            
    Else
            
        MsgBox "Recepient does not exist. Plese try a different one.", vbInformation, "Ims"
            
    End If
End If
End Sub

Private Sub UserControl_Initialize()
opt_Email.Value = True
Call opt_Email_Click
End Sub
