VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNavigators.ocx"
Begin VB.Form frm_StockRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Record"
   ClientHeight    =   5670
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7620
   Tag             =   "02010100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   2310
      TabIndex        =   40
      Top             =   5130
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelToolTipText=   "Undo the changes made to the current record"
      CloseToolTipText=   "Closes the current window"
      EMailEnabled    =   0   'False
      EmailToolTipText=   "Send current record via email"
      FirstToolTipText=   "Moves to the first record"
      LastToolTipText =   "Moves to the last record"
      NewEnabled      =   -1  'True
      NewToolTipText  =   "Adds a new record"
      NextToolTipText =   "Moves to the next record"
      PreviousToolTipText=   "Moves to the previous record"
      PrintToolTipText=   "Prints current record"
      SaveToolTipText =   "Save the changes made to the current record"
      DeleteToolTipText=   ""
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   7455
      TabIndex        =   33
      Top             =   120
      Width           =   7455
      Begin VB.OptionButton optPool 
         Alignment       =   1  'Right Justify
         Caption         =   "Pool"
         Height          =   255
         Left            =   5220
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optSpecific 
         Alignment       =   1  'Right Justify
         Caption         =   "Specific"
         Height          =   255
         Left            =   6315
         TabIndex        =   45
         Top             =   480
         Width           =   1065
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboCategory 
         Bindings        =   "frm_StockRecord.frx":0000
         DataField       =   "stk_catecode"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   44
         Top             =   450
         Width           =   2235
         DataFieldList   =   "cate_catecode"
         AllowInput      =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         stylesets(0).Picture=   "frm_StockRecord.frx":0016
         stylesets(0).AlignmentText=   0
         stylesets(1).Name=   "ColHeader"
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frm_StockRecord.frx":0032
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3942
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "cate_name"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "cate_catename"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1270
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "cate_catecode"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "cate_catecode"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Object.DataMember      =   "CATEGORY"
         DataFieldToDisplay=   "cate_catename"
      End
      Begin MSDataListLib.DataCombo dcboman3 
         Bindings        =   "frm_StockRecord.frx":004E
         DataField       =   "stk_man3"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         Top             =   2445
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "man_name"
         BoundColumn     =   "man_code"
         Text            =   ""
         Object.DataMember      =   "MANUFACTURER"
      End
      Begin MSDataListLib.DataCombo dcboman2 
         Bindings        =   "frm_StockRecord.frx":0064
         DataField       =   "stk_man2"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   2115
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "man_name"
         BoundColumn     =   "man_code"
         Text            =   ""
         Object.DataMember      =   "MANUFACTURER"
      End
      Begin MSDataListLib.DataCombo dcboman1 
         Bindings        =   "frm_StockRecord.frx":007A
         DataField       =   "stk_man1"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   41
         Top             =   1785
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "man_name"
         BoundColumn     =   "man_code"
         Text            =   ""
         Object.DataMember      =   "MANUFACTURER"
      End
      Begin VB.TextBox txt_Estimate 
         DataField       =   "stk_estmprice"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5220
         TabIndex        =   35
         Top             =   780
         Width           =   2160
      End
      Begin VB.TextBox txt_StockNum 
         DataField       =   "stk_stcknumb"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Top             =   120
         Width           =   2235
      End
      Begin VB.TextBox txt_Maximum 
         DataField       =   "stk_maxi"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5220
         TabIndex        =   28
         Top             =   3105
         Width           =   2160
      End
      Begin VB.TextBox txt_LongDescript 
         DataField       =   "stk_desc"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   1092
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   3765
         Width           =   6080
      End
      Begin VB.TextBox txt_ShortDescript 
         DataField       =   "stk_shrtdesc"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   3435
         Width           =   6080
      End
      Begin VB.TextBox txt_Standard 
         DataField       =   "stk_stdrcost"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5220
         TabIndex        =   24
         Top             =   2775
         Width           =   2160
      End
      Begin VB.TextBox txt_Minimum 
         DataField       =   "stk_mini"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Top             =   3105
         Width           =   2235
      End
      Begin VB.TextBox txt_Estimated1 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5610
         TabIndex        =   15
         Top             =   1785
         Width           =   1770
      End
      Begin VB.TextBox txt_MfctrNum1 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   3570
         TabIndex        =   14
         Top             =   1785
         Width           =   2020
      End
      Begin VB.TextBox txt_Estimated3 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5610
         TabIndex        =   22
         Top             =   2445
         Width           =   1770
      End
      Begin VB.TextBox txt_Estimated2 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5610
         TabIndex        =   18
         Top             =   2115
         Width           =   1770
      End
      Begin VB.TextBox txt_MfctrNum3 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   3570
         TabIndex        =   21
         Top             =   2445
         Width           =   2020
      End
      Begin VB.TextBox txt_MfctrNum2 
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   3570
         TabIndex        =   17
         Top             =   2115
         Width           =   2020
      End
      Begin MSDataListLib.DataCombo dcboPrimUnit 
         Bindings        =   "frm_StockRecord.frx":0090
         DataField       =   "stk_primuon"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Top             =   780
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "uni_desc"
         BoundColumn     =   "uni_code"
         Text            =   ""
         Object.DataMember      =   "UNIT"
      End
      Begin MSDataListLib.DataCombo dcboSecUnit 
         Bindings        =   "frm_StockRecord.frx":00A6
         DataField       =   "stk_secouom"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   37
         Top             =   1110
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "uni_desc"
         BoundColumn     =   "uni_code"
         Text            =   ""
         Object.DataMember      =   "UNIT"
      End
      Begin MSDataListLib.DataCombo dcboChargeAccount 
         Bindings        =   "frm_StockRecord.frx":00BC
         DataField       =   "stk_characctcode"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   2775
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cha_acctname"
         BoundColumn     =   "cha_acctcode"
         Text            =   ""
         Object.DataMember      =   "CHARGE"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboStockType 
         Bindings        =   "frm_StockRecord.frx":00D2
         DataField       =   "stk_stcktype"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5220
         TabIndex        =   39
         Top             =   120
         Width           =   2160
         DataFieldList   =   "sty_stcktype"
         _Version        =   196617
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
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   4154
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "sty_desc"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "sty_desc"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Capital/Expense"
         Columns(1).Name =   "sty_cenc"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "sty_cenc"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1032
         Columns(2).Caption=   "Own"
         Columns(2).Name =   "sty_owle"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "sty_owle"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Style=   2
         Columns(3).Width=   900
         Columns(3).Caption=   "Idea"
         Columns(3).Name =   "sty_idaeflag"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "sty_idaeflag"
         Columns(3).DataType=   11
         Columns(3).FieldLen=   256
         Columns(3).Style=   2
         _ExtentX        =   3810
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Object.DataMember      =   "STOCKTYPE"
         DataFieldToDisplay=   "sty_desc"
      End
      Begin VB.Label lbl_CompFactor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "stk_compfctr"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5220
         TabIndex        =   8
         Top             =   1110
         Width           =   2160
      End
      Begin VB.Label lbl_Category 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Left            =   45
         TabIndex        =   2
         Top             =   525
         Width           =   630
      End
      Begin VB.Label lbl_PrimaryUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Unit"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   870
         Width           =   840
      End
      Begin VB.Label lbl_SecondaryUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Unit"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl_StockNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Number"
         Height          =   195
         Left            =   45
         TabIndex        =   0
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label lbl_Charge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Account"
         Height          =   195
         Left            =   45
         TabIndex        =   19
         Top             =   2790
         Width           =   1155
      End
      Begin VB.Label lbl_Minimum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         Height          =   195
         Left            =   45
         TabIndex        =   25
         Top             =   3105
         Width           =   615
      End
      Begin VB.Label lbl_ShortDescript 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Description"
         Height          =   225
         Left            =   45
         TabIndex        =   29
         Top             =   3465
         Width           =   1260
      End
      Begin VB.Label lbl_Long 
         BackStyle       =   0  'Transparent
         Caption         =   "Long Description"
         Height          =   225
         Left            =   45
         TabIndex        =   31
         Top             =   3825
         Width           =   1260
      End
      Begin VB.Label lbl_Maximum 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         Height          =   225
         Left            =   3795
         TabIndex        =   27
         Top             =   3105
         Width           =   1170
      End
      Begin VB.Label lbl_Computed 
         BackStyle       =   0  'Transparent
         Caption         =   "Computed Factor"
         Height          =   225
         Left            =   3795
         TabIndex        =   7
         Top             =   1155
         Width           =   1305
      End
      Begin VB.Label lbl_Standard 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Cost"
         Height          =   225
         Left            =   3795
         TabIndex        =   23
         Top             =   2775
         Width           =   1170
      End
      Begin VB.Label lbl_Estimate 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimated Price"
         Height          =   225
         Left            =   3795
         TabIndex        =   5
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label lbl_PoolSpecific 
         BackStyle       =   0  'Transparent
         Caption         =   "Pool/Specific"
         Height          =   225
         Left            =   3795
         TabIndex        =   3
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label lbl_StockType 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Type"
         Height          =   225
         Left            =   3795
         TabIndex        =   1
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lbl_Estimated 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimated Price"
         Height          =   195
         Left            =   5940
         TabIndex        =   12
         Top             =   1545
         Width           =   1095
      End
      Begin VB.Label lbl_Manufacturer 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   270
         Left            =   1740
         TabIndex        =   10
         Top             =   1545
         Width           =   1170
      End
      Begin VB.Label lbl_MfcNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
         Height          =   225
         Left            =   3795
         TabIndex        =   11
         Top             =   1545
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer 1"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   13
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer 2"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   16
         Top             =   2145
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer 3"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   20
         Top             =   2505
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturers"
         Height          =   195
         Index           =   3
         Left            =   45
         TabIndex        =   9
         Top             =   1500
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frm_StockRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents stk As ADODB.Recordset
Attribute stk.VB_VarHelpID = -1

Private Sub dcboman1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 dcboman1.locked = False
  
End Sub

Private Sub dcboman1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  dcboman1.locked = Not NavBar1.NewEnabled
End Sub

Private Sub dcboman2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 dcboman2.locked = False
  
End Sub

Private Sub dcboman2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
  dcboman2.locked = Not NavBar1.NewEnabled
End Sub

Private Sub dcboPrimUnit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 dcboPrimUnit.locked = False
End Sub

Private Sub dcboPrimUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  dcboPrimUnit.locked = Not NavBar1.NewEnabled
End Sub

Private Sub dcboSecUnit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 dcboSecUnit.locked = False
  
End Sub

Private Sub dcboSecUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  dcboSecUnit.locked = Not NavBar1.NewEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
If open_forms <= 5 Then frmNavigator.Visible = True
    deIms.rsSTOCKMASTER.Filter = 0
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub Form_Load()
    Set NavBar1.Recordset = deIms.rsSTOCKMASTER
    Call DisableButtons(Me, NavBar1)
End Sub

Private Sub optPool_Click()
    If deIms.rsSTOCKMASTER!stk_poolspec = 0 Then _
       deIms.rsSTOCKMASTER!stk_poolspec = 1
End Sub

Private Sub optSpecific_Click()
    If deIms.rsSTOCKMASTER!stk_poolspec <> 0 Then _
       deIms.rsSTOCKMASTER!stk_poolspec = 0
End Sub

Private Sub stk_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If Not ((stk.EOF) Or (stk.BOF)) Then
        If stk!stk_poolspec = False Then
           optSpecific = True
        ElseIf stk!stk_poolspec = True Then
            optPool = True
        Else
            optPool = False
            optSpecific = False
        End If
    End If
End Sub


Private Sub cmd_Compute_Click()
'    MsgBox (cbo_PrimaryUnit.Text & "=" & cbo_SecondaryUnit.Text & "?")
End Sub

Private Sub lst_PoolSpecific_ItemCheck(Item As Integer)

End Sub

