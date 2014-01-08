VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frmTrackManifest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracking Message for Manifest"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   9570
   Tag             =   "02030300"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmTrackManifast.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label20"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LblMessaDate"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblStatu"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "SScboRecip5"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "SScboRecip4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "SScboRecip3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "SScboRecip2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "SScboRecip1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "NavBar1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "SScmbMessage"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "SScboSubject"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "SScboForwarder"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "SScboSupRecipFax"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "DTPickArrival"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "DTPickEstimatedate"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "DTPickNewDelivery"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "SScboPriority"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtOperator"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "chkYesorNo"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmbPOnumber"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "Text"
      TabPicture(1)   =   "frmTrackManifast.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRemark"
      Tab(1).ControlCount=   1
      Begin VB.ComboBox cmbPOnumber 
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkYesorNo 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6720
         TabIndex        =   29
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox txtRemark 
         Height          =   4215
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   480
         Width           =   8175
      End
      Begin VB.TextBox txtOperator 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboPriority 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1680
         Width           =   1815
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   3096
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1905
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker DTPickNewDelivery 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   72089603
         CurrentDate     =   36549
      End
      Begin MSComCtl2.DTPicker DTPickEstimatedate 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   72089603
         CurrentDate     =   36549
      End
      Begin MSComCtl2.DTPicker DTPickArrival 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   72089603
         CurrentDate     =   36549
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboSupRecipFax 
         Height          =   315
         Left            =   5880
         TabIndex        =   7
         Top             =   960
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   4260
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Fax Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4233
         Columns(1).Caption=   "Supplier"
         Columns(1).Name =   "Supplier"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1720
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboForwarder 
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboSubject 
         Height          =   315
         Left            =   5880
         TabIndex        =   27
         Top             =   600
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   4233
         Columns(0).Caption=   "Subject"
         Columns(0).Name =   "Subject"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScmbMessage 
         Height          =   315
         Left            =   2160
         TabIndex        =   31
         Top             =   960
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
         Columns(1).Caption=   "Message Date"
         Columns(1).Name =   "PO-Number"
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
         Left            =   1800
         TabIndex        =   33
         Top             =   4080
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   661
         ButtonHeight    =   329.953
         ButtonWidth     =   345.26
         Style           =   1
         MouseIcon       =   "frmTrackManifast.frx":0038
         PreviousVisible =   0   'False
         LastVisible     =   0   'False
         NextVisible     =   0   'False
         FirstVisible    =   0   'False
         EMailVisible    =   -1  'True
         PrintEnabled    =   0   'False
         SaveEnabled     =   0   'False
         CancelEnabled   =   0   'False
         DeleteEnabled   =   -1  'True
         EditEnabled     =   -1  'True
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip1 
         Height          =   315
         Left            =   5880
         TabIndex        =   34
         Top             =   1680
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip2 
         Height          =   315
         Left            =   5880
         TabIndex        =   35
         Top             =   2040
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip3 
         Height          =   315
         Left            =   5880
         TabIndex        =   36
         Top             =   2400
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip4 
         Height          =   315
         Left            =   5880
         TabIndex        =   37
         Top             =   2760
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip5 
         Height          =   315
         Left            =   5880
         TabIndex        =   38
         Top             =   3120
         Width           =   2535
         DataFieldList   =   "Column 0"
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
         Columns.Count   =   3
         Columns(0).Width=   3836
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3466
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1323
         Columns(2).Caption=   "Code"
         Columns(2).Name =   "Code"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
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
         Left            =   4800
         TabIndex        =   39
         Top             =   3960
         Width           =   3540
      End
      Begin VB.Label LblMessaDate 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Include Original Message"
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Estim. Arrival"
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   3120
         Width           =   1900
      End
      Begin VB.Label Label16 
         Caption         =   "Estim. Departure"
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   1900
      End
      Begin VB.Label Label15 
         Caption         =   "New Delivery Date"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1900
      End
      Begin VB.Label Label14 
         Caption         =   "Operator"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1900
      End
      Begin VB.Label Label13 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   20
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   18
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Forwarder Recipient"
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Supplier Recipient"
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Subject"
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Message Date"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1900
      End
      Begin VB.Label Label4 
         Caption         =   "Ship Via"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1900
      End
      Begin VB.Label Label3 
         Caption         =   "Message #"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1900
      End
      Begin VB.Label Label2 
         Caption         =   "Manifest #"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1900
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo6 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      Text            =   "SSOleDBCombo1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tracking Message For Manifest"
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
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmTrackManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tracklist As imsTrackingMessage
Dim Form As FormMode
Dim NAV_NEW As Boolean
Dim NAV_SAVE As Boolean

Private Sub chkYesorNo_GotFocus()
Call HighlightBackground(chkYesorNo)
End Sub

Private Sub chkYesorNo_LostFocus()
Call NormalBackground(chkYesorNo)
End Sub

'get tracking manifest list for form, and set navbar button

Private Sub cmbPOnumber_Click()
Dim exist As Integer

        Call ClaerForm
        SScmbMessage.RemoveAll
        SScmbMessage = ""
    If Len(cmbPOnumber) Then Call GetOBSMessagelist
    
        exist = CheckMessageNumber(cmbPOnumber)
        
        If exist > 0 Then
            EnableControls (False)
            SScmbMessage.Enabled = True
            NavBar1.CancelEnabled = False
            NavBar1.SaveEnabled = False
        Else
            EnableControls (False)
            NavBar1.PrintEnabled = False
            NavBar1.EMailEnabled = False
        End If
'      NavBar1.EditEnabled = True
     
End Sub

Private Sub cmbPOnumber_GotFocus()
Call HighlightBackground(cmbPOnumber)
End Sub

Private Sub cmbPOnumber_LostFocus()
Call NormalBackground(cmbPOnumber)
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set Tracklist = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'click cancel button clear the form

Private Sub NavBar1_OnCancelClick()
' If Not Len(cmbPOnumber) = 0 And Not Len(SScmbMessage.Columns(0).Text) = 0 _
'       And Not Len(SScmbMessage.Columns(1).Text) = 0 Then
'
'            Call GetOBSList
'      Else
            SScboPriority = ""
            txtOperator = ""
            DTPickNewDelivery = ""
            DTPickEstimatedate = ""
            DTPickArrival = ""
            SScboSubject = ""
            SScboSupRecipFax = ""
            SScboForwarder = ""
            SScboRecip1 = ""
            SScboRecip2 = ""
            SScboRecip3 = ""
            SScboRecip4 = ""
            SScboRecip5 = ""
            chkYesorNo.Value = 0
            txtRemark = ""
'    End If
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
   
    Unload Me
    
End Sub

'click new button set date, new manifest number

Private Sub NavBar1_OnEditClick()
Dim str As String
   If Not Len(SScmbMessage.Columns(0).text) = 0 And Not Len(SScmbMessage.Columns(1).text) = 0 Then
            
        NavBar1.EditEnabled = True
       
        
        str = Format$(Now(), "mm/dd/yyyy")

        Call ClaerForm
        Call GetMessageNumber
        
        LblMessaDate = str
           
    Else
        
        NavBar1.NewEnabled = True
        
    End If
End Sub

'click email button,get report path and reciptent list
Private Sub NavBar1_OnEMailClick()
Dim at() As String
Dim IFile As IMSFile
Dim FileName(1) As String
Dim rsr As ADODB.Recordset
Dim rptinfo As RPTIFileInfo


    FileName(1) = FixDir(App.Path) & "po.rpti"
    FileName(0) = FixDir(App.Path) & "obs.rpti"
    
    Set rsr = GetObsRecipients(deIms.NameSpace, cmbPOnumber, SScmbMessage.Columns(0).text)
    
    Set IFile = New IMSFile
    
    'at(1) = 'FixDir(App.Path) & "po.rpti"
    'at(0) = 'FixDir(App.Path) & "Obs.rpti"
    
    SetOBsReportParam
    
    
    ReDim at(2)
    With rptinfo
        at(2) = "ponumb=" + Trim$(cmbPOnumber)
        at(0) = "namespace=" + deIms.NameSpace
        .ReportFileName = ReportPath & "obs.rpt"
        at(1) = "mesgnumb=" + Trim$(SScmbMessage.Columns("MessageNumber").text)
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("obs.rpt") 'J added
        '---------------------------------------------
        
        .Parameters = at
    End With
    
    Call WriteRPTIFile(rptinfo, FileName(0))
    
    
    If chkYesorNo.Value = 1 Then
        
        
        SetPOReportParam
        
        With rptinfo
            ReDim .Parameters(1)
            .Parameters(0) = "namespace=" + deIms.NameSpace
            .Parameters(1) = "manifestnumb=" + Trim$(cmbPOnumber)
            .ReportFileName = ReportPath & "packinglist.rpt"
            
            'Modified by Juan (8/28/2000) for Multilingual
            Call translator.Translate_Reports("packinglist.rpt") 'J added
            '---------------------------------------------
            
            '.Parameters = Params
        End With
        
        Call WriteRPTIFile(rptinfo, FileName(1))

        'Call MDI_IMS.SaveReport(at(1), crptWinWord)
        '
        'If IFile.FileExists(at(0)) = False Then _
            MsgBox "Error Saving report": Exit Sub
            
        'If IFile.FileExists(at(1)) = False Then _
            MsgBox "Error Saving report": Exit Sub
        
    End If
    
    ReDim at(1)
    
    If chkYesorNo.Value = 0 Then
        Call SendEmailAndFax(rsr, "Recipient", Me.Caption, "", FileName(0))
    Else
        Call SendEmailAndFaxWithAttachments(rsr, "Recipient", Me.Caption, "", FileName, False)
    End If

End Sub

'click new set date and new message number, navbar button

Private Sub NavBar1_OnNewClick()
Dim str As String
    'If Len(SScmbMessage.Columns(0).Text) = 0 And Len(SScmbMessage.Columns(1).Text) = 0 Then
    
        str = Format$(Now(), "mm/dd/yyyy")

        Call ClaerForm
        Call ChangeMode(mdCreation)
        Call GetMessageNumber
        
        LblMessaDate = str
        txtOperator = CurrentUser
        NavBar1.EditEnabled = False
        NavBar1.SaveEnabled = True
        NavBar1.CancelEnabled = True
        Call EnableControls(True)
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
    'Else
     '   MsgBox "The message already exist, Please click message number to see"
        
    'End If

End Sub

'click save check recipient list, send data to tracking message class
'and reset navbar button

Private Sub NavBar1_OnSaveClick()
Set Tracklist = New imsTrackingMessage

     If Len(SScboSupRecipFax) = 0 And Len(SScboForwarder) = 0 And Len(SScboRecip1) = 0 _
       And Len(SScboRecip2) = 0 And Len(SScboRecip3) = 0 And Len(SScboRecip4) = 0 _
       And Len(SScboRecip5) = 0 Then
       
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("M00371") 'J added
        MsgBox IIf(msg1 = "", "You must choose one recipients", msg1) 'J modified
        '---------------------------------------------
        
        Exit Sub
    End If
    
    Call Tracklist.InsertandUpdateTable(cmbPOnumber, deIms.NameSpace, SScmbMessage, SScboSubject, _
    0, LblMessaDate, SScboSupRecipFax, SScboForwarder, SScboRecip1, SScboRecip2, SScboRecip3, _
    SScboRecip4, SScboRecip5, txtOperator, DTPickNewDelivery, DTPickEstimatedate, DTPickArrival, _
    SScboPriority.Columns("code").text, chkYesorNo, txtRemark, CurrentUser, deIms.cnIms)
    
    EnableControls False
    Call ChangeMode(mdVisualization)
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.PrintEnabled = True
    NavBar1.EMailEnabled = True
'    Call ClaerForm
'    cmbPOnumber = ""
End Sub

Private Sub SScboForwarder_GotFocus()
Call HighlightBackground(SScboForwarder)
End Sub

Private Sub SScboForwarder_LostFocus()
Call NormalBackground(SScboForwarder)
End Sub

Private Sub SScboPriority_GotFocus()
Call HighlightBackground(SScboPriority)
End Sub

Private Sub SScboPriority_LostFocus()
Call NormalBackground(SScboPriority)
End Sub

Private Sub SScboRecip1_GotFocus()
Call HighlightBackground(SScboRecip1)
End Sub

Private Sub SScboRecip1_LostFocus()
Call NormalBackground(SScboRecip1)
End Sub

Private Sub SScboRecip2_GotFocus()
Call HighlightBackground(SScboRecip2)
End Sub

Private Sub SScboRecip2_LostFocus()
Call NormalBackground(SScboRecip2)
End Sub

Private Sub SScboRecip3_GotFocus()
Call HighlightBackground(SScboRecip3)
End Sub

Private Sub SScboRecip3_LostFocus()
Call NormalBackground(SScboRecip3)
End Sub

Private Sub SScboRecip4_GotFocus()
Call HighlightBackground(SScboRecip4)
End Sub

Private Sub SScboRecip4_LostFocus()
Call NormalBackground(SScboRecip4)
End Sub

Private Sub SScboRecip5_GotFocus()
Call HighlightBackground(SScboRecip5)
End Sub

Private Sub SScboRecip5_LostFocus()
Call NormalBackground(SScboRecip5)
End Sub

Private Sub SScboSubject_GotFocus()
Call HighlightBackground(SScboSubject)
End Sub

Private Sub SScboSubject_LostFocus()
Call NormalBackground(SScboSubject)
End Sub

Private Sub SScboSupRecipFax_GotFocus()
Call HighlightBackground(SScboSupRecipFax)
End Sub

Private Sub SScboSupRecipFax_LostFocus()
Call NormalBackground(SScboSupRecipFax)
End Sub

'get manifest list information and set button

Private Sub SScmbMessage_Click()
 If Not Len(SScmbMessage.Columns(0).text) = 0 And Not Len(SScmbMessage.Columns(1).text) = 0 Then
        Call GetOBSList
'        Call EnableControls(True)
        NavBar1.EditEnabled = True
        NavBar1.NewEnabled = True
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
    Else
         If Not Len(SScmbMessage.Columns(0).text) = 0 Then
            NavBar1.SaveEnabled = True
            NavBar1.CancelEnabled = True
        End If
'        NavBar1.EditEnabled = False
'        NavBar1.NewEnabled = True
    End If
            NavBar1.NewEnabled = True
End Sub

'load form set date, get lists for all combo and set button
Private Sub Form_Load()
Dim str As String
    
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmTrackManifest")
    '------------------------------------------
    
    str = Format$(Now(), "mm/dd/yyyy")
    DTPickArrival.Value = str
    DTPickEstimatedate.Value = str
    DTPickNewDelivery.Value = str
'    deIms.cnIms.Open
'    deIms.Namespace = "SAKHA"
    
    GetPackingnumber
    
    
    GetSubjectCode
    getPRIORITYlist
    GetSupplierFax
    GetSupplierEmail
    GetForwarderFax
    GetForwarderEmail
    GetPhoneDriFax
    GetPhoneDirMail
    LblMessaDate = str
 
    EnableControls (False)
    NavBar1.EMailEnabled = True
    NavBar1.EMailVisible = True
    
       Call DisableButtons(Me, NavBar1)
'    Call PopuLateFromRecordSet(cmbMessage, Tracklist.GetMessageNumberlist(cmbPOnumber, deIms.Namespace, deIms.cnIms), "ob_ponumb", True)
'    If cmbMessage.ListCount Then cmbMessage.ListIndex = 0
    Caption = Caption + " - " + Tag
    
    NAV_NEW = NavBar1.NewEnabled
    NAV_SAVE = NavBar1.SaveEnabled
    
    Form = mdVisualization
End Sub

'SQL statement get priority list

Public Sub getPRIORITYlist()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT pri_code, pri_desc"
        .CommandText = .CommandText & " From PRIORITY "
        .CommandText = .CommandText & " WHERE pri_npecode = '" & deIms.NameSpace & "'"
         Set rst = .Execute
    End With
    

    str = Chr$(1)
    SScboPriority.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SScboPriority.AddItem rst!pri_desc & str & (rst!pri_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub cmdExit_Click()
    'Unload Me
End Sub


'SQL statement get supplier fax list

Public Sub GetSupplierFax()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT   sup_faxnumb,sup_name,sup_code "
        .CommandText = .CommandText & " From Supplier "
        .CommandText = .CommandText & " WHERE sup_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and sup_faxnumb  IS NOT NULL"
        .CommandText = .CommandText & " order by sup_name "
    
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboSupRecipFax.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboSupRecipFax.AddItem rst!sup_faxnumb & "" & str & rst!sup_name & "" & str & rst!sup_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    
End Sub

'SQL statement, get supplier email list

Public Sub GetSupplierEmail()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT   sup_mail,sup_name,sup_code "
        .CommandText = .CommandText & " From Supplier "
        .CommandText = .CommandText & " WHERE sup_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and sup_mail  IS NOT NULL"
        .CommandText = .CommandText & " order by sup_name "
    
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboSupRecipFax.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboSupRecipFax.AddItem rst!sup_mail & "" & str & rst!sup_name & "" & str & rst!sup_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'SQL statement,get forwarder fax number list

Public Sub GetForwarderFax()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT forw_faxnumb, forw_name, forw_code "
        .CommandText = .CommandText & " From FORWARDER "
        .CommandText = .CommandText & " WHERE forw_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and forw_faxnumb  IS NOT NULL"
        .CommandText = .CommandText & " order by forw_name "
    
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboForwarder.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboForwarder.AddItem rst!forw_faxnumb & "" & str & rst!forw_name & "" & str & rst!forw_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

'SQL statement get forwarder email list number

Public Sub GetForwarderEmail()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT forw_mail, forw_name, forw_code "
        .CommandText = .CommandText & " From FORWARDER "
        .CommandText = .CommandText & " WHERE forw_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and forw_mail  IS NOT NULL"
        .CommandText = .CommandText & " order by forw_name "
    
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboForwarder.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboForwarder.AddItem rst!forw_mail & "" & str & rst!forw_name & "" & str & rst!forw_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

'SQL statement, get manifest number list

Public Sub GetPackingnumber()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT pl_manfnumb From PACKINGLIST "
        .CommandText = .CommandText & " WHERE pl_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by pl_manfnumb "

        Set rst = .Execute
    End With
    
'    str = Chr$(1)
'    cmbPOnumber.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        cmbPOnumber.AddItem rst!pl_manfnumb & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing



End Sub

'Call store procedure, get new message number and return result

Public Function GetMessageNumber() As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GetMessageNumber"
        Set .ActiveConnection = deIms.cnIms
        
        
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cmbPOnumber)
        .Parameters.Append .CreateParameter("@STRING", adVarChar, adParamOutput, 20, GetMessageNumber)
        
        .Execute , , adExecuteNoRecords
        
        GetMessageNumber = .Parameters("@STRING").Value & ""
    End With
        
        SScmbMessage = GetMessageNumber
    
    Set cmd = Nothing
   
End Function

'add subject to suject combo

Public Sub GetSubjectCode()
Dim str As String
    str = SScboSubject.FieldSeparator
    
    SScboSubject.AddItem " Suppliers Follow Up "
    SScboSubject.AddItem " Shipping Information "
    SScboSubject.AddItem " MIscellaneous "
End Sub

'clear manifest form

Public Sub ClaerForm()
'    cmbPOnumber = ""
    SScmbMessage = ""
    LblMessaDate = ""
    SScboPriority = ""
    txtOperator = ""
    DTPickNewDelivery = ""
    DTPickEstimatedate = ""
    DTPickArrival = ""
    SScboSubject = ""
    SScboSupRecipFax = ""
    SScboForwarder = ""
    SScboRecip1 = ""
    SScboRecip2 = ""
    SScboRecip3 = ""
    SScboRecip4 = ""
    SScboRecip5 = ""
    chkYesorNo.Value = 0
    txtRemark = ""
End Sub

'SQL statement, get manifest message list information

Public Sub GetOBSMessagelist()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT OBS.ob_mesgnumb, OBS.ob_mesgdate"
        .CommandText = .CommandText & " FROM OBS INNER JOIN PACKINGLIST "
        .CommandText = .CommandText & " ON OBS.ob_ponumb = PACKINGLIST.pl_manfnumb AND "
        .CommandText = .CommandText & " OBS.ob_npecode = PACKINGLIST.pl_npecode "
        .CommandText = .CommandText & " WHERE (OBS.ob_ponumb = '" & cmbPOnumber & "') AND "
        .CommandText = .CommandText & " OBS.ob_npecode  = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND OBS.ob_flag = 0 "
         Set rst = .Execute
    End With
    
    str = Chr$(1)
    
'    cmdEdit.Enabled = True
    EnableControls True
    NavBar1.EditEnabled = False

    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = True
    NavBar1.PrintEnabled = True
    NavBar1.SaveEnabled = True
    
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

'check recipient list, give message

Public Sub CheckRecipient()

    If Len(SScboSupRecipFax) = 0 And Len(SScboForwarder) = 0 And Len(SScboRecip1) = 0 _
       And Len(SScboRecip2) = 0 And Len(SScboRecip3) = 0 And Len(SScboRecip4) = 0 _
       And Len(SScboRecip5) = 0 Then
       
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("M00369") 'J added
        MsgBox IIf(msg1 = "", "You must choose one", msg1) 'J modified
        '---------------------------------------------
        
    End If

End Sub

'enable controls, without lable, sstab, navbar button

Public Sub EnableControls(bEnable As Boolean)
On Error Resume Next
Dim ctl As Control

    For Each ctl In Controls
        If (Not ((TypeOf ctl Is Label) Or (TypeOf ctl Is SSTab) Or (TypeOf ctl Is NavBar))) Then
            ctl.Enabled = bEnable
'            NavBar1.Enabled = False
        
           If Err Then Err.Clear
        End If
            
    Next ctl
    
    cmbPOnumber.Enabled = True
    NavBar1.CloseEnabled = True
'    NavBar1.EditEnabled = False
'    NavBar1.NewEnabled = False
'    NavBar1.CancelEnabled = False
'    NavBar1.PrintEnabled = False
'    NavBar1.SaveEnabled = False
'
    
'    cmdExit.Enabled = True
End Sub

'Call store procedure get manifest list information

Public Sub GetOBSList()
Dim cmd As ADODB.Command
Dim rst As Recordset

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GetOBSListManifest"
        Set .ActiveConnection = deIms.cnIms
        
        .Parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, cmbPOnumber)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@mesgnumb", adVarChar, adParamInput, 15, SScmbMessage)
        
         Set rst = .Execute
'
'        GetMessageNumber = .Parameters("@STRING").Value & ""
    End With
        
'        cmbPOnumber = RST!ob_ponumb & ""
'        SScmbMessage = RST!ob_mesgnumb & ""
        LblMessaDate = rst!ob_mesgdate & ""
        SScboSubject = rst!ob_subj & ""
        SScboSupRecipFax = rst!ob_suppreci & ""
        SScboForwarder = rst!ob_forwreci & ""
        SScboRecip1 = rst!ob_rec1 & ""
        SScboRecip2 = rst!ob_rec2 & ""
        SScboRecip3 = rst!ob_rec3 & ""
        SScboRecip4 = rst!ob_rec4 & ""
        SScboRecip5 = rst!ob_rec5 & ""
        txtOperator = rst!ob_oper & ""
        DTPickNewDelivery = rst!ob_newdelvdate & ""
        DTPickEstimatedate = rst!ob_etd & ""
        DTPickArrival = rst!ob_eta & ""
        chkYesorNo.Value = IIf((rst!ob_inclmesg), 1, 0)
        SScboPriority = rst!ob_shipvia & ""
        txtRemark = rst!ob_remk & ""
          
    Set cmd = Nothing
     
   
End Sub

'SQL statement,get phone directory fax number

Public Sub GetPhoneDriFax()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT phd_name, phd_faxnumb, phd_code "
        .CommandText = .CommandText & " From PHONEDIR "
        .CommandText = .CommandText & " WHERE phd_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and phd_faxnumb  IS NOT NULL"
        .CommandText = .CommandText & " order by phd_faxnumb "
    
 
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboRecip1.FieldSeparator = str
    SScboRecip2.FieldSeparator = str
    SScboRecip3.FieldSeparator = str
    SScboRecip4.FieldSeparator = str
    SScboRecip5.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboRecip1.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip2.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip3.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip4.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip5.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

'SQL statement,get phone directory email number

Public Sub GetPhoneDirMail()
Dim str As String
Dim cmd As Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT phd_name, phd_mail, phd_code "
        .CommandText = .CommandText & " From PHONEDIR "
        .CommandText = .CommandText & " WHERE phd_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "  and phd_mail  IS NOT NULL"
        .CommandText = .CommandText & " order by phd_mail "
    
 
        Set rst = .Execute
    End With
    
    str = Chr$(1)
    SScboRecip1.FieldSeparator = str
    SScboRecip2.FieldSeparator = str
    SScboRecip3.FieldSeparator = str
    SScboRecip4.FieldSeparator = str
    SScboRecip5.FieldSeparator = str
    
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        SScboRecip1.AddItem rst!phd_mail & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip2.AddItem rst!phd_mail & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip3.AddItem rst!phd_mail & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip4.AddItem rst!phd_mail & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip5.AddItem rst!phd_mail & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'get parament for crystal report and application path

Private Sub SetOBsReportParam()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\obs.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("obs.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "mesgnumb;" + SScmbMessage.text + ";TRUE"
        .ParameterFields(2) = "ponumb;" + cmbPOnumber + ";true"
    End With
End Sub

'get po parameter for crystal report and appliction path

Private Sub SetPOReportParam()

    If chkYesorNo.Value = 1 Then
        With MDI_IMS.CrystalReport1
            .Reset
            .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
            
            'Modified by Juan (8/28/2000) for Multilingual
            Call translator.Translate_Reports("packinglist.rpt") 'J added
            '---------------------------------------------
            
            .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            .ParameterFields(1) = "manifestnumb;" + cmbPOnumber + ";true"
        End With
    End If
    
End Sub

'print crystal report

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    
    SetOBsReportParam
'    MDI_IMS.CrystalReport1.Reset
    MDI_IMS.CrystalReport1.WindowTitle = "Tracking Message"
    MDI_IMS.CrystalReport1.Action = 1: MDI_IMS.CrystalReport1.Reset
   
    If chkYesorNo.Value = 1 Then
        SetPOReportParam
'        MDI_IMS.CrystalReport1.Reset
        MDI_IMS.CrystalReport1.WindowTitle = "Tracking Message"
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

'SQL statement check message number exist or not

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
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckMessageNumber = rst!rt
    End With
        
 
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckMessageNumber", Err.Description, Err.number, True)
End Function

Private Sub SScmbMessage_GotFocus()
Call HighlightBackground(SScmbMessage)
End Sub

Private Sub SScmbMessage_LostFocus()
Call NormalBackground(SScmbMessage)
End Sub

'on tab click set navbar button

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Dim iEditMode As String, blFlag As Boolean

  blFlag = SSTab1.Tab = 0 And Form = mdCreation
    
    With NavBar1
        .SaveEnabled = SSTab1.Tab = 0
        .CloseEnabled = SSTab1.Tab = 0
        .NewEnabled = SSTab1.Tab = 0
        .CancelEnabled = SSTab1.Tab = 0
        .EditEnabled = SSTab1.Tab = 0
        .PrintEnabled = cmbPOnumber.ListIndex <> CB_ERR
        .EMailEnabled = SSTab1.Tab = 0
    End With
    
    If Form = mdVisualization And SSTab1.Tab = 0 Then
        If Len(cmbPOnumber) <> 0 And Len(SScmbMessage) <> 0 Then
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
            NavBar1.CancelEnabled = False
            NavBar1.SaveEnabled = False
        End If
        
            NavBar1.NewEnabled = NAV_NEW
            NavBar1.CancelEnabled = NAV_NEW
        
    End If
    
    If Form = mdCreation And SSTab1.Tab = 0 Then
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
    End If
    
    If Form = mdVisualization Then
         NavBar1.CancelEnabled = False
         NavBar1.SaveEnabled = False
    End If
    
End Sub

'Change form mode show caption text

Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next

    
    If FMode = mdCreation Then
        lblStatu.ForeColor = vbRed
        
        'Modified by Juan (10/10/2000) for Multilingual
        msg1 = translator.Trans("L00125")  'J added
        lblStatu.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '----------------------------------------------
        
        ChangeMode = True
'    ElseIf FMode = mdModification Then
'        lblStatu.ForeColor = vbBlue
'        lblStatu.Caption = "Modification"
  
    ElseIf FMode = mdVisualization Then
        lblStatu.ForeColor = vbGreen
        
        'Modified by Juan (10/10/2000) for Multilingual
        msg1 = translator.Trans("L00125")  'J added
        lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '----------------------------------------------
        
    End If
    
       
    Form = FMode

End Function
