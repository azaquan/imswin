VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmTracking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracking Message for PO"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   9285
   Tag             =   "02020200"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frm_TrackingPO.frx":0000
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
      TabPicture(1)   =   "frm_TrackingPO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRemark"
      Tab(1).ControlCount=   1
      Begin VB.ComboBox cmbPOnumber 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkYesorNo 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6600
         TabIndex        =   15
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox txtRemark 
         Height          =   3495
         Left            =   -74400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   600
         Width           =   7455
      End
      Begin VB.TextBox txtOperator 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1815
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboPriority 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
         DataFieldList   =   "Column 1"
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
         DataFieldToDisplay=   "Column 0"
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
         Format          =   61800451
         CurrentDate     =   36553
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
         Format          =   61800451
         CurrentDate     =   36553
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
         Format          =   61800451
         CurrentDate     =   36553
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboSupRecipFax 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   960
         Width           =   2535
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
         Columns(0).Width=   4260
         Columns(0).Caption=   "Number"
         Columns(0).Name =   "Fax Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboForwarder 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
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
         DataFieldToDisplay=   "Column 2"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboSubject 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
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
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScmbMessage 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
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
         TabIndex        =   38
         Top             =   4080
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   661
         ButtonHeight    =   329.953
         ButtonWidth     =   345.26
         Style           =   1
         MouseIcon       =   "frm_TrackingPO.frx":0038
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip1 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
         DataFieldList   =   "Column 1"
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Number"
         Columns(1).Name =   "Number"
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
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip2 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   2040
         Width           =   2535
         DataFieldList   =   "Column 2"
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 1"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Number"
         Columns(1).Name =   "Number"
         Columns(1).DataField=   "Column 0"
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
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip3 
         Height          =   315
         Left            =   6000
         TabIndex        =   12
         Top             =   2400
         Width           =   2535
         DataFieldList   =   "Column 2"
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 1"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Number"
         Columns(1).Name =   "Number"
         Columns(1).DataField=   "Column 0"
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
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip4 
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   2760
         Width           =   2535
         DataFieldList   =   "Column 2"
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 1"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Number"
         Columns(1).Name =   "Number"
         Columns(1).DataField=   "Column 0"
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
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SScboRecip5 
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   3120
         Width           =   2535
         DataFieldList   =   "Column 2"
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 1"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Number"
         Columns(1).Name =   "Number"
         Columns(1).DataField=   "Column 0"
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
         DataFieldToDisplay=   "Column 0"
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
         Width           =   3660
      End
      Begin VB.Label LblMessaDate 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         TabIndex        =   37
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Include Original Message"
         Height          =   315
         Left            =   4200
         TabIndex        =   36
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "ETA Destination"
         Height          =   315
         Left            =   360
         TabIndex        =   35
         Top             =   3120
         Width           =   1800
      End
      Begin VB.Label Label16 
         Caption         =   "ETD Customs"
         Height          =   315
         Left            =   360
         TabIndex        =   34
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Label Label15 
         Caption         =   "New Delivery Date"
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label14 
         Caption         =   "Operator"
         Height          =   315
         Left            =   360
         TabIndex        =   32
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label Label13 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   31
         Top             =   3120
         Width           =   1800
      End
      Begin VB.Label Label12 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   30
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label10 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Recipients"
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label8 
         Caption         =   "Forwarder Recipient"
         Height          =   315
         Left            =   4200
         TabIndex        =   26
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "Supplier Recipient"
         Height          =   315
         Left            =   4200
         TabIndex        =   25
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Subject"
         Height          =   315
         Left            =   4200
         TabIndex        =   24
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "Message Date"
         Height          =   315
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Ship Via"
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Message #"
         Height          =   315
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Transaction #"
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   1800
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo6 
      Height          =   315
      Left            =   1920
      TabIndex        =   20
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
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tracklist As imsTrackingMessage
Dim Form As FormMode

Dim SaveEnabled As Boolean
'Change form mode show caption text

Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next

    
    If FMode = mdCreation Then
        lblStatu.ForeColor = vbRed
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00125") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '---------------------------------------------
        
        ChangeMode = True
'    ElseIf FMode = mdModification Then
'        lblStatu.ForeColor = vbBlue
'        lblStatu.Caption = "Modification"
  
    ElseIf FMode = mdVisualization Then
        lblStatu.ForeColor = vbGreen
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00092") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
        
    End If
    
       
    Form = FMode

End Function

Private Sub chkYesorNo_GotFocus()
Call HighlightBackground(chkYesorNo)
End Sub

Private Sub chkYesorNo_LostFocus()
Call NormalBackground(chkYesorNo)
End Sub

'SQL statement get po date and load messege list
'check meessege number exist or not

Private Sub cmbPOnumber_Click()

Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


#If DBUG = 0 Then
    On Error Resume Next
#End If
Dim exist As Integer
Dim rsPO As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim query As String
    Call ClaerForm
    SScmbMessage.RemoveAll
    If Len(cmbPOnumber) Then Call GetOBSMessagelist
     
     
     query = "SELECT po_date, po_suppcode, sup_mail,sup_faxnumb"
     query = query & " From PO, supplier "
     query = query & " WHERE po_ponumb = '" & cmbPOnumber & "' AND"
     query = query & " po_npecode = '" & deIms.NameSpace & " ' AND po_suppcode = sup_code AND"
     query = query & "  po_npecode = sup_npecode"
     
      rsPO.ActiveConnection = deIms.cnIms
      rsPO.Open query
      
       
'''    Set cmd = New ADODB.Command
'''
'''    With cmd
'''        .CommandType = adCmdText
'''        Set .ActiveConnection = deIms.cnIms
'''        .CommandText = "SELECT ? = po_date FROM po"
'''        .CommandText = .CommandText & " WHERE po_ponumb = '" & cmbPOnumber & "'"
'''        .CommandText = .CommandText & " AND po_npecode = '" & deIms.NameSpace & "'"
'''        .Parameters.Append .CreateParameter("", adDBTimeStamp, adParamOutput, Date)
'''       ' .Parameters.Append .CreateParameter("@SUPP", adVarChar, adParamOutput, 20)
'''        .Execute 0, , adExecuteNoRecords
        'DTPickNewDelivery.Value = .Parameters(0).Value
        
        DTPickNewDelivery.Value = rsPO!PO_Date
        SScboSupRecipFax.RemoveAll
        
        If Len((rsPO!Sup_mail) & "") > 0 Then
            SScboSupRecipFax.AddItem rsPO!Sup_mail
            SScboSupRecipFax = rsPO!Sup_mail
        Else
            SScboSupRecipFax.AddItem rsPO!sup_faxnumb
            SScboSupRecipFax = rsPO!sup_faxnumb
        End If
        Set rsPO = Nothing
    'End With
    
          exist = CheckMessageNumber(cmbPOnumber)
        
        If exist > 0 Then
            EnableControls (False)
            SScmbMessage.Enabled = True
        Else
            EnableControls (False)
            NavBar1.PrintEnabled = False
            NavBar1.EMailEnabled = False
        End If
        If cmbPOnumber <> "" Then NavBar1.NewEnabled = SaveEnabled 'J added
End Sub

Public Sub cmbPOnumber_DropDown()

End Sub

Private Sub cmbPOnumber_GotFocus()
Call HighlightBackground(cmbPOnumber)
End Sub

Private Sub cmbPOnumber_KeyPress(KeyAscii As Integer)
Dim i
    With cmbPOnumber
        Call cmbPOnumber_DropDown
        For i = 0 To .ListCount - 1
            If .text = Left(.list(i), Len(.text)) Then
                .TopIndex = i
                Exit Sub
            End If
        Next
    End With
    
''''    Dim text
''''    If KeyAscii = 13 Then
''''        Call cmbPOnumber_Validate(False)
''''        If cmbPOnumber <> "" And cmbPOnumber <> "Error" Then SendKeys ("{tab}")
''''        Exit Sub
''''    End If
''''    With cmbPOnumber
''''        text = .text
''''        For i = 0 To .ListCount - 1
''''            If text Like .list(i) Then
''''                .ListIndex = i
''''                Exit For
''''            End If
''''        Next
''''    End With
    
End Sub


Private Sub cmbPOnumber_LostFocus()
Dim i
    With cmbPOnumber
        For i = 0 To .ListCount - 1
            If .text = Left(.list(i), Len(.text)) Then
                .ListIndex = i
                Exit Sub
            End If
        Next
        .text = ""
        NavBar1.NewEnabled = False
    End With
End Sub

Private Sub cmbPOnumber_Scroll()
Dim imsLock As imsLock.lock
Set imsLock = New imsLock.lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

End Sub

Private Sub cmbPOnumber_Validate(Cancel As Boolean)
'
End Sub

'unload form and free memory

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set Tracklist = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'clear form

Private Sub NavBar1_OnCancelClick()
'    If Not Len(cmbPOnumber) = 0 And Not Len(SScmbMessage.Columns(0).Text) = 0 _
'       And Not Len(SScmbMessage.Columns(1).Text) = 0 Then
'
'            'Call GetOBSList
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



Private Sub NavBar1_OnEditClick()
Dim str As String
'    If Not Len(SScmbMessage.Columns(0).Text) = 0 And Not Len(SScmbMessage.Columns(1).Text) = 0 Then
'
''        MsgBox "Please click add buttom to add new message"
''        Call EnableControls(True)
'
'        NavBar1.EditEnabled = True
'
'        str = Format$(Now(), "mm/dd/yyyy")
'        Call ClaerForm
'        Call GetMessageNumber
'
'        LblMessaDate = str
'
'    Else
'
'        NavBar1.NewEnabled = True
'
'    End If
End Sub

'call function get new messege number and assign values to lable
'set navbar buttom

Private Sub NavBar1_OnNewClick()
Dim str As String
    'If Len(SScmbMessage.Columns(0).Text) = 0 And Len(SScmbMessage.Columns(1).Text) = 0 Then
    
        str = Format$(Now(), "mm/dd/yyyy")

        Call ClaerForm
        Call ChangeMode(mdCreation)
        Call GetMessageNumber
        LblMessaDate = str
        txtOperator = CurrentUser
        NavBar1.SaveEnabled = SaveEnabled
        NavBar1.CancelEnabled = True
        Call EnableControls(True)
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
'        NavBar1.EditEnabled = False
   ' Else
        'MsgBox "The message exist, Please click message number to see"
        
   ' End If
SScboPriority.SetFocus
End Sub

'before save a record check data fields
'and set store procedure parameters

Private Sub NavBar1_OnSaveClick()
Set Tracklist = New imsTrackingMessage

     If Len(SScboSupRecipFax) = 0 And Len(SScboForwarder) = 0 And Len(SScboRecip1) = 0 _
       And Len(SScboRecip2) = 0 And Len(SScboRecip3) = 0 And Len(SScboRecip4) = 0 _
       And Len(SScboRecip5) = 0 Then
       
            'Modified by Juan (9/25/2000) for Multilingual
            msg1 = translator.Trans("M00369") 'J added
            MsgBox IIf(msg1 = "", "You must choose one", msg1) 'J modified
            '---------------------------------------------

        Exit Sub
    End If
    
    Call Tracklist.InsertandUpdateTable(cmbPOnumber, deIms.NameSpace, SScmbMessage, SScboSubject, _
    1, LblMessaDate, SScboSupRecipFax, SScboForwarder, SScboRecip1, SScboRecip2, SScboRecip3, _
    SScboRecip4, SScboRecip5, txtOperator, DTPickNewDelivery, DTPickEstimatedate, DTPickArrival, _
    SScboPriority.Columns("code").text, chkYesorNo, txtRemark, CurrentUser, deIms.cnIms)
    
    EnableControls False
    Call ChangeMode(mdVisualization)
    
    If Len(Trim(cmbPOnumber)) <> 0 And Len(Trim(SScmbMessage)) <> 0 Then
        NavBar1.CancelEnabled = False
        NavBar1.SaveEnabled = False
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
    End If
'    Call ClaerForm
'    cmbPOnumber = ""
End Sub

Private Sub SScboForwarder_Click()
SScboForwarder.SelLength = 0
SScboForwarder.SelStart = 0
End Sub

Private Sub SScboForwarder_DropDown()
'
End Sub

Private Sub SScboForwarder_GotFocus()
Call HighlightBackground(SScboForwarder)
End Sub

Private Sub SScboForwarder_KeyDown(KeyCode As Integer, Shift As Integer)
SScboForwarder.DroppedDown = True
End Sub

Private Sub SScboForwarder_LostFocus()
Call NormalBackground(SScboForwarder)
End Sub

Private Sub SScboPriority_Click()
SScboPriority.SelLength = 0
SScboPriority.SelStart = 0
End Sub

Private Sub SScboPriority_DropDown()
'Call SScboPriority_Click
End Sub

Private Sub SScboPriority_GotFocus()
Call HighlightBackground(SScboPriority)
End Sub

Private Sub SScboPriority_KeyDown(KeyCode As Integer, Shift As Integer)
If Not KeyCode = 13 Then
  SScboPriority.DroppedDown = True
End If
End Sub

Private Sub SScboPriority_KeyPress(KeyAscii As Integer)
''''If Not KeyAscii = 13 Then
'''  SScboPriority.DroppedDown = True
''''End If
End Sub

Private Sub SScboPriority_LostFocus()
Call NormalBackground(SScboPriority)
End Sub

Private Sub SScboRecip1_Click()


SScboRecip1.text = SScboRecip1.Columns(1).text

SScboRecip1.SelStart = 0
SScboRecip1.SelLength = 0

End Sub

Private Sub SScboRecip1_DropDown()
'
End Sub

Private Sub SScboRecip1_GotFocus()
Call HighlightBackground(SScboRecip1)
End Sub

Private Sub SScboRecip1_KeyDown(KeyCode As Integer, Shift As Integer)
SScboRecip1.DroppedDown = True
End Sub

Private Sub SScboRecip1_LostFocus()
Call NormalBackground(SScboRecip1)
End Sub

Private Sub SScboRecip2_Click()

SScboRecip2.text = SScboRecip2.Columns(1).text

SScboRecip2.SelStart = 0
SScboRecip2.SelLength = 0

End Sub

Private Sub SScboRecip2_DropDown()
'
End Sub

Private Sub SScboRecip2_GotFocus()
Call HighlightBackground(SScboRecip2)
End Sub

Private Sub SScboRecip2_KeyDown(KeyCode As Integer, Shift As Integer)
SScboRecip2.DroppedDown = True
End Sub

Private Sub SScboRecip2_LostFocus()
Call NormalBackground(SScboRecip2)
End Sub

Private Sub SScboRecip3_Click()

SScboRecip3.text = SScboRecip3.Columns(1).text

SScboRecip3.SelStart = 0
SScboRecip3.SelLength = 0

End Sub

Private Sub SScboRecip3_DropDown()
'
End Sub

Private Sub SScboRecip3_GotFocus()
Call HighlightBackground(SScboRecip3)
End Sub

Private Sub SScboRecip3_KeyDown(KeyCode As Integer, Shift As Integer)
SScboRecip3.DroppedDown = True
End Sub

Private Sub SScboRecip3_LostFocus()
Call NormalBackground(SScboRecip3)
End Sub

Private Sub SScboRecip4_Click()

SScboRecip4.text = SScboRecip4.Columns(1).text

SScboRecip4.SelStart = 0
SScboRecip4.SelLength = 0
End Sub

Private Sub SScboRecip4_DropDown()
'
End Sub

Private Sub SScboRecip4_GotFocus()
Call HighlightBackground(SScboRecip4)
End Sub

Private Sub SScboRecip4_KeyDown(KeyCode As Integer, Shift As Integer)
SScboRecip4.DroppedDown = True
End Sub

Private Sub SScboRecip4_LostFocus()
Call NormalBackground(SScboRecip4)
End Sub

Private Sub SScboRecip5_Click()

SScboRecip5.text = SScboRecip5.Columns(1).text

SScboRecip5.SelStart = 0
SScboRecip5.SelLength = 0

End Sub

Private Sub SScboRecip5_DropDown()
'
End Sub

Private Sub SScboRecip5_GotFocus()
Call HighlightBackground(SScboRecip5)
End Sub

Private Sub SScboRecip5_KeyDown(KeyCode As Integer, Shift As Integer)
SScboRecip5.DroppedDown = True
End Sub

Private Sub SScboRecip5_LostFocus()
Call NormalBackground(SScboRecip5)
End Sub

Private Sub SScboSubject_Click()
SScboSubject.SelLength = 0
SScboSubject.SelStart = 0
End Sub

Private Sub SScboSubject_DropDown()
'
End Sub

Private Sub SScboSubject_GotFocus()
Call HighlightBackground(SScboSubject)
End Sub

Private Sub SScboSubject_KeyDown(KeyCode As Integer, Shift As Integer)
SScboSubject.DroppedDown = True
End Sub

Private Sub SScboSubject_LostFocus()
Call NormalBackground(SScboSubject)
End Sub

Private Sub SScboSupRecipFax_Click()
SScboSupRecipFax.SelStart = 0
SScboSupRecipFax.SelLength = 0
End Sub

Private Sub SScboSupRecipFax_DropDown()
'
End Sub

Private Sub SScboSupRecipFax_GotFocus()
Call HighlightBackground(SScboSupRecipFax)
End Sub

Private Sub SScboSupRecipFax_KeyDown(KeyCode As Integer, Shift As Integer)
SScboSupRecipFax.DroppedDown = True
End Sub

Private Sub SScboSupRecipFax_LostFocus()
Call NormalBackground(SScboSupRecipFax)
End Sub

' call  function get list and set navbar buttom

Private Sub SScmbMessage_Click()
    If Not Len(SScmbMessage.Columns(0).text) = 0 And Not Len(SScmbMessage.Columns(1).text) = 0 Then
        Call GetOBSList
'        Call EnableControls(True)
'        NavBar1.EditEnabled = False
'        NavBar1.NewEnabled = True
         NavBar1.PrintEnabled = True
         NavBar1.EMailEnabled = True
    Else
        If Not Len(SScmbMessage.Columns(0).text) = 0 Then
            NavBar1.SaveEnabled = SaveEnabled
            NavBar1.CancelEnabled = True
        End If
'        NavBar1.EditEnabled = False
'        NavBar1.NewEnabled = True
    End If
    
End Sub

'call function get  datas for data grids and set navbar buttom

Private Sub Form_Load()
Dim str As String
    
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmTracking")
    '------------------------------------------
    
    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    
    str = Format$(Date, "mm/dd/yyyy")
    DTPickArrival.Value = str
    DTPickEstimatedate.Value = str
    DTPickNewDelivery.Value = str
    
    GetPOnumber
       
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
'   NavBar1.EMailEnabled = True
'   NavBar1.EMailVisible = True
   
   Call ChangeMode(mdVisualization)
'    Call PopuLateFromRecordSet(cmbMessage, Tracklist.GetMessageNumberlist(cmbPOnumber, deIms.Namespace, deIms.cnIms), "ob_ponumb", True)
'    If cmbMessage.ListCount Then cmbMessage.ListIndex = 0
    Caption = Caption + " - " + Tag
   ' cmbPOnumber.SetFocus
   
   NavBar1.NewEnabled = SaveEnabled
   NavBar1.SaveEnabled = SaveEnabled
End Sub

'SQL statement get priority list and populate data grid

Public Sub getPRIORITYlist()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT pri_code, pri_desc"
        .CommandText = .CommandText & " From PRIORITY "
        .CommandText = .CommandText & " WHERE pri_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by pri_code "
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

'SQL statemet get supplier fax and populate data grid

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

'SQL statement to get supplier email and populate data grid

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
        SScboSupRecipFax.AddItem rst!Sup_mail & "" & str & rst!sup_name & "" & str & rst!sup_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'SQL statement get forwarder fax number and populate data grid

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

'SQL statement get forwarder email number and populate data grid

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

'SQL statement get po number  and populate data grid

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
    
'    str = Chr$(1)
'    cmbPOnumber.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo clearup
    
    rst.MoveFirst
    Do While ((Not rst.EOF))
        cmbPOnumber.AddItem rst!PO_PONUMB & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing



End Sub

'set store procedure parameters

Public Function GetMessageNumber() As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GetMessageNumber"
        Set .ActiveConnection = deIms.cnIms
        
        
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cmbPOnumber)
'        .Parameters.Append .CreateParameter("@MESSAGE", adVarChar, adParamInput, 15, SScmbMessage.Columns(0).Text)
        .Parameters.Append .CreateParameter("@STRING", adVarChar, adParamOutput, 15, GetMessageNumber)
        
        .Execute , , adExecuteNoRecords
        
        GetMessageNumber = .Parameters("@STRING").Value & ""
    End With
        
        SScmbMessage = GetMessageNumber
    
    Set cmd = Nothing
   
End Function

'add data to data grid

Public Sub GetSubjectCode()
Dim str As String
    str = SScboSubject.FieldSeparator
    
    SScboSubject.AddItem " Suppliers Follow Up "
    SScboSubject.AddItem " Shipping Information "
    SScboSubject.AddItem " MIscellaneous "
End Sub

'clear form

Public Sub ClaerForm()
'    cmbPOnumber = ""
    SScmbMessage = ""
'    LblMessaDate = ""
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
        .CommandText = .CommandText & " WHERE (OBS.ob_ponumb = '" & cmbPOnumber & "') AND "
        .CommandText = .CommandText & " OBS.ob_npecode  = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND OBS.ob_flag  = 1 "
         Set rst = .Execute
    End With
    
    str = Chr$(1)
    
'    cmdEdit.Enabled = True
    EnableControls True
'    NavBar1.EditEnabled = False
'
'    NavBar1.NewEnabled = True
'    NavBar1.CancelEnabled = True
'    NavBar1.PrintEnabled = True
'    NavBar1.SaveEnabled = True
    
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

'check data combo and show messega

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

'enable or disable text boxse and navbar buttom controls

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
    
'    NavBar1.Enabled = True
    cmbPOnumber.Enabled = True
'    NavBar1.CloseEnabled = True
'    NavBar1.EditEnabled = False
'    NavBar1.NewEnabled = False
'    NavBar1.CancelEnabled = False
'    NavBar1.PrintEnabled = False
'    NavBar1.SaveEnabled = False
'
    
'    cmdExit.Enabled = True
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

'SQL statement to get phone directory list and populate data grid

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
        'SScboRecip1.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        SScboRecip1.AddItem rst!phd_name & "" & str & rst!phd_faxnumb & "" & str & rst!phd_code & ""
        SScboRecip2.AddItem rst!phd_name & "" & str & rst!phd_faxnumb & "" & str & rst!phd_code & ""
        SScboRecip3.AddItem rst!phd_name & "" & str & rst!phd_faxnumb & "" & str & rst!phd_code & ""
        SScboRecip4.AddItem rst!phd_name & "" & str & rst!phd_faxnumb & "" & str & rst!phd_code & ""
        SScboRecip5.AddItem rst!phd_name & "" & str & rst!phd_faxnumb & "" & str & rst!phd_code & ""
        
        
        'SScboRecip2.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        'SScboRecip3.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        'SScboRecip4.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        'SScboRecip5.AddItem rst!phd_faxnumb & "" & str & rst!phd_name & "" & str & rst!phd_code & ""
        rst.MoveNext
    Loop
    
clearup:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

'SQL statement to get phone directory email number
'and populate data grid

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

'get parameters for email function and send email

Private Sub NavBar1_OnEMailClick()
Dim IFile As IMSFile
Dim FileName(1) As String
Dim Recepients() As String
Dim rsr As ADODB.Recordset
Dim rptinfo As RPTIFileInfo
Dim Subject As String
Dim attention As String
    
    Set rsr = GetObsRecipients(deIms.NameSpace, cmbPOnumber, SScmbMessage.Columns(0).text)
        
    FileName(0) = "Report-" & "Tracking Message" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
    FileName(1) = "Report-" & "PO" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
    With MDI_IMS.CrystalReport1
        SetOBsReportParam
        .PrintFileType = crptRTF
        .Destination = crptToFile
        .PrintFileName = "c:\IMSRequests\IMSRequests\OUT\" + FileName(0)
        .Action = 1
        If chkYesorNo.Value = 1 Then
            SetPOReportParam
            .PrintFileType = crptRTF
            .Destination = crptToFile
            .PrintFileName = "c:\IMSRequests\IMSRequests\OUT\" + FileName(1)
            .Action = 1
        End If
    End With

    Set IFile = New IMSFile
    Subject = "Tracking message of PO " + cmbPOnumber
    attention = "Attention Please"
    Recepients = ToArrayFromRecordset(rsr)
    

    Call WriteParameterFiles(Recepients, "", FileName, Subject, attention)

'    If chkYesorNo.Value = 0 Then
'        Call SendEmailAndFax(rsr, "Recipient", "Tracking Message for PO -" & cmbPOnumber, "", FileName(0))
'    Else
'        Call SendEmailAndFaxWithAttachments(rsr, "Recipient", "Tracking Message for PO -" & cmbPOnumber, "", FileName, False)
'    End If
End Sub

'call  function to print crystal report

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    
    SetOBsReportParam
    
    'Modified by Juan (9/25/2000) for Multilingual
    msg1 = translator.Trans("M00370") 'J added
    MDI_IMS.CrystalReport1.WindowTitle = IIf(msg1 = "", "Tracking Message", msg1) 'J modified
    '------------------------------------------
    
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset
   
    If chkYesorNo.Value = 1 Then
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

'get crystal report parameters and application path

Private Sub SetOBsReportParam()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\obs.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("obs.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "mesgnumb;" + SScmbMessage + ";TRUE"
        .ParameterFields(2) = "ponumb;" + cmbPOnumber + ";true"
    End With
End Sub

'get po report parameters to print po report

Private Sub SetPOReportParam()

    If chkYesorNo.Value = 1 Then
        With MDI_IMS.CrystalReport1
            .Reset
            .ReportFileName = FixDir(App.Path) + "CRreports\po.rpt"
            
            'Modified by Juan (8/28/2000) for Multilingual
            Call translator.Translate_Reports("po.rpt") 'J added
            Call translator.Translate_SubReports 'J added
            '---------------------------------------------
            
            .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            .ParameterFields(1) = "ponumb;" + cmbPOnumber + ";true"
        End With
    End If
    
End Sub

Private Sub SScmbMessage_DropDown()
'
End Sub

Private Sub SScmbMessage_GotFocus()
Call HighlightBackground(SScmbMessage)
End Sub

Private Sub SScmbMessage_KeyPress(KeyAscii As Integer)
 SScmbMessage.DroppedDown = True
End Sub

Private Sub SScmbMessage_LostFocus()
Call NormalBackground(SScmbMessage)
End Sub

Private Sub SSOleDBCombo6_KeyPress(KeyAscii As Integer)
SSOleDBCombo6.DroppedDown = True
End Sub

'depend on tab  and set navbar buttom

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
    End If
    
    If Form = mdCreation And SSTab1.Tab = 0 Then
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
    End If
    
    If Form = mdVisualization Then
         NavBar1.CancelEnabled = False
         NavBar1.SaveEnabled = False
    End If
    If cmbPOnumber = "" Then NavBar1.NewEnabled = False 'J added
    If SSTab1.Tab = 1 Then txtRemark.SetFocus
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
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        CheckMessageNumber = rst!rt
    End With
        
 
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckMessageNumber", Err.Description, Err.number, True)
End Function

Private Sub SSTab1_GotFocus()
    If cmbPOnumber = "" Then NavBar1.NewEnabled = False 'J added
End Sub


