VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "IMSMAI~1.OCX"
Begin VB.Form frmWHInitialAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initial Adjustment "
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   10035
   Tag             =   "02050100"
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   180
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   180
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Initial Adjustment"
      TabPicture(0)   =   "WHInitialAdjustment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUser"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDesc(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDesc(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblType"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDesc(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDesc(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDesc(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDesc(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDesc(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ssdcboCompany"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ssdcboWarehouse"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCommodity"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ssdbStockInfo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbo_Transaction"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdStockSearch"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Line Items to be Adjusted"
      TabPicture(1)   =   "WHInitialAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPrimUnit"
      Tab(1).Control(1)=   "lblDesc(23)"
      Tab(1).Control(2)=   "lblCurrency"
      Tab(1).Control(3)=   "lblDesc(26)"
      Tab(1).Control(4)=   "lblDesc(22)"
      Tab(1).Control(5)=   "lblDesc(25)"
      Tab(1).Control(6)=   "lblDesc(24)"
      Tab(1).Control(7)=   "lblCurrencyValu"
      Tab(1).Control(8)=   "lblDesc(7)"
      Tab(1).Control(9)=   "lblDesc(10)"
      Tab(1).Control(10)=   "lblDesc(19)"
      Tab(1).Control(11)=   "lblDesc(18)"
      Tab(1).Control(12)=   "lblDesc(15)"
      Tab(1).Control(13)=   "lblDesc(14)"
      Tab(1).Control(14)=   "lblCommodity"
      Tab(1).Control(15)=   "lblDesc(11)"
      Tab(1).Control(16)=   "lblDesc(28)"
      Tab(1).Control(17)=   "lblSecQnty"
      Tab(1).Control(18)=   "lblSecUnit"
      Tab(1).Control(19)=   "lblDesc(8)"
      Tab(1).Control(20)=   "lblDesc(16)"
      Tab(1).Control(21)=   "lblDesc(17)"
      Tab(1).Control(22)=   "Frame1"
      Tab(1).Control(23)=   "ssdcboCondition"
      Tab(1).Control(24)=   "ssdcboCountry"
      Tab(1).Control(25)=   "ssdcboSubLocation"
      Tab(1).Control(26)=   "ssdcboLogicalWHouse"
      Tab(1).Control(27)=   "txtprimUnit"
      Tab(1).Control(28)=   "txtLeaseComp"
      Tab(1).Control(29)=   "optLease"
      Tab(1).Control(30)=   "optOwn"
      Tab(1).Control(31)=   "cboSerialNumb"
      Tab(1).Control(32)=   "txtDesc"
      Tab(1).Control(33)=   "txtUnitPrice"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Remarks"
      TabPicture(2)   =   "WHInitialAdjustment.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRemarks"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recipients"
      TabPicture(3)   =   "WHInitialAdjustment.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl_Recipients"
      Tab(3).Control(1)=   "ssdbRecepientList"
      Tab(3).Control(2)=   "cmd_Add"
      Tab(3).Control(3)=   "cmd_Remove"
      Tab(3).Control(4)=   "Picture1"
      Tab(3).ControlCount=   5
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   -74880
         ScaleHeight     =   3015
         ScaleWidth      =   9255
         TabIndex        =   52
         Top             =   2280
         Width           =   9255
         Begin ImsMailVB.Imsmail Imsmail1 
            Height          =   3015
            Left            =   480
            TabIndex        =   60
            Top             =   120
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5318
         End
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74595
         TabIndex        =   51
         Top             =   1785
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74595
         TabIndex        =   50
         Top             =   1455
         Width           =   1335
      End
      Begin VB.TextBox txtUnitPrice 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73080
         TabIndex        =   11
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdStockSearch 
         Caption         =   "…"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3730
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1590
         Width           =   195
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   -73080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   4080
         Width           =   6735
      End
      Begin VB.ComboBox cboSerialNumb 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68160
         TabIndex        =   14
         Text            =   "cboSerialNumb"
         Top             =   3660
         Width           =   2415
      End
      Begin VB.OptionButton optOwn 
         Alignment       =   1  'Right Justify
         Caption         =   "Own"
         Height          =   315
         Left            =   -73080
         TabIndex        =   6
         Top             =   1500
         Width           =   1035
      End
      Begin VB.OptionButton optLease 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease"
         Height          =   315
         Left            =   -71700
         TabIndex        =   31
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox txtLeaseComp 
         DataField       =   "ird_leasecomp"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73080
         TabIndex        =   7
         Top             =   1860
         Width           =   2415
      End
      Begin VB.TextBox txtprimUnit 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68640
         TabIndex        =   12
         Top             =   2580
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
         Height          =   4815
         Left            =   -74880
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   420
         Width           =   8535
      End
      Begin VB.ComboBox cbo_Transaction 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2760
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbStockInfo 
         Height          =   3015
         Left            =   240
         TabIndex        =   26
         Top             =   2100
         Width           =   9195
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldSeparator  =   ";"
         Col.Count       =   5
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":0070
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":008C
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         SelectTypeRow   =   1
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         ExtraHeight     =   26
         Columns.Count   =   5
         Columns(0).Width=   1905
         Columns(0).Caption=   "Codty #"
         Columns(0).Name =   "Commodity"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).Case =   2
         Columns(0).FieldLen=   256
         Columns(1).Width=   2011
         Columns(1).Caption=   "Stock Type"
         Columns(1).Name =   "StockType"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1296
         Columns(2).Caption=   "Pool"
         Columns(2).Name =   "Pool"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Style=   2
         Columns(3).Width=   2514
         Columns(3).Caption=   "Category Code"
         Columns(3).Name =   "CategoryCode"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   10742
         Columns(4).Caption=   "Description"
         Columns(4).Name =   "Description"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   16219
         _ExtentY        =   5318
         _StockProps     =   79
         BackColor       =   -2147483643
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboLogicalWHouse 
         Height          =   315
         Left            =   -68160
         TabIndex        =   8
         Top             =   1140
         Width           =   2415
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196617
         DataMode        =   2
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":00A8
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":00C4
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   5
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboSubLocation 
         Height          =   315
         Left            =   -68760
         TabIndex        =   9
         Top             =   1500
         Width           =   3015
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196617
         DataMode        =   2
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":00E0
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":00FC
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   5
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCountry 
         Height          =   315
         Left            =   -73080
         TabIndex        =   5
         Top             =   780
         Width           =   2475
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196617
         DataMode        =   2
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":0118
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":0134
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   5
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
         Height          =   315
         Left            =   -68760
         TabIndex        =   10
         Top             =   1860
         Width           =   3015
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196617
         DataMode        =   2
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":0150
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":016C
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   5
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   -73320
         TabIndex        =   35
         Top             =   3480
         Width           =   2715
         Begin VB.OptionButton optSpecific 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific"
            Height          =   315
            Left            =   1380
            TabIndex        =   36
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optPool 
            Alignment       =   1  'Right Justify
            Caption         =   "Pool"
            Height          =   315
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.TextBox txtCommodity 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   2760
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   1605
         Left            =   -72720
         TabIndex        =   53
         Top             =   480
         Width           =   6810
         _Version        =   196617
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   11562
         Columns(0).Caption=   "Recipients"
         Columns(0).Name =   "Recp"
         Columns(0).DataField=   "Recipients"
         Columns(0).FieldLen=   256
         _ExtentX        =   12012
         _ExtentY        =   2831
         _StockProps     =   79
         Caption         =   "Recipient List"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboWarehouse 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   2760
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196617
         DataMode        =   2
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
         stylesets(0).Picture=   "WHInitialAdjustment.frx":0188
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
         stylesets(1).Picture=   "WHInitialAdjustment.frx":01A4
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4868
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCompany 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   2760
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
         BackColorOdd    =   16771818
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4445
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4868
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblDesc 
         Caption         =   "Lease Company"
         Height          =   315
         Index           =   17
         Left            =   -74880
         TabIndex        =   59
         Top             =   1860
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Own / Lease"
         Height          =   315
         Index           =   16
         Left            =   -74880
         TabIndex        =   58
         Top             =   1500
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Quantities"
         Height          =   195
         Index           =   8
         Left            =   -68640
         TabIndex        =   57
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label lblSecUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67200
         TabIndex        =   56
         Top             =   2940
         Width           =   1335
      End
      Begin VB.Label lblSecQnty 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68640
         TabIndex        =   55
         Top             =   2940
         Width           =   1335
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74595
         TabIndex        =   54
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   25
         Top             =   1560
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "Description"
         Height          =   315
         Index           =   28
         Left            =   -74880
         TabIndex        =   46
         Top             =   4080
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Commodity"
         Height          =   315
         Index           =   11
         Left            =   -74880
         TabIndex        =   27
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label lblCommodity 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   28
         Top             =   420
         Width           =   2475
      End
      Begin VB.Label lblDesc 
         Caption         =   "Country of Origin"
         Height          =   315
         Index           =   14
         Left            =   -74880
         TabIndex        =   29
         Top             =   780
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Logical Warehouse"
         Height          =   315
         Index           =   15
         Left            =   -70260
         TabIndex        =   30
         Top             =   1140
         Width           =   2100
      End
      Begin VB.Label lblDesc 
         Caption         =   "Sub Location"
         Height          =   315
         Index           =   18
         Left            =   -70260
         TabIndex        =   32
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Condition"
         Height          =   195
         Index           =   19
         Left            =   -70260
         TabIndex        =   33
         Top             =   1860
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Caption         =   "Pool / Specific"
         Height          =   315
         Index           =   10
         Left            =   -74880
         TabIndex        =   34
         Top             =   3660
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Serial Number"
         Height          =   315
         Index           =   7
         Left            =   -70260
         TabIndex        =   37
         Top             =   3660
         Width           =   2055
      End
      Begin VB.Label lblCurrencyValu 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.0000"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   43
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Secondary"
         Height          =   195
         Index           =   24
         Left            =   -70260
         TabIndex        =   41
         Top             =   2940
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Primary"
         Height          =   195
         Index           =   25
         Left            =   -70260
         TabIndex        =   39
         Top             =   2580
         Width           =   1590
      End
      Begin VB.Label lblDesc 
         Caption         =   "Currency Value"
         Height          =   315
         Index           =   22
         Left            =   -74880
         TabIndex        =   42
         Top             =   2940
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Unit Price"
         Height          =   315
         Index           =   26
         Left            =   -74880
         TabIndex        =   40
         Top             =   2580
         Width           =   1800
      End
      Begin VB.Label lblCurrency 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "US DOLLARS"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71880
         TabIndex        =   38
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Index           =   23
         Left            =   -67200
         TabIndex        =   44
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label lblPrimUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67200
         TabIndex        =   45
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Label lblDesc 
         Caption         =   "Warehouse"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   49
         Top             =   840
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "Transac #"
         Height          =   315
         Index           =   2
         Left            =   5100
         TabIndex        =   18
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         Caption         =   "Date"
         Height          =   315
         Index           =   5
         Left            =   5100
         TabIndex        =   23
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6720
         TabIndex        =   24
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label lblDesc 
         Caption         =   "Type"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   1200
         Width           =   1600
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INITIAL ADJUSTMENT "
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Label lblDesc 
         Caption         =   "Company"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "User"
         Height          =   315
         Index           =   1
         Left            =   5100
         TabIndex        =   19
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6720
         TabIndex        =   20
         Top             =   840
         Width           =   1545
      End
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   3840
      TabIndex        =   48
      Top             =   5640
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "WHInitialAdjustment.frx":01C0
      CancelVisible   =   0   'False
      PreviousVisible =   0   'False
      NewVisible      =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      EMailVisible    =   -1  'True
      CloseEnabled    =   0   'False
      PrintEnabled    =   0   'False
      NewEnabled      =   0   'False
      SaveEnabled     =   0   'False
      CancelEnabled   =   0   'False
      NextEnabled     =   0   'False
      LastEnabled     =   0   'False
      FirstEnabled    =   0   'False
      PreviousEnabled =   0   'False
      EditEnabled     =   -1  'True
   End
End
Attribute VB_Name = "frmWHInitialAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefStr A-Z
Dim fm As FormMode
Dim CompCode As String
Dim Requery As Boolean
Dim Transnumb As String
Dim InvtIss As imsWhareIssue
Dim WithEvents CommSearch As frm_StockSearch
Attribute CommSearch.VB_VarHelpID = -1
Dim rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim SaveEnabled As Boolean
'set navbar buttom

Private Sub cbo_Transaction_Click()
On Error Resume Next

    Transnumb = cbo_Transaction.text
    
    If Len(cbo_Transaction) <> 0 Then
    
        Imsmail1.Enabled = True
        NavBar1.PrintEnabled = True
        ssdbRecepientList.Enabled = True
        NavBar1.EMailEnabled = ssdbRecepientList.Rows
    Else
        NavBar1.EMailEnabled = False
        NavBar1.PrintEnabled = False
    End If
    
    If Err Then Err.Clear
End Sub

'unlock transaction combo

Private Sub cbo_Transaction_DropDown()
    cbo_Transaction.locked = False
End Sub

'call function get reception number

Private Sub cbo_Transaction_GotFocus()
    AddReceptionNumber
    Call HighlightBackground(cbo_Transaction)
End Sub

'disallow enter data to transaction combo

Private Sub cbo_Transaction_KeyPress(KeyAscii As Integer)
    If NavBar1.NewEnabled = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo_Transaction_LostFocus()
Call NormalBackground(cbo_Transaction)
End Sub

Private Sub cboSerialNumb_Click()
    txtprimUnit = 1
    lblSecQnty = 1
End Sub

Private Sub cboSerialNumb_GotFocus()
Call HighlightBackground(cboSerialNumb)
End Sub

Private Sub cboSerialNumb_LostFocus()
Call NormalBackground(cboSerialNumb)
End Sub

'call function to add current receptient to receptient list

Private Sub cmd_Add_Click()
On Error Resume Next
    Imsmail1.AddCurrentRecipient
End Sub

'delete current receptient from receptient list

Private Sub cmd_Remove_Click()
On Error Resume Next

    rsReceptList.Delete
    rsReceptList.Update
    
    If Err Then Err.Clear
End Sub

'call function get stock numbers

Private Sub cmdStockSearch_Click()

    txtCommodity.SetFocus
    cmdStockSearch.Refresh
    Set CommSearch = New frm_StockSearch
    
    CommSearch.Execute
        
End Sub

'populate stock information to data grid

Private Sub CommSearch_Completed(Cancelled As Boolean, sStockNumber As String)
    
    If Not Cancelled Then
        txtCommodity = sStockNumber
        Call AddStockInfo(GetSpecificStockInfo(sStockNumber, deIms.NameSpace, deIms.cnIms))
    End If
    CommSearch.Hide
    Set CommSearch = Nothing
End Sub

'call function datas and set navbar buttom

Private Sub Form_Load()
Dim np As String
Dim FCompany As String
Dim cn As ADODB.Connection
Dim i As Integer
    SSTab1.TabVisible(3) = False
    'Added by Juan (9/26/2000) for Multilingual
    Call translator.Translate_Forms("frmWHInitialAdjustment")
    '------------------------------------------

    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    NavBar1.SaveEnabled = SaveEnabled
    
    For i = 1 To 2
        SSTab1.TabVisible(i) = SaveEnabled
    Next

    np = deIms.NameSpace
    Set cn = deIms.cnIms
    fm = mdVisualization
    ssdcboWarehouse.RemoveAll
    
    Requery = True
    FCompany = GetCompany(np, "PE", cn)
    CompCode = GetCompanyCode(np, FCompany, cn)
    
    AddCompanies
    AddCondition
    AddCountries
    
    AddSubLocations
    AddLogicalWhareHouse
    EnableControls (False)
    Set InvtIss = New imsWhareIssue
    
    Imsmail1.NameSpace = deIms.NameSpace
    'IMSMail1.Connected = True  'M
    Imsmail1.SetActiveConnection deIms.cnIms   'M
    Imsmail1.Language = Language 'M
    Call DisableButtons(Me, NavBar1)
    NavBar1.CloseEnabled = True
    'Call AddStockType(GetStockType(np, cn))
    
    frmWHInitialAdjustment.Caption = frmWHInitialAdjustment.Caption + " - " + frmWHInitialAdjustment.Tag
    cbo_Transaction.locked = False
    cbo_Transaction.Enabled = True
End Sub

'populate receptient recordset

Public Sub AddReceptionNumber()
Dim rst As ADODB.Recordset

    Set rst = deIms.rsReceptionNumber
    If rst.State And adStateOpen = adStateOpen Then rst.Close
    Call deIms.ReceptionNumber(deIms.NameSpace, CompCode, "IA")
    
    Call PopuLateFromRecordSet(cbo_Transaction, rst, rst.Fields(0).Name, True)
    
    rst.Close
    Set rst = Nothing
End Sub

'SQL statement get sub location recordset and populate data grid

Public Sub AddSubLocations()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select sb_code Code, sb_desc Description from SUBLOCATION"
        .CommandText = .CommandText & " where sb_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by sb_desc "
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboSubLocation.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboSubLocation.AddItem rst!Description & ";" & rst!Code
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'populate warehouse data grid

Private Sub AddWhareHouses(rst As ADODB.Recordset)
On Error Resume Next
'
'
'    Set rst = New ADODB.Recordset
'    With rst
'        .LockType = adLockReadOnly
'        .CursorLocation = adUseServer
'        .CursorType = adOpenForwardOnly
'        .ActiveConnection = deIms.cnIms
'        .Source = "SELECT loc_name, loc_locacode From Location"
'        .Source = .Source & " WHERE (UPPER(loc_name) <> 'OTHER') AND"
'        .Source = .Source & " (UPPER(loc_npecode) = '" & deIms.NameSpace & "')"
'
'        .Open
'    End With
    
    ssdcboWarehouse.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboWarehouse.AddItem rst!loc_name & "" & ";" & rst!loc_locacode & ""
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set rst = Nothing
    If Err Then Err.Clear
End Sub

'Public Sub AddStockType(rst As ADODB.Recordset)
'
'    If rst Is Nothing Then Exit Sub
'    If rst.State And adStateOpen = adStateClosed Then Exit Sub
'
'    If rst.RecordCount = 0 Then GoTo CleanUp
'
'    rst.MoveFirst
'
'    ssdcboStockType.RemoveAll
'
'    Do While ((Not rst.EOF))
'        ssdcboStockType.AddItem (rst!Description & "") & ";" & rst!Code & ""
'        rst.MoveNext
'    Loop
'
'CleanUp:
'
'    rst.Close
'    Set rst = Nothing
'End Sub

'SQL statement get country data and populate data grid

Public Sub AddCountries()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select ctry_code Code, ctry_name Description from COUNTRY"
        .CommandText = .CommandText & " where ctry_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by ctry_name "
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboCountry.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboCountry.AddItem rst!Description & ";" & rst!Code
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'unload form free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    If rs.State And adStateOpen = adStateOpen Then
        rs.CancelBatch
        rs.Close
    End If
    
    Set rs = Nothing
    If Err Then Err.Clear

    'Imsmail1.Connected = False 'M
    If open_forms <= 5 Then ShowNavigator
End Sub

'call function add receptient to send email or fax

Private Sub IMSMail1_OnAddClick(ByVal address As String)
On Error Resume Next

    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        
        rsReceptList.Open
    End If
    
'Modified by Muzammil 08/14/00
'Reason - To Add "INTERNET!" before email.
If (InStr(1, address, "@") > 0) And InStr(1, UCase(address), "INTERNET!") = 0 Then address = "INTERNET!" & UCase(address)
    
    
    If Not IsInList(address, "Recipients", rsReceptList) Then _
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))

    Set ssdbRecepientList.DataSource = rsReceptList
    ssdbRecepientList.Columns(0).DataField = "Recipients"
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'call function send email and fax

Private Sub NavBar1_OnEMailClick()
Dim FileName As String
BeforePrint
    Call WriteRPTIFile(CreateRpti, FileName)
    Call SendEmailAndFax(rsReceptList, "Recipients", "Warehouse Initial Adjustment", "", FileName)

    Set rsReceptList = Nothing
    Set ssdbRecepientList.DataSource = Nothing
End Sub

'call function to print crystal report

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    BeforePrint
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'before save validate data format and get store procedure parameters

Private Sub NavBar1_OnSaveClick()
Screen.MousePointer = 11
Dim retval As Boolean
Dim np As String
Dim ToWH As String
Dim FromWH As String

Dim SecUnit As Double
Dim PrimUnit As Double
Dim StockNumb As String
Dim cn As ADODB.Connection


    cbo_Transaction.ListIndex = CB_ERR
    Screen.MousePointer = 11
    If CheckMasterFields And CheckDetl Then
        Screen.MousePointer = 11
        'doevents
        Call BeginTransaction(deIms.cnIms)
        
        retval = PutReturnData
        
        'doevents
        If retval = False Then GoTo RollBack
        
        Set cn = deIms.cnIms
        np = deIms.NameSpace
        FromWH = rs!ird_ware
        ToWH = ssdcboWarehouse.Columns("Code").text
           
        rs.MoveFirst
        Do While Not rs.EOF
            retval = PutDataInsert
            Screen.MousePointer = 11
            'doevents
            If retval = False Then GoTo RollBack
            
            StockNumb = rs!ird_stcknumb
            PrimUnit = rs!ird_primqty
            SecUnit = IIf(IsNull(rs!ird_secoqty), 0, rs!ird_secoqty)
            
            retval = Update_Sap(np, CompCode, StockNumb, ToWH, PrimUnit, 1, rs!ird_unitpric, rs!ird_newcond, CurrentUser, cn)
            retval = retval And Quantity_In_stock1_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, txtDesc, CurrentUser, cn)
            retval = retval And Quantity_In_stock2_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, CurrentUser, cn)
            retval = retval And Quantity_In_stock3_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, rs!ird_tosubloca, CurrentUser, cn)
            retval = retval And Quantity_In_stock4_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, rs!ird_tosubloca, rs!ird_newcond, CurrentUser, cn)
            
            'doevents
            If rs!ird_ps Then
                retval = retval And Quantity_In_stock5_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, rs!ird_tosubloca, rs!ird_newcond, Transnumb, rs!ird_transerl, rs!ird_ware, "IA", CompCode, FromWH, Transnumb, CompCode, rs!ird_transerl, CurrentUser, cn)
            Else
                 retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, rs!ird_tosubloca, rs!ird_newcond, rs!ird_serl, CurrentUser, cn)
                 retval = retval And Quantity_In_stock7_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!ird_tologiware, rs!ird_tosubloca, rs!ird_newcond, Transnumb, FromWH, rs!ird_transerl, rs!ird_ware, "IA", CompCode, Transnumb, CompCode, rs!ird_transerl, rs!ird_serl, CurrentUser, cn)
            
            End If
            
            'doevents
            If retval = False Then GoTo RollBack
            If retval = False Then GoTo RollBack
            rs.MoveNext
        Loop
        
        If Len(txtRemarks) Then _
            Call InvtReceiptRem_Insert(np, CompCode, ToWH, _
                                       Transnumb, txtRemarks, CurrentUser, deIms.cnIms)
        
        If retval Then Call CommitTransaction(deIms.cnIms)
        'If retval Then Call CommitTransaction(deIms.cnIms)
        
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00018") 'J added
        MsgBox IIf(msg1 = "", "Please note that your transaction # is ", msg1) & Transnumb 'J modified
        '---------------------------------------------
    
        'doevents
        Call cbo_Transaction.AddItem(Transnumb, cbo_Transaction.ListCount)
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, Transnumb)
        
        BeforePrint
        Call SendWareHouseMessage(deIms.NameSpace, "Automatic Distribution", _
                                  lblType, deIms.cnIms, CreateRpti)
        
        Screen.MousePointer = 11
        On Error Resume Next
        lblUser = ""
        lblDate = ""
        txtCommodity = ""
        
        
        ssdbStockInfo.RemoveAll
        ssdcboWarehouse.text = ""
        Call EnableControls(False)
        
        Requery = True
        rs.CancelBatch
        Call ClearFields
        fm = mdVisualization
        rs.Close: Set rs = Nothing
        'doevents: 'doevents: 'doevents
    End If
    
    If Err Then Err.Clear
    Screen.MousePointer = 0
    Exit Sub
    
    
RollBack:
    Call RollbackTransaction(deIms.cnIms)
    Call RollbackTransaction(deIms.cnIms)
    Screen.MousePointer = 0
    
End Sub

Private Sub optLease_GotFocus()
Call HighlightBackground(optLease)
End Sub

Private Sub optLease_LostFocus()
Call NormalBackground(optLease)
End Sub

Private Sub optOwn_GotFocus()
Call HighlightBackground(optOwn)
End Sub

Private Sub optOwn_LostFocus()
Call NormalBackground(optOwn)
End Sub

Private Sub optPool_GotFocus()
Call HighlightBackground(optPool)
End Sub

Private Sub optPool_LostFocus()
Call NormalBackground(optPool)
End Sub

Private Sub optSpecific_GotFocus()
Call HighlightBackground(optSpecific)
End Sub

Private Sub optSpecific_LostFocus()
Call NormalBackground(optSpecific)
End Sub

'set receptient combo size

Private Sub ssdbRecepientList_InitColumnProps()
    
    With ssdbRecepientList
        .Columns.RemoveAll
        Call .Columns.Add(0)
        
        .Columns(0).Width = 6554.835
        .Columns(0).DataField = "Recipients"
    End With
        
    
End Sub

'assign data to text box

Private Sub ssdbStockInfo_Click()
    txtCommodity = ssdbStockInfo.Columns("Commodity").text
End Sub

'assign values to text boxse, lable and data grid

Private Sub ssdbStockInfo_DblClick()
On Error Resume Next
Dim rst As ADODB.Recordset
Dim WareHouse As String, SU, PU

    
    cbo_Transaction.ListIndex = CB_ERR
    Call txtprimUnit_Validate(False)
    If ssdbStockInfo.Rows < 1 Then Exit Sub
    
    If Err Then Err.Clear
    If ssdcboWarehouse.text = "" Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00330") 'J added
        MsgBox IIf(msg1 = "", "Warehouse cannot be empty", msg1) 'J modified
        '---------------------------------------------
        
        Exit Sub
    End If
        
    If Not (Requery) Then If CheckDetl = False Then Exit Sub
    
    fm = mdCreation
    txtprimUnit.Tag = ""
    
    SSTab1.Tab = 1
    
    rs.AddNew
    ClearFields
    AssignDefValues
    
    Call EnableControls(True)
    WareHouse = ssdcboWarehouse.Columns("Code").text
    txtDesc = ssdbStockInfo.Columns("Description").text
    optPool.Enabled = CBool(ssdbStockInfo.Columns(2).text)
    lblCommodity = ssdbStockInfo.Columns("Commodity").text
    
    Call GetStockUnit(deIms.NameSpace, lblCommodity, PU, SU, deIms.cnIms)
    
    SU = LCase$(SU)
    PU = LCase$(PU)
    
    If SU = PU Then
        lblPrimUnit = deIms.UnitDescription(PU)
        lblSecUnit = lblPrimUnit
    Else
        lblPrimUnit = deIms.UnitDescription(PU)
        lblSecUnit = deIms.UnitDescription(SU)
    End If
    
    txtCommodity = lblCommodity
    ssdcboCountry.Value = "USA"
    Call txtprimUnit_Validate(False)
    Call FindInGrid(ssdcboCountry, "USA", True, 1)
    
    optSpecific.Enabled = Not optPool.Enabled
    cboSerialNumb.Enabled = optSpecific.Enabled
    
    'txtprimUnit = IIf(txtprimUnit.Tag > 0, txtprimUnit.Tag, "")
    
    
    Call txtprimUnit_Validate(False)
    txtprimUnit = FormatNumber((txtprimUnit), 4)
    'Modified by Muzammil 08/15/00
   'Reason - Did not work good for records which are not in the first 9(The first set
   'which it displays)
    'Call ssdbStockInfo.RemoveItem(ssdbStockInfo.Row)'M
    
    'Call ssdbStockInfo.RemoveItem(ssdbStockInfo.AddItemRowIndex(ssdbStockInfo.Bookmark)) 'M
End Sub

'get warehouse data and populate data grid

Private Sub ssdcboCompany_Click()
    cbo_Transaction.ListIndex = CB_ERR
    CompCode = ssdcboCompany.Columns(1).text
    
    ssdcboWarehouse = ""
    ssdcboWarehouse.RemoveAll
    Call AddWhareHouses(GetLocation(deIms.NameSpace, "OTHER", CompCode, deIms.cnIms, False))
End Sub

Private Sub ssdcboCompany_GotFocus()
Call HighlightBackground(ssdcboCompany)
End Sub

Private Sub ssdcboCompany_LostFocus()
Call NormalBackground(ssdcboCompany)
End Sub

'get condition data and populate data grid

Private Sub ssdcboCondition_Click()
On Error Resume Next
Dim l As Double
Dim Cond As String
Dim rst As ADODB.Recordset

    Cond = ssdcboCondition.Columns("Code").text
    Call deIms.GetItemPriceFromStockNumber(deIms.NameSpace, CompCode, lblCommodity, _
                ssdcboWarehouse.Columns("Code").text, Cond, "01", l)

        
    If l < 0 Then l = 0
''    txtUnitPrice = l
''    rs!ird_unitpric = l
                
    l = 0
    l = QuantityOnHand(deIms.NameSpace, CompCode, lblCommodity, _
                       ssdcboWarehouse.Columns("Code").text, _
                       ssdcboCondition.Columns("Code").text, deIms.cnIms)
                       
    If l > 0 Then
        rs.Delete
        SSTab1.Tab = 0
        
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00380") 'J added
        msg2 = translator.Trans("M00381") 'J modified
        MsgBox IIf(msg1 = "", "Quantity on Hand is", msg1) + " " & l & IIf(msg2 = "", ". Select option Adjustment Entry.", msg2) 'J modified
        '---------------------------------------------
        
        Exit Sub
    End If
        
    'optSpecific = cboSerialNumb.Enabled
    'optSpecific.Enabled = cboSerialNumb.Enabled
    'optPool.Enabled = cboSerialNumb.Enabled = False

    If Err Then Err.Clear
End Sub

'populate stock information data grid

Public Sub AddStockInfo(rst As ADODB.Recordset)
On Error Resume Next
Dim str As String
Dim sRecord As String

    If rst Is Nothing Then Exit Sub
    If rst.EOF And rst.BOF Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    
    Do While Not rst.EOF
        'str = IIf(rst!stk_poolspec , "Pool", "Specific")
        sRecord = (rst!stk_stcknumb & "") & ";" & (rst!stk_stcktype & "")
        sRecord = sRecord & ";" & (rst!stk_poolspec & "") & ";"
        sRecord = sRecord & (rst!stk_catecode & "") & ";" & (rst!stk_desc & "")
        
        ssdbStockInfo.AddItem sRecord
        
        rst.MoveNext
    Loop
End Sub

'assign values to lable

Public Sub AssignInvt()
    With InvtIss
        lblUser = .User
        lblDate = .TransactionDate
        
        
    End With
End Sub

'validate data format and assign values to store procedure
'parameters

Private Function PutDataInsert() As Boolean

    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert = False

    Call txtprimUnit_Validate(False)
    Set cmd = deIms.Commands("INVTRECEIPTDETL_INSERT")


    'Check for valid data.
    If Not ValidateData() Then
        Exit Function
    End If

    'Set the parameter values for the command to be executed.
    cmd.Parameters("@ird_curr") = "USD"
    cmd.Parameters("ird_currvalu") = 1
    cmd.Parameters("@ird_ponumb") = Null
    cmd.Parameters("@ird_lirtnumb") = Null
    cmd.Parameters("@ird_compcode") = CompCode
    cmd.Parameters("@ird_trannumb") = Transnumb
    cmd.Parameters("@ird_npecode") = deIms.NameSpace
    cmd.Parameters("@ird_ware") = GetPKValue(rs.Bookmark, "ird_ware")
    cmd.Parameters("@ird_transerl") = GetPKValue(rs.Bookmark, "ird_transerl")
    cmd.Parameters("@ird_stcknumb") = GetPKValue(rs.Bookmark, "ird_stcknumb")
    cmd.Parameters("@ird_ps") = GetPKValue(rs.Bookmark, "ird_ps")
    cmd.Parameters("@ird_serl") = GetPKValue(rs.Bookmark, "ird_serl")
    cmd.Parameters("@ird_newcond") = GetPKValue(rs.Bookmark, "ird_newcond")
    cmd.Parameters("@ird_stcktype") = GetPKValue(rs.Bookmark, "ird_stcktype")
    cmd.Parameters("@ird_ctry") = GetPKValue(rs.Bookmark, "ird_ctry")
    cmd.Parameters("@ird_tosubloca") = GetPKValue(rs.Bookmark, "ird_tosubloca")
    cmd.Parameters("@ird_tologiware") = GetPKValue(rs.Bookmark, "ird_tologiware")
    cmd.Parameters("@ird_owle") = GetPKValue(rs.Bookmark, "ird_owle")
    cmd.Parameters("@ird_leasecomp") = GetPKValue(rs.Bookmark, "ird_leasecomp")
    cmd.Parameters("@ird_primqty") = GetPKValue(rs.Bookmark, "ird_primqty")
    cmd.Parameters("@ird_secoqty") = GetPKValue(rs.Bookmark, "ird_secoqty")
    cmd.Parameters("@ird_unitpric") = GetPKValue(rs.Bookmark, "ird_unitpric")
    cmd.Parameters("ird_stckdesc") = GetPKValue(rs.Bookmark, "ird_stckdesc")
    cmd.Parameters("@ird_fromlogiware") = GetPKValue(rs.Bookmark, "ird_fromlogiware")
    cmd.Parameters("@ird_fromsubloca") = GetPKValue(rs.Bookmark, "ird_fromsubloca")
    cmd.Parameters("@ird_origcond") = GetPKValue(rs.Bookmark, "ird_origcond")
    cmd.Parameters("@ird_reprcost") = GetPKValue(rs.Bookmark, "ird_reprcost")
    cmd.Parameters("@ird_newstcknumb") = GetPKValue(rs.Bookmark, "ird_newstcknumb")
    cmd.Parameters("@ird_newdesc") = GetPKValue(rs.Bookmark, "ird_newdesc")
    cmd.Parameters("@User") = CurrentUser
    'Execute the command.
    Call cmd.Execute(Options:=adExecuteNoRecords)

    PutDataInsert = True

    Exit Function

errPutDataInsert:
    MsgBox Err.Description: Err.Clear
End Function

'validate data format

Public Function ValidateData() As Boolean

    Dim i As Long

    ValidateData = False

    'Verify the field is not null.
    If IsNull(rs("ird_compcode")) Then
        MsgBox "The field ' ird_compcode ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_compcode")) Then
        If Len(Trim(rs("ird_compcode"))) = 0 Then
            MsgBox "The field ' ird_compcode ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("ird_npecode")) Then
        MsgBox "The field ' ird_npecode ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_npecode")) Then
        If Len(Trim(rs("ird_npecode"))) = 0 Then
            MsgBox "The field ' ird_npecode ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("ird_ware")) Then
        MsgBox "The field ' ird_ware ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_ware")) Then
        If Len(Trim(rs("ird_ware"))) = 0 Then
            MsgBox "The field ' ird_ware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the integer field contains a valid value.
'    If Not IsNull(rs("ird_trannumb")) Then
'        If Not IsNumeric(rs("ird_trannumb")) _
'            And InStr(rs("ird_trannumb"), ".") = 0 Then
'            MsgBox "The field ' ird_trannumb ' does not contain a valid number."
'        Exit Function
'        End If
'    End If

    'Verify the field is not null.
    If IsNull(rs("ird_transerl")) Then
        MsgBox "The field ' ird_transerl ' cannot be null."
        Exit Function
    End If

    'Verify the integer field contains a valid value.
    If Not IsNull(rs("ird_transerl")) Then
        If Not IsNumeric(rs("ird_transerl")) _
            And InStr(rs("ird_transerl"), ".") = 0 Then
            MsgBox "The field ' ird_transerl ' does not contain a valid number."
        Exit Function
        End If
    End If


    'Verify the field is not null.
    If IsNull(rs("ird_stcknumb")) Then
        MsgBox "The field ' ird_stcknumb ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_stcknumb")) Then
        If Len(Trim(rs("ird_stcknumb"))) = 0 Then
            MsgBox "The field ' ird_stcknumb ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("ird_ps")) Then
        MsgBox "The field ' ird_ps ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_serl")) Then
        If Len(Trim(rs("ird_serl"))) = 0 Then
            MsgBox "The field ' ird_serl ' does not contain valid text."
            Exit Function
        End If
    End If


    'Verify the text field contains text.
    If Not IsNull(rs("ird_newcond")) Then
        If Len(Trim(rs("ird_newcond"))) = 0 Then
            MsgBox "The field ' ird_newcond ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_stcktype")) Then
        If Len(Trim(rs("ird_stcktype"))) = 0 Then
            MsgBox "The field ' ird_stcktype ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_ctry")) Then
        If Len(Trim(rs("ird_ctry"))) = 0 Then
            MsgBox "The field ' ird_ctry ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_tosubloca")) Then
        If Len(Trim(rs("ird_tosubloca"))) = 0 Then
            MsgBox "The field ' ird_tosubloca ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_tologiware")) Then
        If Len(Trim(rs("ird_tologiware"))) = 0 Then
            MsgBox "The field ' ird_tologiware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_leasecomp")) Then
        If Len(Trim(rs("ird_leasecomp"))) = 0 Then
            MsgBox "The field ' ird_leasecomp ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("ird_primqty")) Then
        If Not IsNumeric(rs("ird_primqty")) Then
            MsgBox "The field ' ird_primqty ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("ird_secoqty")) Then
        If Not IsNumeric(rs("ird_secoqty")) Then
            MsgBox "The field ' ird_secoqty ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("ird_unitpric")) Then
        If Not IsNumeric(rs("ird_unitpric")) Then
            MsgBox "The field ' ird_unitpric ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_curr")) Then
        If Len(Trim(rs("ird_curr"))) = 0 Then
            MsgBox "The field ' ird_curr ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("ird_currvalu")) Then
        If Not IsNumeric(rs("ird_currvalu")) Then
            MsgBox "The field ' ird_currvalu ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_stckdesc")) Then
        If Len(Trim(rs("ird_stckdesc"))) = 0 Then
            MsgBox "The field ' ird_stckdesc ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_fromlogiware")) Then
        If Len(Trim(rs("ird_fromlogiware"))) = 0 Then
            MsgBox "The field ' ird_fromlogiware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_fromsubloca")) Then
        If Len(Trim(rs("ird_fromsubloca"))) = 0 Then
            MsgBox "The field ' ird_fromsubloca ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_origcond")) Then
        If Len(Trim(rs("ird_origcond"))) = 0 Then
            MsgBox "The field ' ird_origcond ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("ird_reprcost")) Then
        If Not IsNumeric(rs("ird_reprcost")) Then
            MsgBox "The field ' ird_reprcost ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_newstcknumb")) Then
        If Len(Trim(rs("ird_newstcknumb"))) = 0 Then
            MsgBox "The field ' ird_newstcknumb ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("ird_newdesc")) Then
        If Len(Trim(rs("ird_newdesc"))) = 0 Then
            MsgBox "The field ' ird_newdesc ' does not contain valid text."
            Exit Function
        End If
    End If


    ValidateData = True

End Function

'assign data to string variable

Private Function PutReturnData() As Boolean
Dim np As String
Dim WH As String
Dim cmd As Command
Dim From As String
On Error GoTo errPutReturnData

    PutReturnData = False

    Set cmd = deIms.Commands("InvtIssue_Insert")

    
    np = deIms.NameSpace
    Transnumb = "IA-" & GetTransNumb(np, deIms.cnIms)
    WH = ssdcboWarehouse.Columns("Code").text
    From = ssdcboWarehouse.Columns("Code").text
    PutReturnData = InvtReceipt_Insert(np, "", "IA", CompCode, WH, CurrentUser, deIms.cnIms, , From, Transnumb)
    
    Exit Function

errPutReturnData:
    MsgBox Err.Description: Err.Clear
End Function

'if recordset column are empty then assign to null value

Private Function GetPKValue(vBookMark As Variant, sColName As String) As Variant
    
    If IsEmpty(rs(sColName)) Then
        GetPKValue = Null
    Else
        GetPKValue = rs(sColName)
    End If
End Function

'validate recordset fields

Public Function CheckMasterFields() As Boolean

    CheckMasterFields = False
    
    If Len(Trim$(ssdcboWarehouse.text)) = 0 Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00328") 'J added
        MsgBox IIf(msg1 = "", "Issue to cannot be left empty", msg1) 'J modified
        Exit Function
        '---------------------------------------------
    End If
        
        
    If Len(Trim$(ssdcboWarehouse.text)) = 0 Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00330") 'J added
        MsgBox IIf(msg1 = "", "Warehouse cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        Exit Function
    End If
        
    CheckMasterFields = True
        
End Function

Private Sub ssdcboCondition_GotFocus()
Call HighlightBackground(ssdcboCondition)
End Sub

Private Sub ssdcboCondition_LostFocus()
Call NormalBackground(ssdcboCondition)
End Sub

Private Sub ssdcboCountry_GotFocus()
Call HighlightBackground(ssdcboCountry)
End Sub

Private Sub ssdcboCountry_LostFocus()
Call NormalBackground(ssdcboCountry)
End Sub

Private Sub ssdcboLogicalWHouse_GotFocus()
Call HighlightBackground(ssdcboLogicalWHouse)
End Sub

Private Sub ssdcboLogicalWHouse_LostFocus()
Call NormalBackground(ssdcboLogicalWHouse)
End Sub

Private Sub ssdcboSubLocation_GotFocus()
Call HighlightBackground(ssdcboSubLocation)
End Sub

Private Sub ssdcboSubLocation_LostFocus()
Call NormalBackground(ssdcboSubLocation)
End Sub

'enable stock search command

Private Sub ssdcboWarehouse_Change()
    cbo_Transaction.ListIndex = CB_ERR
    cmdStockSearch.Enabled = Len(Trim$(ssdcboWarehouse.text))
End Sub

'assign values to lable

Private Sub ssdcboWarehouse_Click()
    txtRemarks = ""
    lblUser = CurrentUser
    lblDate = Format$(Date, "MM/DD/YYYY")
    
    ssdbStockInfo.RemoveAll
    cmdStockSearch.Enabled = True
End Sub

Private Sub ssdcboWarehouse_GotFocus()
Call HighlightBackground(ssdcboWarehouse)
End Sub

Private Sub ssdcboWarehouse_LostFocus()
Call NormalBackground(ssdcboWarehouse)
End Sub

'depend on tab status set navbar buttom

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim iEditMode As Integer
Dim cmd As ADODB.Command
Dim Opened As Boolean
Dim blFlag As Boolean


    blFlag = SSTab1.Tab = 1
    
    With NavBar1
        .SaveEnabled = SSTab1.Tab = 0
        .CloseEnabled = SSTab1.Tab = 0
        .PrintEnabled = .SaveEnabled And cbo_Transaction.ListIndex <> CB_ERR
        .EMailEnabled = ((ssdbRecepientList.Rows) And (.PrintEnabled))
    End With
    
'    If SSTab1.Caption = "Line Items to be Adjusted" Then NavBar1.CancelEnabled = False
'    If SSTab1.Caption = "Remarks" Then NavBar1.CancelEnabled = False
'    If SSTab1.Caption = "Recipients" Then NavBar1.CancelEnabled = False
    

    If SSTab1.Tab = 1 Then
    
'        If PreviousTab = 0 And fm = mdCreation Then _
'            If Not (CheckMasterFields) Then SSTab1.Tab = 0
            
        If Requery Then
        
            If fm <> mdCreation Then Exit Sub
            Set cmd = deIms.Commands("Get_InvtReceiptDetl")
            iEditMode = IIf(IsNumeric(cbo_Transaction), cbo_Transaction, 0)
            
            Opened = deIms.rsGet_InvtReceiptDetl.State And adStateOpen = adStateOpen
            
            If Opened Then
                deIms.rsGet_InvtReceiptDetl.CancelUpdate
                deIms.rsGet_InvtReceiptDetl.CancelBatch
                deIms.rsGet_InvtReceiptDetl.Close
            End If
            
            Call deIms.Get_InvtReceiptDetl(deIms.NameSpace, Transnumb)


            Set rs = deIms.rsGet_InvtReceiptDetl
            
            Requery = False
       End If
    End If

End Sub


Private Sub AddFromSublocation(rst As ADODB.Recordset)
'    If rst Is Nothing Then Exit Sub
'    If rst.State And adStateOpen = adStateClosed Then Exit Sub
'    If rst.RecordCount = 0 Then Exit Sub
'    ssdcboSubLocation.RemoveAll
'
'    rst.MoveFirst
'    Do While Not rst.EOF
'        ssdcboSubLocation.Text = rst!Description
'        ssdcboSubLocation.AddItem rst!Description & "" & ";" & rst!Code & ""
'        rst.MoveNext
'    Loop
'
'    rst.Close
'    Set rst = Nothing

End Sub

'SQL statement get logical warehouse and populate data grid

Public Sub AddLogicalWhareHouse()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select lw_code Code, lw_desc Description from LOGWAR"
        .CommandText = .CommandText & " where lw_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by lw_desc "
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboLogicalWHouse.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboLogicalWHouse.AddItem rst!Description & ";" & rst!Code
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'SQL statement get condition recordset and populate data grid

Public Sub AddCondition()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select cond_condcode Code, cond_desc Description from CONDITION"
        .CommandText = .CommandText & " where cond_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by cond_condcode"
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboCondition.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboCondition.AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'clear form

Private Sub ClearFields()
    
    optOwn = True
    optPool = True
    
    txtDesc = ""
    lblSecQnty = ""
    txtprimUnit = ""
    txtLeaseComp = ""
    txtUnitPrice = ""
    lblCommodity = ""
    cboSerialNumb = ""
    txtprimUnit.Tag = ""
    ssdcboCountry.text = ""
    
    cboSerialNumb.ListIndex = CB_ERR
    
    ssdcboCondition.text = ""
    ssdcboSubLocation.text = ""
    ssdcboLogicalWHouse.text = ""
    
End Sub

'set controls

Private Sub EnableControls(Value As Boolean)
    optOwn.Enabled = Value
    optPool.Enabled = Value
    optLease.Enabled = Value
    optSpecific.Enabled = Value
    
    
    
    txtDesc.Enabled = False
    lblSecQnty.Enabled = True
    txtprimUnit.Enabled = Value
    txtUnitPrice.Enabled = Value
    ssdcboCountry.Enabled = Value
    
    
    ssdcboCondition.Enabled = False
    ssdcboSubLocation.Enabled = Value
    ssdcboLogicalWHouse.Enabled = Value
    
    ssdcboCondition.Enabled = Value
    ssdcboSubLocation.Enabled = Value
    ssdcboLogicalWHouse.Enabled = Value
End Sub

'validate data format

Private Function CheckDetl() As Boolean
Dim l As Long

    
    If rs Is Nothing Then Exit Function
    If rs.State And adStateOpen = adStateClosed Then Exit Function
     
    l = SSTab1.Tab
    SSTab1.Tab = 1
    
    If Len(Trim$(ssdcboCountry.text)) = 0 Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboCountry.SetFocus: Exit Function
    Else
        rs!ird_ctry = RTrim$(ssdcboCountry.Columns("Code").text)
    End If
    
        
'    If Len(Trim$(ssdcboStockType.Text)) = 0 Then
'        MsgBox "Stock Type cannot be left empty":
'        ssdcboStockType.SetFocus: Exit Function
'
'    Else
'        rs!ird_stcktype = RTrim$(ssdcboStockType.Columns("Code").Text)
'
'    End If
    
    If Len(Trim$(ssdcboCondition.text)) = 0 Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00378") 'J added
        MsgBox IIf(msg1 = "", "condition cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        ssdcboCondition.SetFocus: Exit Function

    Else
        rs!ird_newcond = RTrim$(ssdcboCondition.Columns("Code").text)

    End If

        '// To Sub=location
    If Len(Trim$(ssdcboSubLocation.text)) = 0 Then
    
        'Modified by Juan (9/26/2000) for Multilingual
        msg1 = translator.Trans("M00374") 'J added
        MsgBox IIf(msg1 = "", "Sub-Location cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        ssdcboSubLocation.SetFocus: Exit Function

    Else
        rs!ird_tosubloca = RTrim$(ssdcboSubLocation.Columns("Code").text)
    End If


    '// To Logical Warehouse
    If Len(Trim$(ssdcboLogicalWHouse.text)) = 0 Then
        MsgBox "Logical Warehouse cannot be left empty":
        ssdcboLogicalWHouse.SetFocus: Exit Function

    Else
        rs!ird_tologiware = RTrim$(ssdcboLogicalWHouse.Columns("Code").text)
    End If

    If Len(txtprimUnit) > 0 Then

        If IsNumeric(txtprimUnit) Then
            rs!ird_primqty = CDbl(txtprimUnit)
        Else
            MsgBox "Primary unit is not a valid number": txtprimUnit.SetFocus: Exit Function
        End If

    Else
        MsgBox "Primary unit cannot be left empty": txtprimUnit.SetFocus: Exit Function
    End If

    If optSpecific Then

        rs!ird_ps = 0
        If Len(Trim$(cboSerialNumb)) = 0 Then
            MsgBox "Serial number cannot be left empty":
            cboSerialNumb.SetFocus: Exit Function

        Else
            rs!ird_serl = cboSerialNumb

        End If

    Else
        rs!ird_ps = 1
        rs!ird_serl = Null
    End If


    If optLease Then

        rs!ird_owle = 0

        If Len(Trim$(txtLeaseComp)) = 0 Then
            MsgBox "Lease Company cannot be left empty":
            txtLeaseComp.SetFocus: Exit Function

        Else
            rs!ird_leasecomp = Trim$(txtLeaseComp)

        End If

    Else
         rs!ird_owle = 1
         rs!ird_leasecomp = Null
    End If


    If Len(Trim$(lblSecQnty)) Then

        If Not IsNumeric(lblSecQnty) Then
            MsgBox "Secondary Quantity does not have a valid number"
             Exit Function

        Else
            rs!ird_secoqty = CDbl(lblSecQnty)
        End If

    Else
        rs!ird_secoqty = Null
    End If

    If Len(Trim$(txtDesc)) Then rs!ird_stckdesc = Trim$(txtDesc)
    
    If Len(txtUnitPrice) Then
        
        If IsNumeric(txtUnitPrice) Then
            
            If CDbl(txtUnitPrice) < 1 Then
                MsgBox "Unit Price has to be greater than 0"
                 Exit Function
            End If
                
            rs!ird_unitpric = CDbl(txtUnitPrice)
        Else
            MsgBox "Unit Price has an invalid value"
             Exit Function
        End If
        
    Else
        MsgBox "Unit Price cannot be left empty"
         Exit Function
    End If
    If Len(lblCommodity) Then rs!ird_stcknumb = Trim$(lblCommodity)

    SSTab1.Tab = l
    CheckDetl = True

    If Err Then Err.Clear
End Function

Private Sub txtCommodity_GotFocus()
Call HighlightBackground(txtCommodity)
End Sub

Private Sub txtCommodity_LostFocus()
Call NormalBackground(txtCommodity)
End Sub

Private Sub txtLeaseComp_GotFocus()
Call HighlightBackground(txtLeaseComp)
End Sub

Private Sub txtLeaseComp_LostFocus()
Call NormalBackground(txtLeaseComp)
End Sub

'validate primary unit and format to four decimal digit

Private Sub txtprimUnit_Change()
On Error Resume Next
Dim db As Double

    If Len(txtprimUnit) > 0 Then _
        If Not IsNumeric(txtprimUnit) Then MsgBox "Invalid Value": txtprimUnit.SetFocus: Exit Sub
        
    If Len(txtprimUnit) > 0 Then
    
        If IsNumeric(txtprimUnit) Then
        
'            db = FormatNumber((txtprimUnit), 4)
            db = txtprimUnit
            If db < 1 Then MsgBox "Invalid value"
            
            If Len(Trim$(txtprimUnit.Tag)) > 0 Then
            
                If FormatNumber((txtprimUnit.Tag), 4) < db Then
                    txtprimUnit = ""
                    MsgBox "Value is too large":
                    txtprimUnit.SetFocus:
                    'txtprimUnit = FormatNumber$(txtprimUnit.Tag, 4): Exit Sub
'                    txtprimUnit.SetFocus: Exit Sub
                  txtprimUnit = txtprimUnit.Tag: Exit Sub
                End If
                
            End If
            
            rs!ird_primqty = db
        End If
        
        
    Else
        rs!ird_primqty = Null
    End If
     
    
    Call txtprimUnit_Validate(True)
    If Err Then Err.Clear
End Sub

'set value to recordset

Private Sub optLease_Click()
On Error Resume Next

    If rs!ird_owle <> 0 Then _
       rs!ird_owle = 0
    
    rs!ird_owle = 0
    
    txtLeaseComp.Enabled = True
    txtLeaseComp.SetFocus
    
    If Err Then Err.Clear
End Sub

'asisgn value to recordset

Private Sub optOwn_Click()
On Error Resume Next

    If rs!ird_owle <> 1 Then _
       rs!ird_owle = 1
    
    rs!ird_owle = 1
    txtLeaseComp.Enabled = False
    
    If Err Then Err.Clear
End Sub

'set value to recordset

Private Sub optPool_Click()
On Error Resume Next

    If rs!ird_ps <> 1 Then _
        rs!ird_ps = 1
    
    rs!ird_ps = 1
    
    txtprimUnit.Enabled = True
    cboSerialNumb.Enabled = False
    
    Err.Clear
End Sub

'assign value to recordset

Private Sub optSpecific_Click()
On Error Resume Next
    If rs!ird_ps <> 0 Then _
        rs!ird_ps = 0
        
    rs!ird_ps = 0
    cboSerialNumb.Enabled = True
    cboSerialNumb.SetFocus
    
    txtprimUnit.text = 1
    txtprimUnit.Enabled = False
    Err.Clear
End Sub

Private Sub txtprimUnit_GotFocus()
Call HighlightBackground(txtprimUnit)
End Sub

Private Sub txtprimUnit_LostFocus()
txtprimUnit = FormatNumber((txtprimUnit), 4)
Call NormalBackground(txtprimUnit)
End Sub

'validate primary unit and format to four decimal digit

Private Sub txtprimUnit_Validate(Cancel As Boolean)
On Error Resume Next
Dim CompFactor As Double

    If Len(Trim$(txtprimUnit)) = 0 Then Exit Sub

    If lblPrimUnit = lblSecUnit Then
        lblSecQnty = FormatNumber(txtprimUnit, 4)
    Else

        CompFactor = ImsDataX.ComputingFactor(deIms.NameSpace, lblCommodity, deIms.cnIms)

        If CompFactor = 0 Then
            lblSecQnty = FormatNumber(txtprimUnit, 4)
        Else
            lblSecQnty = FormatNumber(txtprimUnit * 10000 / CompFactor, 4)
        End If
    End If
    
    rs!iid_secoqty = lblSecQnty
End Sub

'check secorndary quantity

Private Sub lblSecQnty_Change()
    If Len(lblSecQnty) Then If Not IsNumeric(lblSecQnty) Then MsgBox "Invalid Value"
End Sub

'assign values to recordset

Private Sub AssignDefValues()


    cboSerialNumb.Clear
    If rs Is Nothing Then Exit Sub
    If rs.State And adStateOpen = adStateClosed Then Exit Sub
    
    optOwn = True
    optPool = True
    
    rs!ird_ps = 1
    rs!ird_owle = 1
    rs!ird_curr = "USD"
    rs!ird_currvalu = 1
    rs!ird_compcode = CompCode
    rs!ird_transerl = GetNextSerial
    rs!ird_npecode = deIms.NameSpace
    rs!ird_ware = RTrim$(ssdcboWarehouse.Columns("Code").text)
    rs!ird_stcknumb = Trim$(ssdbStockInfo.Columns("Commodity").text)
    rs!ird_stckdesc = Trim$(ssdbStockInfo.Columns("Description").text)
End Sub

'SQL statement tomake new serail number

Public Function GetNextSerial() As Long
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
    
    If rs Is Nothing Then Exit Function
    If rs.State And adStateOpen = adStateClosed Then Exit Function
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "Select count(*) +  1 serl from INVTRECEIPTDETL where "
        .CommandText = .CommandText & "ird_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND ird_trannumb = '" & cbo_Transaction & "'"
        
        Set rst = .Execute
        GetNextSerial = rst!Serl
        GetNextSerial = IIf(rs.RecordCount > GetNextSerial, rs.RecordCount, GetNextSerial)
        
        
    End With
End Function

'validate unit price

Private Sub txtUnitPrice_Change()
    If Len(txtUnitPrice) > 0 Then
    
        If Len(txtUnitPrice) = 1 And txtUnitPrice = "-" Then Exit Sub
        
        If Not IsNumeric(txtUnitPrice) Then
            MsgBox "Invalid value"
        End If
        
        
    End If
    
   
    
End Sub

'get crystal report parameters

Public Sub BeforePrint()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = ReportPath & "wareAEIA.rpt"
        .ParameterFields(0) = "transnumb;" & cbo_Transaction & ";TRUE"
        .ParameterFields(1) = "namespace;" & deIms.NameSpace & ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("L00518") 'J added
        .WindowTitle = IIf(msg1 = "", "Initial Adjustment", msg1) 'J modified
        Call translator.Translate_Reports("wareAEIA.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
    End With
End Sub

'get company recordset and populate data grid

Private Sub AddCompanies()
On Error Resume Next
Dim rs As ADODB.Recordset

    If deIms.rsCOMPANY.State Then
        Set rs = deIms.rsCOMPANY.Clone
    Else
        deIms.Company (deIms.NameSpace)
        
        Set rs = deIms.rsCOMPANY.Clone
        deIms.rsCOMPANY.Close
    End If
    
    rs.Filter = "com_actvflag <> 0"
    ssdcboCompany.FieldSeparator = Chr$(1)
    
    rs.MoveFirst
    If rs.RecordCount = 0 Then Exit Sub
    
    Do Until rs.EOF
        ssdcboCompany.AddItem rs("com_name") & Chr$(1) & rs("com_compcode")
        
        rs.MoveNext
    Loop
    
End Sub

Private Sub txtUnitPrice_GotFocus()
Call HighlightBackground(txtUnitPrice)
End Sub

Private Sub txtUnitPrice_LostFocus()
Call NormalBackground(txtUnitPrice)
End Sub

'format unit price to four decimal digit

Private Sub txtUnitPrice_Validate(Cancel As Boolean)
    If Len(txtUnitPrice) <> 0 Then
        txtUnitPrice = FormatNumber((txtUnitPrice), 4)
    Else
        MsgBox "Unit Price can not left empty"
    End If
    
End Sub

' get report parrameters

Private Function CreateRpti() As RPTIFileInfo

    With CreateRpti
        ReDim .Parameters(1)
        .ReportFileName = ReportPath & "wareAEIA.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("wareAEIA.rpt") 'J added
        '---------------------------------------------
        
        .Parameters(0) = "transnumb=" & cbo_Transaction
        .Parameters(1) = "namespace=" & deIms.NameSpace
    End With

End Function


