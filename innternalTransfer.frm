VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmIntTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internal Transfer"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9705
   Tag             =   "02040700"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   3240
      TabIndex        =   61
      Top             =   6120
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "innternalTransfer.frx":0000
      NewVisible      =   0   'False
      EMailVisible    =   -1  'True
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      NextEnabled     =   0   'False
      LastEnabled     =   0   'False
      FirstEnabled    =   0   'False
      PreviousEnabled =   0   'False
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   180
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   10292
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Transfer"
      TabPicture(0)   =   "innternalTransfer.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDesc(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUser"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDesc(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDesc(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblType"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDesc(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDesc(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDesc(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ssdcboCompany"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ssdcboWarehouse"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ssdbStockInfo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbo_Transaction"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Line Items to be transfered"
      TabPicture(1)   =   "innternalTransfer.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPrimUnit"
      Tab(1).Control(1)=   "lblDesc(23)"
      Tab(1).Control(2)=   "lblDesc(26)"
      Tab(1).Control(3)=   "lblDesc(22)"
      Tab(1).Control(4)=   "lblDesc(28)"
      Tab(1).Control(5)=   "lblDesc(25)"
      Tab(1).Control(6)=   "lblDesc(24)"
      Tab(1).Control(7)=   "lblDesc(17)"
      Tab(1).Control(8)=   "lblDesc(16)"
      Tab(1).Control(9)=   "lblDesc(14)"
      Tab(1).Control(10)=   "lblDesc(11)"
      Tab(1).Control(11)=   "lblUnitprice"
      Tab(1).Control(12)=   "lblCurrency"
      Tab(1).Control(13)=   "lblCurrencyValu"
      Tab(1).Control(14)=   "lblCommodity"
      Tab(1).Control(15)=   "lblDesc(20)"
      Tab(1).Control(16)=   "lblDesc(21)"
      Tab(1).Control(17)=   "lblDesc(6)"
      Tab(1).Control(18)=   "lblSecUnit"
      Tab(1).Control(19)=   "lblSecQnty"
      Tab(1).Control(20)=   "Label1"
      Tab(1).Control(21)=   "Label3"
      Tab(1).Control(22)=   "lblRecCount"
      Tab(1).Control(23)=   "lblCurrRec"
      Tab(1).Control(24)=   "Frame1"
      Tab(1).Control(25)=   "Frame2(0)"
      Tab(1).Control(26)=   "ssdcboCountry"
      Tab(1).Control(27)=   "txtDesc"
      Tab(1).Control(28)=   "txtprimUnit"
      Tab(1).Control(29)=   "optLease"
      Tab(1).Control(30)=   "optOwn"
      Tab(1).Control(31)=   "txtLeaseComp"
      Tab(1).Control(32)=   "Frame2(1)"
      Tab(1).Control(33)=   "cboSerialNumb"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Remarks"
      TabPicture(2)   =   "innternalTransfer.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRemarks"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recipients"
      TabPicture(3)   =   "innternalTransfer.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl_Recipients"
      Tab(3).Control(1)=   "ssdbRecepientList"
      Tab(3).Control(2)=   "Picture1"
      Tab(3).Control(3)=   "cmd_Remove"
      Tab(3).Control(4)=   "cmd_Add"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   540
         TabIndex        =   66
         Top             =   1740
         Width           =   1095
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74655
         TabIndex        =   58
         Top             =   1455
         Width           =   1095
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74655
         TabIndex        =   57
         Top             =   1785
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   -74940
         ScaleHeight     =   3015
         ScaleWidth      =   8535
         TabIndex        =   56
         Top             =   2280
         Width           =   8535
         Begin ImsMailVB.Imsmail Imsmail1 
            Height          =   3255
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5741
         End
      End
      Begin VB.TextBox txtRemarks 
         Height          =   5295
         Left            =   -74880
         MaxLength       =   7000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   420
         Width           =   8895
      End
      Begin VB.ComboBox cboSerialNumb 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69000
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "To"
         Height          =   1275
         Index           =   1
         Left            =   -70440
         TabIndex        =   36
         Top             =   2640
         Width           =   4395
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboSubLocation 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   540
            Width           =   2895
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
            BorderStyle     =   0
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
            stylesets(0).Picture=   "innternalTransfer.frx":008C
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
            stylesets(1).Picture=   "innternalTransfer.frx":00A8
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
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboLogicalWHouse 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   11
            Top             =   180
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
            stylesets(0).Picture=   "innternalTransfer.frx":00C4
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
            stylesets(1).Picture=   "innternalTransfer.frx":00E0
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
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   900
            Width           =   2895
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
            stylesets(0).Picture=   "innternalTransfer.frx":00FC
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
            stylesets(1).Picture=   "innternalTransfer.frx":0118
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
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin VB.Label lblDesc 
            Caption         =   "Sub Location"
            Height          =   315
            Index           =   12
            Left            =   60
            TabIndex        =   34
            Top             =   540
            Width           =   1380
         End
         Begin VB.Label lblDesc 
            Caption         =   "Logical Warehouse"
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   33
            Top             =   180
            Width           =   1860
         End
         Begin VB.Label lblDesc 
            Caption         =   "Condition"
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   35
            Top             =   900
            Width           =   1380
         End
      End
      Begin VB.TextBox txtLeaseComp 
         DataField       =   "ird_leasecomp"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73080
         TabIndex        =   6
         Top             =   1620
         Width           =   2475
      End
      Begin VB.OptionButton optOwn 
         Alignment       =   1  'Right Justify
         Caption         =   "Own"
         Height          =   315
         Left            =   -73080
         TabIndex        =   4
         Top             =   1260
         Width           =   1155
      End
      Begin VB.OptionButton optLease 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease"
         Height          =   315
         Left            =   -71760
         TabIndex        =   5
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtprimUnit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68640
         TabIndex        =   7
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox cbo_Transaction 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   480
         Width           =   2640
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbStockInfo 
         Height          =   3615
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   8655
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
         FieldSeparator  =   ";"
         Col.Count       =   4
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
         stylesets(0).Picture=   "innternalTransfer.frx":0134
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
         stylesets(1).Picture=   "innternalTransfer.frx":0150
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
         Columns.Count   =   4
         Columns(0).Width=   1905
         Columns(0).Caption=   "Codty #"
         Columns(0).Name =   "Commodity"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).Case =   2
         Columns(0).FieldLen=   256
         Columns(0).HeadStyleSet=   "ColHeader"
         Columns(0).StyleSet=   "RowFont"
         Columns(1).Width=   10345
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).HeadStyleSet=   "ColHeader"
         Columns(1).StyleSet=   "RowFont"
         Columns(2).Width=   1984
         Columns(2).Caption=   "PU Qty"
         Columns(2).Name =   "ReqQnty"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).HeadStyleSet=   "ColHeader"
         Columns(2).StyleSet=   "RowFont"
         Columns(3).Width=   5292
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "Price"
         Columns(3).Name =   "Price"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         _ExtentX        =   15266
         _ExtentY        =   6376
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboWarehouse 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Tag             =   "1"
         Top             =   840
         Width           =   2745
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
         stylesets(0).Picture=   "innternalTransfer.frx":016C
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
         stylesets(1).Picture=   "innternalTransfer.frx":0188
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
         _ExtentX        =   4842
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -73320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   4320
         Width           =   6735
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCountry 
         Height          =   315
         Left            =   -73080
         TabIndex        =   3
         Top             =   900
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
         stylesets(0).Picture=   "innternalTransfer.frx":01A4
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
         stylesets(1).Picture=   "innternalTransfer.frx":01C0
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
         Enabled         =   0   'False
      End
      Begin VB.Frame Frame2 
         Caption         =   "From"
         Height          =   1275
         Index           =   0
         Left            =   -74880
         TabIndex        =   29
         Top             =   2640
         Width           =   4395
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboSubLocation 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   9
            Top             =   540
            Width           =   2895
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
            BorderStyle     =   0
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
            stylesets(0).Picture=   "innternalTransfer.frx":01DC
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
            stylesets(1).Picture=   "innternalTransfer.frx":01F8
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
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboLogicalWHouse 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   8
            Top             =   180
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
            stylesets(0).Picture=   "innternalTransfer.frx":0214
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
            stylesets(1).Picture=   "innternalTransfer.frx":0230
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
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   10
            Top             =   900
            Width           =   2895
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
            stylesets(0).Picture=   "innternalTransfer.frx":024C
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
            stylesets(1).Picture=   "innternalTransfer.frx":0268
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
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin VB.Label lblDesc 
            Caption         =   "Condition"
            Height          =   315
            Index           =   19
            Left            =   60
            TabIndex        =   32
            Top             =   900
            Width           =   1380
         End
         Begin VB.Label lblDesc 
            Caption         =   "Logical Warehouse"
            Height          =   315
            Index           =   15
            Left            =   60
            TabIndex        =   30
            Top             =   180
            Width           =   1860
         End
         Begin VB.Label lblDesc 
            Caption         =   "Sub Location"
            Height          =   315
            Index           =   18
            Left            =   60
            TabIndex        =   31
            Top             =   540
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   -73320
         TabIndex        =   38
         Top             =   3840
         Width           =   2475
         Begin VB.OptionButton optSpecific 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optPool 
            Alignment       =   1  'Right Justify
            Caption         =   "Pool"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1155
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   1605
         Left            =   -73275
         TabIndex        =   59
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCompany 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Tag             =   "0"
         Top             =   480
         Width           =   2760
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
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
      Begin VB.Label Label2 
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   72
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Search Field"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   71
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblCurrRec 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68040
         TabIndex        =   70
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblRecCount 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66840
         TabIndex        =   69
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "of"
         Height          =   255
         Left            =   -67200
         TabIndex        =   68
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Viewing Record"
         Height          =   255
         Left            =   -69840
         TabIndex        =   67
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblSecQnty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68640
         TabIndex        =   64
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label lblSecUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67320
         TabIndex        =   63
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Secondary"
         Height          =   195
         Index           =   6
         Left            =   -70200
         TabIndex        =   62
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74655
         TabIndex        =   60
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblDesc 
         Caption         =   "Serial Number"
         Height          =   315
         Index           =   21
         Left            =   -70560
         TabIndex        =   39
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         Caption         =   "Pool / Specific"
         Height          =   315
         Index           =   20
         Left            =   -74880
         TabIndex        =   37
         Top             =   4020
         Width           =   1620
      End
      Begin VB.Label lblCommodity 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   23
         Top             =   540
         Width           =   2475
      End
      Begin VB.Label lblCurrencyValu 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.00"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   28
         Top             =   2295
         Width           =   1155
      End
      Begin VB.Label lblCurrency 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U.S DOLLAR"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71880
         TabIndex        =   25
         Top             =   1965
         Width           =   1155
      End
      Begin VB.Label lblUnitprice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   26
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Caption         =   "Commodity"
         Height          =   315
         Index           =   11
         Left            =   -74880
         TabIndex        =   55
         Top             =   540
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Country of Origin"
         Height          =   315
         Index           =   14
         Left            =   -74880
         TabIndex        =   54
         Top             =   900
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Own / Lease"
         Height          =   315
         Index           =   16
         Left            =   -74880
         TabIndex        =   24
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Lease Company"
         Height          =   315
         Index           =   17
         Left            =   -74880
         TabIndex        =   53
         Top             =   1620
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Quantities"
         Height          =   195
         Index           =   24
         Left            =   -68640
         TabIndex        =   52
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Index           =   25
         Left            =   -67320
         TabIndex        =   51
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblDesc 
         Caption         =   "Description"
         Height          =   315
         Index           =   28
         Left            =   -74880
         TabIndex        =   50
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         Caption         =   "Currency Value"
         Height          =   315
         Index           =   22
         Left            =   -74880
         TabIndex        =   49
         Top             =   2290
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Unit Price"
         Height          =   315
         Index           =   26
         Left            =   -74880
         TabIndex        =   48
         Top             =   1950
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Primary"
         Height          =   195
         Index           =   23
         Left            =   -70200
         TabIndex        =   47
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblPrimUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67320
         TabIndex        =   27
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblDesc 
         Caption         =   "Warehouse"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   840
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "Transac #"
         Height          =   315
         Index           =   2
         Left            =   4620
         TabIndex        =   44
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         Caption         =   "Date"
         Height          =   315
         Index           =   5
         Left            =   4620
         TabIndex        =   43
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         TabIndex        =   21
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERNAL TRANSFER"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Label lblDesc 
         Caption         =   "Company"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "User"
         Height          =   315
         Index           =   1
         Left            =   4620
         TabIndex        =   40
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         TabIndex        =   19
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         Caption         =   "Type"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   1600
      End
   End
End
Attribute VB_Name = "frmIntTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim fm As FormMode
Dim CompCode As String
Dim Requery, backpass As Boolean
Dim Transnumb As String
Dim InvtIss As imsWhareIssue
Dim rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim SaveEnabled As Boolean
Dim beginning As Boolean
Dim justDBLCLICK As Boolean
Sub refreshQTY()
Dim qty As New ADODB.Recordset
Dim Sql, LW, SL, SC As String
    If justDBLCLICK Then
        justDBLCLICK = False
    Else
        Screen.MousePointer = 11
        justDBLCLICK = False
        LW = ssdcboLogicalWHouse(0).Columns("Code").Text
        SL = ssdcboSubLocation(0).Columns("Code").Text
        SC = ssdcboCondition(0).Columns("Code").Text
        
        If Not (LW = "" Or SL = "" Or SC = "") Then
            Sql = "SELECT sum(qs4_primqty) AS PrimaryQTY FROM QTYST4 WHERE " _
                & "qs4_compcode = '" + CompCode + "' And " _
                & "qs4_npecode = '" + deIms.NameSpace + "' AND " _
                & "qs4_ware = '" + ssdcboWarehouse.Columns("Code").Text + "' AND " _
                & "qs4_stcknumb = '" + lblCommodity + "' AND " _
                & "qs4_logiware = '" + LW + "' AND " _
                & "qs4_subloca = '" + SL + "' AND " _
                & "qs4_cond = '" + SC + "'"
            Set qty = New ADODB.Recordset
            qty.Open Sql, deIms.cnIms, adOpenForwardOnly
            If qty.RecordCount > 0 Then
                txtprimUnit.Tag = FormatNumber(qty!PrimaryQTY, 0)
                txtprimUnit = txtprimUnit.Tag
            Else
                txtprimUnit.Tag = "0"
                txtprimUnit = "0"
            End If
        End If
        Screen.MousePointer = 0
    End If
End Sub

Public Sub MoveRecord(direction As MoveType)
Screen.MousePointer = 11
On Error Resume Next
    If Not rs Is Nothing Then
        If rs.RecordCount <> 0 Then
        
            'Added by Juan to fix navigation
            rs!iid_ctry = ssdcboCountry
            If optOwn Then
                rs!iid_owle = 1
            Else
                rs!iid_owle = 0
            End If
            If txtLeaseComp <> "" Then rs!iid_leasecomp = txtLeaseComp
            rs!iid_fromlogiware = ssdcboLogicalWHouse(0).Columns("Code").value
            rs!iid_tologiware = ssdcboLogicalWHouse(1).Columns("Code").value
            rs!iid_fromsubloca = ssdcboSubLocation(0).Columns("Code").value
            rs!iid_tosubloca = ssdcboSubLocation(1).Columns("Code").value
            rs!iid_origcond = ssdcboCondition(0).Columns("Code").value
            rs!iid_newcond = ssdcboCondition(1).Columns("Code").value
            If optPool Then
                rs!iid_ps = 1
            Else
                rs!iid_ps = 0
            End If
            If cboSerialNumb <> "" Then rs!iid_serl = cboSerialNumb
            If IsNumeric(txtprimUnit) Then
                rs!iid_primqty = CDbl(txtprimUnit)
                rs!iid_secoqty = CDbl(lblSecQnty)
            Else
                Exit Sub
            End If
            '---------------------------------
        
            If direction > 0 Then
                ValidateControls
                CheckDetl
            End If
            Select Case direction
                Case mtFirst
                    rs.MoveFirst
                Case mtNext
                    rs.MoveNext
                Case mtLast
                    rs.MoveLast
                Case mtPrevious
                    rs.MovePrevious
                Case -1
                    rs.CancelUpdate
            End Select
            If rs.EOF Then
                rs.MoveLast
            ElseIf rs.BOF Then
                rs.MoveFirst
            ElseIf rs.RecordCount = 1 Then
                Call DisableNav(True, True)
            End If
            
            If rs.AbsolutePosition = 1 Then
                Call DisableNav(True, False)
                
            ElseIf rs.AbsolutePosition = rs.RecordCount Then
                Call DisableNav(False, True)
            Else
                Call DisableNav(False, False)
            End If
            
            lblRecCount = rs.RecordCount
            lblCurrRec = rs.AbsolutePosition
            If direction > -2 Then
                Call DisplayRecord
            Else
                If direction = -1 Then
                    Call DisplayRecord
                End If
            End If
        End If
    End If
    
    'Modified by Juan (9/26/2000) for Multilingual
    'If SSTab1.Caption = "Receipt" Then 'J hidden
    If SSTab1.Tab = 0 Then 'J added
    '---------------------------------------------
        
        NavBar1.NextEnabled = False
        NavBar1.LastEnabled = False
    End If
    
'    If SSTab1.Caption = "Line Items Received" Then
'        NavBar1.FirstEnabled = True
'        NavBar1.NextEnabled = True
'        NavBar1.LastEnabled = True
'        NavBar1.PreviousEnabled = True
'    End If
    
    'Modified by Juan (9/26/2000) for Multilingual
    'If SSTab1.Caption = "Remarks" Then 'J hidden
    If SSTab1.Tab = 2 Then 'J added
    '---------------------------------------------
    
        NavBar1.NextEnabled = False
        NavBar1.LastEnabled = False
    End If
    
    
    'Modified by Juan (9/26/2000) for Multilingual
    'If SSTab1.Caption = "Recipients" Then 'J hidden
    If SSTab1.Tab = 3 Then 'J added
    '---------------------------------------------
    
        NavBar1.NextEnabled = False
        NavBar1.LastEnabled = False
    End If

    Screen.MousePointer = 0
    If Err Then Call LogErr(Name & "::MoveRecord", Err.Description, Err.number, True)
End Sub

Private Sub DisableNav(BackWard As Boolean, Forward As Boolean)
On Error Resume Next

    Forward = Not Forward
    BackWard = Not BackWard
    NavBar1.LastEnabled = Forward
    NavBar1.NextEnabled = Forward
    NavBar1.FirstEnabled = BackWard
    NavBar1.PreviousEnabled = BackWard
    

End Sub


Public Sub DisplayRecord()
Dim qty1, qty2

    txtprimUnit = ""

    lblCurrencyValu = "1"
    optPool = rs("iid_ps") & ""
    optOwn = rs("iid_owle") & ""
    optLease = Not rs("iid_owle") & ""
    optSpecific = Not rs("iid_ps") & ""
    txtDesc = rs("iid_stckdesc") & ""
    lblCommodity = rs("iid_stcknumb") & ""
    txtLeaseComp = rs("iid_leasecomp") & ""
    qty1 = rs("iid_primqty")
    qty2 = rs("iid_secoqty")
    
    'FillCombos
    Call ssdcboLogicalWHouse_Click(0)
    Call FindInGrid(ssdcboSubLocation(0), rs("iid_fromsubloca") & "", True, 1)
    
    Call ssdcboSubLocation_Click(0)
    Call FindInGrid(ssdcboCondition(0), rs("iid_origcond") & "", True, 1)
    
    Call ssdcboCondition_Click(0)
    Call FindInGrid(ssdcboLogicalWHouse(1), rs("iid_tologiware") & "", True, 1)
    
    Call ssdcboLogicalWHouse_Click(1)
    Call FindInGrid(ssdcboSubLocation(1), rs("iid_tosubloca") & "", True, 1)
    
    Call ssdcboSubLocation_Click(1)
    Call FindInGrid(ssdcboCondition(1), rs("iid_newcond") & "", True, 1)
    
    Call ssdcboCondition_Click(1)
    ssdcboCountry.Text = ssdcboCountry.Columns(0).Text
    ssdcboCondition(0).Text = ssdcboCondition(0).Columns(0).Text
    ssdcboSubLocation(0).Text = ssdcboSubLocation(0).Columns(0).Text
    ssdcboLogicalWHouse(0).Text = ssdcboLogicalWHouse(0).Columns(0).Text
    
    ssdcboCondition(1).Text = ssdcboCondition(1).Columns(0).Text
    ssdcboSubLocation(1).Text = ssdcboSubLocation(1).Columns(0).Text
    ssdcboLogicalWHouse(1).Text = ssdcboLogicalWHouse(1).Columns(0).Text
    
    rs("iid_primqty") = qty1
    rs("iid_secoqty") = qty2
    lblSecQnty = FormatNumber(rs("iid_secoqty") & "", 2)
    txtprimUnit = FormatNumber(rs("iid_primqty") & "", 0)
    lblUnitprice = FormatNumber(rs("iid_unitpric") & "", 2)
    cboSerialNumb.ListIndex = IndexOf(cboSerialNumb, rs("iid_serl") & "")
    
End Sub

'set navbar buttom

Private Sub cbo_Transaction_Click()

     If Len(cbo_Transaction) <> 0 Then
    
        Imsmail1.Enabled = True
        NavBar1.PrintEnabled = True
        ssdbRecepientList.Enabled = True
        NavBar1.EMailEnabled = ssdbRecepientList.Rows
    Else
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
    End If
    
'    If Len(cbo_Transaction) Then
'        IMSMail1.Enabled = True
'        ssdbRecepientList.Enabled = True
'    End If
End Sub

'nulock trancsation combo

Private Sub cbo_Transaction_DropDown()
    cbo_Transaction.locked = False
End Sub

'call function get issue number

Private Sub cbo_Transaction_GotFocus()
    cbo_Transaction.BackColor = &HC0FFFF
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
End Sub

'do not allow enter data from keybroad

Private Sub cbo_Transaction_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then
KeyAscii = 0
End If
End Sub

Private Sub cbo_Transaction_LostFocus()
    cbo_Transaction.BackColor = &H80000005
End Sub

Private Sub cboSerialNumb_Click()
    txtprimUnit = 1
    lblSecQnty = 1
End Sub

Private Sub cboSerialNumb_GotFocus()
    cboSerialNumb.BackColor = &HC0FFFF
End Sub


Private Sub cboSerialNumb_LostFocus()
    cboSerialNumb.BackColor = &H80000005
End Sub


'call function add current reciptient to reciprient list

Private Sub cmd_Add_Click()
    Imsmail1.AddCurrentRecipient
End Sub

'detele a recipient from recipient list

Private Sub cmd_Remove_Click()
On Error Resume Next

    rsReceptList.Delete
    rsReceptList.Update
    
    If Err Then Err.Clear
End Sub

'call functions get datas for combos and data grid

Private Sub Form_Load()
Dim np As String
Dim FCompany As String
Dim cn As ADODB.Connection
Dim rights
Dim i As Integer

    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmIntTransfer")
    '------------------------------------------
    
    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    NavBar1.SaveEnabled = SaveEnabled
    
    For i = 1 To 2
        SSTab1.TabVisible(i) = SaveEnabled
    Next
    

    np = deIms.NameSpace
    Set cn = deIms.cnIms
    fm = mdvisualization
    
    Requery = True
    FCompany = GetCompany(np, "PE", cn)
    CompCode = GetCompanyCode(np, FCompany, cn)
    
    AddCompanies
    AddCountries
    AddSubLocations
    AddLogicalWhareHouse
    EnableControls (False)
    Set InvtIss = New imsWhareIssue
    
    AddIssueNumb
    
    AddCondition
    ssdbStockInfo.FieldSeparator = Chr(1)
    
    Call DisableButtons(Me, NavBar1)
    rights = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    SaveEnabled = rights
    ssdbStockInfo.Enabled = SaveEnabled
    
    Imsmail1.NameSpace = deIms.NameSpace
    
    'IMSMail1.Connected = True 'M
    Imsmail1.SetActiveConnection deIms.cnIms 'M
    Imsmail1.Language = Language 'M
    Imsmail1.Enabled = False
    ssdbRecepientList.Enabled = False
    NavBar1.CloseEnabled = True
    frmIntTransfer.Caption = frmIntTransfer.Caption + " - " + frmIntTransfer.Tag
    SSTab1.TabVisible(3) = False
    
    cbo_Transaction.locked = False
    cbo_Transaction.Enabled = True
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'get recordset for company combo and populate combo

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

'SQL statement get sub location and populate data grid

Private Sub AddSubLocations()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select sb_code Code, sb_desc Description from SUBLOCATION"
        .CommandText = .CommandText & " where sb_npecode = '" & deIms.NameSpace & "' ORDER BY sb_desc"
        
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboSubLocation(1).RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboSubLocation(1).AddItem rst!Description & ";" & rst!Code
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'fill data to warwhouse data grid

Private Sub AddWhareHouses(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboWarehouse.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboWarehouse.AddItem rst!loc_name & "" & ";" & rst!loc_locacode & ""
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set rst = Nothing
End Sub

'Private Sub AddLocations(rst As ADODB.Recordset)
'
'    If rst Is Nothing Then Exit Sub
'    If rst.State And adStateOpen = adStateClosed Then Exit Sub
'
'    If rst.RecordCount = 0 Then GoTo CleanUp
'
'    rst.MoveFirst
'
'
'    Do While ((Not rst.EOF))
'
'        ssdcboLocation.AddItem rst!loc_name & "" & ";" & rst!loc_locacode & ""
'        rst.MoveNext
'    Loop
'
'CleanUp:
'    rst.Close
'    Set rst = Nothing
'End Sub

'Private Sub AddStockType(rst As ADODB.Recordset)
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

'SQL statement get country information and populate data grid

Private Sub AddCountries()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select ctry_code Code, ctry_name Description from COUNTRY"
        .CommandText = .CommandText & " where ctry_npecode = '" & deIms.NameSpace & "' ORDER BY ctry_name"
        
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

'unload form  free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim closing
    If fm <> mdvisualization Then
        closing = MsgBox("Do you really want to close and lose your last record?", vbYesNo)
        If closing = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    Hide
    If rs.State And adStateOpen = adStateOpen Then
        rs.CancelBatch
        rs.Close
    End If
    
    Set rs = Nothing
    Set rsReceptList = Nothing
    'Imsmail1.Connected = False 'M
    If Err Then Err.Clear
    
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub NavBar1_OnCancelClick()
On Error Resume Next
    Call MoveRecord(-1)
    backpass = True
    optSpecific.value = False
    If rs Is Nothing Then
        SSTab1.Tab = 0
        beginning = False
    Else
        If rs.RecordCount = 0 Then
            SSTab1.Tab = 0
            beginning = False
        End If
    End If
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnFirstClick()
On Error Resume Next
    Call MoveRecord(mtFirst)
End Sub

Private Sub NavBar1_OnLastClick()
On Error Resume Next
    Call MoveRecord(mtLast)
End Sub


Private Sub NavBar1_OnNextClick()
On Error Resume Next
    Call MoveRecord(mtNext)
End Sub

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next
    Call MoveRecord(mtPrevious)
End Sub


'call function to print crystal report

Private Sub NavBar1_OnPrintClick()
    BeforePrint
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset
End Sub

'before save call validate function check data format and assign
'values to store procedure parameter

Private Sub NavBar1_OnSaveClick()
NavBar1.SaveEnabled = False
MDI_IMS.StatusBar1.Panels(1).Text = "Saving"
Screen.MousePointer = 11

Dim retval As Boolean
Dim np As String
Dim ToWH As String
Dim FromWH As String
Dim PrimUnit As Double
Dim SecUnit As Double
Dim StockNumb As String
Dim cn As ADODB.Connection

    On Error Resume Next
    NavBar1.SaveEnabled = False
    cbo_Transaction.ListIndex = CB_ERR
    If CheckMasterFields And CheckDetl Then
    
        If Len(Trim(txtRemarks)) = 0 Then
            Screen.MousePointer = 0
            MsgBox "Remarks cannot be empty"
            SSTab1.Tab = 2
            txtRemarks.SetFocus
            Exit Sub
        End If
        
        FrmShowApproving.Show
        FrmShowApproving.Label2 = "Saving Transaction"
        FrmShowApproving.Refresh

        MDI_IMS.StatusBar1.Panels(1).Text = "Beginning Transaction"
        Call BeginTransaction(deIms.cnIms)
        
        retval = PutInvtIssue
        
        'doevents
        If retval = False Then GoTo RollBack
        
        rs.MoveFirst
        MDI_IMS.StatusBar1.Panels(1).Text = "Saving Details"
        Do While Not rs.EOF
            retval = PutDataInsert
            NavBar1.SaveEnabled = False
            If retval = False Then GoTo RollBack
            
            Set cn = deIms.cnIms
            np = deIms.NameSpace
            FromWH = rs!iid_ware
            
            ToWH = FromWH
            StockNumb = rs!iid_stcknumb
            PrimUnit = rs!iid_primqty
            SecUnit = IIf(IsNull(rs!iid_secoqty), 0, rs!iid_secoqty)
            
            retval = Update_Sap(np, CompCode, StockNumb, ToWH, PrimUnit, 1, rs!iid_unitpric, rs!iid_newcond, CurrentUser, cn)
            retval = retval And Quantity_In_stock1_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_stckdesc, CurrentUser, cn)
            retval = retval And Quantity_In_stock2_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, CurrentUser, cn)
            retval = retval And Quantity_In_stock3_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, CurrentUser, cn)
            retval = retval And Quantity_In_stock4_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, CurrentUser, cn)
            
            'doevents
            If rs!iid_ps Then
                retval = retval And Quantity_In_stock5_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, Transnumb, rs!iid_transerl, rs!iid_ware, "IT", CompCode, FromWH, Transnumb, CompCode, rs!iid_transerl, CurrentUser, cn)
            Else
                 retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, rs!iid_serl, CurrentUser, cn)
                 retval = retval And Quantity_In_stock7_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, Transnumb, FromWH, rs!iid_transerl, rs!iid_ware, "IT", CompCode, Transnumb, CompCode, rs!iid_transerl, rs!iid_serl, CurrentUser, cn)
            
            End If
            
            If retval = False Then GoTo RollBack
            
            SecUnit = SecUnit * -1
            PrimUnit = PrimUnit * -1
            retval = retval And Quantity_In_stock1_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_stckdesc, CurrentUser, cn)
            retval = retval And Quantity_In_stock2_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, CurrentUser, cn)
            retval = retval And Quantity_In_stock3_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, CurrentUser, cn)
            retval = retval And Quantity_In_stock4_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, CurrentUser, cn)
            
            'doevents
            If rs!iid_ps Then
                retval = retval And Quantity_In_stock5_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, Transnumb, rs!iid_transerl, rs!iid_ware, "IT", CompCode, ToWH, Transnumb, CompCode, rs!iid_transerl, CurrentUser, cn)
            Else
               'Modified by Muzammil 08/12/00
               'Reason - it was passing rs!ird_transerl instead of rs!ird_serl
               retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, rs!iid_serl, CurrentUser, cn)
'                 retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, rs!iid_transerl, CurrentUser, cn)
                 retval = retval And Quantity_In_stock7_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, Transnumb, ToWH, rs!iid_transerl, rs!iid_ware, "IT", CompCode, Transnumb, CompCode, rs!iid_transerl, rs!iid_serl, CurrentUser, cn)
            
            End If
            
            If retval = False Then GoTo RollBack
            rs.MoveNext
        Loop
        
        MDI_IMS.StatusBar1.Panels(1).Text = "Saving Remarks"
        'doevents
        PutIssueRemarks
        NavBar1.SaveEnabled = SaveEnabled
        If retval Then Call CommitTransaction(deIms.cnIms)
        If retval Then Call CommitTransaction(deIms.cnIms)
        
        Screen.MousePointer = 0
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00018") 'J added
        MsgBox IIf(msg1 = "", "Please note that your transaction # is ", msg1 + " ") & Transnumb
        '---------------------------------------------
        Screen.MousePointer = 11
        
        MDI_IMS.StatusBar1.Panels(1).Text = "Restoring"
        'doevents
        Screen.MousePointer = 11
        Call cbo_Transaction.AddItem(Transnumb, cbo_Transaction.ListCount)
        Screen.MousePointer = 11
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, Transnumb)
        
        'BeforePrint
        'Call SendWareHouseMessage(deIms.NameSpace, "Automatic Distribution" _
                                 , lblType, deIms.cnIms, CreateRpti)

    On Error Resume Next
        lblUser = ""
        lblDate = ""
        
        
        Call EnableControls(False)
        
        Requery = True
        rs.CancelBatch
        Call ClearFields
        fm = mdvisualization
        rs.Close: Set rs = Nothing
        Set rsReceptList = Nothing
        
        Imsmail1.Enabled = True
        ssdbRecepientList.Enabled = True
        ssdbStockInfo.RemoveAll
        ssdcboWarehouse.Text = ""
    End If
    
    If Err Then Err.Clear
    
    MDI_IMS.StatusBar1.Panels(1).Text = ""
    
    ssdcboCompany.Enabled = True
    ssdcboWarehouse.Enabled = True
    'AddIssueNumb
    cbo_Transaction.SetFocus
    Unload FrmShowApproving
    Screen.MousePointer = 0
    Exit Sub
    
    
RollBack:
    NavBar1.SaveEnabled = SaveEnabled
    Call RollbackTransaction(deIms.cnIms)
    Call RollbackTransaction(deIms.cnIms)
    MDI_IMS.StatusBar1.Panels(1).Text = ""
    Unload FrmShowApproving
    Screen.MousePointer = 0
End Sub

'laod data to stock info text boxse

Private Sub ssdbStockInfo_DblClick()

Screen.MousePointer = 11
Dim rst As ADODB.Recordset
Dim WareHouse As String
Dim SU As String
Dim PU As String


    '
    If ssdbStockInfo.Rows < 1 Then Exit Sub
    cbo_Transaction.ListIndex = CB_ERR
    Screen.MousePointer = 11
    justDBLCLICK = True
    
    If Not (rs Is Nothing) Then
        If rs.RecordCount > 0 Then
            If Not (Requery) Then If CheckDetl = False Then Exit Sub
        End If
    End If
    
    fm = mdCreation
    txtprimUnit.Tag = ""
    beginning = True
    SSTab1.Tab = 1
    Screen.MousePointer = 11
    
    rs.AddNew
    
    lblRecCount = rs.RecordCount
    lblCurrRec = rs.AbsolutePosition
    
    ClearFields
    AssignDefValues
    Screen.MousePointer = 11
    Call EnableControls(True)
    WareHouse = ssdcboWarehouse.Columns("Code").Text
    lblCommodity = ssdbStockInfo.Columns("Commodity").Text
    txtDesc = ssdbStockInfo.Columns("Description").Text
    
    ssdcboCountry.value = "USA"
    Screen.MousePointer = 11
    Call FindInGrid(ssdcboCountry, "USA", True, 1)
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
    
    Call AddFromLogicalWharehouse(Get_LogicalWarehouse_FromQTYST(deIms.NameSpace, _
                                  lblCommodity, WareHouse, CompCode, deIms.cnIms))

     'txtprimUnit = IIf(txtprimUnit.Tag > 0, txtprimUnit.Tag, "")
    
    Screen.MousePointer = 11
    Call txtprimUnit_Validate(False)
    If IsNumeric(txtprimUnit) Then
        txtprimUnit = FormatNumber$(txtprimUnit, 0)
    Else
        txtprimUnit = ""
    End If
    'Modified by Muzammil 08/15/00
   'Reason - Did not work good for records which are not in the first 9(The first set
   'which it displays)
    
    'Call ssdbStockInfo.RemoveItem(ssdbStockInfo.Row)  'M
    
    Call ssdbStockInfo.RemoveItem(ssdbStockInfo.AddItemRowIndex(ssdbStockInfo.Bookmark)) 'M

    ssdcboCompany.Enabled = False
    ssdcboWarehouse.Enabled = False
    Screen.MousePointer = 0
End Sub

'call function get warehouse data and populate data grid

Private Sub ssdcboCompany_Click()
    cbo_Transaction.ListIndex = CB_ERR
    CompCode = ssdcboCompany.Columns("Code").Text
    
    ssdcboWarehouse = ""
    ssdcboWarehouse.RemoveAll
    Call AddWhareHouses(GetLocation(deIms.NameSpace, "OTHER", CompCode, deIms.cnIms, False))
    
    ssdcboCompany.SelLength = 0
    ssdcboCompany.SelStart = 0
End Sub

Private Sub ssdcboCompany_GotFocus()
    ssdcboCompany.BackColor = &HC0FFFF
End Sub


Private Sub ssdcboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboCompany.DroppedDown = True
    End If
End Sub


Private Sub ssdcboCompany_LostFocus()
    ssdcboCompany.BackColor = &H80000005
End Sub

'call function get values and populate condition data grid

Private Sub ssdcboCondition_Click(Index As Integer)
Dim l As Double
Dim rst As ADODB.Recordset

    If Index = 0 Then
        l = GetSapValue(deIms.NameSpace, CompCode, lblCommodity, ssdcboWarehouse.Columns("Code").Text, _
                    ssdcboCondition(0).Columns("Code").Text, deIms.cnIms)
                    

        
        rs!iid_unitpric = l
        lblUnitprice = FormatNumber$(l, 2)
        ssdcboCondition(1).Text = ssdcboCondition(0).Text
        
        Call FindInGrid(ssdcboCondition(1), ssdcboCondition(0).Text, True, 1)
        Call deIms.GetItemCount(CompCode, deIms.NameSpace, ssdcboWarehouse.Columns("Code").Text, _
                               lblCommodity, ssdcboLogicalWHouse(0).Columns("Code").Text, _
                                ssdcboSubLocation(0).Columns("Code").Text, ssdcboCondition(0).Columns("Code").Text)
                               
        If deIms.rsGetItemCount.RecordCount > 0 Then
            txtprimUnit.Tag = deIms.rsGetItemCount.Fields(0).value & ""
            If rs Is Nothing Then
                txtprimUnit = txtprimUnit.Tag
            Else
                If rs.RecordCount > 0 Then
                    txtprimUnit = FormatNumber(rs!iid_primqty, 0)
                Else
                    txtprimUnit = txtprimUnit.Tag
                End If
            End If
        End If
        
        deIms.rsGetItemCount.Close
        cboSerialNumb.Clear
        Set rst = Get_SerialNumberFromStockNumber(deIms.NameSpace, lblCommodity, _
                                            ssdcboWarehouse.Columns("Code").Text, _
                                            CompCode, ssdcboLogicalWHouse(0).Columns("Code").Text, _
                                            ssdcboSubLocation(0).Columns("Code").Text, ssdcboCondition(0).Columns("Code").Text, deIms.cnIms)

        
        Call PopuLateFromRecordSet(cboSerialNumb, rst, rst.Fields(0).Name, True)
         
        'optSpecific.Enabled = cboSerialNumb.ListCount
        cboSerialNumb.Enabled = cboSerialNumb.ListCount
        
        'FG 8/15 if several serial, make qty = 1
        If optSpecific.Enabled Then
            txtprimUnit = 1
            txtprimUnit.Tag = 1
            txtprimUnit_Validate (True)
        End If
        
        'optSpecific = cboSerialNumb.Enabled
        'optPool.Enabled = cboSerialNumb.Enabled = False
        
        If IsNumeric(txtprimUnit) Then
            txtprimUnit = FormatNumber((txtprimUnit), 0)
        Else
            txtprimUnit = ""
        End If
        If rst.State And adStateOpen = adStateOpen Then rst.Close
        Set rst = Nothing
        Call refreshQTY
    End If

'    If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'        If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'            If ssdcboCondition(0) = ssdcboCondition(1) Then
'                MsgBox "Invalid Values Between From and To"
'                Exit Sub
'            End If
'        End If
'    End If

End Sub

'assign values to lable

Private Sub ssdcboLocation_Click()

    lblUser = CurrentUser
    
    
    lblDate = Date
    

End Sub

'do not allow enter data to data grid

Private Sub ssdcboLocation_KeyPress(KeyAscii As Integer)
    If NavBar1.NewEnabled = False Then KeyAscii = 0
End Sub

Private Sub ssdcboCondition_GotFocus(Index As Integer)
    ssdcboCondition(Index).BackColor = &HC0FFFF
End Sub

Private Sub ssdcboCondition_LostFocus(Index As Integer)
    ssdcboCondition(Index).BackColor = &H80000005
End Sub

Private Sub ssdcboCountry_GotFocus()
    ssdcboCountry.BackColor = &HC0FFFF
End Sub

Private Sub ssdcboCountry_LostFocus()
    ssdcboCountry.BackColor = &H80000005
End Sub

'call function populate sub location data grid

Private Sub ssdcboLogicalWHouse_Click(Index As Integer)
    If Index = 0 Then
    
        Call AddFromSublocation(Get_SubLocation_FromQTYST(deIms.NameSpace, lblCommodity, _
                                ssdcboWarehouse.Columns("Code").Text, CompCode, _
                                ssdcboLogicalWHouse(0).Columns("Code").Text, deIms.cnIms))

        Call ssdcboSubLocation_Click(0)
        ssdcboSubLocation(0).Enabled = True
        Call refreshQTY
    End If
'    If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'        If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'            If ssdcboCondition(0) = ssdcboCondition(1) Then
'                MsgBox "Invalid Values Between From and To"
'                Exit Sub
'            End If
'        End If
'    End If
End Sub

Private Sub ssdcboLogicalWHouse_GotFocus(Index As Integer)
    ssdcboLogicalWHouse(Index).BackColor = &HC0FFFF
End Sub


Private Sub ssdcboLogicalWHouse_LostFocus(Index As Integer)
    ssdcboLogicalWHouse(Index).BackColor = &H80000005
End Sub

'assign data to text boxse and call function get condition
'data and populate data grid

Private Sub ssdcboSubLocation_Click(Index As Integer)
Dim WareHouse As String
Dim LWareHouse As String

    If Index = 0 Then
        WareHouse = ssdcboWarehouse.Columns("Code").Text
        LWareHouse = ssdcboLogicalWHouse(0).Columns("Code").Text
        If Len(Trim$(ssdcboLogicalWHouse(0).Text)) = 0 Then Exit Sub
        Call AddFromCondition(Get_Condition_FromQTYST(deIms.NameSpace, lblCommodity, WareHouse, CompCode, LWareHouse, ssdcboSubLocation(0).Columns("Code").Text, deIms.cnIms))
    
        ssdcboCondition(0).Enabled = True
        Call refreshQTY
    End If
    
'    If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'        If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'            If ssdcboCondition(0) = ssdcboCondition(1) Then
'                MsgBox "Invalid Values Between From and To"
'                Exit Sub
'            End If
'        End If
'    End If

End Sub

Private Sub ssdcboSubLocation_GotFocus(Index As Integer)
    ssdcboSubLocation(Index).BackColor = &HC0FFFF
End Sub


Private Sub ssdcboSubLocation_LostFocus(Index As Integer)
    ssdcboSubLocation(Index).BackColor = &H80000005
End Sub

'assign values to lable and call function get stock information
'and populate data grid

Private Sub ssdcboWarehouse_Click()
    
    lblUser = CurrentUser
    
    
    lblDate = Date
    
    
    Imsmail1.Enabled = False
    ssdbRecepientList.Enabled = False
    cbo_Transaction.ListIndex = CB_ERR
    Call AddStockInfo(GetStockInformation(deIms.NameSpace, ssdcboWarehouse.Columns("Code").Text, CompCode, deIms.cnIms))

End Sub

'fill dat to data grid

Private Sub AddStockInfo(rst As ADODB.Recordset)
On Error Resume Next

    ssdbStockInfo.RemoveAll
    
    If rst Is Nothing Then Exit Sub
    If rst.EOF And rst.BOF Then Exit Sub
    If rst.RecordCount = 0 Then
        Exit Sub
    Else
        rst.Sort = "qs1_stcknumb asc"
    End If
    
    Do While Not rst.EOF
        ssdbStockInfo.AddItem ((rst!qs1_stcknumb & "") & Chr(1) & (rst!qs1_desc & "") & Chr(1) & FormatNumber((rst!qs1_primqty & ""), 2))
        rst.MoveNext
    Loop
End Sub

' assign values

Private Sub AssignInvt()
    With InvtIss
        lblUser = .User
        lblDate = .TransactionDate
        
        
    End With
End Sub

'SQL statement get logical warehouse data and populate data grid

Private Sub AddLogicalWhareHouse()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select lw_code Code, lw_desc Description from LOGWAR"
        .CommandText = .CommandText & " where lw_npecode = '" & deIms.NameSpace & "' ORDER BY lw_desc"
        
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboLogicalWHouse(1).RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboLogicalWHouse(1).AddItem rst!Description & ";" & rst!Code
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'assign data values to store procedure parameters

Private Function PutDataInsert() As Boolean

    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert = False

    Set cmd = deIms.Commands("InvtIssueDetl_INSERT")


    'Check for valid data.
    If Not ValidateData() Then
        Exit Function
    End If

    'Set the parameter values for the command to be executed.
    cmd.parameters("@iid_trannumb") = Transnumb
    cmd.parameters("@iid_compcode") = GetPKValue(rs.Bookmark, "iid_compcode")
    cmd.parameters("@iid_npecode") = GetPKValue(rs.Bookmark, "iid_npecode")
    cmd.parameters("@iid_ware") = GetPKValue(rs.Bookmark, "iid_ware")
    cmd.parameters("@iid_transerl") = GetPKValue(rs.Bookmark, "iid_transerl")
    cmd.parameters("@iid_stcknumb") = GetPKValue(rs.Bookmark, "iid_stcknumb")
    cmd.parameters("@iid_ps") = GetPKValue(rs.Bookmark, "iid_ps")
    cmd.parameters("@iid_serl") = GetPKValue(rs.Bookmark, "iid_serl")
    cmd.parameters("@iid_newcond") = GetPKValue(rs.Bookmark, "iid_newcond")
    cmd.parameters("@iid_stcktype") = GetPKValue(rs.Bookmark, "iid_stcktype")
    cmd.parameters("@iid_ctry") = GetPKValue(rs.Bookmark, "iid_ctry")
    cmd.parameters("@iid_tosubloca") = GetPKValue(rs.Bookmark, "iid_tosubloca")
    cmd.parameters("@iid_tologiware") = GetPKValue(rs.Bookmark, "iid_tologiware")
    cmd.parameters("@iid_owle") = GetPKValue(rs.Bookmark, "iid_owle")
    cmd.parameters("@iid_leasecomp") = GetPKValue(rs.Bookmark, "iid_leasecomp")
    cmd.parameters("iid_primqty") = GetPKValue(rs.Bookmark, "iid_primqty")
    cmd.parameters("@iid_secoqty") = GetPKValue(rs.Bookmark, "iid_secoqty")
    cmd.parameters("@iid_unitpric") = GetPKValue(rs.Bookmark, "iid_unitpric")
    cmd.parameters("iid_curr") = "USD" 'GetPKValue(rs.Bookmark, "iid_curr")
    cmd.parameters("iid_currvalu") = 1 ' GetPKValue(rs.Bookmark, "iid_currvalu")
    cmd.parameters("iid_stckdesc") = GetPKValue(rs.Bookmark, "iid_stckdesc")
    cmd.parameters("@iid_fromlogiware") = GetPKValue(rs.Bookmark, "iid_fromlogiware")
    cmd.parameters("@iid_fromsubloca") = GetPKValue(rs.Bookmark, "iid_fromsubloca")
    cmd.parameters("@iid_origcond") = GetPKValue(rs.Bookmark, "iid_origcond")
    cmd.parameters("@user") = CurrentUser
    'Execute the command.
    cmd.Execute

    PutDataInsert = True

    Exit Function

errPutDataInsert:
    MsgBox Err.Description: Err.Clear
End Function

'validate data format

Private Function ValidateData() As Boolean

    Dim i As Long

    ValidateData = False

    'Verify the field is not null.
    If IsNull(rs("iid_compcode")) Then
        MsgBox "The field ' iid_compcode ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_compcode")) Then
        If Len(Trim(rs("iid_compcode"))) = 0 Then
            MsgBox "The field ' iid_compcode ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("iid_npecode")) Then
        MsgBox "The field ' iid_npecode ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_npecode")) Then
        If Len(Trim(rs("iid_npecode"))) = 0 Then
            MsgBox "The field ' iid_npecode ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("iid_ware")) Then
        MsgBox "The field ' iid_ware ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_ware")) Then
        If Len(Trim(rs("iid_ware"))) = 0 Then
            MsgBox "The field ' iid_ware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the integer field contains a valid value.
    If Not IsNull(rs("iid_trannumb")) Then
        If Not IsNumeric(rs("iid_trannumb")) _
            And InStr(rs("iid_trannumb"), ".") = 0 Then
            MsgBox "The field ' iid_trannumb ' does not contain a valid number."
        Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("iid_transerl")) Then
        MsgBox "The field ' iid_transerl ' cannot be null."
        Exit Function
    End If

    'Verify the integer field contains a valid value.
    If Not IsNull(rs("iid_transerl")) Then
        If Not IsNumeric(rs("iid_transerl")) _
            And InStr(rs("iid_transerl"), ".") = 0 Then
            MsgBox "The field ' iid_transerl ' does not contain a valid number."
        Exit Function
        End If
    End If


    'Verify the field is not null.
    If IsNull(rs("iid_stcknumb")) Then
        MsgBox "The field ' iid_stcknumb ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_stcknumb")) Then
        If Len(Trim(rs("iid_stcknumb"))) = 0 Then
            MsgBox "The field ' iid_stcknumb ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("iid_ps")) Then
        MsgBox "The field ' iid_ps ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_serl")) Then
        If Len(Trim(rs("iid_serl"))) = 0 Then
            MsgBox "The field ' iid_serl ' does not contain valid text."
            Exit Function
        End If
    End If


    'Verify the text field contains text.
    If Not IsNull(rs("iid_newcond")) Then
        If Len(Trim(rs("iid_newcond"))) = 0 Then
            MsgBox "The field ' iid_newcond ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_stcktype")) Then
        If Len(Trim(rs("iid_stcktype"))) = 0 Then
            MsgBox "The field ' iid_stcktype ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_ctry")) Then
        If Len(Trim(rs("iid_ctry"))) = 0 Then
            MsgBox "The field ' iid_ctry ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_tosubloca")) Then
        If Len(Trim(rs("iid_tosubloca"))) = 0 Then
            MsgBox "The field ' iid_tosubloca ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_tologiware")) Then
        If Len(Trim(rs("iid_tologiware"))) = 0 Then
            MsgBox "The field ' iid_tologiware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_leasecomp")) Then
        If Len(Trim(rs("iid_leasecomp"))) = 0 Then
            MsgBox "The field ' iid_leasecomp ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("iid_primqty")) Then
        If Not IsNumeric(rs("iid_primqty")) Then
            MsgBox "The field ' iid_primqty ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("iid_secoqty")) Then
        If Not IsNumeric(rs("iid_secoqty")) Then
            MsgBox "The field ' iid_secoqty ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("iid_unitpric")) Then
        If Not IsNumeric(rs("iid_unitpric")) Then
            MsgBox "The field ' iid_unitpric ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_curr")) Then
        If Len(Trim(rs("iid_curr"))) = 0 Then
            MsgBox "The field ' iid_curr ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rs("iid_currvalu")) Then
        If Not IsNumeric(rs("iid_currvalu")) Then
            MsgBox "The field ' iid_currvalu ' does not contain a valid numeric value."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_stckdesc")) Then
        If Len(Trim(rs("iid_stckdesc"))) = 0 Then
            MsgBox "The field ' iid_stckdesc ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_fromlogiware")) Then
        If Len(Trim(rs("iid_fromlogiware"))) = 0 Then
            MsgBox "The field ' iid_fromlogiware ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_fromsubloca")) Then
        If Len(Trim(rs("iid_fromsubloca"))) = 0 Then
            MsgBox "The field ' iid_fromsubloca ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_origcond")) Then
        If Len(Trim(rs("iid_origcond"))) = 0 Then
            MsgBox "The field ' iid_origcond ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the decimal field contains a valid value.


    ValidateData = True

End Function

'assign values to store procedure parameters

Private Function PutInvtIssue() As Boolean
Dim np As String
    Dim cmd As Command

On Error GoTo errPutInvtIssue

    PutInvtIssue = False

    Set cmd = deIms.Commands("InvtIssue_Insert")


    np = deIms.NameSpace
    Transnumb = "IT-" & GetTransNumb(np, deIms.cnIms)
    cmd.parameters("@NAMESPACE") = np
    cmd.parameters("@TRANTYPE") = "IT"
    cmd.parameters("@COMPANYCODE") = CompCode
    cmd.parameters("@TRANSNUMB") = Transnumb
    cmd.parameters("@ISSUTO") = ssdcboWarehouse.Columns("Code").Text
    cmd.parameters("@WHAREHOUSE") = ssdcboWarehouse.Columns("Code").Text
    cmd.parameters("@STCKNUMB") = Null
    cmd.parameters("@COND") = Null
    cmd.parameters("@SAP") = Null
    cmd.parameters("@NEWSAP") = Null
    cmd.parameters("@ENTYNUMB") = Null
    cmd.parameters("@SUPPLIERCODE") = Null
    cmd.parameters("@user") = CurrentUser
    
    cmd.Execute

    PutInvtIssue = cmd.parameters(0).value = 0

    Exit Function

errPutInvtIssue:
    MsgBox Err.Description: Err.Clear
End Function

'get recordset columns values

Private Function GetPKValue(vBookMark As Variant, sColName As String) As Variant
    GetPKValue = rs(sColName)
End Function

'check data field value

Private Function CheckMasterFields() As Boolean

    CheckMasterFields = False
           
    If Len(Trim$(ssdcboWarehouse.Text)) = 0 Then _
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00330") 'J added
        MsgBox IIf(msg1 = "", "Warehouse cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
         Exit Function
    End If
    CheckMasterFields = True
        
End Function

Private Sub ssdcboWarehouse_GotFocus()
    ssdcboWarehouse.BackColor = &HC0FFFF
End Sub

Private Sub ssdcboWarehouse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboWarehouse.DroppedDown = True
    End If
End Sub

'do not allow enter data to data combo

Private Sub ssdcboWarehouse_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then KeyAscii = 0
End Sub

Private Sub ssdcboWarehouse_LostFocus()
    ssdcboWarehouse.BackColor = &H80000005
End Sub

'depend on tab set navbar buttom

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim iEditMode As String
Dim blFlag As Boolean
    If Not beginning Then
        beginning = False
        MsgBox "You are not able to go to next tab until stocknumber is selected."
        SSTab1.Tab = 0
        Exit Sub
    End If

    'Added by Juan (1/17/2001)
    Screen.MousePointer = 11
    Dim alarm As Boolean
    NavBar1.CloseEnabled = True
    alarm = False
    If PreviousTab = 1 Then
        Screen.MousePointer = 0
        If Not backpass Then
            If ssdcboCountry = "" Then
                alarm = True
                MsgBox "Country cannot be empty"
                ssdcboCountry.SetFocus
            End If
            If Not IsNumeric(txtprimUnit) Then
                alarm = True
                MsgBox "Invalid Quantity"
                txtprimUnit.SetFocus
                SSTab1.Tab = 1
            End If
            If ssdcboLogicalWHouse(0) = "" Then
                alarm = True
                MsgBox "From Sub Location is incorrect"
                ssdcboLogicalWHouse(0).SetFocus
            End If
            If ssdcboLogicalWHouse(1) = "" Then
                alarm = True
                MsgBox "From Logical Warehouse is incorrect"
                ssdcboLogicalWHouse(1).SetFocus
            End If
            If ssdcboSubLocation(0) = "" Then
                alarm = True
                MsgBox "To Sub Location is incorrect"
                ssdcboSubLocation(0).SetFocus
            End If
            If ssdcboSubLocation(1) = "" Then
                alarm = True
                MsgBox "To Logical Warehouse is incorrect"
                ssdcboSubLocation(1).SetFocus
            End If
            If ssdcboCondition(0) = "" Then
                alarm = True
                MsgBox "From Condition is incorrect"
                ssdcboCondition(0).SetFocus
            End If
            If ssdcboCondition(1) = "" Then
                alarm = True
                MsgBox "To Condition is incorrect"
                ssdcboCondition(1).SetFocus
            End If
            If optSpecific Then
                If cboSerialNumb = "" Then
                    alarm = True
                    MsgBox "Serial Number is empty"
                    cboSerialNumb.SetFocus
                End If
            End If
        End If
'        If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'            If ssdcboSubLocation(0) = ssdcboSubLocation(1) Then
'                If ssdcboCondition(0) = ssdcboCondition(1) Then
'                    backpass = True
'                    SSTab1.Tab = 1
'                    MsgBox "Invalid Values Between From and To"
'                    Exit Sub
'                End If
'            End If
'        End If
        
        
        Screen.MousePointer = 11
    End If
    If alarm Then
        backpass = False
        SSTab1.Tab = 1
        Exit Sub
    End If
    '------------------------



    blFlag = SSTab1.Tab = 1
    
    With NavBar1
        .NextEnabled = blFlag
        .LastEnabled = blFlag
        .FirstEnabled = blFlag
        .CancelEnabled = blFlag
        .PreviousEnabled = blFlag
    
        .SaveEnabled = SSTab1.Tab = 0
        .CloseEnabled = SSTab1.Tab = 0
        .PrintEnabled = .SaveEnabled And cbo_Transaction.ListIndex <> CB_ERR
        .EMailEnabled = ((ssdbRecepientList.Rows) And (.PrintEnabled))
    End With
    

    Screen.MousePointer = 11
    If SSTab1.Tab = 1 Then
        Me.Refresh
        If PreviousTab = 0 And fm = mdCreation Then _
            If Not (CheckMasterFields) Then SSTab1.Tab = 0
            
        If Requery Then
        
            If fm <> mdCreation Then Exit Sub
            iEditMode = IIf(IsNumeric(cbo_Transaction), cbo_Transaction, "")
            Set rs = deIms.GetInvtIssuedetl(CompCode, iEditMode)
            
            Requery = False
       End If
    End If

    If SSTab1.Tab = 2 Then
        txtRemarks.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

'fill data to sub location data grid

Private Sub AddFromSublocation(rst As ADODB.Recordset)
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboSubLocation(0).RemoveAll
    
    rst.MoveFirst
    ssdcboSubLocation(0).Text = rst!Description
    
    Do While Not rst.EOF
        ssdcboSubLocation(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing

End Sub

'fill data to logical warehouse data grid

Private Sub AddFromLogicalWharehouse(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboLogicalWHouse(0).RemoveAll
    
    rst.MoveFirst
    ssdcboLogicalWHouse(0).Text = rst!Description
    
    Do While Not rst.EOF
        ssdcboLogicalWHouse(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    Call ssdcboLogicalWHouse_Click(0)
End Sub

'fill data to condition data grid

Private Sub AddFromCondition(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboCondition(0).RemoveAll
    
    rst.MoveFirst
    ssdcboCondition(0).Text = rst!Description
    
    Do While Not rst.EOF
        ssdcboCondition(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    Call ssdcboCondition_Click(0)
    
End Sub

'SQL statement get condition information and populate data grid

Private Sub AddCondition()
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
        
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "select cond_condcode Code, cond_desc Description from CONDITION"
        .CommandText = .CommandText & " where cond_npecode = '" & deIms.NameSpace & "' ORDER BY cond_condcode"
        
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    ssdcboCondition(1).RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboCondition(1).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub

'clear dat fields

Private Sub ClearFields()
    
    optOwn = True
    optPool = True
    
    txtDesc = ""
    lblSecQnty = ""
    txtprimUnit = ""
    txtLeaseComp = ""
    txtprimUnit.Tag = ""
    ssdcboCountry.Text = ""
    
    cboSerialNumb.ListIndex = CB_ERR
    
    ssdcboCondition(1).Text = ""
    ssdcboSubLocation(1).Text = ""
    ssdcboLogicalWHouse(1).Text = ""
    
End Sub

'set text boxse and lable active

Private Sub EnableControls(value As Boolean)
    optOwn.Enabled = value
    'optPool.Enabled = Value
    optLease.Enabled = value
    'optSpecific.Enabled = False
    
    
    txtDesc.Enabled = value
    lblSecQnty.Enabled = value
    txtprimUnit.Enabled = value
    ssdcboCountry.Enabled = value
    
    
    ssdcboCondition(0).Enabled = False
    ssdcboSubLocation(0).Enabled = False
    ssdcboLogicalWHouse(0).Enabled = value
    
    ssdcboCondition(1).Enabled = False
    ssdcboSubLocation(1).Enabled = value
    ssdcboLogicalWHouse(1).Enabled = value
End Sub

'validate data grid
Private Function CheckDetl() As Boolean
Dim l As Long

    
    If rs Is Nothing Then Exit Function
    If rs.State And adStateOpen = adStateClosed Then Exit Function
     
    l = SSTab1.Tab
    SSTab1.Tab = 1
    
    If Len(Trim$(ssdcboCountry.Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboCountry.SetFocus: Exit Function
    Else
        rs!iid_ctry = RTrim$(ssdcboCountry.Columns("Code").Text)
    End If
    
        
'    If Len(Trim$(ssdcboStockType.Text)) = 0 Then
'        MsgBox "Stock Type cannot be left empty":
'        ssdcboStockType.SetFocus: Exit Function
'
'    Else
'        rs!iid_stcktype = RTrim$(ssdcboStockType.Columns("Code").Text)
'
'    End If
    
        
    If Len(Trim$(ssdcboCondition(0).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00331") 'J added
        MsgBox IIf(msg1, "From condition cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboCondition(0).SetFocus: Exit Function
        
    Else
        rs!iid_origcond = RTrim$(ssdcboCondition(0).Columns("Code").Text)
        
    End If
            
    If Len(Trim$(ssdcboSubLocation(0).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00332") 'J added
        MsgBox IIf(msg1, "From Sub-Location cannot be left empty", msg1) 'J modified
        '---------------------------------------------

        ssdcboSubLocation(0).SetFocus: Exit Function
        
    
        Else
        rs!iid_fromsubloca = RTrim$(ssdcboSubLocation(0).Columns("Code").Text)
        
    End If

    If Len(Trim$(ssdcboLogicalWHouse(0).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00333") 'J added
        MsgBox IIf(msg1, "From Logical Warehouse cannot be left empty", msg1) 'J modified
        '---------------------------------------------
            
        ssdcboLogicalWHouse(0).SetFocus: Exit Function
        
    Else
        
        rs!iid_fromlogiware = RTrim$(ssdcboLogicalWHouse(0).Columns("Code").Text)
    End If
        
        
    If Len(Trim$(ssdcboCondition(1).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00333") 'J added
        MsgBox IIf(msg1, "to condition cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        ssdcboCondition(1).SetFocus: Exit Function
    Else
        rs!iid_newcond = RTrim$(ssdcboCondition(0).Columns("Code").Text)
        
    End If
    
        '// To Sub=location
    If Len(Trim$(ssdcboSubLocation(1).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00335") 'J added
        MsgBox IIf(msg1, "To Sub-Location cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        ssdcboSubLocation(1).SetFocus: Exit Function
    Else
        rs!iid_tosubloca = RTrim$(ssdcboSubLocation(1).Columns("Code").Text)
    End If
    
    
    '// To Logical Warehouse
    If Len(Trim$(ssdcboLogicalWHouse(1).Text)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00336") 'J added
        MsgBox IIf(msg1, "To Logical Warehouse cannot be left empty", msg1) 'J modified
        '---------------------------------------------
    
        ssdcboLogicalWHouse(1).SetFocus: Exit Function
        
    Else
        rs!iid_tologiware = RTrim$(ssdcboLogicalWHouse(1).Columns("Code").Text)
    End If
    
    If Len(txtprimUnit) > 0 Then
    
        If IsNumeric(txtprimUnit) Then
            rs!iid_primqty = CDbl(txtprimUnit)
        Else
        
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00336") 'J added
            MsgBox IIf(msg1, "Primary unit is not a valid number", msg1) 'J modified
            '---------------------------------------------
        
            txtprimUnit.SetFocus: Exit Function
        End If
        
    Else
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00338") 'J added
        MsgBox IIf(msg1, "Primary unit cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtprimUnit.SetFocus: Exit Function
    End If
        
    If optSpecific Then
        
        rs!iid_ps = 0
        If Len(Trim$(cboSerialNumb)) = 0 Then
        
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00339") 'J added
            MsgBox IIf(msg1, "Serial number cannot be left empty", msg1) 'J modified
            '---------------------------------------------
        
            cboSerialNumb.SetFocus: Exit Function
        Else
            rs!iid_serl = cboSerialNumb
            
        End If
        
    Else
        rs!iid_ps = 1
        rs!iid_serl = Null
    End If
    
    
    If optLease Then
    
        rs!iid_owle = 0
        
        If Len(Trim$(txtLeaseComp)) = 0 Then
        
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00340") 'J added
            MsgBox IIf(msg1, "Lease Company cannot be left empty", msg1) 'J modified
            '---------------------------------------------
        
            txtLeaseComp.SetFocus: Exit Function
        Else
            rs!iid_leasecomp = Trim$(txtLeaseComp)
            
        End If
        
    Else
         rs!iid_owle = 1
         rs!iid_leasecomp = Null
    End If


    If Len(Trim$(lblSecQnty)) Then
    
        If Not IsNumeric(lblSecQnty) Then
        
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00341") 'J added
            MsgBox IIf(msg1, "Secondary Quantity does not have a valid number", msg1) 'J modified
            '---------------------------------------------
        
            Exit Function
            
        Else
            rs!iid_secoqty = CDbl(lblSecQnty)
        End If
        
    Else
        rs!iid_secoqty = Null
    End If
            
    If Len(Trim$(txtDesc)) Then rs!iid_stckdesc = Trim$(txtDesc)
    'If Len(lblUnitprice) Then rs!iid_unitpric = CDbl(lblUnitprice)
    If Len(lblCommodity) Then rs!iid_stcknumb = Trim$(lblCommodity)
    
    SSTab1.Tab = l
    CheckDetl = True
    
    If Err Then Err.Clear
End Function

Private Sub AssignRsValues()
'    With rs
'        !
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = &HC0FFFF
End Sub


Private Sub Text1_LostFocus()
    Text1.BackColor = &HC0E0FF
End Sub


Private Sub txtLeaseComp_GotFocus()
    txtLeaseComp.BackColor = &HC0FFFF
End Sub


Private Sub txtLeaseComp_LostFocus()
    txtLeaseComp.BackColor = &H80000005
End Sub


'validate primary unit and set data format to 4 decimal

Private Sub txtprimUnit_Change()
On Error Resume Next
Dim db As Double

    If Len(txtprimUnit) > 0 Then
        If Not IsNumeric(txtprimUnit) Then

            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00122") 'J added
            MsgBox IIf(msg1, "Invalid Value", msg1) 'J modified
            '---------------------------------------------
            
            txtprimUnit.SetFocus: Exit Sub
        End If
    End If
            
        
    If Len(txtprimUnit) > 0 Then
    
        If IsNumeric(txtprimUnit) Then
        
            'db = FormatNumber((txtprimUnit), 4)   'M
            db = txtprimUnit
            
            If db < 1 Then
            
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00122") 'J added
            MsgBox IIf(msg1, "Invalid Value", msg1) 'J modified
            '---------------------------------------------
            
            End If
            
            If Len(Trim$(txtprimUnit.Tag)) > 0 Then
            
                If FormatNumber((txtprimUnit.Tag), 0) < db Then
                    txtprimUnit = ""
                    
                    'Modified by Juan (9/15/2000) for Multilingual
                    msg1 = translator.Trans("M00342") 'J added
                    MsgBox IIf(msg1, "Value is too large", msg1) 'J modified
                    '---------------------------------------------
                    
                    txtprimUnit.SetFocus:
                    'txtprimUnit = FormatNumber$(txtprimUnit.Tag, 4)  'M
                    txtprimUnit = txtprimUnit.Tag: Exit Sub
                End If
                
            End If
            
            rs!iid_primqty = db
        End If
        
    Else
        rs!iid_primqty = Null
    End If
        
    txtprimUnit_Validate (True)
    If Err Then Err.Clear
End Sub

'assign values to data field

Private Sub optLease_Click()
On Error Resume Next

    If rs!iid_owle <> 0 Then _
       rs!iid_owle = 0
    
    rs!iid_owle = 0
    
    txtLeaseComp.Enabled = True
    txtLeaseComp.SetFocus
    
    If Err Then Err.Clear
End Sub

'assign data values to data field

Private Sub optOwn_Click()
On Error Resume Next

    If rs!iid_owle <> 1 Then _
       rs!iid_owle = 1
    
    rs!iid_owle = 1
    txtLeaseComp.Enabled = False
    
    If Err Then Err.Clear
End Sub

'assign values to data field

Private Sub optPool_Click()
On Error Resume Next

    If rs!iid_ps <> 1 Then _
        rs!iid_ps = 1
    
    rs!iid_ps = 1
    
    txtprimUnit.Enabled = True
    cboSerialNumb.Enabled = False
    
    Err.Clear
End Sub

'assign data value to data field

Private Sub optSpecific_Click()
On Error Resume Next
    If rs!iid_ps <> 0 Then _
        rs!iid_ps = 0
        
    rs!iid_ps = 0
    cboSerialNumb.Enabled = True
    cboSerialNumb.SetFocus
    
    txtprimUnit.Text = 1
    txtprimUnit.Enabled = False
    Err.Clear
End Sub

Private Sub txtprimUnit_GotFocus()
    txtprimUnit.BackColor = &HC0FFFF
End Sub

Private Sub txtprimUnit_LostFocus()
    txtprimUnit.BackColor = &H80000005
End Sub

'validate data and set data format to 4 decimal

Private Sub txtprimUnit_Validate(Cancel As Boolean)
On Error Resume Next
Dim CompFactor As Double

    
    If Len(Trim$(txtprimUnit)) = 0 Then Exit Sub

    If lblPrimUnit = lblSecUnit Then
        If IsNumeric(txtprimUnit) Then
            lblSecQnty = FormatNumber(txtprimUnit, 2)
        Else
            lblSecQnty = ""
        End If
    Else

        CompFactor = ImsDataX.ComputingFactor(deIms.NameSpace, lblCommodity, deIms.cnIms)

        If CompFactor = 0 Then
            lblSecQnty = FormatNumber(txtprimUnit, 2)
        Else
            lblSecQnty = FormatNumber(txtprimUnit * 10000 / CompFactor, 2)
        End If
    End If
    
    'txtprimUnit = FormatNumber$(txtprimUnit, 4)    'M
    rs!iid_secoqty = lblSecQnty
End Sub

'check data format show message

Private Sub lblSecQnty_Change()
    If Len(lblSecQnty) Then
        If Not IsNumeric(lblSecQnty) Then
        
            'Modified by Juan (9/15/2000) for Multilingual
            msg1 = translator.Trans("M00122") 'J added
            MsgBox IIf(msg1, "Invalid Value", msg1) 'J modified
            '---------------------------------------------
        
        End If
    End If
End Sub

'assign data to recordset

Private Sub AssignDefValues()



    cboSerialNumb.Clear
    If rs Is Nothing Then Exit Sub
    If rs.State And adStateOpen = adStateClosed Then Exit Sub
    
    optOwn = True
    optPool = True
    
    rs!iid_ps = 1
    rs!iid_owle = 1
    rs!iid_compcode = CompCode
    rs!iid_transerl = GetNextSerial
    rs!iid_npecode = deIms.NameSpace
    rs!iid_ware = RTrim$(ssdcboWarehouse.Columns("Code").Text)
    rs!iid_stcknumb = Trim$(ssdbStockInfo.Columns("Commodity").Text)
    rs!iid_stckdesc = Trim$(ssdbStockInfo.Columns("Description").Text)
    
End Sub

'SQL statement get serrial number

Private Function GetNextSerial() As Long
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = New ADODB.Command
    
    If rs Is Nothing Then Exit Function
    If rs.State And adStateOpen = adStateClosed Then Exit Function
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "Select count(*) +  1 serl from INVTISSUEDETL where "
        .CommandText = .CommandText & "iid_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND iid_trannumb = '" & cbo_Transaction & "'"
        
        Set rst = .Execute
        GetNextSerial = rst!Serl
        GetNextSerial = IIf(rs.RecordCount > GetNextSerial, rs.RecordCount, GetNextSerial)
    End With
End Function

'call function and populate data combo

Private Sub AddIssueNumb()
On Error Resume Next

Dim rst As ADODB.Recordset

    Set rst = deIms.rsIssueNumber
    If rst.State And adStateOpen = adStateOpen Then rst.Close
    Call deIms.IssueNumber(deIms.NameSpace, CompCode, "IT")
    
    Call PopuLateFromRecordSet(cbo_Transaction, rst, rst.Fields(0).Name, False)
    
    rst.Close
    Set rst = Nothing
    If Err Then Err.Clear
End Sub

'get crystal report parameter and path

Private Sub BeforePrint()
On Error Resume Next
    MDI_IMS.CrystalReport1.Reset
    MDI_IMS.CrystalReport1.ReportFileName = reportPath & "wareI.rpt"
    MDI_IMS.CrystalReport1.ParameterFields(0) = "transnumb;" & cbo_Transaction & ";TRUE"
    MDI_IMS.CrystalReport1.ParameterFields(1) = "namespace;" & deIms.NameSpace & ";TRUE"
    
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("L00453") 'J added
    MDI_IMS.CrystalReport1.WindowTitle = IIf(msg1 = "", "Internal Transfer", msg1) 'J modified
    Call translator.Translate_Reports("wareI.rpt") 'J added
    Call translator.Translate_SubReports 'J added
    '---------------------------------------------
    
    If Err Then MsgBox Err.Description: Err.Clear
End Sub

'add  recipients to recipient list

Private Sub IMSMail1_OnAddClick(ByVal address As String)
On Error Resume Next

    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        
        rsReceptList.Open
    End If
    
    If Not IsInList(address, "Recipients", rsReceptList) Then _
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))

    Set ssdbRecepientList.DataSource = rsReceptList
    ssdbRecepientList.Columns(0).DataField = "Recipients"
End Sub

'set parameter values, call function send email

Private Sub NavBar1_OnEMailClick()
Dim Filename As String
    BeforePrint
    Call WriteRPTIFile(CreateRpti, Filename)
    Call SendEmailAndFax(rsReceptList, "Recipients", "Internal Transfer", "", Filename)

    Set rsReceptList = Nothing
    Set ssdbRecepientList.DataSource = Nothing
End Sub

'assign parameters values to store procedure

Private Function PutIssueRemarks() As Boolean
On Error Resume Next
Dim cmd As ADODB.Command

    Set cmd = deIms.Commands("InvtIssuetRem_Insert")
    
    cmd.parameters("@LineNumb") = 1
    cmd.parameters("Remarks") = txtRemarks
    cmd.parameters("@TranNumb") = Transnumb
    cmd.parameters("@CompanyCode") = CompCode
    cmd.parameters("@NameSpace") = deIms.NameSpace
    cmd.parameters("@User") = CurrentUser
    cmd.parameters("@WhareHouse") = ssdcboWarehouse.Columns("Code").Text
    
    Call cmd.Execute(0, , adExecuteNoRecords)
    
    If IsNull(cmd.parameters(0)) Then
        PutIssueRemarks = False
    ElseIf IsEmpty(cmd.parameters(0)) Then
        PutIssueRemarks = False
    Else
        PutIssueRemarks = cmd.parameters(0) = 0
    End If
    
    If Err Then Err.Clear
End Function

'get parameter values for print crystal

Private Function CreateRpti() As RPTIFileInfo

    With CreateRpti
        ReDim .parameters(1)
        .ReportFileName = reportPath & "wareI.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("wareI.rpt") 'J added
        '---------------------------------------------
        
        .parameters(0) = "transnumb=" & cbo_Transaction
        .parameters(1) = "namespace=" & deIms.NameSpace
    
    End With

End Function
