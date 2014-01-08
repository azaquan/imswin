VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmBaseToBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warehouse to Warehouse"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   9750
   Tag             =   "02040600"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   3360
      TabIndex        =   59
      Top             =   5880
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "BaseToBase.frx":0000
      NewVisible      =   0   'False
      EMailVisible    =   -1  'True
      CloseEnabled    =   0   'False
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      NextEnabled     =   0   'False
      LastEnabled     =   0   'False
      FirstEnabled    =   0   'False
      PreviousEnabled =   0   'False
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   180
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Issue"
      TabPicture(0)   =   "BaseToBase.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblValidateBy"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDesc(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUser"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDesc(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDesc(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblType"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDesc(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDesc(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDesc(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDesc(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDesc(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ssdcboCompany"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ssdcboWarehouse"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ssdcboLocation"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ssdbStockInfo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbo_Transaction"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Line Items to be Issued"
      TabPicture(1)   =   "BaseToBase.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtprimUnit"
      Tab(1).Control(1)=   "cboSerialNumb"
      Tab(1).Control(2)=   "Frame2(1)"
      Tab(1).Control(3)=   "txtLeaseComp"
      Tab(1).Control(4)=   "optOwn"
      Tab(1).Control(5)=   "optLease"
      Tab(1).Control(6)=   "txtDesc"
      Tab(1).Control(7)=   "ssdcboCountry"
      Tab(1).Control(8)=   "Frame2(0)"
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(10)=   "lblCurrRec"
      Tab(1).Control(11)=   "lblRecCount"
      Tab(1).Control(12)=   "Label3"
      Tab(1).Control(13)=   "Label1"
      Tab(1).Control(14)=   "lblDesc(8)"
      Tab(1).Control(15)=   "lblPrimUnit"
      Tab(1).Control(16)=   "lblSecUnit"
      Tab(1).Control(17)=   "lblSecQnty"
      Tab(1).Control(18)=   "lblDesc(21)"
      Tab(1).Control(19)=   "lblDesc(20)"
      Tab(1).Control(20)=   "lblCommodity"
      Tab(1).Control(21)=   "lblCurrencyValu"
      Tab(1).Control(22)=   "lblCurrency"
      Tab(1).Control(23)=   "lblUnitprice"
      Tab(1).Control(24)=   "lblDesc(11)"
      Tab(1).Control(25)=   "lblDesc(14)"
      Tab(1).Control(26)=   "lblDesc(16)"
      Tab(1).Control(27)=   "lblDesc(17)"
      Tab(1).Control(28)=   "lblDesc(24)"
      Tab(1).Control(29)=   "lblDesc(25)"
      Tab(1).Control(30)=   "lblDesc(28)"
      Tab(1).Control(31)=   "lblDesc(22)"
      Tab(1).Control(32)=   "lblDesc(26)"
      Tab(1).Control(33)=   "lblDesc(23)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Remarks"
      TabPicture(2)   =   "BaseToBase.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRemarks"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recipients"
      TabPicture(3)   =   "BaseToBase.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1"
      Tab(3).Control(1)=   "cmd_Remove"
      Tab(3).Control(2)=   "cmd_Add"
      Tab(3).Control(3)=   "ssdbRecepientList"
      Tab(3).Control(4)=   "lbl_Recipients"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   600
         TabIndex        =   70
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox txtprimUnit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68760
         TabIndex        =   10
         Top             =   1620
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   -74640
         ScaleHeight     =   3015
         ScaleWidth      =   8535
         TabIndex        =   62
         Top             =   2160
         Width           =   8535
         Begin ImsMailVB.Imsmail Imsmail1 
            Height          =   3015
            Left            =   120
            TabIndex        =   69
            Top             =   0
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5318
         End
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74595
         TabIndex        =   61
         Top             =   1785
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74595
         TabIndex        =   60
         Top             =   1455
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
         Height          =   4935
         Left            =   -74880
         MaxLength       =   7000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   420
         Width           =   9015
      End
      Begin VB.ComboBox cboSerialNumb 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68520
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "To"
         Height          =   1275
         Index           =   1
         Left            =   -70320
         TabIndex        =   36
         Top             =   2640
         Width           =   4395
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboSubLocation 
            Bindings        =   "BaseToBase.frx":008C
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
            stylesets(0).Picture=   "BaseToBase.frx":00CC
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
            stylesets(1).Picture=   "BaseToBase.frx":00E8
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
            Bindings        =   "BaseToBase.frx":0104
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
            stylesets(0).Picture=   "BaseToBase.frx":0144
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
            stylesets(1).Picture=   "BaseToBase.frx":0160
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
            Bindings        =   "BaseToBase.frx":017C
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
            stylesets(0).Picture=   "BaseToBase.frx":01BC
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
            stylesets(1).Picture=   "BaseToBase.frx":01D8
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
         TabIndex        =   5
         Top             =   1260
         Width           =   1035
      End
      Begin VB.OptionButton optLease 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease"
         Height          =   315
         Left            =   -71640
         TabIndex        =   25
         Top             =   1260
         Width           =   1215
      End
      Begin VB.ComboBox cbo_Transaction 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   480
         Width           =   2640
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbStockInfo 
         Height          =   3015
         Left            =   300
         TabIndex        =   22
         Top             =   2400
         Width           =   8775
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
         stylesets(0).Picture=   "BaseToBase.frx":01F4
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
         stylesets(1).Picture=   "BaseToBase.frx":0210
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
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
         Columns(1).Width=   10742
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).HeadStyleSet=   "ColHeader"
         Columns(1).StyleSet=   "RowFont"
         Columns(2).Width=   1826
         Columns(2).Caption=   "PU Qty"
         Columns(2).Name =   "ReqQnty"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   5
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
         _ExtentX        =   15478
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboLocation 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1200
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
         stylesets(0).Picture=   "BaseToBase.frx":022C
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
         stylesets(1).Picture=   "BaseToBase.frx":0248
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboWarehouse 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Tag             =   "1"
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
         stylesets(0).Picture=   "BaseToBase.frx":0264
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
         stylesets(1).Picture=   "BaseToBase.frx":0280
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
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -73320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   4320
         Width           =   7335
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCountry 
         Bindings        =   "BaseToBase.frx":029C
         Height          =   315
         Left            =   -73080
         TabIndex        =   4
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
         stylesets(0).Picture=   "BaseToBase.frx":02DC
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
         stylesets(1).Picture=   "BaseToBase.frx":02F8
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
            Bindings        =   "BaseToBase.frx":0314
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   8
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
            stylesets(0).Picture=   "BaseToBase.frx":0354
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
            stylesets(1).Picture=   "BaseToBase.frx":0370
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
            Bindings        =   "BaseToBase.frx":038C
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   7
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
            stylesets(0).Picture=   "BaseToBase.frx":03CC
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
            stylesets(1).Picture=   "BaseToBase.frx":03E8
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
            Bindings        =   "BaseToBase.frx":0404
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   9
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
            stylesets(0).Picture=   "BaseToBase.frx":0444
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
            stylesets(1).Picture=   "BaseToBase.frx":0460
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
         TabIndex        =   39
         Top             =   3840
         Width           =   2715
         Begin VB.OptionButton optSpecific 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optPool 
            Alignment       =   1  'Right Justify
            Caption         =   "Pool"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1035
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   1605
         Left            =   -72960
         TabIndex        =   63
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
         Left            =   1860
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
         Left            =   1710
         TabIndex        =   76
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Search Field"
         Height          =   255
         Index           =   0
         Left            =   1930
         TabIndex        =   75
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblCurrRec 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67920
         TabIndex        =   74
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblRecCount 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66720
         TabIndex        =   73
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "of"
         Height          =   255
         Left            =   -67080
         TabIndex        =   72
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Viewing Record"
         Height          =   255
         Left            =   -69720
         TabIndex        =   71
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Primary"
         Height          =   195
         Index           =   8
         Left            =   -70200
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblPrimUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67320
         TabIndex        =   67
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label lblSecUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -67320
         TabIndex        =   66
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lblSecQnty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68760
         TabIndex        =   65
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74595
         TabIndex        =   64
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblDesc 
         Caption         =   "Serial Number"
         Height          =   315
         Index           =   21
         Left            =   -70320
         TabIndex        =   40
         Top             =   4020
         Width           =   1815
      End
      Begin VB.Label lblDesc 
         Caption         =   "Pool / Specific"
         Height          =   315
         Index           =   20
         Left            =   -74880
         TabIndex        =   37
         Top             =   4020
         Width           =   1500
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
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Caption         =   "Commodity"
         Height          =   315
         Index           =   11
         Left            =   -74880
         TabIndex        =   58
         Top             =   540
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Country of Origin"
         Height          =   315
         Index           =   14
         Left            =   -74880
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   1620
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Quantities"
         Height          =   195
         Index           =   24
         Left            =   -68760
         TabIndex        =   55
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label lblDesc 
         Caption         =   "Units"
         Height          =   195
         Index           =   25
         Left            =   -67320
         TabIndex        =   54
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label lblDesc 
         Caption         =   "Description"
         Height          =   315
         Index           =   28
         Left            =   -74880
         TabIndex        =   53
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         Caption         =   "Currency Value"
         Height          =   315
         Index           =   22
         Left            =   -74880
         TabIndex        =   52
         Top             =   2290
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         Caption         =   "Unit Price"
         Height          =   315
         Index           =   26
         Left            =   -74880
         TabIndex        =   51
         Top             =   1950
         Width           =   1800
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Secondary"
         Height          =   195
         Index           =   23
         Left            =   -70200
         TabIndex        =   50
         Top             =   2040
         Width           =   1485
      End
      Begin VB.Label lblDesc 
         Caption         =   "Inventory"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   48
         Top             =   840
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "Transac #"
         Height          =   315
         Index           =   2
         Left            =   4800
         TabIndex        =   47
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         Caption         =   "Date"
         Height          =   315
         Index           =   5
         Left            =   4800
         TabIndex        =   46
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6480
         TabIndex        =   20
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         Caption         =   "Issue To"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   45
         Top             =   1200
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "Type"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   44
         Top             =   1580
         Width           =   1600
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRANSFER ISSUE"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1860
         TabIndex        =   19
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label lblDesc 
         Caption         =   "Company"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         Caption         =   "User"
         Height          =   315
         Index           =   1
         Left            =   4800
         TabIndex        =   42
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Validated By"
         Height          =   315
         Index           =   9
         Left            =   4800
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblValidateBy 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6480
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
   End
End
Attribute VB_Name = "frmBaseToBase"
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
Dim sql, LW, SL, SC As String
    If justDBLCLICK Then
        justDBLCLICK = False
    Else
        Screen.MousePointer = 11
        LW = ssdcboLogicalWHouse(0).Columns("Code").text
        SL = ssdcboSubLocation(0).Columns("Code").text
        SC = ssdcboCondition(0).Columns("Code").text
        
        If Not (LW = "" Or SL = "" Or SC = "") Then
            sql = "SELECT sum(qs4_primqty) AS PrimaryQTY FROM QTYST4 WHERE " _
                & "qs4_compcode = '" + CompCode + "' And " _
                & "qs4_npecode = '" + deIms.NameSpace + "' AND " _
                & "qs4_ware = '" + ssdcboWarehouse.Columns("Code").text + "' AND " _
                & "qs4_stcknumb = '" + lblCommodity + "' AND " _
                & "qs4_logiware = '" + LW + "' AND " _
                & "qs4_subloca = '" + SL + "' AND " _
                & "qs4_cond = '" + SC + "'"
            Set qty = New ADODB.Recordset
            qty.Open sql, deIms.cnIms, adOpenForwardOnly
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
            rs!iid_fromlogiware = ssdcboLogicalWHouse(0).Columns("Code").Value
            rs!iid_tologiware = ssdcboLogicalWHouse(1).Columns("Code").Value
            rs!iid_fromsubloca = ssdcboSubLocation(0).Columns("Code").Value
            rs!iid_tosubloca = ssdcboSubLocation(1).Columns("Code").Value
            rs!iid_origcond = ssdcboCondition(0).Columns("Code").Value
            rs!iid_newcond = ssdcboCondition(1).Columns("Code").Value
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

Public Sub DisplayRecord()
Dim qty1, qty2
Dim WareHouse As String

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

    WareHouse = ssdcboWarehouse.Columns("Code").text

    'FillCombos
    
    Call AddFromLogicalWharehouse(Get_LogicalWarehouse_FromQTYST(deIms.NameSpace, _
                                  lblCommodity, WareHouse, CompCode, deIms.cnIms))
        
    Call FindInGrid(ssdcboLogicalWHouse(0), rs("iid_fromlogiware") & "", True, 1)
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
    
    
    
    
    ssdcboCountry.text = ssdcboCountry.Columns(0).text
    ssdcboCondition(0).text = ssdcboCondition(0).Columns(0).text
    ssdcboSubLocation(0).text = ssdcboSubLocation(0).Columns(0).text
    ssdcboLogicalWHouse(0).text = ssdcboLogicalWHouse(0).Columns(0).text
    
    ssdcboCondition(1).text = ssdcboCondition(1).Columns(0).text
    ssdcboSubLocation(1).text = ssdcboSubLocation(1).Columns(0).text
    ssdcboLogicalWHouse(1).text = ssdcboLogicalWHouse(1).Columns(0).text
    
    rs("iid_primqty") = qty1
    rs("iid_secoqty") = qty2
    
    
    txtprimUnit = FormatNumber(rs("iid_primqty") & "", 0)
    lblSecQnty = FormatNumber(rs("iid_secoqty") & "", 2)
    
    lblUnitprice = FormatNumber(rs("iid_unitpric") & "", 2)
    
    cboSerialNumb.ListIndex = IndexOf(cboSerialNumb, rs("iid_serl") & "")
    
    
    
    Dim SU As String
    Dim PU As String
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
    
End Sub

'unlock transcation combo

Private Sub cbo_Transaction_DropDown()
cbo_Transaction.locked = False
End Sub

'call function add issue number

Private Sub cbo_Transaction_GotFocus()
    cbo_Transaction.BackColor = &HC0FFFF
    NavBar1.SaveEnabled = False
    NavBar1.CancelEnabled = False
End Sub

'if not add new status, lock transcation combo

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


'call function add new recipient to list

Private Sub cmd_Add_Click()
    Imsmail1.AddCurrentRecipient
End Sub

'call function remove a recipient from recipient list

Private Sub cmd_Remove_Click()
On Error Resume Next
    rsReceptList.Delete
    rsReceptList.Update
    
    If Err Then Err.Clear
End Sub

'get form combos recordsets and set navbar buttom

Private Sub Form_Load()
Dim np As String
Dim FCompany As String
Dim cn As ADODB.Connection
Dim i As Integer

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frmBaseToBase")
    '------------------------------------------

    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    NavBar1.SaveEnabled = SaveEnabled
    
    For i = 1 To 2
        SSTab1.TabVisible(i) = SaveEnabled
    Next


    np = deIms.NameSpace
    Set cn = deIms.cnIms
    fm = mdVisualization
    ssdcboLocation.RemoveAll
    
    Requery = True
    FCompany = GetCompany(np, "PE", cn)
    CompCode = GetCompanyCode(np, FCompany, cn)
    ssdcboCompany.DataMode = ssDataModeAddItem
    
    AddCompanies
    AddCountries
    AddSubLocations
    AddLogicalWhareHouse
    EnableControls (False)
    Set InvtIss = New imsWhareIssue
    
    AddCondition
    Imsmail1.NameSpace = deIms.NameSpace
    
    AddIssueNumb
    
    ' IMSMail1.Connected = True 'M
    Imsmail1.SetActiveConnection deIms.cnIms  'M
    Imsmail1.Language = Language 'M
    'Call AddStockType(GetStockType(np, cn))
    ssdbStockInfo.FieldSeparator = Chr(1)
    
    Call DisableButtons(Me, NavBar1)
    SaveEnabled = NavBar1.SaveEnabled
    NavBar1.CloseEnabled = True
    frmBaseToBase.Caption = frmBaseToBase.Caption + " - " + frmBaseToBase.Tag
    SSTab1.TabVisible(3) = False
    NavBar1.EMailVisible = False
    
    cbo_Transaction.locked = False
    cbo_Transaction.Enabled = True
End Sub

'SQL statement get location information for combo location
'and populate combo

Public Sub AddSubLocations()
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

'populate warehouses combo

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

'populate recordset for location data grid

Private Sub AddLocations(rst As ADODB.Recordset)

    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    ssdcboLocation.RemoveAll
    Do While ((Not rst.EOF))
    
        ssdcboLocation.AddItem rst!loc_name & "" & ";" & rst!loc_locacode & ""
        rst.MoveNext
    Loop
    
CleanUp:
    rst.Close
    Set rst = Nothing
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

'SQL statement get countries recordset and populate data grid

Public Sub AddCountries()
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

'unload from  free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim closing
    
    If fm <> mdVisualization Then
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
    If Err Then Err.Clear

    If open_forms <= 5 Then ShowNavigator

End Sub

Private Sub Imsmail1_GotFocus()
Call HighlightBackground(Imsmail1)
End Sub

Private Sub Imsmail1_LostFocus()
Call NormalBackground(Imsmail1)
End Sub

'call function to add recepient torecepient list

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

Private Sub NavBar1_OnCancelClick()
On Error Resume Next
    Call MoveRecord(-1)
    backpass = True
    optSpecific.Value = False
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

'call function and set parameter then send email

Private Sub NavBar1_OnEMailClick()
Dim FileName As String
BeforePrint
    Call WriteRPTIFile(CreateRpti, FileName)
    Call SendEmailAndFax(rsReceptList, "Recipients", "Transfer Issue", "", FileName)

    Set rsReceptList = Nothing
    Set ssdbRecepientList.DataSource = Nothing
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

'before save validate data format,set data to recordset

Private Sub NavBar1_OnSaveClick()
NavBar1.SaveEnabled = False
MDI_IMS.StatusBar1.Panels(1).text = "Saving"
Screen.MousePointer = 11

On Error Resume Next

Dim retval As Boolean
Dim np As String
Dim ToWH As String
Dim FromWH As String
Dim PrimUnit As Double

Dim SecUnit As Double
Dim StockNumb As String
Dim cn As ADODB.Connection


    If rs Is Nothing Then Exit Sub
    cbo_Transaction.ListIndex = CB_ERR
    If CheckMasterFields And CheckDetl Then
    Screen.MousePointer = 11
    
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
        
        deIms.cnIms.BeginTrans
        NavBar1.SaveEnabled = False
        MDI_IMS.StatusBar1.Panels(1).text = "Beginning Transaction"
        Call BeginTransaction(deIms.cnIms)
        
        retval = PutInvtIssue
        
        'doevents
        If retval = False Then GoTo RollBack
        Screen.MousePointer = 11
        rs.MoveFirst
        MDI_IMS.StatusBar1.Panels(1).text = "Saving Details"
        Do While Not rs.EOF
            retval = PutDataInsert
            Screen.MousePointer = 11
            If retval = False Then GoTo RollBack
            
            Set cn = deIms.cnIms
            np = deIms.NameSpace
            FromWH = rs!iid_ware
            StockNumb = rs!iid_stcknumb
            PrimUnit = rs!iid_primqty
            ToWH = ssdcboLocation.Columns("Code").text
            SecUnit = IIf(IsNull(rs!iid_secoqty), 0, rs!iid_secoqty)
            
            retval = Update_Sap(np, CompCode, StockNumb, ToWH, PrimUnit, 1, rs!iid_unitpric, rs!iid_newcond, CurrentUser, cn)
            retval = retval And Quantity_In_stock1_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_stckdesc, CurrentUser, cn)
            retval = retval And Quantity_In_stock2_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, CurrentUser, cn)
            retval = retval And Quantity_In_stock3_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, CurrentUser, cn)
            retval = retval And Quantity_In_stock4_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, CurrentUser, cn)
            
            'doevents
            If rs!iid_ps Then
                retval = retval And Quantity_In_stock5_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, Transnumb, rs!iid_transerl, rs!iid_ware, "TI", CompCode, FromWH, Transnumb, CompCode, rs!iid_transerl, CurrentUser, cn)
            Else
                 retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, rs!iid_serl, CurrentUser, cn)
                 retval = retval And Quantity_In_stock7_Insert(np, CompCode, StockNumb, ToWH, PrimUnit, SecUnit, rs!iid_tologiware, rs!iid_tosubloca, rs!iid_newcond, Transnumb, FromWH, rs!iid_transerl, rs!iid_ware, "TI", CompCode, Transnumb, CompCode, rs!iid_transerl, rs!iid_serl, CurrentUser, cn)
            
            End If
            
            If retval = False Then GoTo RollBack
            Screen.MousePointer = 11
            SecUnit = SecUnit * -1
            PrimUnit = PrimUnit * -1
            retval = retval And Quantity_In_stock1_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_stckdesc, CurrentUser, cn)
            retval = retval And Quantity_In_stock2_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, CurrentUser, cn)
            retval = retval And Quantity_In_stock3_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, CurrentUser, cn)
            retval = retval And Quantity_In_stock4_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, CurrentUser, cn)
            
            'doevents
            If rs!iid_ps Then
                retval = retval And Quantity_In_stock5_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, Transnumb, rs!iid_transerl, rs!iid_ware, "TI", CompCode, ToWH, Transnumb, CompCode, rs!iid_transerl, CurrentUser, cn)
            Else
               'Modified by Muzammil 08/12/00
               'Reason - it was passing rs!ird_transerl instead of rs!ird_serl
               retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, rs!iid_serl, CurrentUser, cn)
'                 retval = retval And Quantity_In_stock6_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, rs!iid_transerl, CurrentUser, cn)
                 retval = retval And Quantity_In_stock7_Insert(np, CompCode, StockNumb, FromWH, PrimUnit, SecUnit, rs!iid_fromlogiware, rs!iid_fromsubloca, rs!iid_newcond, Transnumb, ToWH, rs!iid_transerl, rs!iid_ware, "TI", CompCode, Transnumb, CompCode, rs!iid_transerl, rs!iid_serl, CurrentUser, cn)
            
            End If
            
            If retval = False Then GoTo RollBack
            rs.MoveNext
        Loop
        
        'doevents
        
        'Modified by Muzammil 08/11/00
        'Reason - VBCRLFs before the text would block Email Generation.
          
          MDI_IMS.StatusBar1.Panels(1).text = "Saving Remarks"
          Do While InStr(1, txtRemarks, vbCrLf) = 1                   'M
             txtRemarks = Mid(txtRemarks, 3, Len(txtRemarks))         'M
          Loop                                                        'M
             txtRemarks = LTrim$(txtRemarks)                          'M
        
        Screen.MousePointer = 11
                
        PutIssueRemarks
        NavBar1.SaveEnabled = SaveEnabled
        If retval Then deIms.cnIms.CommitTrans
        If retval Then Call CommitTransaction(deIms.cnIms)
        If retval Then Call CommitTransaction(deIms.cnIms)
        
        Screen.MousePointer = 11
        MDI_IMS.StatusBar1.Panels(1).text = "Restoring"
        Call cbo_Transaction.AddItem(Transnumb, cbo_Transaction.ListCount)
        Screen.MousePointer = 11
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, Transnumb)
        Screen.MousePointer = 11
        
        'doevents
        BeforePrint
       ' Call SendWareHouseMessage(deIms.NameSpace, "Automatic Distribution", _
                                 lblType, deIms.cnIms, CreateRpti)
        
        On Error Resume Next
        lblUser = ""
        lblDate = ""
        
        
        ssdcboLocation.text = ""
        Call EnableControls(False)
        
        Requery = True
        rs.CancelBatch
        Call ClearFields
        fm = mdVisualization
        rs.Close: Set rs = Nothing
        
        ssdbStockInfo.RemoveAll
        ssdcboWarehouse.text = ""
        
        Screen.MousePointer = 0
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00018") 'J added
        MsgBox IIf(msg1 = "", "Please note that your transaction # is ", msg1 + " ") & Transnumb 'J modified
        '---------------------------------------------
    End If
    
    If Err Then Err.Clear
    MDI_IMS.StatusBar1.Panels(1).text = ""
    ssdcboCompany.Enabled = True
    ssdcboWarehouse.Enabled = True
    ssdcboLocation.Enabled = True
    'AddIssueNumb
    cbo_Transaction.SetFocus
    Unload FrmShowApproving
    txtRemarks = ""
    Screen.MousePointer = 0
    Exit Sub
    
    
RollBack:
    deIms.cnIms.RollbackTrans
    NavBar1.SaveEnabled = SaveEnabled
    Call RollbackTransaction(deIms.cnIms)
    Call RollbackTransaction(deIms.cnIms)
    Screen.MousePointer = 0
    MDI_IMS.StatusBar1.Panels(1).text = ""
    Unload FrmShowApproving
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

'get stock information assign data to data text boxes

Private Sub ssdbStockInfo_DblClick()
Screen.MousePointer = 0

Dim rst As ADODB.Recordset
Dim WareHouse As String
Dim SU As String
Dim PU As String

    If Not SaveEnabled Then Exit Sub
    
    If ssdbStockInfo.Rows < 1 Then Exit Sub
    Screen.MousePointer = 11
    cbo_Transaction.ListIndex = CB_ERR
    Screen.MousePointer = 11
    If ssdcboLocation.text = "" Then
        Screen.MousePointer = 0
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00328") 'J added
        MsgBox IIf(msg1 = "", "Issue to cannot be empty", msg1) 'J modified
        '---------------------------------------------
        
        Exit Sub
    End If
    Screen.MousePointer = 11
    beginning = True
    justDBLCLICK = True
    
    If Not (rs Is Nothing) Then
        If rs.RecordCount > 0 Then
            If Not (Requery) Then If CheckDetl = False Then Exit Sub
        End If
    End If
    
    fm = mdCreation
    txtprimUnit.Tag = ""
    
    SSTab1.Tab = 1
    
    rs.AddNew
    Call MoveRecord(mtLast)
    Screen.MousePointer = 11
    lblRecCount = rs.RecordCount
    lblCurrRec = rs.AbsolutePosition
    
    ClearFields
    AssignDefValues
    
    'doevents
    Call EnableControls(True)
    Screen.MousePointer = 11
    WareHouse = ssdcboWarehouse.Columns("Code").text
    lblCommodity = ssdbStockInfo.Columns("Commodity").text
    txtDesc = ssdbStockInfo.Columns("Description").text
    Screen.MousePointer = 11
    'doevents
    ssdcboCountry.Value = "USA"
    Call FindInGrid(ssdcboCountry, "USA", True, 1)
    Call GetStockUnit(deIms.NameSpace, lblCommodity, PU, SU, deIms.cnIms)
    Screen.MousePointer = 11
    SU = LCase$(SU)
    PU = LCase$(PU)
    
    If SU = PU Then
        lblPrimUnit = deIms.UnitDescription(PU)
        lblSecUnit = lblPrimUnit
    Else
        lblPrimUnit = deIms.UnitDescription(PU)
        lblSecUnit = deIms.UnitDescription(SU)
    End If
    Screen.MousePointer = 11
    'doevents
    Call AddFromLogicalWharehouse(Get_LogicalWarehouse_FromQTYST(deIms.NameSpace, _
                                  lblCommodity, WareHouse, CompCode, deIms.cnIms))
    Screen.MousePointer = 11
    'txtprimUnit = IIf(txtprimUnit.Tag > 0, txtprimUnit.Tag, "")
    Screen.MousePointer = 1
    Call txtprimUnit_Validate(False)
    If IsNumeric(txtprimUnit) Then
        txtprimUnit = FormatNumber$(txtprimUnit, 0)
    Else
        txtprimUnit = ""
    End If
    
    Screen.MousePointer = 11
    'Modified by Muzammil 08/15/00
   'Reason - Did not work good for records which are not in the first 9(The first set
   'which it displays)
    'Call ssdbStockInfo.RemoveItem(ssdbStockInfo.Row)' M
    Call ssdbStockInfo.RemoveItem(ssdbStockInfo.AddItemRowIndex(ssdbStockInfo.Bookmark)) 'M

    ssdcboCompany.Enabled = False
    ssdcboCompany.BackColor = &H80000005
    ssdcboWarehouse.Enabled = False
    ssdcboWarehouse.BackColor = &H80000005
    ssdcboLocation.Enabled = False
    ssdcboLocation.BackColor = &H80000005
    
    Screen.MousePointer = 0
End Sub

'call function get loaction and warehouse data

Private Sub ssdcboCompany_Click()
    ssdbStockInfo.RemoveAll
    cbo_Transaction.ListIndex = CB_ERR
    CompCode = ssdcboCompany.Columns("Code").text
    
    ssdcboWarehouse = ""
    ssdcboLocation = ""
    ssdcboLocation.RemoveAll
    Call AddLocations(GetLocation(deIms.NameSpace, "BASE", CompCode, deIms.cnIms))
    ssdcboWarehouse.RemoveAll
    Call AddWhareHouses(GetLocation(deIms.NameSpace, "BASE", CompCode, deIms.cnIms))
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

'fill data to data grids, get serail number

Private Sub ssdcboCondition_Click(Index As Integer)
Dim l As Double
Dim rst As ADODB.Recordset
Dim work As Boolean

    If Index = 0 Then
        l = GetSapValue(deIms.NameSpace, CompCode, lblCommodity, ssdcboWarehouse.Columns("Code").text, _
                    ssdcboCondition(0).Columns("Code").text, deIms.cnIms)
                    

        
        rs!iid_unitpric = l
        lblUnitprice = FormatNumber$(l, 0)
        ssdcboCondition(1).text = ssdcboCondition(0).text
        Call FindInGrid(ssdcboLogicalWHouse(0), ssdcboLogicalWHouse(0).Columns("Code").text, True, 1)
        Call FindInGrid(ssdcboSubLocation(0), ssdcboSubLocation(0).Columns("Code").text, True, 1)
        Call FindInGrid(ssdcboCondition(0), ssdcboCondition(0).Columns("Code").text, True, 1)
        Call FindInGrid(ssdcboCondition(1), ssdcboCondition(0).Columns("Code").text, True, 1)
        
        Call deIms.GetItemCount(CompCode, deIms.NameSpace, ssdcboWarehouse.Columns("Code").text, _
                               lblCommodity, ssdcboLogicalWHouse(0).Columns("Code").text, _
                                ssdcboSubLocation(0).Columns("Code").text, ssdcboCondition(0).Columns("Code").text)
                               
        If rs Is Nothing Then
            work = True
        Else
            If rs.RecordCount > 0 Then
                If IsNull(rs!iid_primqty) Then
                    work = True
                Else
                    work = False
                End If
            Else
                work = True
            End If
        End If
        
        txtprimUnit.Tag = deIms.rsGetItemCount.Fields(0).Value & ""
        If work Then
            If deIms.rsGetItemCount.RecordCount > 0 Then
                txtprimUnit = txtprimUnit.Tag
            End If
        Else
            txtprimUnit = FormatNumber(rs!iid_primqty, 0)
        End If
        
        deIms.rsGetItemCount.Close
        cboSerialNumb.Clear
        Set rst = Get_SerialNumberFromStockNumber(deIms.NameSpace, lblCommodity, _
                                            ssdcboWarehouse.Columns("Code").text, _
                                            CompCode, ssdcboLogicalWHouse(0).Columns("Code").text, _
                                            ssdcboSubLocation(0).Columns("Code").text, ssdcboCondition(0).Columns("Code").text, deIms.cnIms)

        ssdcboCondition(1).Enabled = False
        Call PopuLateFromRecordSet(cboSerialNumb, rst, rst.Fields(0).Name, True)
        
        optSpecific.Enabled = cboSerialNumb.ListCount
        cboSerialNumb.Enabled = cboSerialNumb.ListCount
        
        'FG 8/15 if several serial, make qty = 1
        If optSpecific.Enabled Then
            txtprimUnit = 1
            txtprimUnit.Tag = 1
            txtprimUnit_Validate (True)
        End If
        
        optSpecific = cboSerialNumb.Enabled
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

'assign data to lables and set combo to empty position

Private Sub ssdcboLocation_Click()

    lblDate = Date
    lblUser = CurrentUser
    cbo_Transaction.ListIndex = CB_ERR
    
    ssdcboLocation.SelLength = 0
    ssdcboLocation.SelStart = 0
    If ssdcboWarehouse = ssdcboLocation Then
        MsgBox "From and To Well can not be the same Site"
        ssdcboLocation = ""
    End If

End Sub

Private Sub ssdcboLocation_GotFocus()
    ssdcboLocation.BackColor = &HC0FFFF
End Sub


Private Sub ssdcboLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboLocation.DroppedDown = True
    End If
End Sub


'set data grid can not enter character

Private Sub ssdcboLocation_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then KeyAscii = 0
End Sub

'fill data to sub location data grid

Private Sub ssdcboLogicalWHouse_Click(Index As Integer)
    If Index = 0 Then
    
        Call AddFromSublocation(Get_SubLocation_FromQTYST(deIms.NameSpace, lblCommodity, _
                                ssdcboWarehouse.Columns("Code").text, CompCode, _
                                ssdcboLogicalWHouse(0).Columns("Code").text, deIms.cnIms))

        'doevents: 'doevents: 'doevents
        Call ssdcboSubLocation_Click(0)
        ssdcboSubLocation(0).Enabled = True
        Call refreshQTY
    End If
End Sub

Private Sub ssdcboLogicalWHouse_GotFocus(Index As Integer)
    ssdcboLogicalWHouse(Index).BackColor = &HC0FFFF
End Sub


Private Sub ssdcboLogicalWHouse_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboLogicalWHouse(Index).DroppedDown = True
    End If
End Sub

Private Sub ssdcboLogicalWHouse_LostFocus(Index As Integer)
    ssdcboLogicalWHouse(Index).BackColor = &H80000005
End Sub

'call function get condition data and fill data grid

Private Sub ssdcboSubLocation_Click(Index As Integer)
Dim WareHouse As String
Dim LWareHouse As String

    If Index = 0 Then
        WareHouse = ssdcboWarehouse.Columns("Code").text
        LWareHouse = ssdcboLogicalWHouse(0).Columns("Code").text
        If Len(Trim$(ssdcboLogicalWHouse(0).text)) = 0 Then Exit Sub
        Call AddFromCondition(Get_Condition_FromQTYST(deIms.NameSpace, lblCommodity, WareHouse, CompCode, LWareHouse, ssdcboSubLocation(0).Columns("Code").text, deIms.cnIms))
    
        ssdcboCondition(0).Enabled = True
        Call refreshQTY
    End If
    'doevents: 'doevents
End Sub

Private Sub ssdcboSubLocation_GotFocus(Index As Integer)
    ssdcboSubLocation(Index).BackColor = &HC0FFFF
End Sub


Private Sub ssdcboSubLocation_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboLogicalWHouse(Index).DroppedDown = True
    End If
End Sub

Private Sub ssdcboSubLocation_LostFocus(Index As Integer)
    ssdcboSubLocation(Index).BackColor = &H80000005
End Sub

'call function get stock information and fill data grid

Private Sub ssdcboWarehouse_Click()
    Text1 = ""
    cbo_Transaction.ListIndex = CB_ERR
    Call AddStockInfo(GetStockInformation(deIms.NameSpace, ssdcboWarehouse.Columns("Code").text, CompCode, deIms.cnIms))
    ssdcboWarehouse.SelStart = 0
    ssdcboWarehouse.SelLength = 0
    If ssdcboWarehouse = ssdcboLocation Then
        MsgBox "Inventory and Issue to can not be the same"
        ssdcboWarehouse = ""
    End If
End Sub

'fill data grid

Public Sub AddStockInfo(rst As ADODB.Recordset)
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

'assign data to lable

Public Sub AssignInvt()
    With InvtIss
        lblUser = .User
        lblDate = .TransactionDate
        
        
    End With
End Sub

'SQL statement get logical warehouse recordset

Public Sub AddLogicalWhareHouse()
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

'set parameter values for save data

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
    cmd.Parameters("@iid_trannumb") = Transnumb
    cmd.Parameters("@iid_compcode") = GetPKValue(rs.Bookmark, "iid_compcode")
    cmd.Parameters("@iid_npecode") = GetPKValue(rs.Bookmark, "iid_npecode")
    cmd.Parameters("@iid_ware") = GetPKValue(rs.Bookmark, "iid_ware")
    cmd.Parameters("@iid_transerl") = GetPKValue(rs.Bookmark, "iid_transerl")
    cmd.Parameters("@iid_stcknumb") = GetPKValue(rs.Bookmark, "iid_stcknumb")
    cmd.Parameters("@iid_ps") = GetPKValue(rs.Bookmark, "iid_ps")
    cmd.Parameters("@iid_serl") = GetPKValue(rs.Bookmark, "iid_serl")
    cmd.Parameters("@iid_newcond") = GetPKValue(rs.Bookmark, "iid_newcond")
    cmd.Parameters("@iid_stcktype") = GetPKValue(rs.Bookmark, "iid_stcktype")
    cmd.Parameters("@iid_ctry") = GetPKValue(rs.Bookmark, "iid_ctry")
    cmd.Parameters("@iid_tosubloca") = GetPKValue(rs.Bookmark, "iid_tosubloca")
    cmd.Parameters("@iid_tologiware") = GetPKValue(rs.Bookmark, "iid_tologiware")
    cmd.Parameters("@iid_owle") = GetPKValue(rs.Bookmark, "iid_owle")
    cmd.Parameters("@iid_leasecomp") = GetPKValue(rs.Bookmark, "iid_leasecomp")
    cmd.Parameters("iid_primqty") = GetPKValue(rs.Bookmark, "iid_primqty")
    cmd.Parameters("@iid_secoqty") = GetPKValue(rs.Bookmark, "iid_secoqty")
    cmd.Parameters("@iid_unitpric") = GetPKValue(rs.Bookmark, "iid_unitpric")
    cmd.Parameters("iid_curr") = "USD" 'GetPKValue(rs.Bookmark, "iid_curr")
    cmd.Parameters("iid_currvalu") = 1 'GetPKValue(rs.Bookmark, "iid_currvalu")
    cmd.Parameters("iid_stckdesc") = GetPKValue(rs.Bookmark, "iid_stckdesc")
    cmd.Parameters("@iid_fromlogiware") = GetPKValue(rs.Bookmark, "iid_fromlogiware")
    cmd.Parameters("@iid_fromsubloca") = GetPKValue(rs.Bookmark, "iid_fromsubloca")
    cmd.Parameters("@iid_origcond") = GetPKValue(rs.Bookmark, "iid_origcond")
    cmd.Parameters("@user") = CurrentUser

    'Execute the command.
    cmd.Execute

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
    If IsNull(rs("iid_compcode")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00326") 'J added
        MsgBox IIf(msg1 = "", "The field 'Company Code' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("iid_compcode")) Then
        If Len(Trim(rs("iid_compcode"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00329") 'J added
            MsgBox IIf(msg1 = "", "The field ' iid_compcode ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
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

'set parameters for data insert

Private Function PutInvtIssue() As Boolean
Dim np As String
    Dim cmd As Command

On Error GoTo errPutInvtIssue

    PutInvtIssue = False

    Set cmd = deIms.Commands("InvtIssue_Insert")


    np = deIms.NameSpace
    Transnumb = "TI-" & GetTransNumb(np, deIms.cnIms)
    cmd.Parameters("@NAMESPACE") = np
    cmd.Parameters("@TRANTYPE") = "TI"
    cmd.Parameters("@COMPANYCODE") = CompCode
    cmd.Parameters("@TRANSNUMB") = Transnumb
    cmd.Parameters("@ISSUTO") = ssdcboLocation.Columns("Code").text
    cmd.Parameters("@WHAREHOUSE") = ssdcboWarehouse.Columns("Code").text
    cmd.Parameters("@STCKNUMB") = Null
    cmd.Parameters("@COND") = Null
    cmd.Parameters("@SAP") = Null
    cmd.Parameters("@NEWSAP") = Null
    cmd.Parameters("@ENTYNUMB") = Null
    cmd.Parameters("@SUPPLIERCODE") = Null
    cmd.Parameters("@USER") = CurrentUser
    
    cmd.Execute

    PutInvtIssue = cmd.Parameters(0).Value = 0

    Exit Function

errPutInvtIssue:
    MsgBox Err.Description: Err.Clear
End Function

'function get data from recordset

Private Function GetPKValue(vBookMark As Variant, sColName As String) As Variant
    GetPKValue = rs(sColName)
End Function

'check data fields

Public Function CheckMasterFields() As Boolean

    CheckMasterFields = False
    
    If Len(Trim$(ssdcboLocation.text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00328") 'J added
        MsgBox IIf(msg1 = "", "Issue to cannot be left empty", msg1): Exit Function 'J modified
        '---------------------------------------------
    
    End If
        
    If Len(Trim$(ssdcboWarehouse.text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00330") 'J added
        MsgBox IIf(msg1 = "", "Warehouse cannot be left empty", msg1): Exit Function 'J modified
        '---------------------------------------------
        
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

'set combo can not enter data

Private Sub ssdcboWarehouse_KeyPress(KeyAscii As Integer)
If NavBar1.NewEnabled = False Then KeyAscii = 0
End Sub

Private Sub ssdcboWarehouse_LostFocus()
    ssdcboWarehouse.BackColor = &H80000005
End Sub

'depend tab set navbar button

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
                MsgBox "Invalid Quantity"
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
    End If
    If alarm Then
        backpass = False
        SSTab1.Tab = 1
        Exit Sub
    End If
    Screen.MousePointer = 11
    '------------------------



    blFlag = SSTab1.Tab = 1
    Screen.MousePointer = 11
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
    

    If SSTab1.Tab = 1 Then
        Me.Refresh
        If PreviousTab = 0 And fm = mdCreation Then _
            If Not (CheckMasterFields) Then SSTab1.Tab = 0
            
        If Requery Then
        
            If fm <> mdCreation Then Exit Sub
            iEditMode = IIf(Len(cbo_Transaction), cbo_Transaction, "")
            
            Set rs = deIms.GetInvtIssuedetl(CompCode, iEditMode)
            
            Requery = False
       End If
    End If


'    If SSTab1 = 0 Then
'
'        NavBar1.SaveEnabled = True
'
'        NavBar1.CloseEnabled = True
'        NavBar1.PrintEnabled = False
'        NavBar1.EMailEnabled = False
'    Else
'        NavBar1.PrintEnabled = Transnumb <> ""
'        NavBar1.PrintEnabled = True
'        NavBar1.EMailEnabled = NavBar1.PrintEnabled And rsReceptList.RecordCount <> 0
'        NavBar1.EMailEnabled = True
'    End If
'

    If SSTab1.Tab = 2 Then
        txtRemarks.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

'fill data to data grid

Private Sub AddFromSublocation(rst As ADODB.Recordset)
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboSubLocation(0).RemoveAll
    
    rst.MoveFirst
    ssdcboSubLocation(0).text = rst!Description
    
    Do While Not rst.EOF
        ssdcboSubLocation(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing

End Sub

'fill data to data grid

Private Sub AddFromLogicalWharehouse(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboLogicalWHouse(0).RemoveAll
    
    
    rst.MoveFirst
    ssdcboLogicalWHouse(0).text = rst!Description
    
    Do While Not rst.EOF
        ssdcboLogicalWHouse(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    'doevents: 'doevents: 'doevents
    Call ssdcboLogicalWHouse_Click(0)
End Sub

'fill data to data grid

Private Sub AddFromCondition(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    ssdcboCondition(0).RemoveAll
    
    rst.MoveFirst
    ssdcboCondition(0).text = rst!Description
    
    Do While Not rst.EOF
        ssdcboCondition(0).AddItem rst!Description & "" & ";" & rst!Code & ""
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    Call ssdcboCondition_Click(0)
    
End Sub


'SQL statementget condition recordset

Public Sub AddCondition()
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

'clear form

Private Sub ClearFields()
    
    optOwn = True
    optPool = True
    
    txtDesc = ""
    lblSecQnty = ""
    txtprimUnit = ""
    txtLeaseComp = ""
    txtprimUnit.Tag = ""
    ssdcboCountry.text = ""
    
    cboSerialNumb.ListIndex = CB_ERR
    
    ssdcboCondition(1).text = ""
    ssdcboSubLocation(1).text = ""
    ssdcboLogicalWHouse(1).text = ""
    
End Sub

'enable navbar button

Private Sub EnableControls(Value As Boolean)
    optOwn.Enabled = Value
    'optPool.Enabled = Value
    optLease.Enabled = Value
    optSpecific.Enabled = False
    
    
    txtDesc.Enabled = Value
    lblSecQnty.Enabled = Value
    txtprimUnit.Enabled = Value
    ssdcboCountry.Enabled = Value
    
    
    ssdcboCondition(0).Enabled = False
    ssdcboSubLocation(0).Enabled = False
    ssdcboLogicalWHouse(0).Enabled = Value
    
    ssdcboCondition(1).Enabled = False
    ssdcboSubLocation(1).Enabled = Value
    ssdcboLogicalWHouse(1).Enabled = Value
End Sub

'validate data format

Private Function CheckDetl() As Boolean
Dim l As Long

    
    If rs Is Nothing Then Exit Function
    If rs.State And adStateOpen = adStateClosed Then Exit Function
     
    l = SSTab1.Tab
    SSTab1.Tab = 1
    
    If Len(Trim$(ssdcboCountry.text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00006") 'J added
        MsgBox IIf(msg1 = "", "Country cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboCountry.SetFocus: Exit Function
    Else
        rs!iid_ctry = RTrim$(ssdcboCountry.Columns("Code").text)
    End If
    
        
'    If Len(Trim$(ssdcboStockType.Text)) = 0 Then
'        MsgBox "Stock Type cannot be left empty":
'        ssdcboStockType.SetFocus: Exit Function
'
'    Else
'        rs!iid_stcktype = RTrim$(ssdcboStockType.Columns("Code").Text)
'
'    End If
    
        
    If Len(Trim$(ssdcboCondition(0).text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00331") 'J added
        MsgBox IIf(msg1 = "", " From condition cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
       ssdcboCondition(0).SetFocus: Exit Function
        
    Else
        rs!iid_origcond = RTrim$(ssdcboCondition(0).Columns("Code").text)
        
    End If
            
    If Len(Trim$(ssdcboSubLocation(0).text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00332") 'J added
        MsgBox IIf(msg1 = "", "From Sub-Location cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboSubLocation(0).SetFocus: Exit Function
        
    
        Else
        rs!iid_fromsubloca = RTrim$(ssdcboSubLocation(0).Columns("Code").text)
        
    End If

    If Len(Trim$(ssdcboLogicalWHouse(0).text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00333") 'J added
        MsgBox IIf(msg1 = "", "From SLogical Warehousecannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboLogicalWHouse(0).SetFocus: Exit Function
        
    Else
        
        rs!iid_fromlogiware = RTrim$(ssdcboLogicalWHouse(0).Columns("Code").text)
    End If
        
        
    If Len(Trim$(ssdcboCondition(1).text)) = 0 Then
    
        'Modified by Juan (9/14/2000 for Multilingual
        msg1 = translator.Trans("M00334")  'J added
        MsgBox IIf(msg1 = "", "to condition cannot be left empty", msg1) 'J modified
        '--------------------------------------------
        
        ssdcboCondition(1).SetFocus: Exit Function
        
    Else
        rs!iid_newcond = RTrim$(ssdcboCondition(0).Columns("Code").text)
        
    End If
    
        '// To Sub=location
    If Len(Trim$(ssdcboSubLocation(1).text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00335")  'J added
        MsgBox IIf(msg1 = "", "To Sub-Location cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboSubLocation(1).SetFocus: Exit Function
        
    Else
        rs!iid_tosubloca = RTrim$(ssdcboSubLocation(1).Columns("Code").text)
    End If
    
    
    '// To Logical Warehouse
    If Len(Trim$(ssdcboLogicalWHouse(1).text)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00336") 'J added
        MsgBox IIf(msg1 = "", "To SLogical Warehousecannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboLogicalWHouse(1).SetFocus: Exit Function
        
    Else
        rs!iid_tologiware = RTrim$(ssdcboLogicalWHouse(1).Columns("Code").text)
    End If
    
    If Len(txtprimUnit) > 0 Then
    
        If IsNumeric(txtprimUnit) Then
            rs!iid_primqty = CDbl(txtprimUnit)
        Else
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00337") 'J added
            MsgBox IIf(msg1 = "", "Primary unit is not a valid number", msg1) 'J modified
            '---------------------------------------------
            
            txtprimUnit.SetFocus
            Exit Function
        End If
        
    Else
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00338") 'J added
        MsgBox IIf(msg1 = "", "Primary unit cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtprimUnit.SetFocus: Exit Function
    End If
        
    If optSpecific Then
        
        rs!iid_ps = 0
        If Len(Trim$(cboSerialNumb)) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00339") 'J added
            MsgBox "Serial number cannot be left empty" 'J modified
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
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00340") 'J added
            MsgBox IIf(msg1 = "", "Lease Company cannot be left empty", msg1) 'J modified
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
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00341") 'J added
            MsgBox IIf(msg1 = "", "Secondary Quantity does not have a valid number", msg1) 'J modified
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

Private Sub Text1_Change()
Dim i As Integer
Dim text As String
Dim Flag, last As Variant
    With ssdbStockInfo
        If .Rows > 0 Then
            text = Left(.Columns(0).text, Len(Text1))
            Flag = .Bookmark
            .MoveFirst
            last = .GetBookmark(.Rows - 1)
            If text = UCase(Text1) Then
                .Bookmark = Flag
            Else
                For i = 0 To Len(Text1)
                    If Left(text, i) = Left(Text1, i) Then
                        .MoveNext
                        Flag = .Bookmark
                    Else
                        Flag = .FirstRow
                        Exit For
                    End If
                Next
                .MoveFirst
            End If
            Do While True
                If Left(.Columns(0).text, Len(Text1)) = Text1 Then
                    Exit Sub
                End If
                .MoveNext
                If CStr(.Bookmark) = CStr(last) Then Exit Do
            Loop
        End If
    End With
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


'format primary unit to 4 decimal position

Private Sub txtprimUnit_Change()
On Error Resume Next
Dim db As Double

    'Added by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("M00122") 'J added
    '------------------------------------------

    If Len(txtprimUnit) > 0 Then
        If Not IsNumeric(txtprimUnit) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            MsgBox IIf(msg1 = "", "Invalid Value", msg1) 'J modified
            '---------------------------------------------
            
            txtprimUnit.SetFocus: Exit Sub
        End If
    End If
        
    If Len(txtprimUnit) > 0 Then
    
        If IsNumeric(txtprimUnit) Then
        
            'db = FormatNumber((txtprimUnit), 4)
            db = txtprimUnit      'M
            
            'Modified by Juan (9/14/2000) for Multilingual
            If db < 1 Then MsgBox IIf(msg1 = "", "Invalid Value", msg1) 'J modified
            '---------------------------------------------
            
            
            If Len(Trim$(txtprimUnit.Tag)) > 0 Then
            
                If FormatNumber((txtprimUnit.Tag), 2) < db Then
                    txtprimUnit = ""
                    
                    'Modified by Juan (9/14/2000) for Multilingual
                    msg1 = translator.Trans("M00342") 'J added
                    MsgBox IIf(msg1 = "", "Value is too large", msg1) 'J modified
                    '---------------------------------------------
                    
                    txtprimUnit.SetFocus
                   ' txtprimUnit = FormatNumber$(txtprimUnit.Tag, 4)
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

'set text box enable

Private Sub optLease_Click()
On Error Resume Next

    If rs!iid_owle <> 0 Then _
       rs!iid_owle = 0
    
    rs!iid_owle = 0
    
    txtLeaseComp.Enabled = True
    txtLeaseComp.SetFocus
    
    If Err Then Err.Clear
End Sub

'set text box enable

Private Sub optOwn_Click()
On Error Resume Next

    If rs!iid_owle <> 1 Then _
       rs!iid_owle = 1
    
    rs!iid_owle = 1
    txtLeaseComp.Enabled = False
    
    If Err Then Err.Clear
End Sub

'set text box enable

Private Sub optPool_Click()
On Error Resume Next

    If rs!iid_ps <> 1 Then _
        rs!iid_ps = 1
    
    rs!iid_ps = 1
    
    txtprimUnit.Enabled = True
    cboSerialNumb.Enabled = False
    
    Err.Clear
End Sub

'set text box enable

Private Sub optSpecific_Click()
On Error Resume Next
    If rs!iid_ps <> 0 Then _
        rs!iid_ps = 0
        
    rs!iid_ps = 0
    cboSerialNumb.Enabled = True
    cboSerialNumb.SetFocus
    
    txtprimUnit.text = 1
    txtprimUnit.Enabled = False
    Err.Clear
    
    
End Sub

Private Sub txtprimUnit_GotFocus()
    txtprimUnit.BackColor = &HC0FFFF
End Sub

Private Sub txtprimUnit_LostFocus()
    txtprimUnit.BackColor = &H80000005
    If IsNumeric(txtprimUnit) Then
        txtprimUnit = FormatNumber$(txtprimUnit, 0)
    Else
        txtprimUnit = ""
    End If
End Sub

'validate primary unit and format to 4 decimal

Private Sub txtprimUnit_Validate(Cancel As Boolean)
On Error Resume Next
Dim CompFactor As Double

    
    If Len(Trim$(txtprimUnit)) = 0 Then Exit Sub

    If lblPrimUnit = lblSecUnit Then
        lblSecQnty = FormatNumber(txtprimUnit, 2)
    Else

        CompFactor = ImsDataX.ComputingFactor(deIms.NameSpace, lblCommodity, deIms.cnIms)
        
        If IsNumeric(txtprimUnit) Then
            If CompFactor = 0 Then
                lblSecQnty = FormatNumber(txtprimUnit, 2)
            Else
                lblSecQnty = FormatNumber(txtprimUnit * 10000 / CompFactor, 2)
            End If
        Else
            lblSecQnty = txtprimUnit
        End If
    End If
    
'    txtprimUnit = FormatNumber$(txtprimUnit, 4)   'M
    rs!iid_secoqty = lblSecQnty
End Sub

'check secordary quantity data

Private Sub lblSecQnty_Change()
    If Len(lblSecQnty) Then
        If Not IsNumeric(lblSecQnty) Then
        
            'Added by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00122") 'J added
            MsgBox IIf(msg1 = "", "Invalid Value", msg1) 'J modified
            '------------------------------------------

        End If
    End If
End Sub

'assign values to recordset

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
    rs!iid_ware = RTrim$(ssdcboWarehouse.Columns("Code").text)
    rs!iid_stcknumb = Trim$(ssdbStockInfo.Columns("Commodity").text)
    rs!iid_stckdesc = Trim$(ssdbStockInfo.Columns("Description").text)
End Sub

'SQL statement get serial number

Public Function GetNextSerial() As Long
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

'get data and populate combe

Private Sub AddIssueNumb()
On Error Resume Next

Dim rst As ADODB.Recordset

    Set rst = deIms.rsIssueNumber
    If rst.State And adStateOpen = adStateOpen Then rst.Close
    Call deIms.IssueNumber(deIms.NameSpace, CompCode, "TI")
    
    Call PopuLateFromRecordSet(cbo_Transaction, rst, rst.Fields(0).Name, False)
    
    rst.Close
    Set rst = Nothing
    If Err Then Err.Clear
End Sub

'get crystal report parmeters and application path

Public Sub BeforePrint()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = ReportPath & "wareI.rpt"
        .ParameterFields(0) = "transnumb;" & cbo_Transaction & ";TRUE"
        .ParameterFields(1) = "namespace;" & deIms.NameSpace & ";TRUE"
        
        'Modified by Juan (9/14/2000) for Mulilingual
       ' msg1 = translator.Trans("L00311") 'J added
        .WindowTitle = "Warehouse to Warehouse"
        Call translator.Translate_Reports("wareI.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '--------------------------------------------
        
    End With

End Sub

'assign values to parameters

Private Function PutIssueRemarks() As Boolean
Dim cmd As ADODB.Command

    Set cmd = deIms.Commands("InvtIssuetRem_Insert")
     
    cmd.Parameters("@CompanyCode") = CompCode
    cmd.Parameters("@NameSpace") = deIms.NameSpace
    cmd.Parameters("@WhareHouse") = ssdcboWarehouse.Columns("Code").text
    cmd.Parameters("@TranNumb") = Transnumb
    cmd.Parameters("@LineNumb") = 1
    cmd.Parameters("Remarks") = txtRemarks
    cmd.Parameters("@USER") = CurrentUser
    
    Call cmd.Execute(Options:=adExecuteNoRecords)
    PutIssueRemarks = cmd.Parameters(0).Value = 0
End Function


'get company recordset and fill data grid

Private Sub AddCompanies()
'On Error Resume Next
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
    ssdcboCompany.DataMode = ssDataModeAddItem
    
    rs.MoveFirst
    If rs.RecordCount = 0 Then Exit Sub
    
    Do Until rs.EOF
        ssdcboCompany.AddItem rs("com_name") & Chr$(1) & rs("com_compcode")
        
        rs.MoveNext
    Loop
    
End Sub

'get report parameters

Private Function CreateRpti() As RPTIFileInfo

    With CreateRpti
        ReDim .Parameters(1)
        .ReportFileName = ReportPath & "wareAI.rpt"
        .Parameters(0) = "transnumb=" & cbo_Transaction
        .Parameters(1) = "namespace=" & deIms.NameSpace
    
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("areAI.rpt") 'J added
        '---------------------------------------------
    
    End With

End Function

Private Sub txtRemarks_GotFocus()
Call HighlightBackground(txtRemarks)
End Sub

Private Sub txtRemarks_LostFocus()
Call NormalBackground(txtRemarks)
End Sub
