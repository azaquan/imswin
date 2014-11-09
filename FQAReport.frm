VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_FQAReporting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FQA Reporting"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7020
   Begin VB.Frame Fraclosure 
      Height          =   615
      Left            =   240
      TabIndex        =   46
      Top             =   1200
      Width           =   6615
      Begin VB.OptionButton OptYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptNo 
         Caption         =   "No"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label LblClosure 
         Caption         =   "Is this a month closure ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   4935
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   105
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58261505
         CurrentDate     =   37595
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   105
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58261505
         CurrentDate     =   37595
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdbCompany 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   3855
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6800
         _ExtentY        =   503
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdbLocation 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   3855
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6800
         _ExtentY        =   503
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin VB.Label Label2 
         Caption         =   "Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Company:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   855
      End
      Begin VB.Label LblFromDateInvt 
         Caption         =   "From Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   180
         Width           =   855
      End
      Begin VB.Label LblToDateInvt 
         Caption         =   "To Date:"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   240
      TabIndex        =   28
      Top             =   7080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Max             =   10
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   240
      TabIndex        =   30
      Top             =   1875
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Selection"
      TabPicture(0)   =   "FQAReport.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraPo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraInvoice"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraInventory"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Chk_po"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Chk_Supp"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Chk_Invt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "E-Mail List"
      TabPicture(1)   =   "FQAReport.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Recipients"
      Tab(1).Control(1)=   "SSGRDRecepients"
      Tab(1).Control(2)=   "dgRecepients"
      Tab(1).Control(3)=   "TxtEmail"
      Tab(1).Control(4)=   "cmd_Add"
      Tab(1).Control(5)=   "cmd_Remove"
      Tab(1).ControlCount=   6
      Begin VB.CheckBox Chk_Invt 
         Height          =   200
         Left            =   1680
         TabIndex        =   44
         Top             =   3960
         Width           =   200
      End
      Begin VB.CheckBox Chk_Supp 
         Height          =   200
         Left            =   1680
         TabIndex        =   43
         Top             =   2760
         Width           =   200
      End
      Begin VB.CheckBox Chk_po 
         Height          =   200
         Left            =   1680
         TabIndex        =   42
         Top             =   480
         Width           =   200
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74880
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74880
         TabIndex        =   21
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtEmail 
         Height          =   375
         Left            =   -73680
         TabIndex        =   22
         Top             =   2580
         Width           =   5175
      End
      Begin VB.Frame FraInventory 
         Caption         =   "Inventory"
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   3960
         Width           =   6375
         Begin VB.Frame Frame1 
            Height          =   495
            Left            =   1320
            TabIndex        =   45
            Top             =   150
            Width           =   4935
            Begin VB.OptionButton optInvtGeneralLedger 
               Caption         =   "General Ledger load"
               Height          =   255
               Left            =   2880
               TabIndex        =   14
               Top             =   150
               Width           =   1935
            End
            Begin VB.OptionButton optInvtcomplete 
               Caption         =   "Complete view"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   150
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdTransactype 
            Height          =   285
            Left            =   1320
            TabIndex        =   15
            Top             =   720
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8678
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label Label8 
            Caption         =   "Trans. Type"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame FraInvoice 
         Caption         =   "Supplier Invoicing"
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   6375
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdSupplier 
            Height          =   285
            Left            =   1320
            TabIndex        =   11
            Top             =   360
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8811
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdInvoiceNo 
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   720
            Width           =   4935
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
            ColumnHeaders   =   0   'False
            RowHeight       =   423
            Columns(0).Width=   11959
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin VB.Label Label7 
            Caption         =   "Invoice #"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Supplier"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FraPo 
         Caption         =   "Transaction Order"
         Height          =   2175
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   6375
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdPonumb 
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   1800
            Width           =   4935
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
            Columns(0).Width=   12091
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdPODoc 
            Height          =   285
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8837
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdPODel 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8731
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdPOShip 
            Height          =   285
            Left            =   1320
            TabIndex        =   8
            Top             =   1080
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8811
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSGrdPOInvt 
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   1440
            Width           =   4935
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
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8758
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   8705
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label Label6 
            Caption         =   "Inventory"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Shipping"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Delivery"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label LblDoctype 
            Caption         =   "Document"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LblPonumb 
            Caption         =   "Transaction #"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   23
         Top             =   3000
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "phd_code"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   5
            SizeMode        =   1
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3720.189
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGRDRecepients 
         Height          =   1935
         Left            =   -73680
         TabIndex        =   20
         Top             =   600
         Width           =   5175
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
         ColumnHeaders   =   0   'False
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   8493
         Columns(0).Caption=   "Emails"
         Columns(0).Name =   "Emails"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   9128
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Emails"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74880
         TabIndex        =   34
         Top             =   600
         Width           =   1260
      End
   End
End
Attribute VB_Name = "Frm_FQAReporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FPOGridsPopulated As Boolean
Dim FpopulateInventory As Boolean
Dim FPopulatedSupplierInvoiceGrids As Boolean
Private Sub Chk_Invt_Click()
Dim x As Long
Dim i As Integer
If Chk_Invt.value = 1 Then

    FraInventory.Enabled = True
    Call PopulateInventoryGrids
    
Else

    FraInventory.Enabled = False
    
End If
    

End Sub

Private Sub Chk_po_Click()

If Chk_po.value = 1 Then
    FraPo.Enabled = True
    Call PopulatePOGrids(SSdbCompany.Tag, SSdbLocation.Tag, DTPFrom.value, DTPTo.value)
    
Else

    FraPo.Enabled = False

''    SSGrdPODel.ena
''    SSGrdPODoc
''    SSGrdPOInvt
''    SSGrdPonumb

End If
    

End Sub

Private Sub Chk_Supp_Click()

If Chk_Supp.value = 1 Then
    FraInvoice.Enabled = True
    Call PopulateSupplierInvoiceGrids
Else
    FraInvoice.Enabled = False
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub companyCOMBO_Click()

End Sub

Private Sub companyCOMBO_InitColumnProps()

End Sub

Private Sub cmd_Add_Click()
If Len(Trim(TxtEmail)) > 0 Then SSGRDRecepients.AddItem Trim(UCase(TxtEmail))
TxtEmail = ""
End Sub

Private Sub cmd_Remove_Click()

SSGRDRecepients.DeleteSelected

End Sub

Private Sub CmdExport_Click()


If Len(Trim(SSdbCompany.Text)) = 0 Then MsgBox "Please select a company.", vbInformation, "Ims": Exit Sub
If Len(Trim(SSdbLocation.Text)) = 0 Then MsgBox "Please select a Location.", vbInformation, "Ims": Exit Sub
Screen.MousePointer = vbHourglass
If SSGRDRecepients.Rows > 0 Then

ProgressBar1.Visible = True

If Chk_po.value = 1 Then Call ExportPOtoExcel
If Chk_Invt.value = 1 Then Call ExportInventoryToExcel
If Chk_Supp.value = 1 Then Call ExportInvoicetoExcel

If Chk_po.value = 0 And Chk_Invt.value = 0 And Chk_Supp.value = 0 Then

        MsgBox "Please choose atleast one option to generate report.", vbInformation, "Ims"

End If

Call MDI_IMS.WriteStatus("", 1)

ProgressBar1.Visible = False


Else

MsgBox "No Emails recepients. Excel sheets will not be generated.", vbInformation, "Ims"

End If

Screen.MousePointer = vbNormal

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub dgRecepients_DblClick()
SSGRDRecepients.AddItem Trim(UCase(dgRecepients.Columns(1).Text))
End Sub

Private Sub Form_Load()

'deIms.cnIms.Open
ProgressBar1.Visible = False
'deIms.NameSpace = "PECT"
PopulateCompany
Me.Width = 7140
Me.Height = 7815
DTPFrom.value = Date
DTPTo.value = DateAdd("d", 1, Date)
Chk_Invt.value = 1
Chk_po = 1
Chk_Supp = 1
'DTPFrom.SetFocus
    With Frm_FQAReporting
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub locationCOMBO_InitColumnProps()

End Sub

Private Sub Form_Unload(Cancel As Integer)

FPOGridsPopulated = False
FpopulateInventory = False
FPopulatedSupplierInvoiceGrids = False

End Sub

Private Sub OptComplete_Click()

End Sub

Private Sub OptLedger_Click()

End Sub

Private Sub optInvtcomplete_GotFocus()
 Call HighlightBackground(optInvtcomplete)
End Sub

Private Sub optInvtcomplete_LostFocus()

Call NormalBackground(optInvtcomplete)
End Sub

Private Sub optInvtGeneralLedger_Click()
 Call HighlightBackground(optInvtGeneralLedger)
End Sub

Private Sub optInvtGeneralLedger_LostFocus()
Call NormalBackground(optInvtGeneralLedger)
End Sub

Private Sub SSdbCompany_Click()

SSdbCompany.Tag = Trim(SSdbCompany.Columns(0).value)

Call PopulateLocation

End Sub

Private Sub SSdbCompany_GotFocus()
 Call HighlightBackground(SSdbCompany)
End Sub

Private Sub SSdbCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSdbCompany.DroppedDown Then SSdbCompany.DroppedDown = True
End Sub

Private Sub SSdbCompany_LostFocus()
Call NormalBackground(SSdbCompany)
End Sub

Private Sub SSdbLocation_Click()
SSdbLocation.Tag = Trim(SSdbLocation.Columns("code").value)

If Chk_po.value = 1 Then Call PopulatePOGrids(Trim(SSdbCompany.Tag), Trim(SSdbLocation.Tag), DTPFrom.value, DTPTo.value)

If Chk_Invt.value = 1 Then PopulateInventoryGrids

If Chk_Supp.value = 1 Then PopulateSupplierInvoiceGrids



End Sub

Public Function PopulateLocation() As Boolean

Dim RsLocation As New ADODB.Recordset

On Error GoTo ErrHandler

RsLocation.Source = "select loc_locacode , loc_name from location where loc_npecode ='" & deIms.NameSpace & "' and loc_compcode ='" & SSdbCompany.Tag & "'"

RsLocation.Open , deIms.cnIms

SSdbLocation.RemoveAll

SSdbLocation.AddItem "ALL" & vbTab & "ALL"

Do While Not RsLocation.EOF

    SSdbLocation.AddItem RsLocation!loc_locacode & vbTab & RsLocation!loc_name

    RsLocation.MoveNext
    
Loop

Exit Function
ErrHandler:

MsgBox "Errors occurred while populating the location combo. " & Err.Description, vbCritical, "Ims"
Err.Clear
End Function

Public Function PopulatePOGrids(CompCode As String, Location As String, Fromdate As Date, Todate As Date) As Boolean

Dim rsPO As New ADODB.Recordset
Dim rsDOCTYPE As New ADODB.Recordset
Dim RsDelivery As New ADODB.Recordset
Dim RsShipping As New ADODB.Recordset
Dim RsInventory As New ADODB.Recordset

On Error GoTo ErrHand

rsDOCTYPE.Source = "select doc_code, doc_desc from doctype where doc_npecode ='" & deIms.NameSpace & "'"
rsDOCTYPE.Open , deIms.cnIms

SSGrdPODoc.RemoveAll

SSGrdPODoc.AddItem "ALL" & vbTab & "ALL"

Do While Not rsDOCTYPE.EOF

    SSGrdPODoc.AddItem rsDOCTYPE!doc_code & vbTab & rsDOCTYPE!doc_desc

    rsDOCTYPE.MoveNext

Loop

SSGrdPOInvt.RemoveAll
SSGrdPODel.RemoveAll
SSGrdPOShip.RemoveAll


SSGrdPOInvt.AddItem "ALL" & vbTab & "ALL"
SSGrdPOInvt.AddItem "IC" & vbTab & "INVENTORY, COMPLETE"
SSGrdPOInvt.AddItem "IP" & vbTab & "INVENTORY, PARTIAL"
SSGrdPOInvt.AddItem "NI" & vbTab & "NOT IN INVENTORY"


SSGrdPODel.AddItem "ALL" & vbTab & "ALL"
SSGrdPODel.AddItem "NR" & vbTab & "NOT RECEIVED"
SSGrdPODel.AddItem "RC" & vbTab & "RECEPTION, COMPLETE"
SSGrdPODel.AddItem "RP" & vbTab & "RECEPTION, PARTIAL"


SSGrdPOShip.AddItem "ALL" & vbTab & "ALL"
SSGrdPOShip.AddItem "NS" & vbTab & "NOT SHIPPED"
SSGrdPOShip.AddItem "SC" & vbTab & "SHIPPING, COMPLETE"
SSGrdPOShip.AddItem "SP" & vbTab & "SHIPPING, PARTIAL"


SSGrdPOInvt.Text = "ALL"
SSGrdPOInvt.Tag = "ALL"

SSGrdPODel.Text = "ALL"
SSGrdPODel.Tag = "ALL"

SSGrdPOShip.Text = "ALL"
SSGrdPOShip.Tag = "ALL"

SSGrdPODoc.Text = "ALL"
SSGrdPODoc.Tag = "ALL"

SSGrdPonumb.Text = "ALL"
SSGrdPonumb.Tag = "ALL"

PopulatePOGrids = True
FPOGridsPopulated = True
Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the PO FQA grids." & Err.Description, vbCritical, "Ims"


End Function

Public Function PopulateInventoryGrids() As Boolean
Dim RsTransactype As New ADODB.Recordset

On Error GoTo ErrHand

RsTransactype.Source = "select  tty_code, tty_desc    from Transactype where tty_code in ('TI','RT','RR','I','IT','R') and tty_npecode='" & deIms.NameSpace & "'"
RsTransactype.Open , deIms.cnIms

SSGrdTransactype.RemoveAll

SSGrdTransactype.AddItem "ALL" & vbTab & "ALL"

Do While Not RsTransactype.EOF

    SSGrdTransactype.AddItem RsTransactype!tty_code & vbTab & RsTransactype!tty_desc
    
    RsTransactype.MoveNext
    
Loop

SSGrdTransactype.Text = "ALL"
SSGrdTransactype.Tag = "ALL"

PopulateInventoryGrids = True
FpopulateInventory = True
Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the Inventory grid.", vbCritical, "Ims"


End Function

Public Function PopulateSupplierInvoiceGrids() As Boolean
Dim RsSup As New ADODB.Recordset
Dim Rsinvoice As New ADODB.Recordset

On Error GoTo ErrHand

RsSup.Source = "select sup_code, sup_name from supplier where sup_npecode  ='" & deIms.NameSpace & "' order by sup_name"
RsSup.Open , deIms.cnIms

SSGrdSupplier.RemoveAll

SSGrdSupplier.AddItem "ALL" & vbTab & "ALL"

Do While Not RsSup.EOF

    SSGrdSupplier.AddItem RsSup!sup_code & vbTab & RsSup!sup_name
    
    RsSup.MoveNext
    
Loop

SSGrdSupplier.Text = "ALL"
SSGrdSupplier.Tag = "ALL"

SSGrdInvoiceNo.Text = "ALL"
SSGrdInvoiceNo.Tag = "ALL"

PopulateSupplierInvoiceGrids = True
FPopulatedSupplierInvoiceGrids = True
Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the Supplier Invoice grid.", vbCritical, "Ims"


End Function

Private Sub SSOleDBGrid1_InitColumnProps()

End Sub

Public Function PopulateCompany()
On Error GoTo ErrHand
Dim rsCOMPANY As New ADODB.Recordset

rsCOMPANY.Source = "select com_compcode ,com_name  from company where com_npecode ='" & deIms.NameSpace & "'"
rsCOMPANY.Open , deIms.cnIms

SSdbCompany.Tag = ""

SSdbCompany.AddItem "ALL" & vbTab & "ALL"

Do While Not rsCOMPANY.EOF

    SSdbCompany.AddItem rsCOMPANY!com_compcode & vbTab & rsCOMPANY!com_name
    
    rsCOMPANY.MoveNext
    
Loop

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to fill the Company combo. " & Err.Description, vbCritical, "Ims"
End Function

Private Sub SSGrdEmail_InitColumnProps()

End Sub

Private Sub SSdbLocation_GotFocus()
 Call HighlightBackground(SSdbLocation)
End Sub

Private Sub SSdbLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSdbLocation.DroppedDown Then SSdbLocation.DroppedDown = True
End Sub

Private Sub SSdbLocation_LostFocus()
Call NormalBackground(SSdbLocation)
End Sub

Private Sub SSGrdInvoiceNo_DropDown()
If Trim(SSGrdSupplier.Text) > 0 Then

Else

  MsgBox "Please make sure that supplier is selected." & Err.Description, vbInformation, "Ims"
  
End If
End Sub

Private Sub SSGrdInvoiceNo_GotFocus()
 Call HighlightBackground(SSGrdInvoiceNo)
End Sub

Private Sub SSGrdInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdInvoiceNo.DroppedDown Then SSGrdInvoiceNo.DroppedDown = True
End Sub

Private Sub SSGrdInvoiceNo_LostFocus()
Call NormalBackground(SSGrdInvoiceNo)
End Sub

Private Sub SSGrdPODel_Click()
SSGrdPODel.Tag = Trim(SSGrdPODel.Columns(0).Text)
PopulatePOnumbCombo
End Sub

Private Sub SSGrdPODel_GotFocus()
 Call HighlightBackground(SSGrdPODel)
End Sub

Private Sub SSGrdPODel_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdPODel.DroppedDown Then SSGrdPODel.DroppedDown = True
End Sub

Private Sub SSGrdPODel_LostFocus()
Call NormalBackground(SSGrdPODel)
End Sub

Private Sub SSGrdPODoc_Click()
SSGrdPODoc.Tag = Trim(SSGrdPODoc.Columns(0).Text)
PopulatePOnumbCombo
End Sub

Private Sub SSGrdPODoc_GotFocus()
 Call HighlightBackground(SSGrdPODoc)
End Sub

Private Sub SSGrdPODoc_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdPODoc.DroppedDown Then SSGrdPODoc.DroppedDown = True
End Sub

Private Sub SSGrdPODoc_LostFocus()
Call NormalBackground(SSGrdPODoc)
End Sub

Private Sub SSGrdPOInvt_Click()
SSGrdPOInvt.Tag = Trim(SSGrdPOInvt.Columns(0).Text)
Call PopulatePOnumbCombo
End Sub

Private Sub SSGrdPOInvt_GotFocus()
 Call HighlightBackground(SSGrdPOInvt)
End Sub

Private Sub SSGrdPOInvt_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdPOInvt.DroppedDown Then SSGrdPOInvt.DroppedDown = True
End Sub




Private Sub SSGrdPonumb_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

End Sub

Private Sub SSGrdPOInvt_LostFocus()
Call NormalBackground(SSGrdPOInvt)
End Sub

Private Sub SSGrdPonumb_DropDown()

If Trim(SSGrdPODel.Text) > 0 And Trim(SSGrdPOInvt.Text) > 0 And Trim(SSGrdPOShip.Text) > 0 And Trim(SSGrdPODoc.Text) > 0 Then

Else

  MsgBox "Please make sure that all Delivery, Shipping, Inventory and Document values are filled." & Err.Description, vbInformation, "Ims"
  
End If

End Sub

Private Sub SSGrdPonumb_GotFocus()
 Call HighlightBackground(SSGrdPonumb)
End Sub

Private Sub SSGrdPonumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdPonumb.DroppedDown Then SSGrdPonumb.DroppedDown = True

End Sub

Private Sub SSGrdPonumb_LostFocus()
Call NormalBackground(SSGrdPonumb)
End Sub

Private Sub SSGrdPOShip_Click()
SSGrdPOShip.Tag = Trim(SSGrdPOShip.Columns(0).Text)
PopulatePOnumbCombo
End Sub

Private Sub SSGrdPOShip_GotFocus()
 Call HighlightBackground(SSGrdPOShip)
End Sub

Private Sub SSGrdPOShip_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdPOShip.DroppedDown Then SSGrdPOShip.DroppedDown = True
End Sub

Private Sub SSGrdPOShip_LostFocus()
Call NormalBackground(SSGrdPOShip)
End Sub

Private Sub SSGrdSupplier_Click()
SSGrdSupplier.Tag = Trim(SSGrdSupplier.Columns(0).Text)
Call PopulateInvoiceCombo
End Sub

Private Sub SSGrdSupplier_GotFocus()
 Call HighlightBackground(SSGrdSupplier)
End Sub

Private Sub SSGrdSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdSupplier.DroppedDown Then SSGrdSupplier.DroppedDown = True
End Sub

Private Sub SSGrdSupplier_LostFocus()
Call NormalBackground(SSGrdSupplier)
End Sub

Private Sub SSGrdTransactype_Click()
SSGrdTransactype.Tag = Trim(SSGrdTransactype.Columns(0).Text)
End Sub

Private Sub SSGrdTransactype_GotFocus()
 Call HighlightBackground(SSGrdTransactype)
End Sub

Private Sub SSGrdTransactype_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSGrdTransactype.DroppedDown Then SSGrdTransactype.DroppedDown = True
End Sub



Public Function PopulateEmail() As Boolean
Dim RsEmail As New ADODB.Recordset
Dim co As MSDataGridLib.column

On Error GoTo ErrHand
    
    RsEmail.Source = "select phd_name, phd_mail from phonedir where phd_npecode ='" & deIms.NameSpace & "' and len(isnull(phd_mail,'')) > 0 order by phd_name"
    RsEmail.Open , deIms.cnIms

    Set co = dgRecepients.Columns(1)

    co.Caption = "Email Address"

    co.DataField = "phd_mail"

    dgRecepients.Columns(0).DataField = "phd_name"

    Set dgRecepients.DataSource = RsEmail
      
        
   Exit Function
        
     
ErrHand:
     MsgBox "Errors Occurred while trying to populate the emails.", vbCritical, "Ims"
     Err.Clear

End Function



Public Function ExportPOtoExcel() As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo ErrHand
ExportPOtoExcel = False

Call MDI_IMS.WriteStatus("Generating Transaction Order report ...", 1)

rs.Source = " select Ponumb,po_date 'Creation Date',usr_username 'Buyer',pri_desc 'Shipping Mode', po_srvccode 'ServiceCode',poi_afe 'AFE',"
rs.Source = rs.Source & " po_currcode 'Currency',sup_name 'supplier', ItemNo,poi_comm 'Folio#',poi_primreqdqty 'Quantity',poi_stasliit 'Status', poi_stasdlvy 'Delivery Status' ,poi_stasship 'Shipping Status', poi_stasinvt 'Invt Status',poi_unitprice 'Unit Price', FromCompany, FromLocation, FromUsChar, FromStockType, FromCamChar, ToCompany, ToLocation, ToUsChar,  ToStockType, ToCamChar  from pofqa"
rs.Source = rs.Source & " inner join POitem on poi_ponumb =ponumb and poi_liitnumb = itemno and poi_npecode =Npce_code"
rs.Source = rs.Source & " inner join PO on po_ponumb =ponumb and po_npecode =Npce_code "
rs.Source = rs.Source & " inner join supplier on sup_code=po_suppcode and sup_npecode =Npce_code "
rs.Source = rs.Source & " inner join xuserprofile on usr_userid = po_buyr and usr_npecode = npce_code"
rs.Source = rs.Source & " inner join priority on pri_code = po_priocode and pri_npecode = npce_code"
rs.Source = rs.Source & " Where ItemNo <> 0 and npce_code ='" & deIms.NameSpace & "' and po_stas not in ('OH','CA')"
rs.Source = rs.Source & " and datediff(dd,'" & DTPFrom.value & "',po_date) > =0 and datediff(dd,po_date,'" & DTPTo.value & "') >=0   "

If Trim(SSdbCompany.Tag) <> "ALL" Then rs.Source = rs.Source & " and po_compcode='" & Trim(SSdbCompany.Tag) & "'"
If Trim(SSdbLocation.Tag) <> "ALL" Then rs.Source = rs.Source & " and po_invloca='" & Trim(SSdbLocation.Tag) & "'"

If Trim(SSGrdPonumb.Text) <> "ALL" Then rs.Source = rs.Source & " and ponumb='" & SSGrdPonumb.Text & "'"
If Trim(SSGrdPODoc.Text) <> "ALL" Then rs.Source = rs.Source & " and po_docutype  ='" & SSGrdPODoc.Tag & "'"
If Trim(SSGrdPODel.Text) <> "ALL" Then rs.Source = rs.Source & " and po_stasdelv  ='" & SSGrdPODel.Tag & "'"
If Trim(SSGrdPOInvt.Text) <> "ALL" Then rs.Source = rs.Source & " and po_stasinvt = '" & SSGrdPOInvt.Tag & "'"
If Trim(SSGrdPOShip.Text) <> "ALL" Then rs.Source = rs.Source & " and po_stasship ='" & SSGrdPOShip.Tag & "'"

rs.Source = rs.Source & " order by ponumb, cast(itemno as int)"

rs.Open , deIms.cnIms
If rs.RecordCount > 0 Then Call ExportToExcel1(rs, , , ProgressBar1, "POFQA")


ExportPOtoExcel = True
Exit Function
ErrHand:

MsgBox "Errors occurred while trying to export to excel." & Err.Description, vbCritical, "Ims"
Err.Clear
End Function

Public Function ExportInventoryToExcel() As Boolean
Dim rs As New ADODB.Recordset
Dim Errcode As Boolean
On Error GoTo ErrHand

Call MDI_IMS.WriteStatus("Generating Inventory report ...", 1)

If optInvtcomplete.value = True Then

    rs.Source = "select  FromCompany, FromLocation,"
    rs.Source = rs.Source & " FromUsChar, FromStockType, FromCamChar, TransactionNo, ItemNo, TransactionType, Ponumb, PoItemNo, StockNo,"
    rs.Source = rs.Source & " BaseCurrency, ExtendedCurrency, BaseCurUnitPrice, ExtendedUnitPrice, Quantity, ToCondition, ToCompany, ToLocation, ToUsChar, ToStockType, ToCamChar,"
    rs.Source = rs.Source & " CreaDate  from inventoryFQA where npce_code ='" & deIms.NameSpace & "' "
    
    If Trim(SSdbCompany.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromcompany=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' and  Level ='c')"
    If Trim(SSdbLocation.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromlocation=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' AND LOCATIONCODE = '" & Trim(SSdbLocation.Tag) & "' and  Level ='LB')"

    rs.Source = rs.Source & " and  datediff(dd,'" & DTPFrom.value & "',transactiondate) > =0 and datediff(dd,transactiondate,'" & DTPTo.value & "') >=0   "
    If Trim(SSGrdTransactype.Text) <> "ALL" Then rs.Source = rs.Source & " and TransactionType = '" & SSGrdTransactype.Tag & "' "
    rs.Source = rs.Source & " order by fromcompany, FromLocation , TransactionType, TransactionNo, ItemNo"

ElseIf optInvtGeneralLedger.value = True Then
    
    rs.Source = "select  ToCompany company, ToLocation location, ToUsChar USChar, ToStockType StockType, ToCamChar CamChar,"
    rs.Source = rs.Source & " TransactionNo,ItemNo,"
    rs.Source = rs.Source & " (cast(BaseCurUnitPrice as numeric(18,2)) * cast(Quantity as numeric(18,2))) 'Amount', TRANSACTIONDATE"
    rs.Source = rs.Source & " from inventoryFQA where npce_code ='" & deIms.NameSpace & "' "
    
    If Trim(SSdbCompany.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromcompany=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' and  Level ='c')"
    If Trim(SSdbLocation.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromlocation=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' AND LOCATIONCODE = '" & Trim(SSdbLocation.Tag) & "' and  Level ='LB')"

    
    rs.Source = rs.Source & " and  datediff(dd,'" & DTPFrom.value & "',transactiondate) > =0 and datediff(dd,transactiondate,'" & DTPTo.value & "') >=0   "
    If Trim(SSGrdTransactype.Text) <> "ALL" Then rs.Source = rs.Source & " and TransactionType = '" & SSGrdTransactype.Tag & "' "
    rs.Source = rs.Source & " Union"
    rs.Source = rs.Source & " select  FromCompany company, FromLocation location, FromUsChar USChar, FromStockType StockType, FromCamChar CamChar,"
    rs.Source = rs.Source & " TransactionNo,ItemNo,"
    rs.Source = rs.Source & " -cast(BaseCurUnitPrice as numeric(18,2)) * cast(Quantity as numeric(18,2)) 'Amount', TRANSACTIONDATE"
    rs.Source = rs.Source & " from inventoryFQA where npce_code ='" & deIms.NameSpace & "' "
    
    If Trim(SSdbCompany.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromcompany=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' and  Level ='c')"
    If Trim(SSdbLocation.Tag) <> "ALL" Then rs.Source = rs.Source & " and fromlocation=(select fqa from fqa where Namespace='" & deIms.NameSpace & "' and Companycode ='" & Trim(SSdbCompany.Tag) & "' AND LOCATIONCODE = '" & Trim(SSdbLocation.Tag) & "' and  Level ='LB')"
  
    
    rs.Source = rs.Source & " and  datediff(dd,'" & DTPFrom.value & "',transactiondate) > =0 and datediff(dd,transactiondate,'" & DTPTo.value & "') >=0   "
    If Trim(SSGrdTransactype.Text) <> "ALL" Then rs.Source = rs.Source & " and TransactionType = '" & SSGrdTransactype.Tag & "' "
    rs.Source = rs.Source & " order by TransactionNo,ItemNo, company, location  "

End If

rs.Open , deIms.cnIms

If rs.RecordCount > 0 Then

   If optInvtcomplete.value = True Then
   
            Call ExportToExcel1(rs, , , ProgressBar1, "InventoryFQA")
            
   ElseIf optInvtGeneralLedger.value = True Then
   
            Errcode = ExportToFlatFile(rs, , , ProgressBar1, "InventoryFQA")
            'If Errcode = 0 Then Dotheclosure
    End If
   
End If

ExportInventoryToExcel = True
Exit Function
ErrHand:


MsgBox "Errors occurred while trying to create a dump fo the data." & Err.Description, vbCritical, "Ims"

End Function


Public Function ExportInvoicetoExcel() As Boolean
Dim rs As New ADODB.Recordset
Dim Sql As String
On Error GoTo ErrHand

Call MDI_IMS.WriteStatus("Generating Invoice report ...", 1)

Sql = " select invd_ponumb 'PONUMB', invd_invcnumb 'INVOICE',  inv_invcdate 'DATE', invd_liitnumb 'ITEM NO',POI_COMM 'FOLIO#', invd_primreqdqty QUANTITY, invd_primuom UNIT, "
Sql = Sql & "  sup_name 'SUPPLIER NAME', curr_desc CURRENCY, FromCompany, FromLocation, FromUsChar, FromStockType, "
Sql = Sql & " FromCamChar, ToCompany ToLocation ,ToUsChar , ToStockType, ToCamChar,invd_unitpric 'INVOICE UNITPRICE' ,"
Sql = Sql & "  'USD' 'BaseCurrency' ,psys_extendedcurcode 'ExtendedCurrency' ,"
Sql = Sql & " 'BasePOUnitPrice'  ="
Sql = Sql & " case "
Sql = Sql & "    when curr_code= 'USD' then round(POI_unitprice,4)"
Sql = Sql & "    when curr_code <> 'USD' and (len(rtrim( isnull(psys_extendedcurcode,''))) = 0)  then  round(POI_unitprice / a.curd_value,4)"
Sql = Sql & "    else round(POI_unitprice / a.curd_value,4)"
Sql = Sql & " end  "
Sql = Sql & " ,"

Sql = Sql & " 'ExtendedPOUnitPrice'  ="
Sql = Sql & " case "
Sql = Sql & "    when curr_code ='USD' then round(POI_unitprice *  B.curd_value,4)"
Sql = Sql & "    when (len(rtrim( isnull(psys_extendedcurcode,''))) = 0) then null"
Sql = Sql & "    else  round(POI_unitprice,4)"
Sql = Sql & " end  "

Sql = Sql & " ,"
Sql = Sql & " 'BaseInvoiceUnitPrice'  ="
Sql = Sql & " case curr_code"
Sql = Sql & "    when 'USD' then round(invd_unitpric,4)"
Sql = Sql & "    else round(invd_unitpric / a.curd_value,4)"
Sql = Sql & " end  "
Sql = Sql & " ,"

Sql = Sql & " 'ExtendedInvoiceUnitPrice'  ="
Sql = Sql & " case "
Sql = Sql & "    when curr_code= 'USD' then round(invd_unitpric *  B.curd_value,4)"
Sql = Sql & "    when (len(rtrim( isnull(psys_extendedcurcode,''))) = 0) then null"
Sql = Sql & "    else  round(invd_unitpric,4)"
Sql = Sql & " end  "

Sql = Sql & " from invoicedetl "
Sql = Sql & " left join invoice on   inv_invcnumb = invd_invcnumb and inv_npecode = invd_npecode"
Sql = Sql & " left join PO on po_ponumb = invd_ponumb and po_npecode = invd_npecode"
Sql = Sql & " left join POITEM on poi_ponumb =po_ponumb and poi_liitnumb  = invd_liitnumb and poi_npecode =invd_npecode"
Sql = Sql & " left outer join supplier on sup_code = po_suppcode and sup_npecode = invd_npecode"
Sql = Sql & " left join Pesys on psys_npecode =invd_npecode "
Sql = Sql & " left outer join currency on curr_code = po_currcode and curr_npecode =invd_npecode"
Sql = Sql & " left outer join currencydetl A on A.curd_code = po_currcode and A.curd_npecode =invd_npecode and datediff(dd, A.curd_from ,inv_invcdate) > =0 and datediff(dd,A.curd_to ,inv_invcdate) <= 0"
Sql = Sql & " left outer join currencydetl B on B.curd_code =psys_extendedcurcode  and B.curd_npecode =invd_npecode and datediff(dd, B.curd_from ,inv_invcdate) > =0 and datediff(dd,B.curd_to ,inv_invcdate) <= 0"
Sql = Sql & " left join poFQA on Ponumb = invd_ponumb and ItemNo = invd_liitnumb"
Sql = Sql & " where invd_npecode ='" & deIms.NameSpace & "'"





'''''''''Sql = " select invd_ponumb 'PONUMB', invd_invcnumb 'INVOICE',  inv_invcdate 'DATE', invd_liitnumb 'ITEM NO, invd_primreqdqty QUANTITY, invd_primuom UNIT, "
'''''''''Sql = Sql & "  sup_name 'SUPPLIER NAME',curr_code, curr_desc , A.curd_value,FromCompany, FromLocation, FromUsChar, FromStockType, "
'''''''''Sql = Sql & " FromCamChar, ToCompany ToLocation ,ToUsChar , ToStockType, ToCamChar,invd_unitpric, psys_extendedcurcode , B.curd_value,"
'''''''''Sql = Sql & "  'USD' 'BaseCurrency' ,psys_extendedcurcode 'ExtendedCurrency' ,"
'''''''''Sql = Sql & " 'BasePOUnitPrice'  ="
'''''''''Sql = Sql & " case "
'''''''''Sql = Sql & "    when curr_code= 'USD' then round(POI_unitprice,4)"
'''''''''Sql = Sql & "    when curr_code <> 'USD' and (len(rtrim( isnull(psys_extendedcurcode,''))) = 0)  then  round(POI_unitprice / a.curd_value,4)"
'''''''''Sql = Sql & "    else round(POI_unitprice / a.curd_value,4)"
'''''''''Sql = Sql & " end  "
'''''''''Sql = Sql & " ,"
'''''''''
'''''''''Sql = Sql & " 'ExtendedPOUnitPrice'  ="
'''''''''Sql = Sql & " case "
'''''''''Sql = Sql & "    when curr_code ='USD' then round(POI_unitprice *  B.curd_value,4)"
'''''''''Sql = Sql & "    when (len(rtrim( isnull(psys_extendedcurcode,''))) = 0) then null"
'''''''''Sql = Sql & "    else  round(POI_unitprice,4)"
'''''''''Sql = Sql & " end  "
'''''''''
'''''''''Sql = Sql & " ,"
'''''''''Sql = Sql & " 'BaseInvoiceUnitPrice'  ="
'''''''''Sql = Sql & " case curr_code"
'''''''''Sql = Sql & "    when 'USD' then round(invd_unitpric,4)"
'''''''''Sql = Sql & "    else round(invd_unitpric / a.curd_value,4)"
'''''''''Sql = Sql & " end  "
'''''''''Sql = Sql & " ,"
'''''''''
'''''''''Sql = Sql & " 'ExtendedInvoiceUnitPrice'  ="
'''''''''Sql = Sql & " case "
'''''''''Sql = Sql & "    when curr_code= 'USD' then round(invd_unitpric *  B.curd_value,4)"
'''''''''Sql = Sql & "    when (len(rtrim( isnull(psys_extendedcurcode,''))) = 0) then null"
'''''''''Sql = Sql & "    else  round(invd_unitpric,4)"
'''''''''Sql = Sql & " end  "
'''''''''
'''''''''Sql = Sql & " from invoicedetl "
'''''''''Sql = Sql & " left join invoice on   inv_invcnumb = invd_invcnumb and inv_npecode = invd_npecode"
'''''''''Sql = Sql & " left join PO on po_ponumb = invd_ponumb and po_npecode = invd_npecode"
'''''''''Sql = Sql & " left join POITEM on poi_ponumb =po_ponumb and poi_liitnumb  = invd_liitnumb and poi_npecode =invd_npecode"
'''''''''Sql = Sql & " left outer join supplier on sup_code = po_suppcode and sup_npecode = invd_npecode"
'''''''''Sql = Sql & " left join Pesys on psys_npecode =invd_npecode "
'''''''''Sql = Sql & " left outer join currency on curr_code = po_currcode and curr_npecode =invd_npecode"
'''''''''Sql = Sql & " left outer join currencydetl A on A.curd_code = po_currcode and A.curd_npecode =invd_npecode and datediff(dd, A.curd_from ,inv_invcdate) > =0 and datediff(dd,A.curd_to ,inv_invcdate) <= 0"
'''''''''Sql = Sql & " left outer join currencydetl B on B.curd_code =psys_extendedcurcode  and B.curd_npecode =invd_npecode and datediff(dd, B.curd_from ,inv_invcdate) > =0 and datediff(dd,B.curd_to ,inv_invcdate) <= 0"
'''''''''Sql = Sql & " left join poFQA on Ponumb = invd_ponumb and ItemNo = invd_liitnumb"
'''''''''Sql = Sql & " where invd_npecode ='" & deIms.NameSpace & "'"




If Trim(SSGrdSupplier.Text) <> "ALL" Then Sql = Sql & " and sup_code = '" & Trim(SSGrdSupplier.Tag) & "'"
If Trim(SSGrdInvoiceNo.Text) <> "ALL" Then Sql = Sql & " and invd_invcnumb = '" & Trim(SSGrdInvoiceNo.Text) & "'"

rs.Source = Sql

rs.Open , deIms.cnIms

If rs.RecordCount > 0 Then Call ExportToExcel1(rs, , , ProgressBar1, "SupplierInvoiceFQA")

Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the combo boxes." & Err.Description, vbCritical, "Ims"
Err.Clear
End Function

Public Function ExportToExcel1(Optional RsRecord As ADODB.Recordset, Optional Arr As Variant, Optional ArrColumnNames As Variant, Optional ProgressBar1 As Progressbar, Optional Filename As String)

Dim Report As Excel.Application
Dim i As Integer
Dim j As Integer
Dim Sa As Scripting.FileSystemObject
Dim Fld As ADODB.Field
Dim x As Integer
Dim y As Integer
Dim Incr As Integer
Dim m As Integer
Dim subject As String
Dim onlyname As String
    Set Report = New Excel.Application
    Set Sa = New Scripting.FileSystemObject
    
    m = Rnd(20)
    subject = Filename
    
    onlyname = "Report-" & Filename & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".XLS"
    
    Filename = ConnInfo.EmailOutFolder & "Report-" & Filename & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".XLS"
    
    If Sa.FileExists(Filename) = False Then Sa.CreateTextFile Filename
    
    Report.Workbooks.OpenText Filename
    Report.SheetsInNewWorkbook = 1
     
    With Report
       
    If RsRecord Is Nothing Then
    
    'This is executed when an array is passed in here.
    
        x = UBound(ArrColumnNames)
        
                  For i = 0 To x
               
                   .Cells(1, i + 1) = ArrColumnNames(i)
               '     .ActiveCell.NumberFormat = "text"
                    
                  Next i
                  
        .activeCELL.EntireRow.Font.Bold = True
    
    
        x = UBound(Arr, 1)
        
        y = UBound(Arr, 2)
        
       Incr = y / 10
        
            For j = 0 To y
            .Rows(j).NumberFormat = "Text"
              For i = 0 To x
               
                  .Cells(j + 2, i + 1) = Arr(i, j)
                         
              Next i
              
                If j > 0 And j Mod Incr = 0 Then
                 Call IncrementProgreesBar(1, ProgressBar1)
                End If
                    
            Next j
             
             
     ElseIf RsRecord Is Nothing = False Then
        
                
                If RsRecord.RecordCount = 0 Then GoTo Gohome
                
                RsRecord.MoveFirst
                'This is executed when a recordset is passed
                    
                    i = 1
                    j = 1
                    
                'Writing the names of the Fields
                    
                        For Each Fld In RsRecord.Fields
                                   
                                .Cells(i, j) = Fld.Name & ""
                           '     .ActiveCell.NumberFormat = "text"
                                
                                 j = j + 1
                                
                                
                         Next Fld
                    
                    i = i + 1
                    
                 Incr = RsRecord.RecordCount / 10
                    
                   Do While Not RsRecord.EOF
                    
                        j = 1
                    
                         For Each Fld In RsRecord.Fields
                         
                                .Cells(i, j) = Fld.value & ""
                              '   .ActiveCell.EntireRow.NumberFormat = "text"
                                 j = j + 1
                         
                         Next Fld
                        
                        If Incr > 0 Then
                        
                           If i > 0 And i Mod Incr = 0 Then
                             Call IncrementProgreesBar(1, ProgressBar1)
                            End If
                        End If
                        
                        i = i + 1
                        RsRecord.MoveNext
                    
                    Loop
        
   End If
        
        
        
    End With
    
Gohome:
    
    SetProgressbar1ToMax ProgressBar1
        
        Report.Workbooks.Item(1).SAVE
        Report.Workbooks.Item(1).Saved = True
        Report.Workbooks.Item(1).Close
        
        Dim Attachment(0) As String
        Dim Recepients() As String
        
        
        Attachment(0) = onlyname
    
        'If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
    
        
        For i = 0 To SSGRDRecepients.Rows - 1
            
            SSGRDRecepients.row = i
            ReDim Preserve Recepients(i)
            Recepients(i) = SSGRDRecepients.Columns(0).Text
            
        Next i
        
        
        Call WriteParameterFileEmail(Attachment, Recepients, subject, "PECTEN CAMEROON COMPANY", "!ATTENTION")
        
End Function

Public Function IncrementProgreesBar(value As Integer, ProgressBar1 As Progressbar)

ProgressBar1.Max = 10

If ProgressBar1.Visible = False Then ProgressBar1.Visible = True

If ProgressBar1.value < ProgressBar1.Max Then

        ProgressBar1.value = ProgressBar1.value + value
        
ElseIf ProgressBar1.value = ProgressBar1.Max Or ProgressBar1.value > ProgressBar1.Max Then

        ProgressBar1.value = 1
        
End If

End Function

Public Sub SetProgressbar1ToMax(ProgressBar1 As Progressbar)
Dim i As Integer

For i = ProgressBar1.value To ProgressBar1.Max - 1

    ProgressBar1.value = ProgressBar1.value + 1
    
Next
    
ProgressBar1.value = 0



'Progressbar1.Visible = False

End Sub

Private Sub SSGrdTransactype_LostFocus()
Call NormalBackground(SSGrdTransactype)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab

    Case 0
          CmdExport.Default = True
          CmdExport.SetFocus
    Case 1
            
          cmd_Add.Default = True
          Call PopulateEmail
          TxtEmail.SetFocus
End Select

End Sub


Public Function PopulateInvoiceCombo() As Boolean
Dim Rsinvoice As New ADODB.Recordset
On Error GoTo ErrHand

Rsinvoice.Source = "select inv_invcnumb from invoice"
Rsinvoice.Source = Rsinvoice.Source & "  where inv_npecode  ='" & deIms.NameSpace & "'"
If SSGrdSupplier.Tag <> "ALL" Then Rsinvoice.Source = Rsinvoice.Source & "  and inv_ponumb in ( select po_ponumb from po where po_suppcode ='" & SSGrdSupplier.Tag & "' and PO_npecode ='" & deIms.NameSpace & "') "
Rsinvoice.Source = Rsinvoice.Source & " and datediff(dd, '" & DTPFrom.value & "',inv_invcdate) > =  0 and datediff(dd, inv_invcdate,'" & DTPTo.value & "') > =0"
Rsinvoice.Source = Rsinvoice.Source & " order by inv_invcnumb"

Rsinvoice.Open , deIms.cnIms

SSGrdInvoiceNo.RemoveAll

SSGrdInvoiceNo.AddItem "ALL" & vbTab & "ALL"

Do While Not Rsinvoice.EOF

    SSGrdInvoiceNo.AddItem Rsinvoice!inv_invcnumb
    
    Rsinvoice.MoveNext
    
Loop

PopulateInvoiceCombo = True
Exit Function
ErrHand:

MsgBox "Errors occurred while trying to poupulate the Invoice combo boxes." & Err.Description, vbCritical, "Ims"
Err.Clear
End Function

Public Function PopulatePOnumbCombo() As Boolean
Dim rsPO As New ADODB.Recordset

On Error GoTo ErrHand

If Trim(SSGrdPODel.Text) > 0 And Trim(SSGrdPOInvt.Text) > 0 And Trim(SSGrdPOShip.Text) > 0 And Trim(SSGrdPODoc.Text) > 0 Then

        rsPO.Source = "select Po_ponumb from po where po_stas not in ('OH','CL') and po_npecode ='" & deIms.NameSpace & "' and po_compcode ='" & Trim(SSdbCompany.Tag) & "' and po_invloca ='" & Trim(SSdbLocation.Tag) & "' and  datediff(dd,'" & DTPFrom.value & "',po_date) > =0 and  datediff(dd,po_date,'" & DTPTo.value & "') > =0 "
        
        If SSGrdPODoc.Tag <> "ALL" Then rsPO.Source = rsPO.Source & " AND PO_DOCUTYPE ='" & SSGrdPODoc.Tag & "'"
        If SSGrdPODel.Tag <> "ALL" Then rsPO.Source = rsPO.Source & " AND po_stasdelv ='" & SSGrdPODel.Tag & "'"
        If SSGrdPOShip.Tag <> "ALL" Then rsPO.Source = rsPO.Source & " AND po_stasship ='" & SSGrdPOShip.Tag & "'"
        If SSGrdPOInvt.Tag <> "ALL" Then rsPO.Source = rsPO.Source & " AND po_stasinvt='" & SSGrdPOInvt.Tag & "'"
        rsPO.Source = rsPO.Source & " order by Po_ponumb"
        
        rsPO.Open , deIms.cnIms
        
        SSGrdPonumb.RemoveAll
        
        SSGrdPonumb.AddItem "ALL" & vbTab & "ALL"
        
        Do While Not rsPO.EOF
        
            SSGrdPonumb.AddItem rsPO!po_ponumb
        
            rsPO.MoveNext
            
        Loop
        
Else

  '  MsgBox "Please make sure that all Delivery, Shipping, Inventory and Document values are filled." & Err.Description, vbInformation, "Ims"

End If

Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the Ponumb combo boxes." & Err.Description, vbCritical, "Ims"
End Function

Private Sub TxtEmail_GotFocus()
 Call HighlightBackground(TxtEmail)
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = vbEnter Then Call cmd_Add_Click
End Sub

Public Function ExportToFlatFile(Optional RsRecord As ADODB.Recordset, Optional Arr As Variant, Optional ArrColumnNames As Variant, Optional ProgressBar1 As Progressbar, Optional Filename As String) As Integer
Dim Report As Scripting.TextStream
Dim i As Integer
Dim j As Integer
Dim Sa As Scripting.FileSystemObject
Dim Fld As ADODB.Field
Dim x As Integer
Dim y As Integer
Dim Incr As Integer
Dim m As Integer
Dim subject As String
Dim onlyname As String
Dim str As String

On Error GoTo ErrHand

    Set Sa = New Scripting.FileSystemObject
    
    subject = Filename
    
    onlyname = "Report-" & Filename & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
    
    Filename = ConnInfo.EmailOutFolder & "Report-" & Filename & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
    
    If Sa.FileExists(Filename) = False Then Set Report = Sa.CreateTextFile(Filename)
    
    
    With Report
       
    If RsRecord Is Nothing Then
    
    'This is executed when an array is passed in here.
    
        x = UBound(ArrColumnNames)
        
                  For i = 0 To x
               
                        str = srt & ArrColumnNames(i) & vbTab
                    
                  Next i
    
    
        x = UBound(Arr, 1)
        
        y = UBound(Arr, 2)
        
       Incr = y / 10
        
            For j = 0 To y

              For i = 0 To x
               
                  str = srt & Arr(i, j) & vbTab
                         
              Next i
              
                If j > 0 And j Mod Incr = 0 Then
                 
                  Call IncrementProgreesBar(1, ProgressBar1)
                
                End If
                    
            Next j
             
             
     ElseIf RsRecord Is Nothing = False Then
        
                
                If RsRecord.RecordCount = 0 Then GoTo Gohome
                
                RsRecord.MoveFirst
                'This is executed when a recordset is passed
                    
                        For Each Fld In RsRecord.Fields
                                   
                                str = str & Fld.Name & "" & vbTab
            
                        Next Fld
                    
                        str = str & vbCrLf
                    
                   Incr = RsRecord.RecordCount / 10
                    
                   Do While Not RsRecord.EOF
                    
                         For Each Fld In RsRecord.Fields
                         
                                str = str & Fld.value & "" & vbTab
            
                         Next Fld
                        
                        str = str & vbCrLf
                        
                        If Incr > 0 Then
                        
                           If i > 0 And i Mod Incr = 0 Then Call IncrementProgreesBar(1, ProgressBar1)
                        
                        End If
                        
                        i = i + 1
                        
                        RsRecord.MoveNext
                    
                   Loop
   End If
        
   End With
    
Gohome:
    
    Report.Write str
    Report.Close
    
    SetProgressbar1ToMax ProgressBar1
        
        Dim Attachment(0) As String
        Dim Recepients() As String
        
        
        Attachment(0) = onlyname
    
        'If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
    
        
        For i = 0 To SSGRDRecepients.Rows - 1
            
            SSGRDRecepients.row = i
            ReDim Preserve Recepients(i)
            Recepients(i) = SSGRDRecepients.Columns(0).Text
            
        Next i
        
        
If WriteParameterFileEmail(Attachment, Recepients, subject, "PECTEN CAMEROON COMPANY", "!ATTENTION") <> 1 Then GoTo ErrHand

ExportToFlatFile = 1

Exit Function

ErrHand:

MsgBox "Errros occurred while trying to export the data to flat files. Error description :" & Err.Description, "Imswin", vbCritical

Err.Clear

End Function

Private Sub TxtEmail_LostFocus()
Call NormalBackground(TxtEmail)
End Sub
