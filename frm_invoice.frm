VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#8.0#0"; "LRNavigators.ocx"
Begin VB.Form frm_invoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Invoice"
   ClientHeight    =   6630
   ClientLeft      =   4650
   ClientTop       =   5115
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   7920
   Tag             =   "02050700"
   Begin TabDlg.SSTab sst_PO 
      Height          =   5745
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   10134
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "Invoice"
      TabPicture(0)   =   "frm_invoice.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbl_Company"
      Tab(0).Control(1)=   "lbl_PO"
      Tab(0).Control(2)=   "lbl_dateinvoice"
      Tab(0).Control(3)=   "lbl_Date(0)"
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(5)=   "lbl_User(1)"
      Tab(0).Control(6)=   "Label1(2)"
      Tab(0).Control(7)=   "Label1(3)"
      Tab(0).Control(8)=   "lbl_Date(1)"
      Tab(0).Control(9)=   "lbl_User(0)"
      Tab(0).Control(10)=   "Label1(0)"
      Tab(0).Control(11)=   "cbo_PO"
      Tab(0).Control(12)=   "Frame1(0)"
      Tab(0).Control(13)=   "Frame1(1)"
      Tab(0).Control(14)=   "ssdcboInvoiceNumb"
      Tab(0).Control(15)=   "txt_dateinvoice"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Line Items Invoiced"
      TabPicture(1)   =   "frm_invoice.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtInvcPrice"
      Tab(1).Control(1)=   "txtInvcQnty"
      Tab(1).Control(2)=   "ssdcbolineitem"
      Tab(1).Control(3)=   "lblPrice(3)"
      Tab(1).Control(4)=   "Label1(32)"
      Tab(1).Control(5)=   "lblPrice(1)"
      Tab(1).Control(6)=   "lbl_Icost(1)"
      Tab(1).Control(7)=   "lbl_Cost"
      Tab(1).Control(8)=   "lbl_Icost(0)"
      Tab(1).Control(9)=   "lbl_Total"
      Tab(1).Control(10)=   "lblAmount(2)"
      Tab(1).Control(11)=   "lblUnit(1)"
      Tab(1).Control(12)=   "Label1(28)"
      Tab(1).Control(13)=   "lbl_Prime(1)"
      Tab(1).Control(14)=   "lbl_qtpri(1)"
      Tab(1).Control(15)=   "lblAmount(1)"
      Tab(1).Control(16)=   "lblUnit(0)"
      Tab(1).Control(17)=   "Label1(26)"
      Tab(1).Control(18)=   "lblPrice(2)"
      Tab(1).Control(19)=   "lblPrice(0)"
      Tab(1).Control(20)=   "lblAmount(0)"
      Tab(1).Control(21)=   "lbl_Commodity"
      Tab(1).Control(22)=   "lblComm"
      Tab(1).Control(23)=   "lbl_qtpri(0)"
      Tab(1).Control(24)=   "lbl_Prime(0)"
      Tab(1).Control(25)=   "lbl_LI"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Remarks"
      TabPicture(2)   =   "frm_invoice.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "rtbRemarks"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recipients"
      TabPicture(3)   =   "frm_invoice.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Line1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lbl_New"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lbl_Recipients"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "dgRecepients"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fra_FaxSelect"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmd_Add"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmd_Remove"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txt_Recipient"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "ssdbRecepientList"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   2115
         Left            =   1740
         TabIndex        =   79
         Top             =   480
         Width           =   5655
         _Version        =   196617
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
         stylesets(0).Picture=   "frm_invoice.frx":0070
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
         stylesets(1).Picture=   "frm_invoice.frx":008C
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
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
         Columns(0).Width=   9419
         Columns(0).Caption=   "Recipients"
         Columns(0).Name =   "Recipients"
         Columns(0).DataField=   "Recipients"
         Columns(0).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   9975
         _ExtentY        =   3731
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
      Begin VB.TextBox txt_Recipient 
         Height          =   288
         Left            =   1740
         TabIndex        =   75
         Top             =   3090
         Width           =   5670
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   165
         TabIndex        =   74
         Top             =   2055
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   165
         TabIndex        =   73
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Frame fra_FaxSelect 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   120
         TabIndex        =   70
         Top             =   3465
         Width           =   1575
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   60
            TabIndex        =   72
            Top             =   165
            Width           =   1460
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   60
            TabIndex        =   71
            Top             =   660
            Width           =   1460
         End
      End
      Begin VB.TextBox txt_dateinvoice 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69390
         TabIndex        =   69
         Top             =   435
         Width           =   1845
      End
      Begin VB.TextBox rtbRemarks 
         Enabled         =   0   'False
         Height          =   5055
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   480
         Width           =   7120
      End
      Begin VB.TextBox txtInvcPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69240
         TabIndex        =   51
         Top             =   2520
         Width           =   1500
      End
      Begin VB.TextBox txtInvcQnty 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72900
         TabIndex        =   50
         Top             =   1590
         Width           =   930
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboInvoiceNumb 
         Height          =   315
         Left            =   -73200
         TabIndex        =   38
         Top             =   765
         Width           =   2055
         DataFieldList   =   "Column 0"
         AllowNull       =   0   'False
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         stylesets(0).Picture=   "frm_invoice.frx":00A8
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
         stylesets(1).Picture=   "frm_invoice.frx":00C4
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3281
         Columns(0).Caption=   "Invoice Number"
         Columns(0).Name =   "InvcNumber"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   15
         Columns(0).HeadStyleSet=   "ColHeader"
         Columns(0).StyleSet=   "RowFont"
         Columns(1).Width=   2434
         Columns(1).Caption=   "PO Number"
         Columns(1).Name =   "PoNumb"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).HeadStyleSet=   "ColHeader"
         Columns(1).StyleSet=   "RowFont"
         Columns(2).Width=   2593
         Columns(2).Caption=   "Invoice Date"
         Columns(2).Name =   "invcDate"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   7
         Columns(2).FieldLen=   256
         Columns(2).HeadStyleSet=   "ColHeader"
         Columns(2).StyleSet=   "RowFont"
         _ExtentX        =   3619
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Supplier Information"
         Height          =   2175
         Index           =   1
         Left            =   -74880
         TabIndex        =   21
         Top             =   3420
         Width           =   7340
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr1"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   21
            Left            =   1260
            TabIndex        =   37
            Top             =   1740
            Width           =   2700
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_ctry"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   17
            Left            =   1260
            TabIndex        =   36
            Top             =   1380
            Width           =   2700
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_city"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   16
            Left            =   1260
            TabIndex        =   35
            Top             =   1020
            Width           =   2700
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr1"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   15
            Left            =   1260
            TabIndex        =   34
            Top             =   660
            Width           =   2700
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_name"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   14
            Left            =   1260
            TabIndex        =   33
            Top             =   300
            Width           =   5940
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone"
            Height          =   315
            Index           =   22
            Left            =   120
            TabIndex        =   32
            Top             =   1740
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   315
            Index           =   27
            Left            =   120
            TabIndex        =   31
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address1"
            Height          =   315
            Index           =   25
            Left            =   120
            TabIndex        =   30
            Top             =   660
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address2"
            Height          =   315
            Index           =   23
            Left            =   4080
            TabIndex        =   29
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr2"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   18
            Left            =   5100
            TabIndex        =   28
            Top             =   660
            Width           =   2100
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            Height          =   315
            Index           =   24
            Left            =   120
            TabIndex        =   27
            Top             =   1020
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   315
            Index           =   12
            Left            =   4080
            TabIndex        =   26
            Top             =   1020
            Width           =   1065
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_stat"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   19
            Left            =   5100
            TabIndex        =   25
            Top             =   1020
            Width           =   1560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
            Height          =   315
            Index           =   13
            Left            =   4080
            TabIndex        =   24
            Top             =   1380
            Width           =   1065
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_zipc"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   20
            Left            =   5100
            TabIndex        =   23
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   22
            Top             =   1380
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Transcation Order"
         Height          =   1395
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   1920
         Width           =   7340
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Buyer"
            Height          =   315
            Index           =   30
            Left            =   3900
            TabIndex        =   20
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataMember      =   "GETPONUMBERSFORRECEPTION"
            DataSource      =   "deIms"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   4980
            TabIndex        =   19
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            Height          =   315
            Index           =   31
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   1600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_name"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   1700
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone"
            Height          =   315
            Index           =   29
            Left            =   3900
            TabIndex        =   16
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr1"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   4980
            TabIndex        =   15
            Top             =   600
            Width           =   2220
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Issued"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   14
            Top             =   660
            Width           =   1600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_city"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   1700
            TabIndex        =   13
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Requested"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   1700
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.ComboBox cbo_PO 
         Height          =   315
         Left            =   -73200
         TabIndex        =   3
         Top             =   435
         Width           =   2052
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   2055
         Left            =   1740
         TabIndex        =   76
         Top             =   3480
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3625
         _Version        =   393216
         Enabled         =   0   'False
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
               ColumnWidth     =   3509.858
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcbolineitem 
         Height          =   315
         Left            =   -72840
         TabIndex        =   80
         Top             =   600
         Width           =   3135
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
         Columns(0).Width=   2540
         Columns(0).Caption=   "Line#"
         Columns(0).Name =   "Line#"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   2
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Quantity"
         Columns(1).Name =   "Quantity"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         Columns(2).Width=   5556
         Columns(2).Caption=   "Desricption"
         Columns(2).Name =   "Desricption"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -73200
         TabIndex        =   41
         Top             =   1095
         Width           =   2055
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   165
         TabIndex        =   78
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label lbl_New 
         Caption         =   "New"
         Height          =   420
         Left            =   105
         TabIndex        =   77
         Top             =   3150
         Width           =   1620
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   7872
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -72900
         TabIndex        =   66
         Top             =   3180
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Difference"
         Height          =   315
         Index           =   32
         Left            =   -74880
         TabIndex        =   65
         Top             =   3180
         Width           =   2000
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -69240
         TabIndex        =   64
         Top             =   2850
         Width           =   1500
      End
      Begin VB.Label lbl_Icost 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Invoice Price"
         Height          =   315
         Index           =   1
         Left            =   -71115
         TabIndex        =   63
         Top             =   2850
         Width           =   1800
      End
      Begin VB.Label lbl_Cost 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O Unit Price"
         Height          =   315
         Left            =   -74880
         TabIndex        =   62
         Top             =   2520
         Width           =   2000
      End
      Begin VB.Label lbl_Icost 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Unit Price"
         Height          =   315
         Index           =   0
         Left            =   -71115
         TabIndex        =   61
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label lbl_Total 
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.O Price"
         Height          =   315
         Left            =   -74880
         TabIndex        =   60
         Top             =   2850
         Width           =   2000
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -72900
         TabIndex        =   59
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label lblUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -69240
         TabIndex        =   58
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit"
         Height          =   195
         Index           =   28
         Left            =   -70500
         TabIndex        =   57
         Top             =   2010
         Width           =   1200
      End
      Begin VB.Label lbl_Prime 
         AutoSize        =   -1  'True
         Caption         =   "of "
         Height          =   315
         Index           =   1
         Left            =   -71835
         TabIndex        =   56
         Top             =   1920
         Width           =   180
      End
      Begin VB.Label lbl_qtpri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity (Secondary)"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   55
         Top             =   1920
         Width           =   2000
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -71580
         TabIndex        =   54
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label lblUnit 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -69240
         TabIndex        =   53
         Top             =   1590
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit"
         Height          =   195
         Index           =   26
         Left            =   -70500
         TabIndex        =   52
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -72900
         TabIndex        =   49
         Top             =   2850
         Width           =   1500
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -72900
         TabIndex        =   48
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -71580
         TabIndex        =   47
         Top             =   1590
         Width           =   930
      End
      Begin VB.Label lbl_Commodity 
         BackStyle       =   0  'Transparent
         Caption         =   "Commodity"
         Height          =   315
         Left            =   -74880
         TabIndex        =   46
         Top             =   1320
         Width           =   2000
      End
      Begin VB.Label lblComm 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -72900
         TabIndex        =   45
         Top             =   1260
         Width           =   2250
      End
      Begin VB.Label lbl_qtpri 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity (Primary)"
         Height          =   315
         Index           =   0
         Left            =   -74880
         TabIndex        =   44
         Top             =   1590
         Width           =   2000
      End
      Begin VB.Label lbl_Prime 
         AutoSize        =   -1  'True
         Caption         =   "of "
         Height          =   315
         Index           =   0
         Left            =   -71835
         TabIndex        =   43
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label lbl_User 
         Caption         =   "Created By"
         Height          =   225
         Index           =   0
         Left            =   -74856
         TabIndex        =   42
         Top             =   1190
         Width           =   1700
      End
      Begin VB.Label lbl_Date 
         Caption         =   "Date Modified"
         Height          =   225
         Index           =   1
         Left            =   -70890
         TabIndex        =   40
         Top             =   1185
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -69390
         TabIndex        =   39
         Top             =   1095
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -73200
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lbl_User 
         Caption         =   "Modified By"
         Height          =   225
         Index           =   1
         Left            =   -74856
         TabIndex        =   8
         Top             =   1530
         Width           =   1700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -69390
         TabIndex        =   7
         Top             =   765
         Width           =   1845
      End
      Begin VB.Label lbl_Date 
         Caption         =   "Date Created"
         Height          =   225
         Index           =   0
         Left            =   -70890
         TabIndex        =   6
         Top             =   855
         Width           =   1500
      End
      Begin VB.Label lbl_dateinvoice 
         Caption         =   "Date of invoice"
         Height          =   225
         Left            =   -70890
         TabIndex        =   5
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label lbl_PO 
         Caption         =   "Purchase Order #"
         Height          =   228
         Left            =   -74856
         TabIndex        =   4
         Top             =   522
         Width           =   1700
      End
      Begin VB.Label lbl_LI 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Item #"
         Height          =   315
         Left            =   -74880
         TabIndex        =   2
         Top             =   585
         Width           =   2000
      End
      Begin VB.Label lbl_Company 
         Caption         =   "Vendor Invoice"
         Height          =   225
         Left            =   -74856
         TabIndex        =   1
         Top             =   858
         Width           =   1700
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8016
         Y1              =   -435
         Y2              =   -435
      End
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   1500
      TabIndex        =   67
      Top             =   6000
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frm_invoice.frx":00E0
      CancelVisible   =   0   'False
      EMailVisible    =   -1  'True
      CloseEnabled    =   0   'False
      PrintEnabled    =   0   'False
      NewEnabled      =   0   'False
      SaveEnabled     =   0   'False
      NextEnabled     =   0   'False
      LastEnabled     =   0   'False
      FirstEnabled    =   0   'False
      PreviousEnabled =   0   'False
      EditEnabled     =   -1  'True
   End
End
Attribute VB_Name = "frm_invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fm As FormMode
Dim inv As imsInvoiceDetl
Dim InvcD As New imsInvoiceDetls
Dim WithEvents poInfo As imsGetPOInfo
Attribute poInfo.VB_VarHelpID = -1

Dim rsReceptList As ADODB.Recordset
Dim WithEvents remk As ADODB.Recordset
Attribute remk.VB_VarHelpID = -1

'call function addrecepients to add a record
'if same recepient number clear it

Private Sub cmd_Add_Click()
    If Len(Trim$(txt_Recipient)) Then
    
      If InStr(1, txt_Recipient, "@") Then
           txt_Recipient = UCase(txt_Recipient)
           If InStr(1, txt_Recipient, "INTERNET!") = 0 Then txt_Recipient = ("INTERNET!" & txt_Recipient)
       
       Else
           txt_Recipient = ("FAX!" & txt_Recipient)
       End If
    
        Call AddRecepients(txt_Recipient)
        txt_Recipient = ""
    Else
        dgRecepients_DblClick
    End If

End Sub

'delete one recepient number

Private Sub cmd_Remove_Click()
On Error Resume Next

    rsReceptList.Delete
    If Err Then Err.Clear
End Sub

'function get recepient fax numbers

Public Sub AddRecepients(Recepient As String)

    If Len(Recepient) = 0 Then Exit Sub
    If opt_FaxNum Then
        If UCase(Left$(Recepient, 4)) <> "FAX!" Then
            Recepient = "FAX!" & Recepient
        End If
    End If
    
    If rsReceptList Is Nothing Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        
        rsReceptList.Open: Set ssdbRecepientList.DataSource = rsReceptList
    End If
        
    If Not IsRecipientInList(Recepient) Then
        Call rsReceptList.AddNew(Array("Recipients"), Array(Recepient))
        
        rsReceptList.Update
        rsReceptList.UpdateBatch adAffectCurrent
    End If
End Sub

'function check recepient number exist or not if the number exist
'show message
Private Function IsRecipientInList(RecepientName As String) As Boolean
On Error Resume Next
Dim BK As Variant
    
    
    If rsReceptList.RecordCount = 0 Then Exit Function
    If Not (rsReceptList.EOF Or rsReceptList.BOF) Then BK = rsReceptList.Bookmark
    
    rsReceptList.MoveFirst
    Call rsReceptList.Find("Recipients = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
    
    If Not (rsReceptList.EOF) Then
        
        If opt_Email Then
        
            'Modified by Juan (9/12/2000) for Multilingual
            msg1 = translator.Trans("M00076") 'J added
            MsgBox IIf(msg1 = "", "Email Address Already in list", msg1) 'J modified
            '---------------------------------------------
            
        ElseIf opt_FaxNum Then
        
            'Modified by Juan (9/12/2000) for Multilingual
            msg1 = translator.Trans("M00077") 'J added
            MsgBox IIf(msg1 = "", "Fax Number Already in list", msg1) 'J modified
            '---------------------------------------------
            
        End If
        IsRecipientInList = True
    End If
    
    rsReceptList.Bookmark = BK
    If Err Then Err.Clear
End Function

'call addrecepients function

Private Sub dgRecepients_DblClick()
    If dgRecepients.ApproxCount > 0 Then _
        Call AddRecepients(dgRecepients.Columns(1).text)
End Sub

'before send email get report parameter
'and application path

Private Sub NavBar1_OnEMailClick()
Dim Params(2) As String
Dim rptinf As RPTIFileInfo

On Error Resume Next


    With rptinf
    
    
        .ReportFileName = ReportPath & "invoice.rpt"
        
        Params(0) = "namespace=" & deIms.NameSpace
        Params(1) = "invnumb=" & ssdcboInvoiceNumb.text
        Params(2) = "ponumb=" & cbo_PO
        .Parameters = Params
    End With
    
    BeforePrint
    Params(0) = ""
    Call WriteRPTIFile(rptinf, Params(0))
    Call SendEmailAndFax(rsReceptList, "Recipients", "Invoice", "", Params(0))
    
    
    Set rsReceptList = Nothing
    Set ssdbRecepientList.DataSource = Nothing
End Sub

'clear form

Private Sub NavBar1_OnNewClick()
        cbo_PO = ""
        ssdcboInvoiceNumb = ""
        Call ClearLineitem
End Sub

'set back ground color

Private Sub opt_Email_GotFocus()
    Call HighlightBackground(opt_Email)
End Sub

'set back ground color

Private Sub opt_Email_LostFocus()
    Call NormalBackground(opt_Email)
End Sub

'call store procedure to get email address

Private Sub opt_Email_Click()
Dim co As MSDataGridLib.Column

    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (9/12/2000) for Multilinguaje
    msg1 = translator.Trans("L00121") 'J added
    co.Caption = IIf(msg1 = "", "Email Address", msg1) 'J modified
    '----------------------------------------------
    
    co.DataField = "phd_mail"
    
    dgRecepients.Columns(0).DataField = "phd_name"
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atEmail)
End Sub

'call store procedure to get fax numbers

Private Sub opt_FaxNum_Click()
On Error Resume Next
Dim co As MSDataGridLib.Column
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (9/12/2000) for Multilinguaje
    msg1 = translator.Trans("L00122") 'J added
    co.Caption = IIf(msg1 = "", "Fax Number", msg1) 'J modified
    '----------------------------------------------
    
    co.DataField = "phd_faxnumb"
    
    dgRecepients.Columns(0).DataField = "phd_name"
     
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atFax)
End Sub

'set back ground color

Private Sub opt_FaxNum_GotFocus()
    Call HighlightBackground(opt_FaxNum)
End Sub

'set back ground color

Private Sub opt_FaxNum_LostFocus()
    Call NormalBackground(opt_FaxNum)
End Sub

'set navbar button

Private Sub ssdcboInvoiceNumb_Change()
    If Len(ssdcboInvoiceNumb.text) = 0 Then
        NavBar1.EMailEnabled = False
        NavBar1.PrintEnabled = False
        NavBar1.SaveEnabled = False
    End If
        
End Sub

'set inventory combo allow input data

Private Sub ssdcboInvoiceNumb_DropDown()
    If NavBar1.NewEnabled = False Then ssdcboInvoiceNumb.AllowInput = False
End Sub

'when click tab, set date, fill line item to form, and set navbar button

Private Sub cbo_PO_Click()
On Error Resume Next
Dim STR As String
Dim ctl As Control
    
    'If Len(Trim$(ssdcboInvoiceNumb.Text)) Then _
        SaveAll
        
        
    cbo_PO.Tag = cbo_PO
    
    If NavBar1.NewEnabled Then
        If Len(Trim$(cbo_PO)) Then Call DisableAllControls(False)
        
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
    
        rsReceptList.Open
        Set ssdbRecepientList.DataSource = rsReceptList
    End If
        
    fm = mdCreation
    If poInfo.PO_Number <> cbo_PO.text Then _
        Call poInfo.GetValues(cbo_PO)
        
    STR = Format(Date, "mm/dd/yyyy")
    Label1(1) = STR
    Label1(3) = STR
    txt_dateinvoice = STR
    
    GetInvoices
    FillLineItems
    Set inv = Nothing
    txt_dateinvoice.Enabled = True
    If Err Then MsgBox Err.Description:  Err.Clear
    
    lblComm = ""
    txtInvcQnty = ""
    lblPrice(0) = ""
    lblPrice(1) = ""
    lblPrice(2) = ""
    lblPrice(3) = ""
    
    txtInvcQnty = ""
    lblAmount(0) = ""
    lblAmount(2) = ""
    txtInvcPrice = ""
    txtInvcPrice = ""
    rtbRemarks.text = ""
    ssdcbolineitem.text = ""
    rtbRemarks.Enabled = True
    txtInvcQnty.Enabled = False
    txtInvcPrice.Enabled = False
    ssdcboInvoiceNumb.text = ""
    If Err Then Err.Clear
End Sub

'SQL statement get po information,populate data to form
'and set navbar button


Private Sub Form_Load()
Dim Rs As ADODB.Recordset

    'Added by Juan (9/12/2000) for Multilingual
    Call translator.Translate_Forms("frm_invoice")
    '------------------------------------------

    Set Rs = New ADODB.Recordset
    Set rsReceptList = New ADODB.Recordset
    
    Rs.LockType = adLockReadOnly
    Rs.CursorLocation = adUseServer
    Rs.CursorType = adOpenForwardOnly
    Rs.ActiveConnection = deIms.cnIms
    
    Rs.Source = "Select po_ponumb, po_suppcode,po_buyr from po"
    Rs.Source = Rs.Source & " where po_stas = 'OP' and po_npecode = '"
    Rs.Source = Rs.Source & deIms.NameSpace & "'"
    Rs.Open

    Call PopuLateFromRecordSet(cbo_PO, Rs, "po_ponumb", False)
    
    Set poInfo = New imsGetPOInfo
    poInfo.NameSpace = deIms.NameSpace
    Set poInfo.Connection = deIms.cnIms
    Call DisableButtons(Me, NavBar1)
    NavBar1.NewEnabled = True
    NavBar1.CloseEnabled = True
    frm_invoice.Caption = frm_invoice.Caption + " - " + frm_invoice.Tag
End Sub

'unload form free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    Set inv = Nothing
    Set InvcD = Nothing
    Set poInfo = Nothing
    deIms.rsGetPoitemFromPoForInvc.Close
    deIms.rsGet_Invoiece_Numbers_For_PO.Close
    
     If open_forms <= 5 Then ShowNavigator
    If Err Then Err.Clear
    
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
On Error Resume Next

    If sst_PO.Tab = 0 Then
    
        If Len(Trim$(ssdcboInvoiceNumb.text)) > 0 Then
            If remk Is Nothing Then: Beep: Exit Sub
            remk.MoveFirst
        End If
    ElseIf sst_PO.Tab = 1 Then
        deIms.rsGet_Invoice.MoveFirst
        Call FillInvoiceInfo(deIms.rsGet_Invoice)
    End If
    
    If ssdcboInvoiceNumb.Rows > 0 Then _
        ssdcboInvoiceNumb.MoveFirst
        
    If Err Then Err.Clear
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
On Error Resume Next

    If sst_PO.Tab = 0 Then
    
        If Len(Trim$(ssdcboInvoiceNumb.text)) > 0 Then
            If remk Is Nothing Then: Beep: Exit Sub
            remk.MoveLast
        End If
        
    ElseIf sst_PO.Tab = 1 Then
        deIms.rsGet_Invoice.MoveLast
        Call FillInvoiceInfo(deIms.rsGet_Invoice)
    End If
    
    If Err Then Err.Clear
    If ssdcboInvoiceNumb.Rows > 0 Then _
        ssdcboInvoiceNumb.MoveLast
End Sub

'move  recordset to next positon

Private Sub NavBar1_OnNextClick()
On Error Resume Next

    If sst_PO.Tab = 2 Then
    
        If Len(Trim$(ssdcboInvoiceNumb.text)) > 0 Then
            If remk Is Nothing Then: Beep: Exit Sub
            remk.MoveNext
            If remk.EOF Then remk.MoveLast
        End If
    ElseIf sst_PO.Tab = 1 Then
        deIms.rsGet_Invoice.MoveNext
        If deIms.rsGet_Invoice.EOF Then deIms.rsGet_Invoice.MoveLast
        
        Call FillInvoiceInfo(deIms.rsGet_Invoice)
    End If
    
    With ssdcboInvoiceNumb
        'If .Rows > 0 Then .Row
    End With
    
    If Err Then Err.Clear
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next

    If sst_PO.Tab = 0 Then
    
        If Len(Trim$(ssdcboInvoiceNumb.text)) > 0 Then
            If remk Is Nothing Then: Beep: Exit Sub
            remk.MovePrevious
            If remk.BOF Then remk.MoveFirst
        End If
    ElseIf sst_PO.Tab = 1 Then
        deIms.rsGet_Invoice.MovePrevious
        If deIms.rsGet_Invoice.EOF Then deIms.rsGet_Invoice.MoveFirst
        
        Call FillInvoiceInfo(deIms.rsGet_Invoice)
    End If
    
    If ssdcboInvoiceNumb.Rows > 0 Then _
        ssdcboInvoiceNumb.MovePrevious
        
    If Err Then Err.Clear
End Sub

'call before print function

Private Sub NavBar1_OnPrintClick()
        BeforePrint
    
        MDI_IMS.CrystalReport1.Action = 1
        MDI_IMS.CrystalReport1.Reset
End Sub

'call function saveall, after then free memory

Private Sub NavBar1_OnSaveClick()
'    deIms.rsGet_Invoice!inv_modiuser = CurrentUser
    SaveAll
    Set InvcD = Nothing
    Set inv = Nothing
End Sub

'set data to invoice form

Private Sub poInfo_FindComplete(ByVal Found As Boolean)

Dim f As imsGetSupplierInfo

    With poInfo
        Set f = .Supplier
        Label1(16) = f.City
        Label1(19) = f.State
        Label1(20) = f.ZipCode
        Label1(17) = f.Country
        Label1(15) = f.address1
        Label1(18) = f.address2
        Label1(21) = f.Telephone
        Label1(14) = f.SupplierName
        
        Label1(5) = .Buyer
        Label1(4) = .POCurrency
        Label1(7) = .BuyerTelephone
        
        On Error Resume Next
        Label1(6) = Format(.PO_Date, "MM/DD/YYYY")
        Label1(8) = Format(.DateRequested, "MM/DD/YYYY")
        
    End With
End Sub

'call store procedure to get invoice data for combo

Public Sub GetInvoices()
On Error Resume Next
Dim lng As Long
Dim Rs As ADODB.Recordset

    deIms.rsGet_Invoiece_Numbers_For_PO.Close
    
    ssdcboInvoiceNumb.RemoveAll
    Call deIms.Get_Invoiece_Numbers_For_PO(cbo_PO.text, deIms.NameSpace)

    If Err Then Err.Clear
    Set Rs = deIms.rsGet_Invoiece_Numbers_For_PO
    
    lng = Rs.RecordCount
    
    If lng = 0 Then Exit Sub
    Rs.MoveFirst
    
    ssdcboInvoiceNumb.RemoveAll
    If Err Then Err.Clear
    
    Do While Not Rs.EOF
        ssdcboInvoiceNumb.AddItem (CStr(Rs(0) & "" & ";" & Rs(1) & "" & ";" & CStr(Rs(2)) & ""))
        
        Rs.MoveNext
        
    Loop
    
If Err Then MsgBox Err.Description:  Err.Clear
End Sub

'call store procedure to get po line item information for combo

Public Sub FillLineItems()
On Error Resume Next
Dim Rs As ADODB.Recordset

    ssdcbolineitem.RemoveAll
    Set Rs = deIms.rsGetPoitemFromPoForInvc
    If ((Rs.State And adStateOpen) = adStateOpen) Then Rs.Close
    
    'deIms.rsGetPoitemFromPoForInvc.Close
    Call deIms.GetPoitemFromPoForInvc(cbo_PO, deIms.NameSpace)
    
    Call AddLineItems(deIms.rsGetPoitemFromPoForInvc)
    
End Sub
    
'add line item information to combo

Public Sub AddLineItems(Rs As ADODB.Recordset)
Dim STR As String

    If Rs Is Nothing Then Exit Sub
    If Rs.EOF And Rs.BOF Then Exit Sub
    If Rs.RecordCount = 0 Then Exit Sub
    
    STR = Chr$(1)
    ssdcbolineitem.FieldSeparator = STR
    Rs.MoveFirst
    If Err Then Err.Clear
    
    Do While Not Rs.EOF
        ssdcbolineitem.AddItem Rs!poi_liitnumb & "" & STR & Rs!poi_primreqdqty & "" & STR & Rs!poi_desc & ""
        Rs.MoveNext
    Loop
    
End Sub

'check move record status

Private Sub remk_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Select Case adReason
    
        Case adRsnMoveFirst, adRsnMoveLast, adRsnMoveNext, adRsnMovePrevious
            If remk.EOF Or remk.BOF Then rtbRemarks.text = "":  Exit Sub
            rtbRemarks.text = remk!invr_rem
    
    End Select
    
End Sub

'get invoice number list and set navbar button

Private Sub ssdcboInvoiceNumb_Click()
On Error Resume Next
   
    If Len(ssdcboInvoiceNumb.text) <> 0 Then
        
'        IMSMail1.Enabled = True
        NavBar1.PrintEnabled = True
        ssdbRecepientList.Enabled = True
        
'        invnumb = ssdcboInvoiceNumb.Text
        NavBar1.EMailEnabled = ssdbRecepientList.Rows
    Else
        NavBar1.PrintEnabled = False
        NavBar1.EMailEnabled = False
        
    End If
    
    rtbRemarks.text = ""
    fm = mdVisualization
    If Len(Trim$(ssdcboInvoiceNumb.text)) Then GetInvoiceInfo
    
    If Err Then Err.Clear
End Sub

'set value to invoice combo

Private Sub ssdcboInvoiceNumb_KeyPress(KeyAscii As Integer)
Dim l As Long

    fm = mdCreation
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii > 31 Then
        If (Len(Trim$(ssdcboInvoiceNumb.text)) > 15) Then
            Beep
            l = ssdcboInvoiceNumb.SelStart
            ssdcboInvoiceNumb.text = Trim$(Left$(ssdcboInvoiceNumb.text, 14))
            ssdcboInvoiceNumb.SelStart = l
        End If
    End If
End Sub

'load invoice line item numbers to combo

Private Sub ssdcboLineItem_Click()
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim STR As String
    
 

'    If Not inv Is Nothing Then
'        If txtInvcQnty_Validate(True) Then
'        End If
'        If Len(txtInvcPrice) Then
'
'            If Not IsNumeric(txtInvcPrice) Then
'                MsgBox "Invoice price is not correct)"
'                txtInvcPrice.SetFocus: Exit Sub
'            End If
'
'        Else
'
'            MsgBox "Invoice price cannot be left empty"
'                txtInvcPrice.SetFocus: Exit Sub
'        End If
'
'
'        If Len(txtInvcQnty) Then
'
'            If Not IsNumeric(txtInvcQnty) Then
'                MsgBox "Invoice Quantity is not correct"
'                txtInvcQnty.SetFocus: Exit Sub
'            End If
'
'        Else
'            MsgBox "Invoice Quantity cannot be left empty"
'                txtInvcQnty.SetFocus: Exit Sub
'        End If
'
'    End If
'
'    If Err Then Err.Clear
    If Len(Trim$(ssdcbolineitem)) <> 0 Then
         Call Getlineitem(cbo_PO, ssdcbolineitem)
'        Call GetPoLineItem(ssdcboPoNumb.Text, CoBLineitem)
    Else
    
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00263") 'J added
        MsgBox IIf(msg1 = "", "The line item cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcbolineitem.SetFocus: Exit Sub
    End If

'    Call Getlineitem(cbo_PO, ssdcbolineitem.Columns("liitnumb").Value)
'    Set rs = deIms.rsGetPoitemFromPoForInvc
'
'    If ((rs.State And adStateOpen) <> adStateOpen) Then _
'        Call deIms.GetPoitemFromPoForInvc(cbo_PO, deIms.NameSpace)
'    rs.Filter = 0
'    rs.Filter = "poi_liitnumb = '" & ssdcbolineitem.Columns("liitnumb").Value & "'"
'
'    On Error Resume Next
'    With rs
'        lblPrice(3) = ""
'        lblComm = !poi_comm
'        lblUnit(0) = !poi_primuom & ""
'        lblUnit(1) = !poi_secouom & ""
'        lblAmount(0) = FormatNumber$(CDbl(!poi_primreqdqty), 4)
'        lblAmount(1) = FormatNumber$(CDbl(IIf(IsNull(!poi_secoreqdqty), 0, !poi_secoreqdqty)), 4)
'        lblPrice(0) = FormatCurrency$(CDbl(!poi_unitprice), 4)
'        lblPrice(2) = FormatCurrency$(CDbl(!poi_totaprice), 2)
'    End With
'
'    txtInvcQnty.Enabled = True
'    txtInvcPrice.Enabled = True
'
'    txtInvcQnty.Text = ""
'    txtInvcPrice.Text = ""
'
'    If Err Then Err.Clear
'
''    If InvcD Is Nothing Then Set InvcD = New imsInvoiceDetls
'
'    Set inv = InvcD.item(ssdcbolineitem.Text)
'    ssdcboInvoiceNumb.Text = IIf(fm = mdCreation, ssdcboInvoiceNumb.Text, "")
'
'    If Not inv Is Nothing Then AssignValues: Exit Sub
'
'
'
'    Set inv = InvcD.Add(cbo_PO, deIms.NameSpace, ssdcboInvoiceNumb.Text, _
'                        ssdcbolineitem.Text, 0, 0, 0, ssdcbolineitem.Text)
'
'
''     Set inv.Connection = deIms.cnIms
'
'    Set rs = Nothing
'
            
End Sub

'call store procedure to save invoice data

Private Sub SaveAll()
Dim cmd As ADODB.Command
    
    If Not ValidateData Then Exit Sub
    Set cmd = New ADODB.Command
    
    cmd.CommandText = "INVOICEINSERT"
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = deIms.cnIms

    cmd.Parameters.Append _
        cmd.CreateParameter("Return_Value", adInteger, adParamReturnValue)
        
    cmd.Parameters.Append _
        cmd.CreateParameter("@inv_ponumb", adVarChar, adParamInput, 15, cbo_PO)

    cmd.Parameters.Append _
        cmd.CreateParameter("@inv_npecode", adVarChar, adParamInput, 5, deIms.NameSpace)


    cmd.Parameters.Append _
        cmd.CreateParameter("@inv_invcnumb", adVarChar, adParamInput, 15, ssdcboInvoiceNumb.text)


    cmd.Parameters.Append _
        cmd.CreateParameter("@inv_invcdate", adDBTimeStamp, adParamInput, , txt_dateinvoice)
        
    cmd.Parameters.Append _
        cmd.CreateParameter("@user", adVarChar, adParamInput, 20, CurrentUser)


    On Error Resume Next
    
    cmd.Execute
    
    If cmd.Parameters("Return_Value") <> 0 Then MsgBox Err.Description
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "if @@trancount > 0 commit"
    cmd.Execute
    
    Dim ctl As Object

    For Each ctl In InvcD
        If Len(Trim$(ctl.InvoiceNumber)) = 0 Then _
            ctl.InvoiceNumber = ssdcboInvoiceNumb.text
            
        Call SaveDetl(ctl)
    Next ctl
            
    SaveRemarks
    cmd.Execute
    
    'Modified by Juan (9/11/2000) for Multilingual
    msg1 = translator.Trans("M00264") 'J added
    MsgBox IIf(msg1 = "", "Insert into Invoice was completed successfully", msg1) 'J modified
    '---------------------------------------------
    
    Set inv = Nothing
'    Set InvcD = Nothing
    InvcD.RemoveAll
    
CleanUp:
    
    Set cmd = Nothing
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear: GoTo CleanUp
    End If
    
End Sub

' validate invoice data

Public Function ValidateData() As Boolean
Dim STR As String

    ValidateData = False
    STR = txt_dateinvoice
    
    If Len(Trim$(STR)) = 0 Then
    
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00265") 'J added
        MsgBox IIf(msg1 = "", "Invoice date cannot be left empty", msg1) 'J modified
        '----------------------------------------------
        
        Exit Function
        
    ElseIf Not IsDate(STR) Then
    
            'Modified by Juan (9/12/2000) for Multilingual
            msg1 = translator.Trans("M00266") 'J added
            MsgBox IIf(msg1 = "", "Invoice date is not a correct date", msg1)
            '---------------------------------------------
            
            Exit Function
    End If
        
    STR = ssdcboInvoiceNumb.text
    
    If Len(Trim$(STR)) = 0 Then
    
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00267")
        MsgBox IIf(msg1 = "", "Invoice Number cannot be left empty", msg1)
        '---------------------------------------------
        
        Exit Function
    End If

    ValidateData = True
End Function

'on tab click set navbar button

Private Sub sst_PO_Click(PreviousTab As Integer)
Dim l As Long
Dim iEditMode As String, blFlag As Boolean


    blFlag = sst_PO.Tab = 1
    
    With NavBar1
        .NextEnabled = blFlag
        .LastEnabled = blFlag
        .FirstEnabled = blFlag
        .CancelEnabled = blFlag
        .PreviousEnabled = blFlag
        .SaveEnabled = sst_PO.Tab = 0
        .CloseEnabled = sst_PO.Tab = 0
        .NewEnabled = sst_PO.Tab = 0
        .PrintEnabled = .SaveEnabled And ssdcboInvoiceNumb.Rows <> 0
        .EMailEnabled = ((ssdbRecepientList.Rows) And (.PrintEnabled))
    End With
    
    
    
    l = sst_PO.Tab
    
    If PreviousTab = 0 Then
        
        If Len(Trim$(cbo_PO)) Then
            sst_PO.Tab = IIf(ValidateData, l, PreviousTab)
        End If
        
    End If
    
End Sub

'check invoice price format, if wrong type, show message

Private Sub txtInvcPrice_Change()

    If inv Is Nothing Then Exit Sub
    If Len(txtInvcPrice) = 0 Then Exit Sub

    If IsNumeric(txtInvcPrice) Then
        inv.UnitPrice = txtInvcPrice

    Else
    
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00268") 'J added
        MsgBox IIf(msg1 = "", "Cannot add non numerical characters", msg1) 'J modified
        '----------------------------------------------
        
        Exit Sub
    End If

  If Len(Trim$(txtInvcPrice)) Then

        Dim db As Double
        If IsNumeric(txtInvcPrice) Then
            inv.TotalPrice = txtInvcQnty * txtInvcPrice
            lblPrice(1) = Format(inv.TotalPrice, "currency")

            db = CDbl(CDbl((lblPrice(0)) * txtInvcQnty))
            db = CDbl(lblPrice(1) - db)
            lblPrice(3) = Format$(db, "currency")
        End If
    End If


End Sub

'check over set invoice price, show message

Private Sub txtInvcPrice_Validate(Cancel As Boolean)
Dim msg, Style, Title
Dim Num1 As Double
Dim Num2 As Double

'Modified by Juan (9/12/2000) for Multilingual
msg1 = translator.Trans("L00186") 'J added
msg2 = translator.Trans("M00269") 'J added
msg = IIf(msg1 = "", " Lineitem# ", msg1 + " ") & ssdcbolineitem & IIf(msg2 = "", " is being over priced. Do you want to continue ?", " " + msg2) 'J modified
'---------------------------------------------

Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Imswin"

'    Cancel = False

    Dim db As Double


    If Len(Trim$(txtInvcPrice)) Then
        If IsNumeric(txtInvcPrice) Then

            inv.TotalPrice = txtInvcQnty * txtInvcPrice
            lblPrice(1) = Format(inv.TotalPrice, "currency")

            db = CDbl(CDbl((lblPrice(0)) * txtInvcQnty))
            db = CDbl(lblPrice(1)) - db
            lblPrice(3) = Format$(db, "currency")
        End If
    End If




    If Not Len(Trim(txtInvcPrice)) = 0 Then

    Num1 = lblPrice(0).Caption
    Num2 = txtInvcPrice

         If Num2 > Num1 Then
            If MsgBox(msg, Style, Title) = vbNo Then
                txtInvcPrice = ""
                txtInvcPrice.SetFocus: Exit Sub
            End If
        Else
        End If
    End If

    If Len(Trim$(txtInvcPrice)) <> 0 Then
       txtInvcPrice = FormatNumber((txtInvcPrice), 4)
        
    End If
'    Cancel = True

End Sub

'Private Sub txtInvcQnty_Change()

'    If inv Is Nothing Then Exit Sub
'    If Len(txtInvcQnty) = 0 Then Exit Sub
'
'    If IsNumeric(txtInvcQnty) Then
'        inv.Quantity = txtInvcQnty
'        lblPrice(2) = Format$(CDbl(txtInvcQnty) * CDbl(lblPrice(0)), "Currency")
'
'    Else
'        MsgBox "cannot add non numerical characters"
'        Exit Sub
'
'    End If
    
'    If Len(Trim$(txtInvcPrice)) Then
'
'        Dim db As Double
'        If IsNumeric(txtInvcPrice) Then
'            inv.TotalPrice = txtInvcQnty * txtInvcPrice
'            lblPrice(1) = Format(inv.TotalPrice, "currency")
'
'            db = CDbl(CDbl((lblPrice(0)) * txtInvcQnty))
'            db = CDbl(lblPrice(1)) - db
'            lblPrice(3) = Format$(db, "currency")
'        End If
'    End If
'End Sub

'calculate invoice total price and quantity

Public Sub AssignValues()
    With inv
        txtInvcQnty = .Quantity
        txtInvcPrice = .TotalPrice \ .Quantity
    End With
End Sub

'set data for invoice table insert

Private Sub SaveDetl(dt As imsInvoiceDetl)
'    Dim cmd As ADODB.Command
'
'    Set cmd = New ADODB.Command
'
'    cmd.CommandType = adCmdStoredProc
'    Set cmd.ActiveConnection = deIms.cnIms
'    cmd.CommandText = "invoicedetl_insert_sp"
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("Retval", adInteger, adParamReturnValue)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("(@invd_ponumb", adVarChar, adParamInput, 15, cbo_PO)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_invcnumb", adVarChar, adParamInput, 15, ssdcboInvoiceNumb.Text)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_liitnumb", adChar, adParamInput, 6, dt.LineItem)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_qty", adInteger, adParamInput, , dt.Quantity)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_totapric", adNumeric, adParamInput, , dt.TotalPrice)
'
'    cmd.Parameters.Append _
'        cmd.CreateParameter("@invd_unitpric", adCurrency, adParamInput, , dt.UnitPrice)
'
'    cmd.Execute
'    Set cmd = Nothing
        
    Dim l As Long
    l = deIms.invoicedetl_insert_sp(cbo_PO.Tag, deIms.NameSpace, ssdcboInvoiceNumb.text _
                                , dt.LineItem, dt.Quantity, dt.TotalPrice, dt.UnitPrice, CurrentUser)
End Sub

'set data for invoice remark table insert

Public Sub SaveRemarks()
Dim l As Long


'Modified by Muzammil 08/11/00
       'Reason - VBCRLFs before the text would block Email Generation.
          
          Do While InStr(1, rtbRemarks, vbCrLf) = 1                   'M
             rtbRemarks = Mid(rtbRemarks, 3, Len(rtbRemarks))         'M
          Loop                                                        'M
             rtbRemarks = LTrim$(rtbRemarks)                          'M
        

    If Len(Trim$(rtbRemarks.text)) = 0 Then Exit Sub
    l = deIms.Invoice_Remark_insert(deIms.NameSpace, cbo_PO.Tag, ssdcboInvoiceNumb.text, rtbRemarks.text, CurrentUser)
End Sub

'call function get invoice information

Private Sub GetInvoiceInfo()
On Error Resume Next
Dim Rs As ADODB.Recordset

    Set Rs = deIms.rsGet_Invoice
    
    If Rs.State And adStateOpen = adStateOpen Then Rs.Close
    
    Call deIms.Get_Invoice(ssdcboInvoiceNumb.text, deIms.NameSpace, cbo_PO)
    
    Call FillInvoiceInfo(deIms.rsGet_Invoice)
    
    Set remk = Nothing
    GetInvoiceRemarks
    If Err Then Err.Clear
    Call poInfo.GetValues(cbo_PO)
    Call DisableAllControls(True)

End Sub

'fill invoice information to form

Private Sub FillInvoiceInfo(Rs As ADODB.Recordset)
Dim db As Double
On Error Resume Next

    If Rs Is Nothing Then Exit Sub

    ssdcbolineitem.text = Rs!invd_liitnumb
    
    With Rs
    
        lblPrice(3) = ""
        lblComm = !poi_comm & ""
        txtInvcQnty = !invd_qty & ""
        Label1(1) = !invd_creadate
        lblUnit(0) = !poi_primuom & ""
        lblUnit(1) = !poi_secouom & ""
        
        Label1(0) = !invd_creauser & ""
        Label1(2) = !invd_modiuser & ""
        lblAmount(0) = CDbl(!poi_primreqdqty)
        
        Label1(1) = Format$(!invd_creadate, "mm/dd/yyyy")
        Label1(3) = Format$(!invd_modidate, "mm/dd/yyyy")
        
        
        txt_dateinvoice = Label1(1)
        txt_dateinvoice.Enabled = False
        lblPrice(0) = Format$(CDbl(!poi_unitprice), "currency")
        txtInvcPrice = Format$(CDbl(!invd_unitpric), "currency")
        lblAmount(1) = CDbl(IIf(IsNull(!poi_secoreqdqty), 0, !poi_secoreqdqty))
        lblPrice(2) = Format$(CDbl(!poi_unitprice * !invd_qty), "currency")
  End With
  
    If Len(Trim$(txtInvcQnty)) Then
        If IsNumeric(txtInvcQnty) Then
            
            
            lblPrice(1) = Format(txtInvcQnty * txtInvcPrice, "currency")
            
            db = CDbl(CDbl((lblPrice(0)) * txtInvcQnty))
            db = CDbl(lblPrice(1)) - db
            lblPrice(3) = Format$(db, "currency")
        End If
    End If
    
    If Err Then Err.Clear
End Sub

'call stor procedure to get invoice remarks information

Private Sub GetInvoiceRemarks()
Dim cmd As ADODB.Command
Dim l As Long

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        .CommandText = "Get_Invoice_Remarks"
    
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .Parameters.Append .CreateParameter("@INVPOICENUMBER", adVarChar, adParamInput, 15, ssdcboInvoiceNumb.text)
        
        Set remk = .Execute
        
        l = .Parameters("RetVal").Value
        If l = 0 Then Set remk = Nothing
        
        rtbRemarks.Enabled = False
        If l Then rtbRemarks.text = remk!invr_rem
    End With
End Sub

'get crystal report parameters and application path

Public Sub BeforePrint()
On Error GoTo ErrHandler

      With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\invoice.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "invnumb;" + ssdcboInvoiceNumb.text + ";TRUE"
        .ParameterFields(2) = "ponumb;" + cbo_PO + ";TRUE"
        
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00176") 'J added
        .WindowTitle = IIf(msg1 = "", "Invoice", msg1) 'J modified
        Call translator.Translate_Reports("invoice.rpt") 'J added
        '---------------------------------------------
        
    End With
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'disable controls, without lable, sstab

Private Sub DisableAllControls(Disable As Boolean)
Dim ctl As Control
On Error Resume Next

    For Each ctl In Me.Controls
        If Not TypeOf ctl Is Label Or TypeOf ctl Is SSTab Then
            ctl.Enabled = Not Disable
            If Err Then Err.Clear
        End If
    Next ctl
        
    sst_PO.Enabled = True
    cbo_PO.Enabled = True
    NavBar1.Enabled = True
    
    cmd_Add.Enabled = True
    cmd_Remove.Enabled = True
    opt_Email.Enabled = True
    opt_FaxNum.Enabled = True
    dgRecepients.Enabled = True
    fra_FaxSelect.Enabled = True
    txt_Recipient.Enabled = True
    ssdbRecepientList.Enabled = True
    ssdcboInvoiceNumb.Enabled = True
End Sub

'check quantity over set, show message

Private Sub txtInvcQnty_Validate(Cancel As Boolean)
Dim msg, Style, Title
Dim Num1 As Double
Dim Num2 As Double

'Modified by Juan (9/12/2000) for Multilingual
msg1 = translator.Trans("L00186") 'J added
msg2 = translator.Trans("M00270") 'J added
msg = IIf(msg1 = "", " Lineitem# ", msg1 + " ") & ssdcbolineitem & IIf(msg2 = "", " is being over invoiced. Do you want to continue ?", " " + msg2)
'---------------------------------------------

Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Imswin"



    If inv Is Nothing Then Exit Sub
    If Len(txtInvcQnty) = 0 Then Exit Sub

    If IsNumeric(txtInvcQnty) Then
        inv.Quantity = txtInvcQnty
        lblPrice(2) = Format$(CDbl(txtInvcQnty) * CDbl(lblPrice(0)), "Currency")
        txtInvcQnty = CDbl(txtInvcQnty)
    Else
    
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00268") 'J added
        MsgBox IIf(msg1 = "", "cannot add non numerical characters", msg1) 'J modified
        '---------------------------------------------
        
        Exit Sub

    End If
    
    Num1 = lblAmount(0).Caption
    Num2 = txtInvcQnty

    If Not Len(Trim(txtInvcQnty)) = 0 Then
         If txtInvcQnty > Num1 Then
            If MsgBox(msg, Style, Title) = vbNo Then
'                txtInvcQnty = ""
                txtInvcQnty.SetFocus: Exit Sub
            End If
        Else
        End If
    End If
    
    If Len(Trim$(txtInvcQnty)) <> 0 Then
        txtInvcQnty = FormatNumber((txtInvcQnty), 4)
        
    End If
    
    txtInvcPrice.SetFocus
End Sub

'Public Function checkquantity() As Boolean
'Dim msg, Style, Title
'Dim Num1 As Double
'Dim Num2 As Double
'
'msg = " Lineitem# " & ssdcbolineitem & " is being over received, Do you want to continue ?"
'Style = vbYesNo + vbCritical + vbDefaultButton2
'Title = "Imswin"
'
'    checkquantity = False
'    Num1 = lblAmount(0).Caption
'    Num2 = txtInvcQnty
'
'    If Not Len(Trim(txtInvcQnty)) = 0 Then
'         If txtInvcQnty > Num1 Then
'            If MsgBox(msg, Style, Title) = vbNo Then
'                txtInvcQnty = ""
'                txtInvcQnty.SetFocus: Exit Function
'            End If
'        Else
'        End If
'    End If
'
'    checkquantity = True
'
'
'End Function

'Public Function checkprice() As Boolean
'Dim msg, Style, Title
'Dim Num1 As Double
'Dim Num2 As Double
'
'msg = " Lineitem# " & ssdcbolineitem & " price is being over received, Do you want to continue ?"
'Style = vbYesNo + vbCritical + vbDefaultButton2
'Title = "Imswin"
'
'    checkprice = False
'
'   If Len(Trim$(txtInvcPrice)) Then
'
'        Dim db As Double
'        If IsNumeric(txtInvcPrice) Then
'            inv.TotalPrice = txtInvcQnty * txtInvcPrice
'            lblPrice(1) = Format(inv.TotalPrice, "currency")
'
'            db = CDbl(CDbl((lblPrice(0)) * txtInvcQnty))
'            db = CDbl(lblPrice(1)) - db
'            lblPrice(3) = Format$(db, "currency")
'        End If
'    End If
'
'    Num1 = lblPrice(0).Caption
'    Num2 = txtInvcPrice
'
'    If Not Len(Trim(txtInvcPrice)) = 0 Then
'         If txtInvcPrice > Num1 Then
'            If MsgBox(msg, Style, Title) = vbNo Then
'                txtInvcPrice = ""
'                txtInvcPrice.SetFocus: Exit Function
'            End If
'        Else
'        End If
'    End If
'
'    checkprice = True
'End Function

'SQL statement get po line item information

Public Sub Getlineitem(Ponumb As String, Linenumb As String)
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

'    Call ClearLineitem

     Set cmd = MakeCommand(deIms.cnIms, adCmdText)
     
     With cmd
        .CommandText = " SELECT poi_primreqdqty,poi_primuom , poi_unitprice, poi_comm,"
        .CommandText = .CommandText & "poi_totaprice, poi_secoreqdqty, poi_secouom "
        .CommandText = .CommandText & " FROM POITEM WHERE poi_ponumb = '" & Ponumb & "'"
        .CommandText = .CommandText & "and poi_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and poi_liitnumb = '" & Linenumb & " '"
     
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then GoTo clearup
    
    If rst.EOF And rst.BOF Then
    
        lblPrice(3) = ""
        lblComm = ""
        lblUnit(0) = ""
        lblUnit(1) = ""
        lblAmount(0) = ""
        lblAmount(1) = ""
        lblPrice(0) = ""
        lblPrice(2) = ""
    Else
    
        lblPrice(3) = ""
        lblComm = rst!poi_comm & ""
        lblUnit(0) = rst!poi_primuom & ""
        lblUnit(1) = rst!poi_secouom & ""
        lblAmount(0) = FormatNumber$(CDbl(rst!poi_primreqdqty), 4)
        lblAmount(1) = FormatNumber$(CDbl(IIf(IsNull(rst!poi_secoreqdqty), 0, rst!poi_secoreqdqty)), 4)
        lblPrice(0) = FormatCurrency$(CDbl(rst!poi_unitprice), 4)
        lblPrice(2) = FormatCurrency$(CDbl(rst!poi_totaprice), 2)
    End If
    
        txtInvcQnty.Enabled = True
        txtInvcPrice.Enabled = True

        txtInvcQnty.text = ""
        txtInvcPrice.text = ""
    
        If InvcD Is Nothing Then Set InvcD = New imsInvoiceDetls
        

       Set inv = InvcD.Add(cbo_PO, deIms.NameSpace, ssdcboInvoiceNumb.text, _
                        ssdcbolineitem.text, 0, 0, 0, ssdcbolineitem.text)
    
    
clearup:
    Set rst = Nothing
    Set cmd = Nothing
    
End Sub

'clear invoice line items form

Public Sub ClearLineitem()
     
        lblPrice(3) = ""
        lblComm = ""
        lblUnit(0) = ""
        lblUnit(1) = ""
        lblAmount(0) = ""
        lblAmount(1) = ""
        lblPrice(0) = ""
        lblPrice(2) = ""
End Sub
