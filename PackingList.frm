VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frm_PackingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List / Manifest Management"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10200
   Tag             =   "02030200"
   Begin LRNavigators.NavBar NavBar1 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   3360
      TabIndex        =   11
      Top             =   6960
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "PackingList.frx":0000
      PreviousVisible =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      EMailVisible    =   -1  'True
      CloseToolTipText=   ""
      PrintToolTipText=   ""
      EmailToolTipText=   ""
      NewToolTipText  =   ""
      SaveToolTipText =   ""
      CancelToolTipText=   ""
      NextToolTipText =   ""
      LastToolTipText =   ""
      FirstToolTipText=   ""
      PreviousToolTipText=   ""
      DeleteToolTipText=   ""
      EditToolTipText =   ""
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   33
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Packing List"
      TabPicture(0)   =   "PackingList.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "PackingList.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "lbl_New"
      Tab(1).Control(2)=   "lbl_Recipients"
      Tab(1).Control(3)=   "dgRecepients"
      Tab(1).Control(4)=   "dgRecepientList"
      Tab(1).Control(5)=   "fra_FaxSelect"
      Tab(1).Control(6)=   "cmd_Add"
      Tab(1).Control(7)=   "cmd_Remove"
      Tab(1).Control(8)=   "txt_Recipient"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Line Item"
      TabPicture(2)   =   "PackingList.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LlbShipTo"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LlbManifest"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label9(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label9(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label9(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "ssdcboPoNumb"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "CoBLineitem"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame4"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Remark"
      TabPicture(3)   =   "PackingList.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtRemarks"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Line Item List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -73680
         TabIndex        =   89
         Top             =   1920
         Width           =   7095
         Begin VB.TextBox TxtBoxNumber 
            Height          =   315
            Left            =   5880
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtDescription 
            Height          =   2115
            Left            =   2040
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox TxtBeShipped 
            Height          =   315
            Left            =   2040
            TabIndex        =   2
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label LblTobeInven 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5880
            TabIndex        =   103
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblQtyInv 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5880
            TabIndex        =   102
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblQtyDelv 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5880
            TabIndex        =   101
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblAmount 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   100
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblUnitPrice 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   99
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblReqQty 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   98
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Qty. To Ship"
            Height          =   255
            Left            =   3960
            TabIndex        =   97
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Qty Already Shipped"
            Height          =   255
            Left            =   3960
            TabIndex        =   96
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Qty. Already Delivered"
            Height          =   255
            Left            =   3960
            TabIndex        =   95
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Unit Price"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Requested Qty"
            Height          =   315
            Left            =   240
            TabIndex        =   92
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Box Number"
            Height          =   255
            Left            =   3960
            TabIndex        =   91
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Quantity Being Shipped"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Marks"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   480
         TabIndex        =   67
         Top             =   4680
         Width           =   8775
         Begin VB.TextBox TxtMark4 
            Height          =   315
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   21
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Txtviacarr 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   14
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Txtnumbpiec 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   15
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Txtgrosweig 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   16
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox Txttotavolu 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   17
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox TxtMark1 
            Height          =   315
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   18
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtMark2 
            Height          =   315
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   19
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtMark3 
            Height          =   315
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   20
            Top             =   960
            Width           =   2175
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboTermDelivery 
            Height          =   315
            Left            =   1560
            TabIndex        =   32
            Top             =   240
            Width           =   2175
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":008C
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
            stylesets(1).Picture=   "PackingList.frx":00A8
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Label Label4 
            Caption         =   "Shipping Terms"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   78
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Via Carrier"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   77
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Destination"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   76
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Number Pieces"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   75
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Gross Weight Kg"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   74
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Total Volume"
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   73
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 1"
            Height          =   255
            Index           =   6
            Left            =   4440
            TabIndex        =   72
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 2"
            Height          =   255
            Index           =   7
            Left            =   4440
            TabIndex        =   71
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 3"
            Height          =   255
            Index           =   8
            Left            =   4440
            TabIndex        =   70
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 4"
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   69
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label LblDestination 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   68
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.TextBox TxtRemarks 
         Enabled         =   0   'False
         Height          =   4575
         Left            =   -74520
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   600
         Width           =   8895
      End
      Begin VB.TextBox txt_Recipient 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72990
         TabIndex        =   64
         Top             =   3510
         Width           =   7230
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74520
         TabIndex        =   63
         Top             =   2745
         Width           =   972
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74505
         TabIndex        =   62
         Top             =   2430
         Width           =   972
      End
      Begin VB.Frame fra_FaxSelect 
         Enabled         =   0   'False
         Height          =   1290
         Left            =   -74505
         TabIndex        =   59
         Top             =   3885
         Width           =   1410
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   60
            TabIndex        =   61
            Top             =   285
            Width           =   1275
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   60
            TabIndex        =   60
            Top             =   780
            Width           =   684
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Shipping Information"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   480
         TabIndex        =   46
         Top             =   2520
         Width           =   8775
         Begin MSComCtl2.DTPicker DTPicker1etd 
            Height          =   315
            Left            =   4560
            TabIndex        =   26
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24707073
            CurrentDate     =   36524
         End
         Begin VB.TextBox Txtawbnumb 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Txtflig1 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   9
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Txtflig2 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   12
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox TxtRemark 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   1800
            Width           =   7095
         End
         Begin VB.TextBox Txtfrom1 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   10
            Top             =   1080
            Width           =   2175
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDestination 
            Height          =   315
            Left            =   4560
            TabIndex        =   28
            Top             =   1080
            Width           =   1695
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":00C4
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
            stylesets(1).Picture=   "PackingList.frx":00E0
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Destination"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDestinationTo 
            Height          =   315
            Left            =   6960
            TabIndex        =   29
            Top             =   1080
            Width           =   1695
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":00FC
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
            stylesets(1).Picture=   "PackingList.frx":0118
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Destination"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDesnationFrom 
            Height          =   315
            Left            =   4560
            TabIndex        =   30
            Top             =   1440
            Width           =   1695
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":0134
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
            stylesets(1).Picture=   "PackingList.frx":0150
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Destination"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDestinationTo1 
            Height          =   315
            Left            =   6960
            TabIndex        =   31
            Top             =   1440
            Width           =   1695
            DataFieldList   =   "Column 0"
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
            stylesets(0).Picture=   "PackingList.frx":016C
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
            stylesets(1).Picture=   "PackingList.frx":0188
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Destination"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin MSComCtl2.DTPicker DTPicker2eta 
            Height          =   315
            Left            =   6960
            TabIndex        =   27
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24707073
            CurrentDate     =   36524
         End
         Begin VB.Label Label2 
            Caption         =   "AWB / BL"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "HAWB / TBL"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   56
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Flight / Voyage"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   55
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Flight / Voyage"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   54
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Remark"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   53
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "ETD"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   52
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "ETA"
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   51
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "From"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   50
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   255
            Index           =   3
            Left            =   6360
            TabIndex        =   49
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "From"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   48
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   255
            Index           =   5
            Left            =   6360
            TabIndex        =   47
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Packing List / Manifest Management"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   480
         TabIndex        =   34
         Top             =   360
         Width           =   8775
         Begin VB.ComboBox cboPackingNumber 
            Height          =   315
            Left            =   1800
            TabIndex        =   107
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Txtcustrefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "4"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox Txtforwrefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "5"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox Txtshprefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            MaxLength       =   20
            TabIndex        =   5
            Tag             =   "3"
            Top             =   1080
            Width           =   2175
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboShipper 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Tag             =   "1"
            Top             =   720
            Width           =   2175
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":01A4
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
            stylesets(1).Picture=   "PackingList.frx":01C0
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3519
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2646
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboPriority 
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   1800
            Width           =   1335
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":01DC
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
            stylesets(1).Picture=   "PackingList.frx":01F8
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo5 
            Height          =   315
            Left            =   4095
            TabIndex        =   35
            Top             =   4140
            Width           =   2175
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":0214
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
            stylesets(1).Picture=   "PackingList.frx":0230
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 1"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 0"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboShipto 
            Height          =   315
            Left            =   5880
            TabIndex        =   24
            Top             =   360
            Width           =   2175
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":024C
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
            stylesets(1).Picture=   "PackingList.frx":0268
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboSoldTo 
            Height          =   315
            Left            =   5880
            TabIndex        =   25
            Top             =   720
            Width           =   2175
            DataFieldList   =   "Column 0"
            AllowInput      =   0   'False
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
            stylesets(0).Picture=   "PackingList.frx":0284
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
            stylesets(1).Picture=   "PackingList.frx":02A0
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3096
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1958
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).HeadStyleSet=   "ColHeader"
            Columns(1).StyleSet=   "RowFont"
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPshidate 
            Height          =   315
            Left            =   1800
            TabIndex        =   105
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24707073
            CurrentDate     =   36524
         End
         Begin MSComCtl2.DTPicker DTPDocudate 
            Height          =   315
            Left            =   1800
            TabIndex        =   106
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24707073
            CurrentDate     =   36524
         End
         Begin VB.Label Label1 
            Caption         =   "Packing / Manifest"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Shipper"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Document Date"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Shipping Date"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Air / Sea / Other"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   41
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Ship To Code"
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   40
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Sold To Code"
            Height          =   255
            Index           =   6
            Left            =   4200
            TabIndex        =   39
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Shipper's Ref"
            Height          =   255
            Index           =   7
            Left            =   4200
            TabIndex        =   38
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Customers Ref"
            Height          =   255
            Index           =   8
            Left            =   4200
            TabIndex        =   37
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Forwarder's Ref"
            Height          =   255
            Index           =   9
            Left            =   4200
            TabIndex        =   36
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.ComboBox CoBLineitem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68640
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid dgRecepientList 
         Height          =   2535
         Left            =   -72960
         TabIndex        =   58
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Recipient"
            Caption         =   "Recipient List"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   7004.977
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboPoNumb 
         Height          =   315
         Left            =   -71880
         TabIndex        =   0
         Top             =   1440
         Width           =   1920
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
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
         stylesets(0).Picture=   "PackingList.frx":02BC
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
         stylesets(1).Picture=   "PackingList.frx":02D8
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   2302
         Columns(0).Caption=   "PO-Number"
         Columns(0).Name =   "PO-Number"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).HeadStyleSet=   "ColHeader"
         Columns(0).StyleSet=   "RowFont"
         Columns(1).Width=   2037
         Columns(1).Caption=   "Date"
         Columns(1).Name =   "Date"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   7
         Columns(1).FieldLen=   256
         Columns(1).HeadStyleSet=   "ColHeader"
         Columns(1).StyleSet=   "RowFont"
         Columns(2).Width=   1826
         Columns(2).Caption=   "Status"
         Columns(2).Name =   "Status"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).HeadStyleSet=   "ColHeader"
         Columns(2).StyleSet=   "RowFont"
         Columns(3).Width=   3889
         Columns(3).Caption=   "Supplier"
         Columns(3).Name =   "Supplier"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).HeadStyleSet=   "ColHeader"
         Columns(3).StyleSet=   "RowFont"
         Columns(4).Width=   1958
         Columns(4).Caption=   "Code"
         Columns(4).Name =   "Code"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).HeadStyleSet=   "ColHeader"
         Columns(4).StyleSet=   "RowFont"
         _ExtentX        =   3387
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   2775
         Left            =   -72960
         TabIndex        =   66
         Top             =   3960
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4895
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   3225.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3690.142
            EndProperty
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Line item"
         Height          =   255
         Index           =   2
         Left            =   -69720
         TabIndex        =   88
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Packing List / Manifest"
         Height          =   315
         Left            =   -73680
         TabIndex        =   87
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Ship To Code"
         Height          =   255
         Index           =   0
         Left            =   -69720
         TabIndex        =   86
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "PO Number"
         Height          =   255
         Index           =   1
         Left            =   -73680
         TabIndex        =   85
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label LlbManifest 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -71880
         TabIndex        =   84
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label LlbShipTo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68640
         TabIndex        =   83
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Line Items List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72240
         TabIndex        =   82
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74520
         TabIndex        =   81
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lbl_New 
         Caption         =   "New"
         Height          =   300
         Left            =   -74520
         TabIndex        =   80
         Top             =   3570
         Width           =   660
      End
      Begin VB.Label Label5 
         Caption         =   "Visualization"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68280
         TabIndex        =   79
         Top             =   7440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_PackingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReceptList As ADODB.Recordset

Dim AddingRecord As Boolean
Dim pl As imsPackinListDetl
Dim WithEvents rec As imsPackingListRecp
Attribute rec.VB_VarHelpID = -1
Dim WithEvents pld As PackingListDetls
Attribute pld.VB_VarHelpID = -1

Private Sub cboPackingNumber_Click()

    dgRecepientList.Tag = ""
    Set rsReceptList = Nothing
    'Call EnableControls(True)

    If Len(Trim$(cboPackingNumber)) <> 0 Then
'        Call cboPackingNumber_Validate(True)
        Call GetPackingAlloflist(cboPackingNumber)
        If GetPackingNumber(cboPackingNumber) Then
            
            cboPackingNumber.Enabled = False
            
            NavBar1.NewEnabled = True
            NavBar1.EMailEnabled = True
            NavBar1.CloseEnabled = True
            NavBar1.CancelEnabled = True
            NavBar1.PrintEnabled = True
            NavBar1.SaveEnabled = False
            Call EnableControls(False)
            'MsgBox "Packing List Entered Number is already exist"
            
            
            
       '     cboPackingNumber.SetFocus: Exit Sub
        Else
           
            AssignDefault
            NavBar1.NewEnabled = True
            NavBar1.EMailEnabled = True
            NavBar1.CloseEnabled = True
            NavBar1.CancelEnabled = True
            NavBar1.PrintEnabled = True
            NavBar1.SaveEnabled = False
            Call EnableControls(True)
        End If
            
    End If
'    NavBar1.EMailEnabled = True
'    NavBar1.CloseEnabled = True
'    NavBar1.CancelEnabled = False
'    NavBar1.PrintEnabled = False
'    NavBar1.NewEnabled = False
'    NavBar1.SaveEnabled = False
End Sub



Private Sub CoBLineitem_Click()
    
    If Len(Trim$(CoBLineitem)) <> 0 Then
        'if pld.Count then
        Call GetPoLineItem(ssdcboPoNumb.Text, CoBLineitem)
    Else
        MsgBox "The line item cannot be left empty"
        CoBLineitem.SetFocus: Exit Sub
    End If
        
        
        'Call EnableControlsLine(True)
'    Else
'        Call CoBLineitem_DblClick
'    End If
End Sub



'Private Sub CoBLineitem_DblClick()
'     Call EnableControlsLine(True)
'
'    If Len(ssdcboPoNumb) <> 0 Then
'        Call GetPoLineItem(ssdcboPoNumb.Text, CoBLineitem)
'        Call EnableControlsLine(False)
'    End If
'
'
'End Sub

Private Sub Form_Load()
'    deIms.cnIms.Open
'    deIms.NameSpace = "SAKHA"
    
    GetPriorityList
    GetShipperName
    GetShiptoName
    GetSoldToName
    GetDestinationName
    GetTermOfDelivery
    GetPoInfoForPackinglist
'    Call DisableButtons(Me, NavBar1)
    GetManifestNumberList
'   GetPackingAlloflist (cboPackingNumber)
   
    
    NavBar1.NewEnabled = False
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.PrintEnabled = False
    NavBar1.EMailEnabled = False
    NavBar1.CloseEnabled = False
    NavBar1.CancelEnabled = False
    
    Call EnableControls(False)
End Sub

Private Sub AddPoNumb(rst As ADODB.Recordset)
Dim str As String

    ssdcboPoNumb.RemoveAll
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        str = (rst!po_ponumb & "") & ";" & (rst!PO_Date & "") & ";" & (rst!po_priocode & "") & ";"
        str = str & (rst!po_suppcode & "") & ";" & (rst!po_stas & "")
        
        ssdcboPoNumb.AddItem str
        rst.MoveNext
    Loop
    
CleanUp:
    rst.Close
    Set rst = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rsReceptList = Nothing
    If open_forms <= 5 Then frmNavigator.Visible = True
End Sub



Private Sub NavBar1_OnCancelClick()
    Select Case SSTab1.Tab
        
        Case 0
            'kin Clear all controls on the form and disable them except the one for the Packing list number
            Call GetCancelSSTab1
            Call GetCancelSSTabLine
            Call GetCancelSSTabRecp
            Call GetCancelSSTabRemark
           ' Call rsReceptList.Delete(adAffectAllChapters)
        Case 1
            rec.Remove (rsReceptList.Fields(0).Value)
            Call rsReceptList.Delete(adAffectCurrent)
            'rsReceptList.Delete
            'rec.Remove (rsReceptList.Fields(0).Value)
            'rsReceptList.Update
        Case 2
            'Kin Clear all controls for this tab
            Call GetCancelSSTabLine
            
            If Not IsNothing(pld) Then
            
                If pld.Count Then
                
                    Call pld.Remove(pld.Count)
                    Call EnableControlsLine(True)
                    
                End If
                
            End If
                
        Case 3
            TxtRemarks = ""
    End Select
End Sub

Private Sub GetCancelSSTab1()
    cboPackingNumber.ListIndex = CB_ERR
    'DTPDocudate = ""
    ssdcboShipper = ""
    'DTPshidate = ""
    SSdcboPriority = ""
    SSdcboShipto = ""
    SSdcboSoldTo = ""
    Txtshprefe = ""
    Txtcustrefe = ""
    Txtforwrefe = ""
    Txtawbnumb = ""
    Txtflig1 = ""
    Txtfrom1 = ""
    Txtflig2 = ""
    TxtRemark = ""
    'DTPicker1etd = ""
    SSdcboDestination = ""
    SSdcboDesnationFrom = ""
    'DTPicker2eta = ""
    SSdcboDestinationTo = ""
    SSdcboDestinationTo1 = ""
    SSdcboTermDelivery = ""
    Txtviacarr = ""
    LblDestination = ""
    Txtnumbpiec = ""
    Txtgrosweig = ""
    Txttotavolu = ""
    TxtMark1 = ""
    TxtMark2 = ""
    TxtMark3 = ""
    TxtMark4 = ""
    cboPackingNumber.SetFocus: Exit Sub
    
End Sub

Private Sub GetCancelSSTabLine()
    ssdcboPoNumb = ""
    CoBLineitem = ""
    TxtBeShipped = ""
    lblReqQty = ""
    lblUnitPrice = ""
    lblAmount = ""
    TxtBoxNumber = ""
    lblQtyDelv = ""
    lblQtyInv = ""
    LblTobeInven = ""
    TxtDescription = ""
    ssdcboPoNumb.SetFocus: Exit Sub

End Sub

Private Sub GetCancelSSTabRecp()
On Error Resume Next

    If IsNothing(rsReceptList) Then Exit Sub
    If rsReceptList.RecordCount = 0 Then Exit Sub
    If rsReceptList.EOF And rsReceptList.BOF Then Exit Sub
    
    rsReceptList.MoveFirst
    
    Do Until rsReceptList.EOF
        rec.Remove (rsReceptList.Fields(0).Value)
        Call rsReceptList.Delete(adAffectCurrent)
        
        rsReceptList.Update
        rsReceptList.MoveFirst
        
        If Err Then Err.Clear
    Loop
    
End Sub
Private Sub GetCancelSSTabRemark()
    TxtRemarks = ""
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnNewClick()
  Call EnableControls(True)
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
        .ParameterFields(0) = "namespace;" + deIms.Namespace + ";TRUE"
        .ParameterFields(1) = "manifestnumb;" + cboPackingNumber + ";true"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If

End Sub

Private Sub NavBar1_OnSaveClick()
 Dim Ponumb As String
 Dim cmd As ADODB.Command
 Dim cn As ADODB.Connection
  
'kin validate the data
'kin check the tab and if it is the first tab saveall else

    Set cmd = New ADODB.Command
    
    If SSTab1.Tab = 0 Then
        

        
'        Call CheckCombFields
'        Call CheckLIFields
        
        If Not CheckCombFields = True Then
        Exit Sub
             
        ElseIf Not CheckLIFields = True Then
        Exit Sub

        End If
        
       
        Call InsertPackingList
        
        Call InsertPackRem
        pl.RequestedQty = TxtBeShipped
        If Not (pld Is Nothing) Then pld.UpdateAll
            'MsgBox "Insert into Packing List Detail is completed"
        If Not IsNothing(rec) Then Call rec.UpdateAll(deIms.cnIms)
            'MsgBox "Insert into Packing List Receipients is completed"
        
        Set pld = Nothing
        Set rec = Nothing
        
    ElseIf SSTab1.Tab = 2 Then
        'SSTab1.Tab
               
    ElseIf SSTab1.Tab = 1 Then
       ' Call InsertPackRecip
    End If
      

End Sub
Private Sub InsertPackingList()
Dim cmd As ADODB.Command
On Error GoTo Noinsert


 Dim Shipper As String
 Dim Priority As String
 Dim ShipTo As String
 Dim SoldTo As String
 Dim From1 As String
 Dim to1 As String
 Dim from2 As String
 Dim to2 As String
 Dim shipterm As String
 Dim Ponumb As String
 
 

    Shipper = ssdcboShipper.Columns("Code").Text
    Priority = SSdcboPriority.Columns("Code").Text
    ShipTo = SSdcboShipto.Columns("Code").Text
    SoldTo = SSdcboSoldTo.Columns("Code").Text
    From1 = SSdcboDestination.Columns("Code").Text
    to1 = SSdcboDestinationTo.Columns("Code").Text
    from2 = SSdcboDesnationFrom.Columns("Code").Text
    to2 = SSdcboDestinationTo1.Columns("Code").Text
    shipterm = SSdcboTermDelivery.Columns("Code").Text
    
    Ponumb = ssdcboPoNumb.Columns("po-number").Text

  Set cmd = New ADODB.Command
  
    With cmd
        .CommandText = "Upd_Ins_PACKLIST"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@manfnumb", adVarChar, adParamInput, 10, cboPackingNumber)
        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.Namespace)
        .Parameters.Append .CreateParameter("@shipcode", adVarChar, adParamInput, 10, Shipper)
        .Parameters.Append .CreateParameter("@shipdate", adDBTimeStamp, adParamInput, 15, DTPshidate)
        .Parameters.Append .CreateParameter("@shipterm", adVarChar, adParamInput, 24, shipterm)
        .Parameters.Append .CreateParameter("@viacarrier", adVarChar, adParamInput, 20, Txtviacarr)
        .Parameters.Append .CreateParameter("@shiprefe", adVarChar, adParamInput, 20, Txtshprefe)
        .Parameters.Append .CreateParameter("@custrefe", adVarChar, adParamInput, 20, Txtcustrefe)
        .Parameters.Append .CreateParameter("@dest", adVarChar, adParamInput, 15, LblDestination)
        .Parameters.Append .CreateParameter("@numbpiec", adInteger, adParamInput, 4, Txtnumbpiec)
        .Parameters.Append .CreateParameter("@grosweig", adDouble, adParamInput, 10, Txtgrosweig)
        .Parameters.Append .CreateParameter("@totavolu", adDouble, adParamInput, 20, Txttotavolu)
        .Parameters.Append .CreateParameter("@mark1", adVarChar, adParamInput, 20, TxtMark1)
        .Parameters.Append .CreateParameter("@mark2", adVarChar, adParamInput, 20, TxtMark2)
        .Parameters.Append .CreateParameter("@mark3", adVarChar, adParamInput, 20, TxtMark3)
        .Parameters.Append .CreateParameter("@mark4", adVarChar, adParamInput, 20, TxtMark4)
        .Parameters.Append .CreateParameter("@docudate", adDBTimeStamp, adParamInput, 10, DTPDocudate)
        .Parameters.Append .CreateParameter("@shtocode", adVarChar, adParamInput, 10, ShipTo)
        .Parameters.Append .CreateParameter("@sltcode", adVarChar, adParamInput, 20, SoldTo)
        .Parameters.Append .CreateParameter("@priocode", adVarChar, adParamInput, 10, Priority)
        .Parameters.Append .CreateParameter("@awbnumb", adVarChar, adParamInput, 20, Txtawbnumb)
        .Parameters.Append .CreateParameter("@hawbnumb", adVarChar, adParamInput, 20, to1)
        .Parameters.Append .CreateParameter("@fig1", adVarChar, adParamInput, 25, Txtflig1)
        .Parameters.Append .CreateParameter("@from1", adVarChar, adParamInput, 25, From1)
        .Parameters.Append .CreateParameter("@to1", adVarChar, adParamInput, 25, From1)
        .Parameters.Append .CreateParameter("@fig2", adVarChar, adParamInput, 25, Txtflig2)
        .Parameters.Append .CreateParameter("@from2", adVarChar, adParamInput, 25, from2)
        .Parameters.Append .CreateParameter("@to2", adVarChar, adParamInput, 25, to1)
        .Parameters.Append .CreateParameter("@etd", adDate, adParamInput, 10, DTPicker1etd)
        .Parameters.Append .CreateParameter("@etda", adDate, adParamInput, 10, DTPicker2eta)
        .Parameters.Append .CreateParameter("@forwrefe", adVarChar, adParamInput, 20, Txtforwrefe)
        .Parameters.Append .CreateParameter("@remk", adVarChar, adParamInput, 2000, TxtRemark)
        .Execute , , adExecuteNoRecords

      End With
      
    Set cmd = Nothing
        'MsgBox "Insert into Packinglist is completed"
    Exit Sub
    
Noinsert:
    MsgBox "Insert into Packinglist is failure"

End Sub


Private Sub InsertPackRem()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "Upd__Ins_PACKINGREMARK"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        .Parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, cboPackingNumber)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.Namespace)
        .Parameters.Append .CreateParameter("@LINENUMB", adInteger, adParamInput, 4, CoBLineitem)
        .Parameters.Append .CreateParameter("@remk", adVarChar, adParamInput, 400, TxtRemarks)
        .Execute , , adExecuteNoRecords
    
    End With
    
    Set cmd = Nothing
        'MsgBox "Insert into Packinglist Remark is completed"
    Exit Sub
    
Noinsert:
        MsgBox "Insert into Packinglist Remark is failure "
        
End Sub



Private Sub pld_SaveError(sError As String, bContinue As Boolean)

    sError = sError & vbCrLf & "Continue ?"
    bContinue = MsgBox(sError, vbYesNo Or vbQuestion) = vbYes
End Sub

Private Sub rec_UpdateError(sError As String, bContinue As Boolean)
    MsgBox sError
End Sub

Private Sub SSdcboDestination_Click()
    
'    If SSdcboDestinationTo1.Text <> "" Then
'        LblDestination = SSdcboDestinationTo1.Text
'    ElseIf SSdcboDestination.Text <> "" Then
'        LblDestination = SSdcboDestination.Text
'    End If
    
End Sub

Private Sub SSdcboDestinationTo_Click()
    If SSdcboDestinationTo1.Text <> "" Then
        LblDestination = SSdcboDestinationTo1.Text
    ElseIf SSdcboDestinationTo.Text <> "" Then
        LblDestination = SSdcboDestinationTo.Text
    End If
End Sub

Private Sub SSdcboDestinationTo1_Click()
    If SSdcboDestinationTo1.Text <> "" Then
        LblDestination = SSdcboDestinationTo1.Text
    ElseIf SSdcboDestinationTo.Text <> "" Then
        LblDestination = SSdcboDestinationTo.Text
    End If
    
End Sub

Private Sub ssdcboPoNumb_Click()
    Call GetPoLineNumber(ssdcboPoNumb.Text)
    CoBLineitem.SetFocus: Exit Sub
End Sub



Private Sub GetPoInfoForPackinglist()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "SELECT PO.po_ponumb, PO.po_date, STATUS.sts_name,"
        .CommandText = .CommandText & "    SUPPLIER.sup_name , PRIORITY.pri_desc"
        .CommandText = .CommandText & " FROM PO INNER JOIN"
        .CommandText = .CommandText & "    STATUS ON PO.po_stas = STATUS.sts_code AND"
        .CommandText = .CommandText & "    PO.po_npecode = STATUS.sts_npecode INNER JOIN"
        .CommandText = .CommandText & "    SUPPLIER ON"
        .CommandText = .CommandText & "    PO.po_suppcode = SUPPLIER.sup_code AND"
        .CommandText = .CommandText & "    PO.po_npecode = SUPPLIER.sup_npecode INNER JOIN"
        .CommandText = .CommandText & "    PRIORITY ON"
        .CommandText = .CommandText & "    PO.po_priocode = PRIORITY.pri_code AND"
        .CommandText = .CommandText & "    PO.po_npecode = PRIORITY.pri_npecode"
        .CommandText = .CommandText & " WHERE (UPPER(PO.po_stas) = 'OP') AND"
        .CommandText = .CommandText & "    (PO.po_npecode = '" & deIms.Namespace & "')"
            
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    str = Chr$(1)
    ssdcboPoNumb.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    With rst
        Do While ((Not rst.EOF))
            ssdcboPoNumb.AddItem !po_ponumb & str & !PO_Date & str & !sts_name & str & !sup_name & str & !pri_desc & ""
            rst.MoveNext
        Loop
    End With
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub GetPoLineNumber(Ponumb As String)
Dim str As String
Dim cmd  As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
          
        .CommandText = " SELECT POITEM.poi_liitnumb"
        .CommandText = .CommandText & " FROM POITEM "
        .CommandText = .CommandText & " WHERE (UPPER(POITEM.poi_stasdlvy) IN ('RP', 'RC')) AND"
        .CommandText = .CommandText & "     (UPPER(POITEM.poi_stasship) IN ('NS', 'SP')) AND"
        .CommandText = .CommandText & "     (UPPER(POITEM.poi_stasliit) = 'OP') AND"
        .CommandText = .CommandText & "     (POITEM.poi_ponumb = '" & Ponumb & "')"
    
        Set rst = .Execute
        
    End With
    
    Call PopuLateFromRecordSet(CoBLineitem, rst, "poi_liitnumb", True)
End Sub


Private Sub GetPoLineItem(Ponumb As String, Linenumb As String)
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim Box As Integer

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT POITEM.poi_liitnumb, POITEM.poi_stasliit,POITEM.poi_primreqdqty,"
        .CommandText = .CommandText & "     POITEM.poi_qtytobedlvd,POITEM.poi_unitprice, POITEM.poi_totaprice, "
        .CommandText = .CommandText & "     POITEM.poi_qtyship, POITEM.poi_qtydlvd, "
        .CommandText = .CommandText & "     POITEM.poi_qtyinvt,POITEM.poi_desc"
        .CommandText = .CommandText & " FROM POITEM INNER JOIN"
        .CommandText = .CommandText & "     PO ON POITEM.poi_ponumb = PO.po_ponumb AND"
        .CommandText = .CommandText & "     POITEM.poi_npecode = PO.po_npecode INNER JOIN"
        .CommandText = .CommandText & "     STATUS ON PO.po_stas = STATUS.sts_code AND"
        .CommandText = .CommandText & "     PO.po_npecode = Status.sts_npecode"
        .CommandText = .CommandText & " WHERE (UPPER(POITEM.poi_stasdlvy) IN ('RP', 'RC')) AND"
        .CommandText = .CommandText & "     (UPPER(POITEM.poi_stasship) IN ('NS', 'SP')) AND"
        .CommandText = .CommandText & "     (UPPER(POITEM.poi_stasliit) = 'OP') AND"
        .CommandText = .CommandText & "     (POITEM.poi_ponumb = '" & Ponumb & "') AND"
        .CommandText = .CommandText & "     (POITEM.poi_liitnumb = '" & Linenumb & "')"
        Set rst = .Execute
    End With

    If pld Is Nothing Then Set pld = New PackingListDetls

    TxtBoxNumber = pld.Count + 1
    TxtDescription = rst!poi_desc
    lblReqQty = rst!poi_primreqdqty
    lblUnitPrice = rst!poi_unitprice
    lblAmount = rst!poi_totaprice
    lblQtyDelv = rst!poi_qtydlvd
    lblQtyInv = rst!poi_qtyship
    LblTobeInven = (rst!poi_qtydlvd - rst!poi_qtyship)
    TxtBeShipped = LblTobeInven
    
    
    Set pl = pld.Add(cboPackingNumber, deIms.Namespace, TxtBoxNumber, ssdcboPoNumb.Text, CoBLineitem, TxtBoxNumber, LblTobeInven, lblUnitPrice, lblAmount)
    
    Set pl.Connection = deIms.cnIms
    
    
        
    
           'If rst Is Nothing Then Exit Sub
          ' If rst.State And adStateOpen = adStateClosed Then Exit Sub
        
'    str = Chr$(1)
'    CoBLineitem.FieldSeparator = str
'    If rst.RecordCount = 0 Then GoTo CleanUp
'
'    rst.MoveFirst
'
'    With rst
'        Do While ((Not rst.EOF))
'            CoBLineitem.AddItem !poi_liitnumb & str & !poi_stasliit & str _
'            & !poi_primreqdqty & str & !poi_unitprice & str & !poi_totaprice & str _
'            & !poi_qtyship & str & !poi_qtydlvd & str & !poi_qtyinvt & ""
'            rst.MoveNext
'        Loop
'    End With
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub


Private Sub GetShipperName()
    Dim str As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT SHIPPER.shi_code, SHIPPER.shi_name"
        .CommandText = .CommandText & " From SHIPPER"
        .CommandText = .CommandText & " WHERE SHI_NPECODE = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " order by SHIPPER.shi_code "
         Set rst = .Execute
    End With
    
         
    str = Chr$(1)
    ssdcboShipper.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        ssdcboShipper.AddItem rst!shi_name & str & (rst!shi_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub GetPriorityList()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT pri_code, pri_desc"
        .CommandText = .CommandText & " From PRIORITY "
        .CommandText = .CommandText & " WHERE pri_npecode = '" & deIms.Namespace & "'"
         Set rst = .Execute
    End With
    

    str = Chr$(1)
    SSdcboPriority.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SSdcboPriority.AddItem rst!pri_desc & str & (rst!pri_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub GetShiptoName()
Dim str As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT sht_code, sht_name"
        .CommandText = .CommandText & " From SHIPTO "
        .CommandText = .CommandText & " WHERE sht_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " ORDER BY sht_code "
         Set rst = .Execute
    End With
    

    str = Chr$(1)
    SSdcboShipto.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        SSdcboShipto.AddItem rst!sht_name & str & (rst!sht_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub GetSoldToName()
Dim str As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT slt_code, slt_name "
        .CommandText = .CommandText & " From SOLDTO "
        .CommandText = .CommandText & " WHERE slt_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " order by slt_code "
         Set rst = .Execute
    End With
    

    str = Chr$(1)
    SSdcboSoldTo.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SSdcboSoldTo.AddItem rst!slt_name & str & (rst!slt_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing


End Sub

Private Sub GetDestinationName()
Dim str As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT des_destcode, des_destname "
        .CommandText = .CommandText & " From Destination "
        .CommandText = .CommandText & " WHERE des_npecode = '" & deIms.Namespace & "'"

         Set rst = .Execute
    End With
    

    str = Chr$(1)
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    SSdcboDestination.FieldSeparator = str
    SSdcboDestinationTo.FieldSeparator = str
    SSdcboDesnationFrom.FieldSeparator = str
    SSdcboDestinationTo1.FieldSeparator = str
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SSdcboDestination.AddItem rst!des_destname & str & (rst!des_destcode & "")
        SSdcboDestinationTo.AddItem rst!des_destname & str & (rst!des_destcode & "")
        SSdcboDesnationFrom.AddItem rst!des_destname & str & (rst!des_destcode & "")
        SSdcboDestinationTo1.AddItem rst!des_destname & str & (rst!des_destcode & "")
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing


End Sub


Private Sub GetTermOfDelivery()
Dim str As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT tod_termcode, tod_desc "
        .CommandText = .CommandText & " From TERMOFDELIVERY "
        .CommandText = .CommandText & " WHERE tod_npecode = '" & deIms.Namespace & "'"

         Set rst = .Execute
    End With
    
    str = Chr$(1)
    SSdcboTermDelivery.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SSdcboTermDelivery.AddItem rst!tod_desc & str & (rst!tod_termcode & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing



End Sub



Private Sub SaveAll()
Dim cmd As ADODB.Command
    
    
    Set cmd = New ADODB.Command
    
    cmd.CommandText = "UPDATE_POITEM_PACKING"
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = deIms.cnIms

    cmd.Parameters.Append _
        cmd.CreateParameter("Return_Value", adInteger, adParamReturnValue)
        
    cmd.Parameters.Append _
        cmd.CreateParameter("@Namespace", adVarChar, adParamInput, 5, deIms.Namespace)

    cmd.Parameters.Append _
        cmd.CreateParameter("@PONUMB", adVarChar, adParamInput, 15, ssdcboPoNumb)


    cmd.Parameters.Append _
        cmd.CreateParameter("@LINENUMB", adVarChar, adParamInput, 6, CoBLineitem)


    cmd.Parameters.Append _
        cmd.CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, cboPackingNumber)
    
    cmd.Parameters.Append _
        cmd.CreateParameter("@MANFSRL", adInteger, adParamInput, 10, CoBLineitem)

    On Error Resume Next
    
    cmd.Execute
    
    If cmd.Parameters("Return_Value") <> 0 Then MsgBox Err.Description
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "if @@trancount > 0 commit"
    cmd.Execute
   
CleanUp:
    
    Set cmd = Nothing
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear: GoTo CleanUp
    End If
       
End Sub


Private Function GetPackingNumber(ManifestNumb As String) As Boolean
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From PACKINGLIST "
        .CommandText = .CommandText & " Where pl_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND pl_manfnumb = '" & ManifestNumb & "'"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetPackingNumber = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
End Function


Private Function CheckPackingNumber() As Boolean

'    CheckPackingNumber = False
'
'    If Len(Trim$(cboPackingNumber)) = 0 Then
'
'        Call EnableControls(True)
'    Else
''        MsgBox "Packing List Entered Number do not exist": Exit Function
'
' '   If Len(Trim$(cboPackingNumber.Text)) = 0 Then _
' '       MsgBox "Packing List Entered Number cannot left empty": Exit Function
'
'
'    CheckPackingNumber = True
'    End If
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim I As Integer

    
    'kin validate the data for the Previous Tab
    If SSTab1.Tab = 0 Then
        
        Call EnableControls(True)
'        If Not CheckCombFields = True Then
'                SSTab1.Tab = PreviousTab
'            ElseIf Not CheckLIFields = True Then
'                SSTab1.Tab = PreviousTab
'        End If
        
    ElseIf SSTab1.Tab = 1 Then
'        Call GetRecipientList
        
    '    If dgRecepientList.Tag = "" Then
        
        If dgRecepientList.Tag <> "1" And SSTab1.Tab <> 1 Then
             SSTab1.Tab = 1
    
        Else
                If Not CheckCombFields = True Then
                    SSTab1.Tab = PreviousTab
                ElseIf Not CheckLIFields = True Then
                    SSTab1.Tab = PreviousTab
            End If
        
           
        End If

'            Call CheckCombFields
'            Call CheckLIFields
            
'            SSTab1 =
'        End If
    ElseIf SSTab1.Tab = 2 Then
        If Len(Trim$(cboPackingNumber)) <> 0 Then
            LlbManifest = cboPackingNumber
            LlbShipTo = SSdcboShipto
        
        End If
    
'            If SSTab1.Tab = 2 Then
'                SSTab1.TabEnabled(2) = False
'            End If
'        LlbManifest = cboPackingNumber
'        Call GetPoLineItem(ssdcboPoNumb.Text, CoBLineitem)
    End If
    
'    If dgRecepientList.Tag <> "" And SSTab1.Tab <> 1 Then
'        If SSTab1.Tab Then
'            LlbManifest = cboPackingNumber
'            LlbShipTo = SSdcboShipto
'        End If
'
'        If PreviousTab = 0 Then
'
'            If Not CheckCombFields = True Then
'                SSTab1.Tab = PreviousTab
'            ElseIf Not CheckLIFields = True Then
'                SSTab1.Tab = PreviousTab
'            End If
'
'         End If
'
'
'        If Err Then Err.Clear
'    Else
'        GetRecipientList
'        Call EnableControls(False)
'    End If
End Sub


Private Sub txt_Recipient_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim$(txt_Recipient)) Then cmd_Add_Click
    End If
End Sub



Private Sub TxtBeShipped_Change()
    If TxtBeShipped >= lblQtyInv Then
        MsgBox "You are over shipped Material": Exit Sub
    End If
End Sub

Private Sub TxtBoxNumber_Validate(Cancel As Boolean)
    Cancel = True
    
    TxtBoxNumber = Trim$(TxtBoxNumber)
    
    If Len(TxtBoxNumber) Then
        If Not IsNumeric(TxtBoxNumber) Then
   
            MsgBox "Box Number must be numeric"
            TxtBoxNumber.SetFocus: Exit Sub
        ElseIf Len(TxtBoxNumber) > 0 Then
            pl.BoxNumber = CInt(TxtBoxNumber)
            
        End If
    
    Else
        MsgBox "Box Number cannot be left empty"
        TxtBoxNumber.SetFocus: Exit Sub

    End If
    Cancel = False

End Sub

Private Sub cboPackingNumber_Change()
On Error GoTo Noitem


    If dgRecepientList.Tag <> "" Then
        Set dgRecepientList.DataSource = Nothing
    
        dgRecepientList.Tag = ""
        Set rsReceptList = Nothing
    End If
    
    If IsNothing(pld) Then Set pld = New PackingListDetls

'    If Len(Trim$(cboPackingNumber)) Then
'
'        Call IndexOf(cboPackingNumber, cboPackingNumber)
       
        
        cboPackingNumber.Tag = ""
        NavBar1.SaveEnabled = True
        
        Call EnableControls(True)
        Call Clearform
        
        
        If Not Len(Trim$(ssdcboShipper)) = 0 Then
            Call CheckCombFields
            Call CheckLIFields
        End If
'    Else
'        MsgBox "Please enter new number"

'    End If

    Exit Sub

Noitem:
  If Err Then Err.Clear

End Sub
 Private Sub cboPackingNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cboPackingNumber_Validate(Cancel As Boolean)

'    Cancel = False
'
'    If (Len(cboPackingNumber) And (Len(cboPackingNumber.Tag) = 0)) Then
'
'
'        If GetPackingNumber(cboPackingNumber) Then
'            cboPackingNumber.ListIndex = CB_ERR
'            GetRecipientList
'            Call EnableControls(False)
'            MsgBox "Packing List Entered Number is already exist": Exit Sub
            
'            If SSTab1.Tab = 2 Then
'                SSTab1.TabEnabled(2) = False
'            End If
           
'        Else
'            Cancel = False
'            AssignDefault
'        End If
'
'    End If
End Sub

Private Sub Txtnumbpiec_Change()
On Error Resume Next


    If Len(Txtnumbpiec) Then
    
      If Len(Trim$(Txtnumbpiec)) = 0 Then
            MsgBox "Number of Pieces cannot be left empty"
            Txtnumbpiec.SetFocus: Exit Sub
      
      ElseIf Not IsNumeric(Txtnumbpiec) Then
            MsgBox "Number of Pieces must be numeric"
            Txtnumbpiec.SetFocus: Exit Sub
      End If
      
    End If
End Sub



Private Sub Txttotavolu_Change()
    
    If Len(Txttotavolu) Then
    
        If Not IsNumeric(Txttotavolu) Then
            MsgBox "Total Volume must be numeric"
            Txttotavolu.SetFocus: Exit Sub
        End If
        
    End If
    
End Sub


Private Function CheckLIFields() As Boolean
On Error Resume Next

    CheckLIFields = False
    
    If Len(Trim$(cboPackingNumber)) = 0 Then
        MsgBox "The Manifest Number cannot be left empty"
        cboPackingNumber.SetFocus: Exit Function
    End If
    
    If Len(Trim$(DTPDocudate)) = 0 Then
         DTPDocudate.SetFocus: Exit Function

    ElseIf Not IsDate(DTPDocudate) Then
        MsgBox "The document date must be date type"
        DTPDocudate.SetFocus: Exit Function
    End If


    If Len(Trim$(DTPshidate)) = 0 Then
         DTPshidate.SetFocus: Exit Function
    ElseIf Not IsDate(DTPshidate) Then
        MsgBox "The ship date must be date type"
        DTPshidate.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtawbnumb)) = 0 Then
        MsgBox "AIR WAY BILL canot be left empty"
        Txtawbnumb.SetFocus: Exit Function
    End If
    
        
'    If Len(Trim$(Txtflig1)) = 0 Then
'        MsgBox "Flight or Voyage canot be left empty"
'        Txtflig1.SetFocus: Exit Function
'    End If
    
    If Len(Trim$(DTPicker1etd)) = 0 Then
        MsgBox "Edit Date canot be left empty"
        DTPicker1etd.SetFocus: Exit Function
    ElseIf Not IsDate(DTPicker1etd) Then
        MsgBox "The Edit date must be date type"
        
    End If
    
    If Len(Trim$(DTPicker2eta)) = 0 Then
        MsgBox "Edit Date canot be left empty"
        DTPicker2eta.SetFocus: Exit Function
    ElseIf Not IsDate(DTPicker2eta) Then
        MsgBox "The Edit Date must be date type"
    End If
    
    If Len(Trim$(Txtviacarr)) = 0 Then
        MsgBox "Vai Carrier canot be left empty"
        Txtviacarr.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtgrosweig)) = 0 Then
        MsgBox "Gross Weight cannot be left empty"
        Txtgrosweig.SetFocus: Exit Function
     
    ElseIf Not IsNumeric(Txtgrosweig) Then
           MsgBox "Gross Weight must be numeric"
           Txtgrosweig.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtnumbpiec)) = 0 Then
           MsgBox "Number of Pieces cannot be left empty"
           Txtnumbpiec.SetFocus: Exit Function
      
    ElseIf Not IsNumeric(Txtnumbpiec) Then
        MsgBox "Number of Pieces must be numeric"
        Txtnumbpiec.SetFocus: Exit Function
    End If
    
   
    CheckLIFields = True
    
    If Not IsNumeric(Txttotavolu) Then
          MsgBox "Total Volume must be numeric and entry optional"
          Txttotavolu.SetFocus: Exit Function
    End If
        
    If Len(Trim$(Txtflig2)) = 0 Then
        
        Txtflig2.SetFocus: Exit Function
        
    End If
            
    If Len(Trim$(TxtRemark)) = 0 Then
        
        TxtRemark.SetFocus: Exit Function
        
    End If
    
    If Len(Trim$(TxtMark1)) = 0 Then
        
        TxtMark1.SetFocus: Exit Function
        
    End If
    
    If Len(Trim$(TxtMark2)) = 0 Then
        
        TxtMark2.SetFocus: Exit Function
        
    End If
    
    If Len(Trim$(TxtMark3)) = 0 Then
        
        TxtMark3.SetFocus: Exit Function
        
    End If
    
    If Len(Trim$(TxtMark4)) = 0 Then
        
        TxtMark4.SetFocus: Exit Function
        
    End If
    
        CheckLIFields = True: Err.Clear
End Function

Private Function CheckLineitemFlied()
On Error Resume Next
    CheckLineitemFlied = False
    
    If Len(ssdcboPoNumb.Text) = 0 Then
        
        MsgBox "PO Number canot be left empty"
        ssdcboPoNumb.SetFocus: Exit Function
    End If
            
    If Len(CoBLineitem.Text) = 0 Then
        
        MsgBox "Line Item Number canot be left empty"
        CoBLineitem.SetFocus: Exit Function
    End If
            
        'ElseIf Len(TxtBeShipped) = 0 Then
        
    If Len(TxtBeShipped) > 0 Then
        pl.RequestedQty = CInt(TxtBeShipped)
        MsgBox "Quantity Been shipped cannot be left empty"
    ElseIf Not IsNumeric(TxtBeShipped) Then
        MsgBox "Quantity been shipped must be numeric"
        TxtBeShipped.SetFocus: Exit Function
    End If
    
            
        'ElseIf Len(TxtBoxNumber) = 0 Then
        
   If Len(TxtBoxNumber) > 0 Then
       pl.BoxNumber = CInt(TxtBoxNumber)
       MsgBox "Box Number cannot be left empty"
   ElseIf Not IsNumeric(TxtBoxNumber) Then
       MsgBox "Box Number must be numeric"
       TxtBoxNumber.SetFocus: Exit Function
   End If
    
        CheckLineitemFlied = True: Err.Clear
            
End Function


Private Function CheckCombFields() As Boolean
On Error Resume Next
    CheckCombFields = False
    
    If Len(ssdcboShipper.Text) = 0 Then
        MsgBox "Shipper Name canot be left empty"
        If ssdcboShipper.Enabled Then ssdcboShipper.SetFocus:
        Exit Function
    End If
    

    If Len(SSdcboPriority.Text) = 0 Then
        MsgBox "Priority Name canot be left empty"
        SSdcboPriority.SetFocus: Exit Function
    End If
   
    If Len(SSdcboShipto.Text) = 0 Then
        MsgBox "Ship To Name canot be left empty"
        SSdcboShipto.SetFocus: Exit Function
    End If
    
    
    If Len(SSdcboSoldTo.Text) = 0 Then
        MsgBox "Sold To Name canot be left empty"
        SSdcboSoldTo.SetFocus: Exit Function
    End If
    
    
    If Len(SSdcboTermDelivery.Text) = 0 Then
        MsgBox "Term of Delivery canot be left empty"
        SSdcboTermDelivery.SetFocus: Exit Function
    End If
    
    If Len(SSdcboDestination.Text) = 0 Then
        MsgBox "Destination cannot be left empty"
        SSdcboDestination.SetFocus: Exit Function
        
    End If
    
    If Len(SSdcboDestinationTo.Text) = 0 Then
        MsgBox "Destination cannot be left empty"
        SSdcboDestinationTo.SetFocus: Exit Function
        
    End If
   
     CheckCombFields = True
    
    If Len(SSdcboDesnationFrom.Text) = 0 Then
        
        SSdcboDesnationFrom.SetFocus: Exit Function
        
    End If
    
    If Len(SSdcboDestinationTo1.Text) = 0 Then
        
        SSdcboDestinationTo1.SetFocus: Exit Function
        
    End If
    
    If Err Then Err.Clear
End Function


Private Sub cmd_Add_Click()

    
    If Len(Trim$(txt_Recipient)) Then
        Call AddRecepients(txt_Recipient)
        txt_Recipient = ""
    Else
        dgRecepients_DblClick
    End If

End Sub

Private Sub cmd_Remove_Click()
On Error Resume Next
    rec.Remove (rsReceptList.Fields(0).Value)
    Call rsReceptList.Delete(adAffectCurrent)
    'rsReceptList.Delete
    If Err Then Err.Clear
End Sub

Private Sub AddRecepients(Recepient As String)
Dim repnumber As Integer

    Recepient = Trim$(Recepient)
    If Len(Recepient) = 0 Then Exit Sub
    If opt_FaxNum Then Recepient = "FAX!" & Recepient
    If IsNumeric(Recepient) Then Recepient = "FAX!" & Recepient
    
    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        
        rsReceptList.LockType = adLockOptimistic
        Call rsReceptList.Fields.Append("Recipient", adVarChar, 60, adFldUpdatable)
        
        rsReceptList.Open
        Set rec = New imsPackingListRecp
        Set dgRecepientList.DataSource = rsReceptList
        
    End If
        
    If Not IsRecipientInList(Recepient) Then
        Call rsReceptList.AddNew(Array("Recipient"), Array(Recepient))
        
        rsReceptList.Update
        Call rec.Add(deIms.Namespace, cboPackingNumber, Recepient, Recepient)
    End If
End Sub


Private Function IsRecipientInList(RecepientName As String) As Boolean
On Error Resume Next
Dim BK As Variant
    
    
    If rsReceptList.RecordCount = 0 Then Exit Function
    If Not (rsReceptList.EOF Or rsReceptList.BOF) Then BK = rsReceptList.Bookmark
    
    rsReceptList.MoveFirst
    Call rsReceptList.Find("Recipient = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
    
    If Not (rsReceptList.EOF) Then
        
        If opt_Email Then
            MsgBox "Email Address Already in list"
        ElseIf opt_FaxNum Then
            MsgBox "Fax Number Already in list"
        End If
        IsRecipientInList = True
    End If
    
    rsReceptList.Bookmark = BK
    If Err Then Err.Clear
End Function

Private Sub dgRecepients_DblClick()
    If dgRecepients.ApproxCount > 0 Then _
        Call AddRecepients(dgRecepients.Columns(1).Text)
End Sub

Private Sub NavBar1_OnEMailClick()
Dim FileName As String
Dim Addresses() As String
Dim Attachments(1) As String

On Error Resume Next
    BeforePrint
    Call SendEmailAndFax(rsReceptList, "Recipient", Caption, "")
    
End Sub

Private Sub BeforePrint()
     With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
        .ParameterFields(0) = "namespace;" + deIms.Namespace + ";TRUE"
        .ParameterFields(1) = "manifestnumb;" + cboPackingNumber + ";true"
    End With
    
End Sub
Private Sub opt_Email_GotFocus()
    Call HighlightBackground(opt_Email)
End Sub

Private Sub opt_Email_LostFocus()
    Call NormalBackground(opt_Email)
End Sub

Private Sub opt_Email_Click()
Dim co As MSDataGridLib.Column

    Set co = dgRecepients.Columns(1)
    co.Caption = "Email Address"
    co.DataField = "phd_mail"
    
    dgRecepients.Columns(0).DataField = "phd_name"
    Set dgRecepients.DataSource = GetAddresses(deIms.Namespace, deIms.cnIms, adLockReadOnly, atEmail)
End Sub

Private Sub opt_FaxNum_Click()
On Error Resume Next
Dim co As MSDataGridLib.Column
    
    Set co = dgRecepients.Columns(1)
    co.Caption = "Fax Number"
    co.DataField = "phd_faxnumb"
    
    dgRecepients.Columns(0).DataField = "phd_name"
     
    Set dgRecepients.DataSource = GetAddresses(deIms.Namespace, deIms.cnIms, adLockReadOnly, atFax)
End Sub

Private Sub opt_FaxNum_GotFocus()
    Call HighlightBackground(opt_FaxNum)
End Sub

Private Sub opt_FaxNum_LostFocus()
    Call NormalBackground(opt_FaxNum)
End Sub


Private Sub EnableControls(bEnable As Boolean)
On Error Resume Next
Dim ctl As Control

    For Each ctl In Controls
        If (Not ((TypeOf ctl Is Label) Or (TypeOf ctl Is SSTab))) Then ctl.Enabled = bEnable
        If Err Then Err.Clear
    Next ctl
    
    cboPackingNumber.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub AssignDefault()
Dim str As String
    
'    str = Format$(Now(), "mm/dd/yyyy")
    
    DTPDocudate.Value = Date
    DTPshidate.Value = Date
    Txttotavolu = 0

    DTPicker1etd.Value = Date
    DTPicker2eta.Value = Date
    
    cboPackingNumber.Tag = cboPackingNumber.Text
    
    If IsNothing(pld) Then Set pld = New PackingListDetls
    If IsNothing(rec) Then Set rec = New imsPackingListRecp
    
End Sub

Private Sub EnableControlsLine(bEnable As Boolean)
On Error Resume Next

    ssdcboPoNumb.Enabled = bEnable
    CoBLineitem.Enabled = bEnable
End Sub

Private Sub GetManifestNumberList()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
    
    .CommandText = " SELECT pl_manfnumb "
    .CommandText = .CommandText & " From PACKINGLIST"
    .CommandText = .CommandText & " WHERE pl_npecode = '" & deIms.Namespace & "'"
    .CommandText = .CommandText & " order by pl_manfnumb "
    Set rst = .Execute
   End With
    
    If rst.RecordCount = 0 Then GoTo CleanUp

    rst.MoveFirst

    
        Do While ((Not rst.EOF))
            cboPackingNumber.AddItem rst!pl_manfnumb
            rst.MoveNext
        Loop
CleanUp:
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
End Sub

Private Sub GetRecipientList()
On Error Resume Next
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = deIms.cnIms
    
    With cmd
        .CommandText = "SELECT plrc_rec Recipient"
        .CommandText = .CommandText & " FROM PACKINGREC"
        .CommandText = .CommandText & " WHERE plrc_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND plrc_manfnumb = '" & cboPackingNumber & "'"
        
        Set rsReceptList = .Execute
        
        dgRecepientList.Tag = "1"
        Set dgRecepientList.DataSource = rsReceptList
        If Err Then MsgBox Err.Description: Err.Clear
    End With
    
    Set cmd = Nothing
    If rsReceptList.BOF And rsReceptList.EOF Then Set rsReceptList = Nothing
End Sub

'Private Function CheckAddingRecord(Number As String) As Boolean
'Dim str As String
'
'       Cancel = False
'
'    If (Len(cboPackingNumber) And (Len(cboPackingNumber.Tag) = 0)) Then
'
'            cboPackingNumber.DataChanged = True
'End Function

Public Sub GetPackingAlloflist(Manunumber As String)
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = New ADODB.Command
        
    With cmd
        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = "SELECT  pl_shipcode,pl_shipdate, pl_shipterm, pl_viacarr, "
        .CommandText = .CommandText & " pl_shiprefe,pl_custrefe, pl_dest, pl_numbpiec, "
        .CommandText = .CommandText & " pl_grosweig,pl_totavolu, pl_mark1, pl_mark2, "
        .CommandText = .CommandText & " pl_mark3, pl_mark4,pl_docudate, pl_shtocode, pl_sltcode,"
        .CommandText = .CommandText & " pl_priocode,pl_awbnumb, pl_hawbnumb, pl_fig1, "
        .CommandText = .CommandText & " pl_from1, pl_fig2,pl_from2 , pl_etd, pl_eta, "
        .CommandText = .CommandText & " pl_forwrefe , pl_remk, pl_to2, pl_to1"
        .CommandText = .CommandText & " From PACKINGLIST "
        .CommandText = .CommandText & " WHERE pl_manfnumb = '" & Manunumber & "' AND "
        .CommandText = .CommandText & " pl_npecode = '" & deIms.Namespace & "' "
        
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    
    ssdcboShipper = rst!pl_shipcode
    DTPDocudate = rst!pl_docudate
    DTPshidate = rst!pl_shipdate
    SSdcboPriority = rst!pl_priocode
    SSdcboShipto = rst!pl_shtocode
    SSdcboSoldTo = rst!pl_sltcode
    Txtshprefe = rst!pl_shiprefe
    Txtcustrefe = rst!pl_custrefe
    Txtforwrefe = rst!pl_forwrefe
    Txtawbnumb = rst!pl_awbnumb
    Txtflig1 = rst!pl_fig1
    Txtfrom1 = rst!pl_from1
    Txtflig2 = rst!pl_fig2
    TxtRemark = rst!pl_remk
    DTPicker1etd = rst!pl_etd
    DTPicker2eta = rst!pl_eta
    SSdcboDestination = rst!pl_dest
    SSdcboDestinationTo = rst!pl_hawbnumb
    SSdcboDesnationFrom = rst!pl_from2
    SSdcboDestinationTo1 = rst!pl_to2
    SSdcboTermDelivery = rst!pl_shipterm
    Txtviacarr = rst!pl_viacarr
    LblDestination = rst!pl_to1
    Txtnumbpiec = rst!pl_numbpiec
    Txtgrosweig = rst!pl_grosweig
    Txttotavolu = rst!pl_totavolu
    TxtMark1 = rst!pl_mark1
    TxtMark2 = rst!pl_mark2
    TxtMark3 = rst!pl_mark3
    TxtMark4 = rst!pl_mark4
        
End Sub

Public Sub Clearform()
    ssdcboShipper = ""
'    DTPDocudate = ""
'    DTPshidate = ""
    SSdcboPriority = ""
    SSdcboShipto = ""
    SSdcboSoldTo = ""
    Txtshprefe = ""
    Txtcustrefe = ""
    Txtforwrefe = ""
    Txtawbnumb = ""
    Txtflig1 = ""
    Txtfrom1 = ""
    Txtflig2 = ""
    TxtRemark = ""
'    DTPicker1etd = ""
'    DTPicker2eta = ""
    SSdcboDestination = ""
    SSdcboDestinationTo = ""
    SSdcboDesnationFrom = ""
    SSdcboDestinationTo1 = ""
    SSdcboTermDelivery = ""
    Txtviacarr = ""
    LblDestination = ""
    Txtnumbpiec = ""
    Txtgrosweig = ""
    Txttotavolu = ""
    TxtMark1 = ""
    TxtMark2 = ""
    TxtMark3 = ""
    TxtMark4 = ""

End Sub
