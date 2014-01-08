VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#8.0#0"; "LRNavigators.ocx"
Begin VB.Form frm_PackingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List / Manifest Management "
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   10200
   Tag             =   "02030200"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Packing List"
      TabPicture(0)   =   "PackingListnew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "PackingListnew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl_New"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl_Recipients"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dgRecepients"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dgRecepientList"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fra_FaxSelect"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmd_Add"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmd_Remove"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt_Recipient"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Line Item"
      TabPicture(2)   =   "PackingListnew.frx":0038
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
      TabPicture(3)   =   "PackingListnew.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtRemarks"
      Tab(3).Control(0).Enabled=   0   'False
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
         Left            =   -74160
         TabIndex        =   61
         Top             =   1920
         Width           =   7935
         Begin VB.TextBox TxtBoxNumber 
            Height          =   315
            Left            =   6600
            TabIndex        =   68
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtDescription 
            Height          =   2115
            Left            =   2520
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   1920
            Width           =   5055
         End
         Begin VB.TextBox TxtBeShipped 
            Height          =   315
            Left            =   2520
            TabIndex        =   63
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   1920
            Width           =   2100
         End
         Begin VB.Label LblTobeInven 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6600
            TabIndex        =   71
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblQtyInv 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6600
            TabIndex        =   70
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblQtyDelv 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6600
            TabIndex        =   69
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblAmount 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   66
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblUnitPrice 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   65
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblReqQty 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   64
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Qty. To Ship"
            Height          =   255
            Left            =   4440
            TabIndex        =   106
            Top             =   1560
            Width           =   2100
         End
         Begin VB.Label Label18 
            Caption         =   "Qty Already Shipped"
            Height          =   255
            Left            =   4440
            TabIndex        =   105
            Top             =   1200
            Width           =   2100
         End
         Begin VB.Label Label17 
            Caption         =   "Qty. Already Delivered"
            Height          =   255
            Left            =   4440
            TabIndex        =   104
            Top             =   840
            Width           =   2100
         End
         Begin VB.Label Label16 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   1560
            Width           =   2100
         End
         Begin VB.Label Label12 
            Caption         =   "Unit Price"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   1200
            Width           =   2100
         End
         Begin VB.Label Label11 
            Caption         =   "Requested Qty"
            Height          =   315
            Left            =   240
            TabIndex        =   101
            Top             =   840
            Width           =   2100
         End
         Begin VB.Label Label10 
            Caption         =   "Box Number"
            Height          =   255
            Left            =   4440
            TabIndex        =   100
            Top             =   480
            Width           =   2100
         End
         Begin VB.Label Label7 
            Caption         =   "Quantity Being Shipped"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   480
            Width           =   2100
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
         Height          =   1695
         Left            =   120
         TabIndex        =   80
         Top             =   4680
         Width           =   9495
         Begin VB.TextBox txtShippingterms 
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   23
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtMark4 
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   33
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Txtviacarr 
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   25
            Top             =   960
            Width           =   975
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
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   26
            Top             =   1320
            Width           =   975
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
            Left            =   4440
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1320
            Width           =   1095
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
            Left            =   4440
            MaxLength       =   40
            TabIndex        =   27
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox TxtMark1 
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   30
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TxtMark2 
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   31
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox TxtMark3 
            Height          =   315
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   32
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label4 
            Caption         =   "Shipping Terms"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   1600
         End
         Begin VB.Label Label4 
            Caption         =   "Via Carrier"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   1600
         End
         Begin VB.Label Label4 
            Caption         =   "Destination"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   1600
         End
         Begin VB.Label Label4 
            Caption         =   "Number Pieces"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   1320
            Width           =   1600
         End
         Begin VB.Label Label4 
            Caption         =   "Gross Weight Kg"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   86
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Total Volume"
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   85
            Top             =   990
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 1"
            Height          =   255
            Index           =   6
            Left            =   5640
            TabIndex        =   84
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 2"
            Height          =   255
            Index           =   7
            Left            =   5640
            TabIndex        =   83
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 3"
            Height          =   255
            Index           =   8
            Left            =   5640
            TabIndex        =   82
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Marks 4"
            Height          =   255
            Index           =   9
            Left            =   5640
            TabIndex        =   81
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label LblDestination 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.TextBox TxtRemarks 
         Enabled         =   0   'False
         Height          =   4575
         Left            =   -74520
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.TextBox txt_Recipient 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72720
         TabIndex        =   48
         Top             =   3510
         Width           =   7230
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74640
         TabIndex        =   46
         Top             =   2745
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74625
         TabIndex        =   45
         Top             =   2430
         Width           =   1320
      End
      Begin VB.Frame fra_FaxSelect 
         Enabled         =   0   'False
         Height          =   1290
         Left            =   -74625
         TabIndex        =   49
         Top             =   3840
         Width           =   1770
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   60
            TabIndex        =   50
            Top             =   285
            Width           =   1635
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   60
            TabIndex        =   51
            Top             =   780
            Width           =   1635
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
         Left            =   120
         TabIndex        =   53
         Top             =   2520
         Width           =   9495
         Begin MSComCtl2.DTPicker DTPicker1etd 
            Height          =   315
            Left            =   4560
            TabIndex        =   14
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   22937603
            CurrentDate     =   36524
         End
         Begin VB.TextBox Txtawbnumb 
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   12
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Txthawbnum 
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   13
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Txtflig2 
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   19
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox TxtRemark 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   1800
            Width           =   7575
         End
         Begin VB.TextBox Txtflight1 
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   16
            Top             =   1080
            Width           =   1935
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboFrom1 
            Height          =   315
            Left            =   4560
            TabIndex        =   17
            Top             =   1080
            Width           =   2055
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
            stylesets(0).Picture=   "PackingListnew.frx":0070
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
            stylesets(1).Picture=   "PackingListnew.frx":008C
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
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDestinationTo 
            Height          =   315
            Left            =   7320
            TabIndex        =   18
            Top             =   1080
            Width           =   2055
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
            stylesets(0).Picture=   "PackingListnew.frx":00A8
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
            stylesets(1).Picture=   "PackingListnew.frx":00C4
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
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboFrom2 
            Height          =   315
            Left            =   4560
            TabIndex        =   20
            Top             =   1440
            Width           =   2055
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
            stylesets(0).Picture=   "PackingListnew.frx":00E0
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
            stylesets(1).Picture=   "PackingListnew.frx":00FC
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
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboDestinationTo1 
            Height          =   315
            Left            =   7320
            TabIndex        =   21
            Top             =   1440
            Width           =   2055
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
            stylesets(0).Picture=   "PackingListnew.frx":0118
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
            stylesets(1).Picture=   "PackingListnew.frx":0134
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
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin MSComCtl2.DTPicker DTPicker2eta 
            Height          =   315
            Left            =   7320
            TabIndex        =   15
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   22937603
            CurrentDate     =   36524
         End
         Begin VB.Label Label2 
            Caption         =   "AWB / BL"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   1700
         End
         Begin VB.Label Label2 
            Caption         =   "HAWB / TBL"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   1700
         End
         Begin VB.Label Label2 
            Caption         =   "Flight / Voyage"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   1700
         End
         Begin VB.Label Label2 
            Caption         =   "Flight / Voyage"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   1440
            Width           =   1700
         End
         Begin VB.Label Label2 
            Caption         =   "Remark"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   74
            Top             =   1800
            Width           =   1700
         End
         Begin VB.Label Label3 
            Caption         =   "ETD"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   72
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "ETA"
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   67
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "From"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   62
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   255
            Index           =   3
            Left            =   6720
            TabIndex        =   59
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "From"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   56
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "To"
            Height          =   255
            Index           =   5
            Left            =   6720
            TabIndex        =   54
            Top             =   1440
            Width           =   600
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
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   9495
         Begin VB.ComboBox cboPackingNumber 
            Height          =   315
            Left            =   2040
            TabIndex        =   1
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox Txtcustrefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   20
            TabIndex        =   10
            Tag             =   "4"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Txtforwrefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   20
            TabIndex        =   11
            Tag             =   "5"
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Txtshprefe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "3"
            Top             =   1080
            Width           =   2535
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboShipper 
            Height          =   315
            Left            =   2040
            TabIndex        =   2
            Tag             =   "1"
            Top             =   720
            Width           =   2535
            DataFieldList   =   "Column 0"
            AutoRestore     =   0   'False
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
            stylesets(0).Picture=   "PackingListnew.frx":0150
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
            stylesets(1).Picture=   "PackingListnew.frx":016C
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
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboPriority 
            Height          =   315
            Left            =   2040
            TabIndex        =   5
            Top             =   1800
            Width           =   2535
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
            stylesets(0).Picture=   "PackingListnew.frx":0188
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
            stylesets(1).Picture=   "PackingListnew.frx":01A4
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
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo5 
            Height          =   315
            Left            =   4095
            TabIndex        =   34
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
            stylesets(0).Picture=   "PackingListnew.frx":01C0
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
            stylesets(1).Picture=   "PackingListnew.frx":01DC
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
            Left            =   6840
            TabIndex        =   6
            Top             =   360
            Width           =   2535
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
            stylesets(0).Picture=   "PackingListnew.frx":01F8
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
            stylesets(1).Picture=   "PackingListnew.frx":0214
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
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboSoldTo 
            Height          =   315
            Left            =   6840
            TabIndex        =   7
            Top             =   720
            Width           =   2535
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
            stylesets(0).Picture=   "PackingListnew.frx":0230
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
            stylesets(1).Picture=   "PackingListnew.frx":024C
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
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPshidate 
            Height          =   315
            Left            =   2040
            TabIndex        =   4
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   22937603
            CurrentDate     =   36524
         End
         Begin MSComCtl2.DTPicker DTPDocudate 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   22937603
            CurrentDate     =   36524
         End
         Begin VB.Label Label1 
            Caption         =   "Packing / Manifest"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "Shipper"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "Document Date"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "Shipping Date"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "Air / Sea / Other"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "Ship To Code"
            Height          =   255
            Index           =   5
            Left            =   4800
            TabIndex        =   39
            Top             =   360
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Sold To Code"
            Height          =   255
            Index           =   6
            Left            =   4800
            TabIndex        =   38
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Shipper's Ref"
            Height          =   255
            Index           =   7
            Left            =   4800
            TabIndex        =   37
            Top             =   1080
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Customer's Ref"
            Height          =   255
            Index           =   8
            Left            =   4800
            TabIndex        =   36
            Top             =   1440
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Forwarder's Ref"
            Height          =   255
            Index           =   9
            Left            =   4800
            TabIndex        =   35
            Top             =   1800
            Width           =   1995
         End
      End
      Begin VB.ComboBox CoBLineitem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -68160
         TabIndex        =   60
         Top             =   1440
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid dgRecepientList 
         Height          =   2535
         Left            =   -72720
         TabIndex        =   47
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
         TabIndex        =   57
         Top             =   1440
         Width           =   1920
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
         stylesets(0).Picture=   "PackingListnew.frx":0268
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
         stylesets(1).Picture=   "PackingListnew.frx":0284
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
         Height          =   2295
         Left            =   -72720
         TabIndex        =   52
         Top             =   3960
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4048
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
         TabIndex        =   98
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Packing List / Manifest"
         Height          =   315
         Left            =   -74160
         TabIndex        =   97
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Ship To Code"
         Height          =   255
         Index           =   0
         Left            =   -69720
         TabIndex        =   96
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "PO Number"
         Height          =   255
         Index           =   1
         Left            =   -74160
         TabIndex        =   95
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label LlbManifest 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -71880
         TabIndex        =   55
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label LlbShipTo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68160
         TabIndex        =   58
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
         Left            =   -74160
         TabIndex        =   94
         Top             =   480
         Width           =   7935
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74640
         TabIndex        =   93
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label lbl_New 
         Caption         =   "New"
         Height          =   300
         Left            =   -74640
         TabIndex        =   92
         Top             =   3570
         Width           =   1380
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
         TabIndex        =   91
         Top             =   7440
         Width           =   2055
      End
   End
   Begin LRNavigators.NavBar NavBar1 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "PackingListnew.frx":02A0
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
      Left            =   6000
      TabIndex        =   108
      Top             =   6480
      Width           =   3900
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
Dim Rstlist As ADODB.Recordset
Dim Rstitem As ADODB.Recordset
Dim Rstnum As ADODB.Recordset
Dim Update_insert As String
Dim Form As FormMode
'Change form mode show caption text
Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next

    
    If FMode = mdCreation Then
        lblStatu.ForeColor = vbRed
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00125") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '---------------------------------------------
        
        ChangeMode = True
'    ElseIf FMode = mdModification Then
'        lblStatu.ForeColor = vbBlue
'        lblStatu.Caption = "Modification"
  
    ElseIf FMode = mdVisualization Then
        lblStatu.ForeColor = vbGreen
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00092") 'J added
        lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
        
    End If
    
       
    Form = FMode

End Function
'depend on form mode and combo value disable butten or load packinglist numbers
Public Sub cboPackingNumber_Click()
On Error Resume Next


    dgRecepientList.Tag = ""
    Set rsReceptList = Nothing
    'Call EnableControls(True)

    If Len(Trim$(cboPackingNumber)) <> 0 And Form = mdVisualization Then
           Navbar1.PrintEnabled = True
           Navbar1.EMailEnabled = True
    Else
          Navbar1.PrintEnabled = False
          Navbar1.EMailEnabled = False
    End If
    
    
    If Len(Trim$(cboPackingNumber)) <> 0 Then
    
    
'       Call cboPackingNumber_Validate(True)
'            If GetPackingnumber(cboPackingNumber) Then
'                MsgBox "Packing List Entered Number is already exist"
'                cboPackingNumber.SetFocus: Exit Sub
'            End If
        
'        Call cboPackingNumber_Validate(True)
        Call GetPackingAlloflist(cboPackingNumber)
'        If GetPackingnumber(cboPackingNumber) Then
            
        Else
           If Len(Trim$(cboPackingNumber)) <> 0 And Form = mdCreation Then
                   AssignDefault

                  Call EnableControls(True)
            End If
        End If
            
    If Err Then Call LogErr(Name & "::cboPackingNumber_Click", Err.Description, Err.number, True)

End Sub

Public Sub cboPackingNumber_DropDown()

End Sub

Private Sub cboPackingNumber_GotFocus()
    cboPackingNumber.SetFocus
    cboPackingNumber.SelStart = 1
    cboPackingNumber.Refresh
End Sub

'Load po line item number and call function to load po line item information
Private Sub CoBLineitem_Click()
Dim Result As Boolean
On Error Resume Next


'    Set pld.Item(10) = TxtBeShipped
'        'Set variable = pld.Item(1)
'        pl.Tobeship = TxtBeShipped
'    Set pl = pld.Add(cboPackingNumber, deIms.NameSpace, TxtBoxNumber, ssdcboPoNumb.Text, CoBLineitem, TxtBoxNumber, LblTobeInven, lblUnitPrice, lblAmount, TxtBeShipped)
    If Not Len(CoBLineitem) = 0 Then
        
        If Lineitemcheck = False Then
            Exit Sub: CoBLineitem.SetFocus
        End If
    End If
 

        
    If Len(Trim$(CoBLineitem)) <> 0 Then

        Call GetPoLineItem(ssdcboPoNumb.text, CoBLineitem)
'        Call ClearPOitemForm
'        Call AssignDefaulValues
    Else
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00263") 'J added
        MsgBox IIf(msg1 = "", "The line item cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        CoBLineitem.SetFocus: Exit Sub
    End If
    
If Err Then Call LogErr(Name & "::CoBLineitem_Click", Err.Description, Err.number, True)
    CoBLineitem.SelStart = Len(cboPackingNumber)
    CoBLineitem.SelLength = Len(cboPackingNumber)
End Sub

Private Sub CoBLineitem_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan for Alpha Search (11/14/2000)
    With cboPackingNumber
        text = .text
        For i = 0 To .ListCount - 1
            If text Like .list(i) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '-------------------------------------------

End Sub


Private Sub CoBLineitem_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With CoBLineitem
        text = .text
        If text <> "" Then
            If text = .list(.ListIndex) Then Exit Sub
            For i = 0 To .ListCount - 1
                If text Like .list(i) Then
                    Exit For
                End If
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub Form_Activate()
    SSTab1.TabVisible(3) = False
    cboPackingNumber.SetFocus
End Sub

' Load first time combo information and form caption text
Private Sub Form_Load()
Update_insert = ""
On Error Resume Next

    'Added by Juan (9/13/2000) for Multilingual
    Call translator.Translate_Forms("frm_PackingList")
    '------------------------------------------
    
    GetPriorityList
    GetShipperName
    GetShiptoName
    GetSoldToName
    GetDestinationName
'    GetTermOfDelivery
    GetPoInfoForPackinglist
'    Call DisableButtons(Me, NavBar1)
    GetManifestNumberList
'   GetPackingAlloflist (cboPackingNumber)
   
    
'    NavBar1.NewEnabled = True
'    NavBar1.EditEnabled = False
'    NavBar1.SaveEnabled = False
'    NavBar1.PrintEnabled = False
'    NavBar1.EMailEnabled = False
'    NavBar1.CloseEnabled = True
'    NavBar1.CancelEnabled = False
'    NavBar1.PreviousEnabled = False
'    NavBar1.FirstEnabled = False
'    NavBar1.LastEnabled = False
'    NavBar1.NextEnabled = False
    
    'Added by Muzammil 1/11/00
    Call DisableButtons(Me, Navbar1)  'M
    
    Navbar1.NewEnabled = True
    Navbar1.CloseEnabled = True
    Call EnableControls(False)
    Call ChangeMode(mdVisualization)
    frm_PackingList.Caption = frm_PackingList.Caption + " - " + frm_PackingList.Tag
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.number, True)
End Sub
'load po data to sheridan combo
Private Sub AddPoNumb(rst As ADODB.Recordset)
Dim STR As String
On Error Resume Next


    ssdcboPoNumb.RemoveAll
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        STR = (rst!po_ponumb & "") & ";" & (rst!PO_Date & "") & ";" & (rst!po_priocode & "") & ";"
        STR = STR & (rst!po_suppcode & "") & ";" & (rst!po_stas & "")
        
        ssdcboPoNumb.AddItem STR
        rst.MoveNext
    Loop
    
CleanUp:
    rst.Close
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::AddPoNumb", Err.Description, Err.number, True)
End Sub

' unload form
Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set rsReceptList = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'depend tab Clear controls on the form and disable them except the one for the Packing list number
Private Sub NavBar1_OnCancelClick()
On Error Resume Next
      Update_insert = ""
    Select Case SSTab1.Tab
        
        Case 0
            'Clear all controls on the form and disable them except the one for the Packing list number
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
            txtRemarks = ""
    End Select
    If Err Then Call LogErr(Name & "::NavBar1_OnCancelClick", Err.Description, Err.number, True)
End Sub
'Sub function clear form
Private Sub GetCancelSSTab1()
On Error Resume Next

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
    Txtflight1 = ""
    Txthawbnum = ""
    Txtflig2 = ""
    TxtRemark = ""
    'DTPicker1etd = ""
    SSdcboFrom1 = ""
    SSdcboFrom2 = ""
    'DTPicker2eta = ""
    SSdcboDestinationTo = ""
    SSdcboDestinationTo1 = ""
    txtShippingterms = ""
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
    
    If Err Then Call LogErr(Name & "::GetCancelSSTab1", Err.Description, Err.number, True)
    
End Sub
'Sub function clear line item form
Private Sub GetCancelSSTabLine()
On Error Resume Next

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
'    ssdcboPoNumb.SetFocus: Exit Sub
    If Err Then Call LogErr(Name & "::GetCancelSSTabLine", Err.Description, Err.number, True)

End Sub
'Sub function clear reception form
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
        
       If Err Then Call LogErr(Name & "::GetCancelSSTabLine", Err.Description, Err.number, True)
    Loop
    
End Sub
'Sub function clear remark form
Private Sub GetCancelSSTabRemark()
    txtRemarks = ""
End Sub
'close packing list form
Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub
'Move recordset to first position
Private Sub NavBar1_OnFirstClick()
On Error Resume Next
    
    
    If Form = mdCreation Then
        Rstitem.MoveFirst
        Call LoadValuesrst
'        Call MoveRecord(mtFirst)
    Else
        Rstlist.MoveFirst
        Call LoadValues
    End If
'    Rstitem
    If Err Then Call LogErr(Name & "::NavBar1_OnFirstClick", Err.Description, Err.number, True)
End Sub
'move recordset to last position
Private Sub NavBar1_OnLastClick()
On Error Resume Next
    
    
    If Form = mdCreation Then
        Rstitem.MoveLast
        Call LoadValuesrst
'        Call MoveRecord(mtLast)
    Else
        Rstlist.MoveLast
        Call LoadValues
    End If
    
    If Err Then Call LogErr(Name & "::NavBar1_OnLastClick", Err.Description, Err.number, True)

End Sub
' before add new record, clear form and enable the control.
Private Sub NavBar1_OnNewClick()
On Error Resume Next
'   Update_insert = "Insert"
    Call Clearform
    Call ClearPOitemForm
    txtRemarks = ""
    cboPackingNumber = ""
    Call EnableControls(True)
    Call AssignDefault
    'Muzammil -just addes this
'    NavBar1.NewEnabled = False
    Navbar1.SaveEnabled = True
    Navbar1.EMailEnabled = False
    Navbar1.PrintEnabled = False
    Call ChangeMode(mdCreation)
    If Err Then Call LogErr(Name & "::NavBar1_OnNewClick", Err.Description, Err.number, True)
    
    cboPackingNumber.SetFocus
End Sub
'Move recordset to next position
Private Sub NavBar1_OnNextClick()
On Error Resume Next
        
    If Form = mdCreation Then
        With Rstitem
        If Not .EOF Then .MoveNext
            If .EOF And .RecordCount > 0 Then
                .MoveLast
            End If
        End With
           Call LoadValuesrst
'        Call MoveRecord(mtNext)
    Else
        With Rstlist
        If Not .EOF Then .MoveNext
            If .EOF And .RecordCount > 0 Then
                .MoveLast
            End If
        End With
        Call LoadValues
    End If
     
    If Err Then Call LogErr(Name & "::NavBar1_OnNextClick", Err.Description, Err.number, True)
End Sub
'Move recordset to previous position
Private Sub NavBar1_OnPreviousClick()
On Error Resume Next
    
    
    If Form = mdCreation Then
    
        With Rstitem
            If Not .BOF Then .MovePrevious
                If .BOF And .RecordCount > 0 Then
                    .MoveFirst
                End If
        End With
        Call LoadValuesrst
'        Call MoveRecord(mtPrevious)
    Else
        With Rstlist
        If Not .BOF Then .MovePrevious
            If .BOF And .RecordCount > 0 Then
                .MoveFirst
            End If
        End With
        
        Call LoadValues
    End If
    
    
    If Err Then Call LogErr(Name & "::NavBar1_OnPreviousClick", Err.Description, Err.number, True)
End Sub
'Print crystal report
Private Sub NavBar1_OnPrintClick()
On Error Resume Next

    With MDI_IMS.CrystalReport1
        .Reset
'        .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
        .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "manifestnumb;" + cboPackingNumber + ";true"
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00213") 'J added
        .WindowTitle = IIf(msg1 = "", "Packing List", msg1) 'J modified
        Call translator.Translate_Reports("packinglist.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
If Err Then Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)

End Sub
'call validata function before save data to database
Private Sub NavBar1_OnSaveClick()
On Error Resume Next
 
 Dim Ponumb As String
 Dim cmd As ADODB.Command
 Dim cn As ADODB.Connection
 Dim i As Integer
 Dim maxFields As Integer
 Dim Result As Boolean
 Dim Item As PackingListDetls
  
'kin validate the data
'kin check the tab and if it is the first tab saveall else

    Set cmd = New ADODB.Command
       
    If SSTab1.Tab = 0 Then
        
        If Not CheckCombFields = True Then
            Exit Sub
             
        ElseIf Not CheckLIFields = True Then
            Exit Sub

        End If
                
     Call InsertPackingList
        
     pl.Tobeship = CDbl(TxtBeShipped)
     
     If Not (pld Is Nothing) Then pld.UpdateAll
            'MsgBox "Insert into Packing List Detail is completed"
        If Not IsNothing(rec) Then Call rec.UpdateAll(deIms.cnIms)
            'MsgBox "Insert into Packing List Receipients is completed"
'        If Not Len(Trim$(TxtRemarks)) = 0 Then
'            Call InsertPackRem
'        End If
        
        'Modified by Juan (9/13/2000) for Multilanguage
        msg1 = translator.Trans("M00306") 'J added
        MsgBox IIf(msg1 = "", "Insert into Packing List is completed successfully", msg1) 'J modified
        '----------------------------------------------
        
        Set pld = Nothing
        Set rec = Nothing
        
        Call ChangeMode(mdVisualization)
        Call EnableControls(False)
'        NavBar1.PrintEnabled = True
        
             If Len(Trim(cboPackingNumber)) <> 0 And Form = mdVisualization Then
            
                Navbar1.PrintEnabled = True
                Navbar1.EMailEnabled = True
                Navbar1.SaveEnabled = False
            End If
        
    ElseIf SSTab1.Tab = 2 Then
        'SSTab1.Tab
               
    ElseIf SSTab1.Tab = 1 Then
       'commented it so that the packinglist no remains in the combo after it is saved
       ' Call InsertPackRecip
    End If
       ' cboPackingNumber.Clear
        Call GetManifestNumberList
        
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub
'Call store procedure insert one record to database table packing list
Private Sub InsertPackingList()
On Error Resume Next
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
 
    Shipper = ssdcboShipper.Columns("Code").text
    Priority = SSdcboPriority.Columns("Code").text
    ShipTo = SSdcboShipto.Columns("Code").text
    SoldTo = SSdcboSoldTo.Columns("Code").text
    From1 = SSdcboFrom1.Columns("Code").text
    to1 = SSdcboDestinationTo.Columns("Code").text
    from2 = SSdcboFrom2.Columns("Code").Value
    to2 = SSdcboDestinationTo1.Columns("Code").text
    shipterm = txtShippingterms
    
    Ponumb = ssdcboPoNumb.Columns("po-number").text

  Set cmd = New ADODB.Command
  
    With cmd
        .CommandText = "Upd_Ins_PACKLIST"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@manfnumb", adVarChar, adParamInput, 15, cboPackingNumber)
        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
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
        .Parameters.Append .CreateParameter("@hawbnumb", adVarChar, adParamInput, 20, Txthawbnum)
        .Parameters.Append .CreateParameter("@fig1", adVarChar, adParamInput, 25, Txtflight1)
        .Parameters.Append .CreateParameter("@from1", adVarChar, adParamInput, 25, From1)
        .Parameters.Append .CreateParameter("@to1", adVarChar, adParamInput, 25, to1)
        .Parameters.Append .CreateParameter("@fig2", adVarChar, adParamInput, 25, Txtflig2)
        .Parameters.Append .CreateParameter("@from2", adVarChar, adParamInput, 25, from2)
        .Parameters.Append .CreateParameter("@to2", adVarChar, adParamInput, 25, to2)
        .Parameters.Append .CreateParameter("@etd", adDate, adParamInput, 10, DTPicker1etd)
        .Parameters.Append .CreateParameter("@etda", adDate, adParamInput, 10, DTPicker2eta)
        .Parameters.Append .CreateParameter("@forwrefe", adVarChar, adParamInput, 20, Txtforwrefe)
        .Parameters.Append .CreateParameter("@remk", adVarChar, adParamInput, 2000, TxtRemark)
        .Parameters.Append .CreateParameter("@user", adVarChar, adParamInput, 20, CurrentUser)
        .Execute , , adExecuteNoRecords

      End With
      
    Set cmd = Nothing
        'MsgBox "Insert into Packinglist is completed"
    Exit Sub
    
Noinsert:

    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("M00279") 'J added
    MsgBox IIf(msg1 = "", "Insert into Packinglist is failure", msg1) 'J modified
    '---------------------------------------------
    
    If Err Then Call LogErr(Name & "::InsertPackingList", Err.Description, Err.number, True)

End Sub
'insert one record to packing list remark table
Private Sub InsertPackRem()
'On Error GoTo Noinsert
'Dim cmd As ADODB.Command
'
'    Set cmd = New ADODB.Command
'
'    With cmd
'        .CommandText = "Upd__Ins_PACKINGREMARK"
'        .CommandType = adCmdStoredProc
'        .ActiveConnection = deIms.cnIms
'
'        .Parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, cboPackingNumber)
'        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.Namespace)
'        .Parameters.Append .CreateParameter("@LINENUMB", adInteger, adParamInput, 4, CoBLineitem)
'        .Parameters.Append .CreateParameter("@remk", adVarChar, adParamInput, 400, TxtRemarks)
'        .Execute , , adExecuteNoRecords
'
'    End With
'
'    Set cmd = Nothing
'        'MsgBox "Insert into Packinglist Remark is completed"
'    Exit Sub
    
'Noinsert:
'        MsgBox "Insert into Packinglist Remark is failure "
'
End Sub


'call error function
Private Sub pld_SaveError(sError As String, bContinue As Boolean)

    sError = sError & vbCrLf & "Continue ?"
    bContinue = MsgBox(sError, vbYesNo Or vbQuestion) = vbYes
End Sub
'get error message
Private Sub rec_UpdateError(sError As String, bContinue As Boolean)
    MsgBox sError
End Sub
'get value for destination lable
Private Sub SSdcboDestinationTo_Click()
On Error Resume Next
    If SSdcboDestinationTo1.text <> "" Then
        LblDestination = SSdcboDestinationTo1.text
    ElseIf SSdcboDestinationTo.text <> "" Then
        LblDestination = SSdcboDestinationTo.text
    End If
If Err Then Call LogErr(Name & "::SSdcboDestinationTo_Click", Err.Description, Err.number, True)

    SSdcboDestinationTo.SelStart = 0
    SSdcboDestinationTo.SelLength = 0
End Sub

Private Sub SSdcboDestinationTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboDestinationTo.DroppedDown Then SSdcboDestinationTo.DroppedDown = True
End Sub


Private Sub SSdcboDestinationTo_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
    
    'Added by Juan for Alpha Search (11/14/2000)
    With SSdcboDestinationTo
        text = .text
        For i = 0 To .Rows - 1
            If text Like .Columns(0).text Then
                .Row = i
                Exit For
            End If
        Next
    End With
    '-------------------------------------------

End Sub


Private Sub SSdcboDestinationTo_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboDestinationTo
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


'get value for destination lable
Private Sub SSdcboDestinationTo1_Click()
On Error Resume Next

    If SSdcboDestinationTo1.text <> "" Then
        LblDestination = SSdcboDestinationTo1.text
    ElseIf SSdcboDestinationTo.text <> "" Then
        LblDestination = SSdcboDestinationTo.text
    End If
If Err Then Call LogErr(Name & "::SSdcboDestinationTo1_Click", Err.Description, Err.number, True)
    SSdcboDestinationTo1.SelStart = 0
    SSdcboDestinationTo1.SelLength = 0
End Sub

Private Sub SSdcboDestinationTo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboDestinationTo1.DroppedDown Then SSdcboDestinationTo1.DroppedDown = True
End Sub


Private Sub SSdcboDestinationTo1_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
    
    'Added by Juan for Alpha Search (11/14/2000)
    With SSdcboDestinationTo1
        text = .text
        For i = 0 To .Rows - 1
            If text Like .Columns(0).text Then
                .Row = i
                Exit For
            End If
        Next
    End With
    '-------------------------------------------

End Sub


Private Sub SSdcboDestinationTo1_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboDestinationTo1
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub SSdcboFrom1_Click()
    SSdcboFrom1.SelStart = 0
    SSdcboFrom1.SelLength = 0
End Sub

Private Sub SSdcboFrom1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboFrom1.DroppedDown Then SSdcboFrom1.DroppedDown = True
End Sub


Private Sub SSdcboFrom1_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
    
    'Added by Juan for Alpha Search (11/14/2000)
    With SSdcboFrom1
        text = .text
        For i = 0 To .Rows - 1
            If text Like .Columns(0).text Then
                .Row = i
                Exit For
            End If
        Next
    End With
    '-------------------------------------------

End Sub


Private Sub SSdcboFrom1_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboFrom1
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub SSdcboFrom2_Click()
    SSdcboFrom2.SelStart = 0
    SSdcboFrom2.SelLength = 0
End Sub

Private Sub SSdcboFrom2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboFrom2.DroppedDown Then SSdcboFrom2.DroppedDown = True
End Sub


Private Sub SSdcboFrom2_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
    
    'Added by Juan for Alpha Search (11/14/2000)
    With SSdcboFrom2
        text = .text
        For i = 0 To .Rows - 1
            If text Like .Columns(0).text Then
                .Row = i
                Exit For
            End If
        Next
    End With
    '-------------------------------------------

End Sub


Private Sub SSdcboFrom2_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboFrom2
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


'clear po line item form, get po line item details
Private Sub ssdcboPoNumb_Click()
On Error Resume Next

    Call ClearPOitemForm
    CoBLineitem.Clear
    If (Not Len(Trim(pl.PoNumber)) = 0 And Len(Trim(pl.PoNumber)) = Null) Then
        Set pl = Nothing
        Set pld = Nothing
        Set pld = New PackingListDetls
    End If
    
    Call GetPoLineNumber(ssdcboPoNumb.text)
    CoBLineitem.SetFocus: Exit Sub
    If Err Then Call LogErr(Name & "::ssdcboPoNumb_Click", Err.Description, Err.number, True)
    ssdcboPoNumb.SelStart = Len(cboPackingNumber)
    ssdcboPoNumb.SelLength = Len(cboPackingNumber)
End Sub
'SQL statement get po information for packing list
Private Sub GetPoInfoForPackinglist()
On Error Resume Next
Dim STR As String
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
'        .CommandText = .CommandText & " WHERE (UPPER(PO.po_stas) = 'OP') AND"
        .CommandText = .CommandText & " WHERE (UPPER(PO.po_stasship) IN ('NS', 'SP')) AND"
        .CommandText = .CommandText & " (UPPER(PO.po_stasdelv) IN ('RC', 'RP')) AND "
        .CommandText = .CommandText & " (UPPER(PO.po_stas) = 'OP') AND"
        .CommandText = .CommandText & " (PO.po_npecode = '" & deIms.NameSpace & "')"
        .CommandText = .CommandText & " order by PO.po_ponumb "
            
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    STR = Chr$(1)
    ssdcboPoNumb.FieldSeparator = STR
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    With rst
        Do While ((Not rst.EOF))
            ssdcboPoNumb.AddItem !po_ponumb & STR & !PO_Date & STR & !sts_name & STR & !sup_name & STR & !pri_desc & ""
            rst.MoveNext
        Loop
    End With
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
 If Err Then Call LogErr(Name & "::GetPoInfoForPackinglist", Err.Description, Err.number, True)
End Sub
'SQl statement get po line item number for po line item form
Private Sub GetPoLineNumber(Ponumb As String)
On Error Resume Next
Dim STR As String
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
        .CommandText = .CommandText & "     and poi_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & "     order by cast((poi_liitnumb) as int) "
    
        Set rst = .Execute
        
    End With
    
    Call PopuLateFromRecordSet(CoBLineitem, rst, "poi_liitnumb", True)
    
If Err Then Call LogErr(Name & "::GetPoLineNumber", Err.Description, Err.number, True)
End Sub
'SQl statement get po line item information for po line item form
Private Function GetPoLineItem(Ponumb As String, Linenumb As String) As Recordset
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
'Dim rst As ADODB.Recordset
'Dim Box As Integer

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
        Set Rstitem = .Execute
    End With

    
    If Not pl Is Nothing Then pl.Tobeship = CDbl(TxtBeShipped)
    
    If Rstitem Is Nothing Then Exit Function
    If Rstitem.RecordCount = 0 Then GoTo CleanUp
    
  
    
    If Rstitem.BOF And Rstitem.EOF Then
        TxtDescription = ""
        lblReqQty = ""
        lblUnitPrice = ""
        lblAmount = ""
        lblQtyDelv = ""
        lblQtyInv = ""
        LblTobeInven = ""
        TxtBeShipped = ""
        TxtBoxNumber = ""
        Exit Function
    
    Else
        TxtDescription = Rstitem!poi_desc
        lblReqQty = FormatNumber((Rstitem!poi_primreqdqty), 4)
        lblUnitPrice = FormatNumber((Rstitem!poi_unitprice), 4)
        lblAmount = FormatNumber((Rstitem!poi_totaprice), 4)
        lblQtyDelv = FormatNumber((Rstitem!poi_qtydlvd), 4)
        lblQtyInv = FormatNumber((Rstitem!poi_qtyship), 4)
        LblTobeInven = FormatNumber((Rstitem!poi_qtydlvd - Rstitem!poi_qtyship), 4)
        TxtBeShipped = FormatNumber((LblTobeInven), 4)
        TxtBoxNumber = pld.Count + 1
    End If
    
       If pld Is Nothing Then Set pld = New PackingListDetls

    
    Set pl = pld.Add(cboPackingNumber, deIms.NameSpace, TxtBoxNumber, ssdcboPoNumb.text, CoBLineitem, TxtBoxNumber, LblTobeInven, lblUnitPrice, lblAmount, TxtBeShipped, CurrentUser)

    Set pl.Connection = deIms.cnIms
     
'     Rstitem.AddNew
'    If Not Len(Trim$(CoBLineitem)) = 0 Then
'        If Lineitemcheck = False Then
'            Exit Sub: CoBLineitem.SetFocus
'        End If
'    End If

    
CleanUp:
    Rstitem.Close
    Set cmd = Nothing
    Set Rstitem = Nothing
If Err Then Call LogErr(Name & "::GetPoLineItem", Err.Description, Err.number, True)
End Function
'SQL statement get shipper name list for shipper combo
Private Sub GetShipperName()
On Error Resume Next
    Dim STR As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT SHIPPER.shi_code, SHIPPER.shi_name"
        .CommandText = .CommandText & " From SHIPPER"
        .CommandText = .CommandText & " WHERE SHI_NPECODE = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and shi_actvflag = 1"
        .CommandText = .CommandText & " order by SHIPPER.shi_code "
         Set rst = .Execute
    End With
    
         
    STR = Chr$(1)
    ssdcboShipper.FieldSeparator = STR
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        ssdcboShipper.AddItem rst!shi_name & STR & (rst!shi_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::GetShipperName", Err.Description, Err.number, True)
End Sub
'SQL statement get priority list for priority combo
Private Sub GetPriorityList()
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT pri_code, pri_desc"
        .CommandText = .CommandText & " From PRIORITY "
        .CommandText = .CommandText & " WHERE pri_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by pri_code"
         Set rst = .Execute
    End With
    
    STR = Chr$(1)
    SSdcboPriority.FieldSeparator = STR
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
       
    Do While ((Not rst.EOF))
        SSdcboPriority.AddItem rst!pri_desc & STR & (rst!pri_code & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetPriorityList", Err.Description, Err.number, True)
End Sub
'SQL statement get shiptoname list for ship to name combo
Private Sub GetShiptoName()
On Error Resume Next
Dim STR As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = "SELECT sht_code, sht_name"
        .CommandText = .CommandText & " From SHIPTO "
        .CommandText = .CommandText & " WHERE sht_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and sht_actvflag = 1"
        .CommandText = .CommandText & " ORDER BY sht_code "
         Set rst = .Execute
    End With
    

    STR = Chr$(1)
    SSdcboShipto.FieldSeparator = STR
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        SSdcboShipto.AddItem rst!sht_name & STR & (rst!sht_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetShiptoName", Err.Description, Err.number, True)
End Sub
'SQL statement get soldtoname list for soldtoname combo
Private Sub GetSoldToName()
On Error Resume Next
Dim STR As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT slt_code, slt_name "
        .CommandText = .CommandText & " From SOLDTO "
        .CommandText = .CommandText & " WHERE slt_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by slt_code "
         Set rst = .Execute
    End With
    

    STR = Chr$(1)
    SSdcboSoldTo.FieldSeparator = STR
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        SSdcboSoldTo.AddItem rst!slt_name & STR & (rst!slt_code & "")
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

If Err Then Call LogErr(Name & "::GetSoldToName", Err.Description, Err.number, True)
End Sub
'SQL statement get destinationname list for destination combo
Private Sub GetDestinationName()
On Error Resume Next
Dim STR As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    With cmd
        .CommandText = " SELECT des_destcode, des_destname "
        .CommandText = .CommandText & " From Destination "
        .CommandText = .CommandText & " WHERE des_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " Order by des_destname "
         Set rst = .Execute
    End With
    

    STR = Chr$(1)
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    SSdcboFrom1.FieldSeparator = STR
    SSdcboDestinationTo.FieldSeparator = STR
    SSdcboFrom2.FieldSeparator = STR
    SSdcboDestinationTo1.FieldSeparator = STR
    rst.MoveFirst
    
'    SSdcboFrom2.AddItem " ", 0
'    SSdcboDestinationTo1.AddItem "", 0
    Do While ((Not rst.EOF))
        SSdcboFrom1.AddItem rst!des_destname & STR & (rst!des_destcode & "")
        SSdcboDestinationTo.AddItem rst!des_destname & STR & (rst!des_destcode & "")
        SSdcboFrom2.AddItem rst!des_destname & STR & (rst!des_destcode & "")
        SSdcboDestinationTo1.AddItem rst!des_destname & STR & (rst!des_destcode & "")
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

If Err Then Call LogErr(Name & "::GetDestinationName", Err.Description, Err.number, True)
End Sub
'SQL statement get termof delivery list for shipper term combo
Private Sub GetTermOfDelivery()
'On Error Resume Next
'Dim str As String
'    Dim cmd As ADODB.Command
'    Dim rst As ADODB.Recordset
'
'    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
'    With cmd
'        .CommandText = " SELECT tod_termcode, tod_desc "
'        .CommandText = .CommandText & " From TERMOFDELIVERY "
'        .CommandText = .CommandText & " WHERE tod_npecode = '" & deIms.NameSpace & "'"
'
'         Set rst = .Execute
'    End With
'
'    str = Chr$(1)
'    SSdcboTermDelivery.FieldSeparator = str
'    If rst.RecordCount = 0 Then GoTo CleanUp
'
'    rst.MoveFirst
'
'
'    Do While ((Not rst.EOF))
'        SSdcboTermDelivery.AddItem rst!tod_desc & str & (rst!tod_termcode & "")
'
'        rst.MoveNext
'    Loop
'
'
'CleanUp:
'    rst.Close
'    Set cmd = Nothing
'    Set rst = Nothing
'
'    If Err Then Call LogErr(Name & "::GetTermOfDelivery", err.Description ,err.Number , True)

End Sub
'get parameter to exec store procedure
Private Sub SaveAll()
On Error Resume Next
Dim cmd As ADODB.Command
    
    
    Set cmd = New ADODB.Command
    
    cmd.CommandText = "UPDATE_POITEM_PACKING"
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = deIms.cnIms

    cmd.Parameters.Append _
        cmd.CreateParameter("Return_Value", adInteger, adParamReturnValue)
        
    cmd.Parameters.Append _
        cmd.CreateParameter("@Namespace", adVarChar, adParamInput, 5, deIms.NameSpace)

    cmd.Parameters.Append _
        cmd.CreateParameter("@PONUMB", adVarChar, adParamInput, 15, ssdcboPoNumb)


    cmd.Parameters.Append _
        cmd.CreateParameter("@LINENUMB", adVarChar, adParamInput, 6, CoBLineitem)


    cmd.Parameters.Append _
        cmd.CreateParameter("@MANFNUMB", adVarChar, adParamInput, 15, cboPackingNumber)
    
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
 If Err Then Call LogErr(Name & "::SaveAll", Err.Description, Err.number, True)
End Sub
'SQl statement get manifest number for manifest combo
Private Function GetPackingnumber(ManifestNumb As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From PACKINGLIST "
        .CommandText = .CommandText & " Where pl_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND pl_manfnumb = '" & ManifestNumb & "'"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetPackingnumber = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::GetPackingnumber", Err.Description, Err.number, True)
End Function


Private Sub ssdcboPoNumb_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not ssdcboPoNumb.DroppedDown Then ssdcboPoNumb.DroppedDown = True
End Sub

Private Sub ssdcboPoNumb_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub

Private Sub ssdcboPoNumb_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With ssdcboPoNumb
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub SSdcboPriority_Click()
    SSdcboPriority.SelStart = 0
    SSdcboPriority.SelLength = 0
End Sub

Private Sub SSdcboPriority_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboPriority.DroppedDown Then SSdcboPriority.DroppedDown = True
End Sub


Private Sub SSdcboPriority_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub SSdcboPriority_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboPriority
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub ssdcboShipper_Click()
'     Call cboPackingNumber_Validate(True)
'     cboPackingNumber.SetFocus
'     Call EnableControls(False)

    ssdcboShipper.SelStart = 0
    ssdcboShipper.SelLength = 0
End Sub

Private Sub ssdcboShipper_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not ssdcboShipper.DroppedDown Then ssdcboShipper.DroppedDown = True
End Sub


Private Sub ssdcboShipper_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub ssdcboShipper_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With ssdcboShipper
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub

Private Sub SSdcboShipto_Click()
    SSdcboShipto.SelStart = 0
    SSdcboShipto.SelLength = 0
    SSdcboShipto.Refresh
End Sub

Private Sub SSdcboShipto_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboShipto.DroppedDown Then SSdcboShipto.DroppedDown = True
End Sub


Private Sub SSdcboShipto_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub SSdcboShipto_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboShipto
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


Private Sub SSdcboSoldTo_Click()
    SSdcboSoldTo.SelStart = 0
    SSdcboSoldTo.SelLength = 0
End Sub

Private Sub SSdcboSoldTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSdcboSoldTo.DroppedDown Then SSdcboSoldTo.DroppedDown = True
End Sub


Private Sub SSdcboSoldTo_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub SSdcboSoldTo_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    Dim text, i
    With SSdcboSoldTo
        text = .text
        If text <> "" Then
            If text = .Columns(0).text Then Exit Sub
            .MoveFirst
            For i = 0 To .Rows - 1
                If text Like .Columns(0).text Then
                    Exit For
                End If
                .MoveNext
            Next
            .text = ""
        End If
    End With
    '------------------------
End Sub


'depend on tab position to disable or enable navigetor butten
Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Dim iEditMode As String, blFlag As Boolean
Dim i As Integer

    'kin validate the data for the Previous Tab
    
    blFlag = SSTab1.Tab = 0 And Form = mdCreation
    cboPackingNumber.Enabled = True
    With Navbar1
        .NextEnabled = blFlag
        .LastEnabled = blFlag
        .FirstEnabled = blFlag
        .CancelEnabled = blFlag
        .PreviousEnabled = blFlag
        .SaveEnabled = SSTab1.Tab = 0
        .CloseEnabled = SSTab1.Tab = 0
        .NewEnabled = SSTab1.Tab = 0
        .PrintEnabled = cboPackingNumber.ListIndex <> CB_ERR
'        .EMailEnabled = SSTab1.Tab = 0
'        .EMailEnabled = (.PrintEnabled)
'        .PrintEnabled = .SaveEnabled And cboPackingNumber.ListIndex <> CB_ERR
        .EMailEnabled = ((dgRecepientList.Row) And (.PrintEnabled))
    End With
    
    If SSTab1.Tab = 0 And Form = mdVisualization Then


            Navbar1.FirstEnabled = False
            Navbar1.LastEnabled = False
            Navbar1.NextEnabled = False
            Navbar1.PreviousEnabled = False
            Navbar1.SaveEnabled = False
            If Len(Trim(cboPackingNumber)) <> 0 And Form = mdVisualization Then
            
                Navbar1.PrintEnabled = True
                Navbar1.EMailEnabled = True
            End If
    Else
           
            Navbar1.FirstEnabled = False
            Navbar1.LastEnabled = False
            Navbar1.NextEnabled = False
            Navbar1.PreviousEnabled = False
       
        If Len(Trim(cboPackingNumber)) <> 0 And Form = mdCreation Then
            Navbar1.SaveEnabled = True
'            NavBar1.FirstEnabled = False
'            NavBar1.LastEnabled = False
'            NavBar1.NextEnabled = False
'            NavBar1.PreviousEnabled = False
          
        End If
    End If

        If SSTab1.Tab = 1 Then
                Navbar1.NewEnabled = False
                Navbar1.SaveEnabled = False
                Navbar1.PrintEnabled = False
                Navbar1.EMailEnabled = False
             If SSTab1.Tab = 1 And Form = mdVisualization Then
                 dgRecepientList = ""
                 Call GetRecipientList
                 dgRecepientList.Enabled = True
                 dgRecepients.Enabled = True
                 fra_FaxSelect.Enabled = True
                 cmd_Add.Enabled = True
                 cmd_Remove.Enabled = True
                 opt_FaxNum.Enabled = True
                 opt_Email.Enabled = True
                
'                Call EnableControls(True)
    
              Else
                If Not CheckCombFields = True Then
                    SSTab1.Tab = PreviousTab
                ElseIf Not CheckLIFields = True Then
                    SSTab1.Tab = PreviousTab
                End If
             End If
        

    ElseIf SSTab1.Tab = 2 Then
'            ssdcboPoNumb = ""
             Navbar1.NewEnabled = False
             Navbar1.SaveEnabled = False
             Navbar1.PrintEnabled = False
             Navbar1.EMailEnabled = False
             Navbar1.CancelEnabled = True
        If Len(Trim$(cboPackingNumber)) <> 0 Then
            LlbManifest = cboPackingNumber
            LlbShipTo = SSdcboShipto
            If Form = mdVisualization Then
                Call Getpartingdeltlist
            End If
            
            Navbar1.FirstEnabled = True
            Navbar1.LastEnabled = True
            Navbar1.NextEnabled = True
            Navbar1.PreviousEnabled = True
            Navbar1.NewEnabled = False

        End If
        
    ElseIf SSTab1.Tab = 3 Then
          Navbar1.NewEnabled = False
          Navbar1.SaveEnabled = False
          Navbar1.PrintEnabled = False
          Navbar1.EMailEnabled = False
          txtRemarks = ""
        If Len(Trim$(cboPackingNumber)) <> 0 Then
'            Call GetParkingRemarklist
        End If
             Navbar1.NewEnabled = False

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
If Err Then Call LogErr(Name & "::SSTab1_Click", Err.Description, Err.number, True)
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Added by Juan 11/17/2000
    If KeyCode = 9 Then
        Select Case SSTab1.Tab
            Case 0
                cboPackingNumber.SetFocus
            Case 1
                cmd_Add.SetFocus
            Case 2
                If ssdcboPoNumb.Enabled Then ssdcboPoNumb.SetFocus
            Case 3
                If txtRemarks.Enabled Then txtRemarks.SetFocus
        End Select
    End If
    '------------------------
End Sub

'press return key add new record to recipient list
Private Sub txt_Recipient_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim$(txt_Recipient)) Then cmd_Add_Click
    End If
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub

Private Sub Txtawbnumb_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtBeShipped_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub

'check tobe shipped text for over shipped
Private Sub TxtBeShipped_Validate(Cancel As Boolean)
On Error Resume Next
Dim Result As Boolean
        Result = Checkovershiped
        If Result = False Then
            TxtBeShipped = ""
            TxtBeShipped.SetFocus
            Exit Sub
        End If
         
        pl.Tobeship = CDbl(TxtBeShipped)
        TxtBeShipped = FormatNumber((TxtBeShipped), 4)
    If Err Then Call LogErr(Name & "::TxtBeShipped_Validate", Err.Description, Err.number, True)
End Sub

Private Sub TxtBoxNumber_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub

'validate text box number
Private Sub TxtBoxNumber_Validate(Cancel As Boolean)
On Error Resume Next
    Cancel = True
    
    TxtBoxNumber = Trim$(TxtBoxNumber)
    
    If Len(TxtBoxNumber) Then
        If Not IsNumeric(TxtBoxNumber) Then
   
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00280") 'J added
            MsgBox IIf(msg1 = "", "Box Number must be numeric", msg1) 'J modified
            '---------------------------------------------
            
            TxtBoxNumber.SetFocus: Exit Sub
        ElseIf Len(TxtBoxNumber) > 0 Then
            pl.BoxNumber = CDbl(TxtBoxNumber)
            
        End If
    
    Else
    
        'Modified by Juan (9/13/2000) for Mutilingual
        msg1 = translator.Trans("M00281") 'J added
        MsgBox IIf(msg1 = "", "Box Number cannot be left empty", msg1) 'J modified
        '--------------------------------------------
        
        TxtBoxNumber.SetFocus: Exit Sub

    End If
    Cancel = False
If Err Then Call LogErr(Name & "::TxtBoxNumber_Validate", Err.Description, Err.number, True)
End Sub
'reset data grid recepientlist datasource.
Private Sub cboPackingNumber_Change()
On Error Resume Next

    
    If dgRecepientList.Tag <> "" Then
        Set dgRecepientList.DataSource = Nothing
    
        dgRecepientList.Tag = ""
        Set rsReceptList = Nothing
    End If
    
    If IsNothing(pld) Then Set pld = New PackingListDetls

        cboPackingNumber.Tag = ""
    
  
'    If Len(Trim$(cboPackingNumber)) <> 0 And ChangeMode(mdCreation) = True Then
'         NavBar1.PrintEnabled = False
'         NavBar1.EMailEnabled = False
'         Call EnableControls(True)
'         Call Clearform
'         Call AssignDefault
'    Else
'         NavBar1.PrintEnabled = True
'         NavBar1.EMailEnabled = True
'    End If
'        Call EnableControls(True)
'        Call Clearform
'        Call AssignDefault
        
        
'    If Len(Trim$(cboPackingNumber)) <> 0 Then
'
''    Call cboPackingNumber_Validate(True)
''        If GetPackingnumber(cboPackingNumber) Then
''            MsgBox "Packing List Entered Number is already exist"
''            cboPackingNumber.SetFocus: Exit Sub
''        End If
'    Else
'        MsgBox "Please enter new number"
'
'    End If

    Exit Sub


If Err Then Call LogErr(Name & "::cboPackingNumber_Change", Err.Description, Err.number, True)
 
 End Sub
'
Private Sub cboPackingNumber_KeyPress(KeyAscii As Integer)
    Dim i, text 'J added
    
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        Exit Sub
    End If
    '------------------------
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    'Added by Juan for Alpha Search (11/14/2000)
    With cboPackingNumber
        text = .text
        For i = 0 To .ListCount - 1
            If text Like .list(i) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    '-------------------------------------------
    
End Sub
'check enter manifest number exist on database or not
Private Sub cboPackingNumber_Validate(Cancel As Boolean)
On Error Resume Next

    Cancel = False

    If Len(Trim$(cboPackingNumber)) <> 0 And Form = mdCreation Then


        If GetPackingnumber(cboPackingNumber) Then
            cboPackingNumber.ListIndex = CB_ERR
            GetRecipientList
            
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00282") 'J added
            MsgBox IIf(msg1 = "", "Packing List Entered Number is already exist", msg1): 'J modified
            '---------------------------------------------
            
            Call EnableControls(False)
            cboPackingNumber.SetFocus
            Exit Sub

           
        Else
           Call EnableControls(True)
           'Call Clearform
           Call AssignDefault
         
        End If

    End If
    
    If Err Then Call LogErr(Name & "::cboPackingNumber_Validate", Err.Description, Err.number, True)
End Sub

Private Sub Txtcustrefe_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txtflig2_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txtflight1_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txtforwrefe_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txtgrosweig_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txthawbnum_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtMark1_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtMark2_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtMark3_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtMark4_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtMark4_Validate(Cancel As Boolean)
    cboPackingNumber.SetFocus
End Sub


'validate number of pieces text box
Private Sub Txtnumbpiec_Change()
On Error Resume Next


    If Len(Txtnumbpiec) Then
    
      If Len(Trim$(Txtnumbpiec)) = 0 Then
      
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00283") 'J added
            MsgBox IIf(msg1 = "", "Number of Pieces cannot be left empty", msg1) 'J modified
            '---------------------------------------------
            
            Txtnumbpiec.SetFocus: Exit Sub
      
      ElseIf Not IsNumeric(Txtnumbpiec) Then
      
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00284") 'J added
            MsgBox IIf(msg1 = "", "Number of Pieces must be numeric", msg1) 'J modified
            '---------------------------------------------
            
            Txtnumbpiec.SetFocus: Exit Sub
      End If
      
    End If
    
    If Err Then Call LogErr(Name & "::Txtnumbpiec_Change", Err.Description, Err.number, True)
End Sub

Private Sub Txtnumbpiec_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub TxtRemark_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub txtShippingterms_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


Private Sub Txtshprefe_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


'validate total volume text box
Private Sub Txttotavolu_Change()
    
    If Len(Txttotavolu) Then
    
        If Not IsNumeric(Txttotavolu) Then
        
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00285") 'J added
            MsgBox IIf(msg1 = "", "Total Volume must be numeric", msg1) 'J modified
            '---------------------------------------------
            
            If Txttotavolu.Enabled Then Txttotavolu.SetFocus
        End If
        
    End If
    
End Sub
'validate all text box fields
Private Function CheckLIFields() As Boolean
On Error Resume Next

    CheckLIFields = False
    
    If Len(Trim$(cboPackingNumber)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00286") 'J added
        MsgBox IIf(msg1 = "", "The Manifest Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        cboPackingNumber.SetFocus: Exit Function
    End If
    
    If Len(Trim$(DTPDocudate)) = 0 Then
         DTPDocudate.SetFocus: Exit Function

    ElseIf Not IsDate(DTPDocudate) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00287") 'J added
        MsgBox IIf(msg1 = "", "The document date must be date type", msg1) 'J modified
        '---------------------------------------------
        
        DTPDocudate.SetFocus: Exit Function
    End If


    If Len(Trim$(DTPshidate)) = 0 Then
         DTPshidate.SetFocus: Exit Function
    ElseIf Not IsDate(DTPshidate) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00288") 'J added
        MsgBox IIf(msg1 = "", "The ship date must be date type", msg1) 'J modified
        '---------------------------------------------
        
        DTPshidate.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtawbnumb)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00289") 'J added
        MsgBox IIf(msg1 = "", "AIR WAY BILL canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        Txtawbnumb.SetFocus: Exit Function
    End If
    
        
'    If Len(Trim$(Txtflight1)) = 0 Then
'        MsgBox "Flight or Voyage canot be left empty"
'        Txtflight1.SetFocus: Exit Function
'    End If
    
    If Len(Trim$(DTPicker1etd)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00290") 'J added
        MsgBox IIf(msg1 = "", "Edit Date cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        DTPicker1etd.SetFocus: Exit Function
    ElseIf Not IsDate(DTPicker1etd) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00291") 'J added
        MsgBox IIf(msg1 = "", "The Edit date must be date type", msg1) 'J modified
        '---------------------------------------------
        
    End If
    
    If Len(Trim$(DTPicker2eta)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00292") 'J added
        MsgBox IIf(msg1 = "", "Edit Date canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        DTPicker2eta.SetFocus: Exit Function
    ElseIf Not IsDate(DTPicker2eta) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00287") 'J added
        MsgBox IIf(msg1 = "", "The Edit Date must be date type", msg1) 'J modified
        '---------------------------------------------
        
    End If
    
    If Len(Trim$(Txtviacarr)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00292") 'J added
        MsgBox IIf(msg1 = "", "Via Carrier canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        Txtviacarr.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtgrosweig)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00293") 'J added
        MsgBox IIf(msg1 = "", "Gross Weight cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        Txtgrosweig.SetFocus: Exit Function
     
    ElseIf Not IsNumeric(Txtgrosweig) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00294") 'J added
        MsgBox IIf(msg1 = "", "Gross Weight must be numeric", msg1) 'J modified
        '---------------------------------------------
        
           Txtgrosweig.SetFocus: Exit Function
    End If
    
    If Len(Trim$(Txtnumbpiec)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00283") 'J added
        MsgBox IIf(msg1 = "", "Number of Pieces cannot be left empty", msg1) 'J modified
        '---------------------------------------------
           
           Txtnumbpiec.SetFocus: Exit Function
      
    ElseIf Not IsNumeric(Txtnumbpiec) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00284") 'J added
        MsgBox IIf(msg1 = "", "Number of Pieces must be numeric", msg1) 'J modified
        '---------------------------------------------
        
        Txtnumbpiec.SetFocus: Exit Function
    End If
    
   
    CheckLIFields = True
    
    If Not IsNumeric(Txttotavolu) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00295") 'J added
        MsgBox IIf(msg1 = "", "Total Volume must be numeric and entry optional", msg1) 'J modified
        '---------------------------------------------
        
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
    
        CheckLIFields = True
    If Err Then Call LogErr(Name & "::CheckLIFields", Err.Description, Err.number, True)
End Function
'validate all combo fields
Private Function CheckLineitemFlied()
On Error Resume Next
    CheckLineitemFlied = False
    
    If Len(ssdcboPoNumb.text) = 0 Then
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00296") 'J added
        MsgBox IIf(msg1 = "", "PO Number canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboPoNumb.SetFocus: Exit Function
    End If
            
    If Len(CoBLineitem.text) = 0 Then
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00297") 'J added
        MsgBox IIf(msg1 = "", "Line Item Number canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        CoBLineitem.SetFocus: Exit Function
    End If
            
      
        
    If Len(TxtBeShipped) > 0 Then
        pl.RequestedQty = CDbl(TxtBeShipped)
        
        'Modified by Juan (9/13/2000) for Multilangual
        msg1 = translator.Trans("M00298") 'J added
        MsgBox IIf(msg1 = "", "Quantity Been shipped cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
    ElseIf Not IsNumeric(TxtBeShipped) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00299") 'J added
        MsgBox IIf(msg1 = "", "Quantity been shipped must be numeric", msg1) 'J modified
        '---------------------------------------------
        
        TxtBeShipped.SetFocus: Exit Function
    End If
    
        
   If Len(TxtBoxNumber) > 0 Then
       pl.BoxNumber = CDbl(TxtBoxNumber)
       
       'Modified by Juan (9/13/2000) for Multilingual
       msg1 = translator.Trans("M00281") 'J added
       MsgBox IIf(msg1 = "", "Box Number cannot be left empty", msg1) 'J modified
       '---------------------------------------------
       
   ElseIf Not IsNumeric(TxtBoxNumber) Then
   
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00280") 'J added
        MsgBox IIf(msg1 = "", "Box Number must be numeric", msg1) 'J modified
        '---------------------------------------------
       
       TxtBoxNumber.SetFocus: Exit Function
   End If
    
        CheckLineitemFlied = True
   If Err Then Call LogErr(Name & "::CheckLineitemFlied", Err.Description, Err.number, True)
End Function

'validate all combo fields
Private Function CheckCombFields() As Boolean
On Error Resume Next
    CheckCombFields = False
    
    If Len(ssdcboShipper.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00239") 'J added
        MsgBox IIf(msg1 = "", "Shipper Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        If ssdcboShipper.Enabled Then ssdcboShipper.SetFocus:
        Exit Function
    End If
    

    If Len(SSdcboPriority.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00300") 'J added
        MsgBox IIf(msg1 = "", "Priority Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        SSdcboPriority.SetFocus: Exit Function
    End If
   
    If Len(SSdcboShipto.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00301") 'J added
        MsgBox IIf(msg1 = "", "Ship To Name cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        SSdcboShipto.SetFocus: Exit Function
    End If
    
    
    If Len(SSdcboSoldTo.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00302") 'J added
        MsgBox IIf(msg1 = "", "Sold To Name canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        SSdcboSoldTo.SetFocus: Exit Function
    End If
    
    
    If Len(txtShippingterms) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00303") 'J added
        MsgBox IIf(msg1 = "", "Shipping Term canot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtShippingterms.SetFocus: Exit Function
    End If
    
    If Len(SSdcboFrom1.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00304") 'J added
        MsgBox IIf(msg1 = "", "Destination cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        SSdcboFrom1.SetFocus: Exit Function
        
    End If
    
    If Len(SSdcboDestinationTo.text) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00304") 'J added
        MsgBox IIf(msg1 = "", "Destination cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        SSdcboDestinationTo.SetFocus: Exit Function
        
    End If
   
     CheckCombFields = True
    
    If Len(SSdcboFrom2.text) = 0 Then
        
        SSdcboFrom2.SetFocus: Exit Function
        
    End If
    
    If Len(SSdcboDestinationTo1.text) = 0 Then
        
        SSdcboDestinationTo1.SetFocus: Exit Function
        
    End If
    
    CheckCombFields = True
    If Err Then Call LogErr(Name & "::CheckCombFields", Err.Description, Err.number, True)
End Function
'add a recepient to recepient list
Private Sub cmd_Add_Click()
On Error Resume Next

    
    If Len(Trim$(txt_Recipient)) Then
        Call AddRecepients(txt_Recipient)
        txt_Recipient = ""
    Else
        dgRecepients_DblClick
    End If
If Err Then Call LogErr(Name & "::cmd_Add_Click", Err.Description, Err.number, True)
End Sub
'remove a recepient from recepient list
Private Sub cmd_Remove_Click()
On Error Resume Next
    rec.Remove (rsReceptList.Fields(0).Value)
    Call rsReceptList.Delete(adAffectCurrent)
    'rsReceptList.Delete
   If Err Then Call LogErr(Name & "::cmd_Remove_Click", Err.Description, Err.number, True)
End Sub
'function add a new fax number to recepient list
Private Sub AddRecepients(Recepient As String)
On Error Resume Next
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
        Call rec.Add(deIms.NameSpace, cboPackingNumber, Recepient, CurrentUser, Recepient)
    End If
    
    If Err Then Call LogErr(Name & "::AddRecepients", Err.Description, Err.number, True)
End Sub
'check email address and fax number exist on recepient list
Private Function IsRecipientInList(RecepientName As String) As Boolean
On Error Resume Next
Dim BK As Variant
    
    
    If rsReceptList.RecordCount = 0 Then Exit Function
    If Not (rsReceptList.EOF Or rsReceptList.BOF) Then BK = rsReceptList.Bookmark
    
    rsReceptList.MoveFirst
    Call rsReceptList.Find("Recipient = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
    
    If Not (rsReceptList.EOF) Then
        
        If opt_Email Then
            
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00076") 'J added
            MsgBox IIf(msg1 = "", "Email Address Already in list", msg1) 'J modified
            '---------------------------------------------
            
        ElseIf opt_FaxNum Then
        
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("M00077") 'J added
            MsgBox IIf(msg1 = "", "Fax Number Already in list", msg1) 'J modified
            '---------------------------------------------
            
        End If
        IsRecipientInList = True
    End If
    
    rsReceptList.Bookmark = BK
      
    If Err Then Call LogErr(Name & "::AddRecepients", Err.Description, Err.number, True)
      
End Function
'add a recepient to recepient list
Private Sub dgRecepients_DblClick()
On Error Resume Next
    If dgRecepients.ApproxCount > 0 Then _
        Call AddRecepients(dgRecepients.Columns(1).text)
If Err Then Call LogErr(Name & "::dgRecepients_DblClick", Err.Description, Err.number, True)
End Sub
'send email function
Private Sub NavBar1_OnEMailClick()
Dim Params(1) As String
Dim rptinfo As RPTIFileInfo

On Error Resume Next
    BeforePrint
    
    With rptinfo
        Params(0) = "namespace=" & deIms.NameSpace
        Params(1) = "manifestnumb=" & cboPackingNumber
        .ReportFileName = ReportPath & "packinglist.rpt"
'        .ReportFileName = ReportPath & "newpackinglist2.rpt"

        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("packinglist.rpt") 'J added
        '---------------------------------------------
        
        .Parameters = Params
    End With
    
    Params(0) = ""
    Call WriteRPTIFile(rptinfo, Params(0))
    Call SendEmailAndFax(rsReceptList, "Recipient", _
                         "Packing List / Manifest Management " & cboPackingNumber, "", Params(0))
    
If Err Then Call LogErr(Name & "::NavBar1_OnEMailClick", Err.Description, Err.number, True)
End Sub
'print crystal report
Private Sub BeforePrint()
On Error Resume Next
     With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = ReportPath & "packinglist.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("packinglist.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "manifestnumb;" + cboPackingNumber + ";true"
    End With
 If Err Then Call LogErr(Name & "::BeforePrint", Err.Description, Err.number, True)
End Sub
'set email background color
Private Sub opt_Email_GotFocus()
On Error Resume Next
    Call HighlightBackground(opt_Email)
If Err Then Call LogErr(Name & "::opt_Email_GotFocus", Err.Description, Err.number, True)
End Sub
'reset email background color
Private Sub opt_Email_LostFocus()
On Error Resume Next
    Call NormalBackground(opt_Email)
If Err Then Call LogErr(Name & "::opt_Email_LostFocus", Err.Description, Err.number, True)
End Sub
'get email address from data grid recepient list.
Private Sub opt_Email_Click()
On Error Resume Next
Dim co As MSDataGridLib.Column

    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("L00121") 'J added
    co.Caption = IIf(msg1 = "", "Email Address", msg1) 'J modified
    '---------------------------------------------
    
    co.DataField = "phd_mail"
    
    dgRecepients.Columns(0).DataField = "phd_name"
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atEmail)
    If Err Then Call LogErr(Name & "::opt_Email_Click", Err.Description, Err.number, True)
End Sub
'get fax number from data grid recepient list.
Private Sub opt_FaxNum_Click()
On Error Resume Next
Dim co As MSDataGridLib.Column
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("L00122") 'J added
    co.Caption = IIf(msg1 = "", "Fax Number", msg1) 'J added
    '---------------------------------------------
    
    co.DataField = "phd_faxnumb"
    
    dgRecepients.Columns(0).DataField = "phd_name"
     
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atFax)

     If Err Then Call LogErr(Name & "::opt_FaxNum_Click", Err.Description, Err.number, True)
End Sub
'set fax number back ground color
Private Sub opt_FaxNum_GotFocus()
    Call HighlightBackground(opt_FaxNum)
End Sub
'reset fax number back ground color
Private Sub opt_FaxNum_LostFocus()
    Call NormalBackground(opt_FaxNum)
End Sub
'enable form control text box, lable, button
Private Sub EnableControls(bEnable As Boolean)
On Error Resume Next
Dim ctl As Control

    For Each ctl In Controls
        If (Not ((TypeOf ctl Is Label) Or (TypeOf ctl Is SSTab) Or (TypeOf ctl Is NavBar))) Then ctl.Enabled = bEnable
        If Err Then Err.Clear
    Next ctl
    
    cboPackingNumber.Enabled = True
    Frame1.Enabled = True
    If Err Then Call LogErr(Name & "::EnableControls", Err.Description, Err.number, True)
End Sub
'set date data combo
Private Sub AssignDefault()
On Error Resume Next
Dim STR As String
    
'    str = Format$(Now(), "mm/dd/yyyy")
    
    DTPDocudate.Value = Date
    DTPshidate.Value = Date
    Txttotavolu = 0

    DTPicker1etd.Value = Date
    DTPicker2eta.Value = Date
    
    cboPackingNumber.Tag = cboPackingNumber.text
    
    If IsNothing(pld) Then Set pld = New PackingListDetls
    If IsNothing(rec) Then Set rec = New imsPackingListRecp
    
    If Err Then Call LogErr(Name & "::AssignDefault", Err.Description, Err.number, True)
    
End Sub
'set data combo enable
Private Sub EnableControlsLine(bEnable As Boolean)
On Error Resume Next

    ssdcboPoNumb.Enabled = bEnable
    CoBLineitem.Enabled = bEnable
If Err Then Call LogErr(Name & "::EnableControlsLine", Err.Description, Err.number, True)
End Sub
'SQL statement get manifest number list for manifest combo
Private Sub GetManifestNumberList()
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
    
    .CommandText = " SELECT pl_manfnumb "
    .CommandText = .CommandText & " From PACKINGLIST"
    .CommandText = .CommandText & " WHERE pl_npecode = '" & deIms.NameSpace & "'"
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
    
    If Err Then Call LogErr(Name & "::GetManifestNumberList", Err.Description, Err.number, True)
End Sub
'SQl statement get recipient list for recepient data grid
Private Sub GetRecipientList()
On Error Resume Next
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = deIms.cnIms
    
    With cmd
        .CommandText = "SELECT plrc_rec Recipient"
        .CommandText = .CommandText & " FROM PACKINGREC"
        .CommandText = .CommandText & " WHERE plrc_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND plrc_manfnumb = '" & cboPackingNumber & "'"
        
        Set rsReceptList = .Execute
        
        dgRecepientList.Tag = "1"
        Set dgRecepientList.DataSource = rsReceptList
        If Err Then MsgBox Err.Description: Err.Clear
    End With
    
    Set cmd = Nothing
    
    If rsReceptList.BOF And rsReceptList.EOF Then Set rsReceptList = Nothing
If Err Then Call LogErr(Name & "::GetRecipientList", Err.Description, Err.number, True)
End Sub
'SQL statement get packing list record data information
Public Sub GetPackingAlloflist(Manunumber As String)
On Error Resume Next
Dim STR As String
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
        .CommandText = .CommandText & " pl_npecode = '" & deIms.NameSpace & "' "
        
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
    Txtflight1 = rst!pl_fig1
    Txthawbnum = rst!pl_hawbnumb
    Txtflig2 = rst!pl_fig2
    TxtRemark = rst!pl_remk
    DTPicker1etd = rst!pl_etd
    DTPicker2eta = rst!pl_eta
    SSdcboFrom1 = rst!pl_from1
    SSdcboDestinationTo = rst!pl_to1
    SSdcboFrom2 = rst!pl_from2
    SSdcboDestinationTo1 = rst!pl_to2
    txtShippingterms = rst!pl_shipterm
    Txtviacarr = rst!pl_viacarr
    LblDestination = rst!pl_dest
    Txtnumbpiec = rst!pl_numbpiec
    Txtgrosweig = rst!pl_grosweig
    Txttotavolu = rst!pl_totavolu
    TxtMark1 = rst!pl_mark1
    TxtMark2 = rst!pl_mark2
    TxtMark3 = rst!pl_mark3
    TxtMark4 = rst!pl_mark4
 
 If Err Then Call LogErr(Name & "::GetPackingAlloflist", Err.Description, Err.number, True)
        
End Sub
'clear packing list form
Public Sub Clearform()
On Error Resume Next
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
    Txtflight1 = ""
    Txthawbnum = ""
    Txtflig2 = ""
    TxtRemark = ""
'    DTPicker1etd = ""
'    DTPicker2eta = ""
    SSdcboFrom1 = ""
    SSdcboDestinationTo = ""
    SSdcboFrom2 = ""
    SSdcboDestinationTo1 = ""
    txtShippingterms = ""
    Txtviacarr = ""
    LblDestination = ""
    Txtnumbpiec = ""
    Txtgrosweig = ""
    Txttotavolu = ""
    TxtMark1 = ""
    TxtMark2 = ""
    TxtMark3 = ""
    TxtMark4 = ""
 If Err Then Call LogErr(Name & "::Clearform", Err.Description, Err.number, True)
           

End Sub
'clear packing list line item form
Public Sub ClearPOitemForm()
On Error Resume Next
    TxtBeShipped = ""
    lblReqQty = ""
    lblUnitPrice = ""
    lblAmount = ""
    TxtBoxNumber = ""
    lblQtyDelv = ""
    lblQtyInv = ""
    LblTobeInven = ""
    TxtDescription = ""
    CoBLineitem = ""
'    ssdcboPoNumb = ""
If Err Then Call LogErr(Name & "::ClearPOitemForm", Err.Description, Err.number, True)
           
End Sub
'SQL statement get packing list details record
Public Function Getpartingdeltlist() As Recordset
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
'Dim rst As ADODB.Recordset


    
    Set cmd = New ADODB.Command
        
    With cmd
        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT  PACKINGDETL.pld_manfnumb,PACKINGDETL.pld_manfsrl, "
        .CommandText = .CommandText & " PACKINGDETL.pld_ponum,PACKINGDETL.pld_liitnumb,"
        .CommandText = .CommandText & " PACKINGDETL.pld_boxnumb, PACKINGDETL.pld_reqdqty,"
        .CommandText = .CommandText & " PACKINGDETL.pld_unitpric,PACKINGDETL.pld_totaprice,"
        .CommandText = .CommandText & " POITEM.poi_desc,POITEM.poi_primreqdqty,"
        .CommandText = .CommandText & " POITEM.poi_totaprice,POITEM.poi_qtydlvd,"
        .CommandText = .CommandText & " POITEM.poi_qtyship"
        .CommandText = .CommandText & " FROM PACKINGDETL INNER JOIN POITEM ON "
        .CommandText = .CommandText & " PACKINGDETL.pld_ponum = POITEM.poi_ponumb AND"
        .CommandText = .CommandText & " PACKINGDETL.pld_npecode = POITEM.poi_npecode AND"
        .CommandText = .CommandText & " PACKINGDETL.pld_liitnumb = POITEM.poi_liitnumb and"
        .CommandText = .CommandText & " PACKINGDETL.pld_manfnumb = '" & cboPackingNumber & "' AND"
        .CommandText = .CommandText & " PACKINGDETL.pld_npecode = '" & deIms.NameSpace & "' "
        
        
        Set Rstlist = .Execute
    End With
    
    STR = Chr$(1)
    
    If Rstlist Is Nothing Then Exit Function
    If Rstlist.RecordCount = 0 Then GoTo CleanUp
        

        If Rstlist.BOF And Rstlist.EOF Then
                cboPackingNumber = ""
                ssdcboPoNumb = ""
                CoBLineitem = ""
                TxtBoxNumber = ""
                TxtDescription = ""
                lblReqQty = ""
                lblUnitPrice = ""
                lblAmount = ""
                lblQtyDelv = ""
                lblQtyInv = ""
                LblTobeInven = ""
                TxtBeShipped = ""
            Exit Function
        Else
                cboPackingNumber = Rstlist!pld_manfnumb
                ssdcboPoNumb = Rstlist!pld_ponum
                CoBLineitem = Rstlist!pld_liitnumb
                TxtBoxNumber = Rstlist!pld_boxnumb
                TxtDescription = Rstlist!poi_desc
                lblReqQty = FormatNumber((Rstlist!poi_primreqdqty), 4)
                lblUnitPrice = FormatNumber((Rstlist!pld_unitpric), 4)
                lblAmount = FormatNumber((Rstlist!poi_totaprice), 4)
                lblQtyDelv = FormatNumber((Rstlist!poi_qtydlvd), 4)
                lblQtyInv = FormatNumber((Rstlist!poi_qtyship), 4)
                LblTobeInven = FormatNumber((Rstlist!pld_reqdqty), 4)
                TxtBeShipped = FormatNumber((Rstlist!pld_reqdqty), 4)
        End If
CleanUp:
    
    Set cmd = Nothing
'    Set rst = Nothing
If Err Then Call LogErr(Name & "::Getpartingdeltlist", Err.Description, Err.number, True)
End Function
'SQL statement get packing list remark list record
Public Sub GetParkingRemarklist()
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset


    
    Set cmd = New ADODB.Command
        
    With cmd
        .Prepared = True
        .CommandType = adCmdText
        .ActiveConnection = deIms.cnIms
        
        .CommandText = " SELECT  plr_remk"
        .CommandText = .CommandText & " From PACKINGREM"
        .CommandText = .CommandText & " WHERE (plr_manfnumb = '" & cboPackingNumber & "') AND"
        .CommandText = .CommandText & " (plr_npecode = '" & deIms.NameSpace & "')"
        Set rst = .Execute
    End With
    
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then GoTo CleanUp
    
        
        txtRemarks = rst!plr_remk

CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    
If Err Then Call LogErr(Name & "::GetParkingRemarklist", Err.Description, Err.number, True)
End Sub
'set values to form text box and lable
Private Sub LoadValues()
On Error Resume Next
                 
                cboPackingNumber = Rstlist!pld_manfnumb & ""
                ssdcboPoNumb = Rstlist!pld_ponum & ""
                CoBLineitem = Rstlist!pld_liitnumb & ""
                TxtBoxNumber = Rstlist!pld_boxnumb & ""
                TxtDescription = Rstlist!poi_desc & ""
                lblReqQty = FormatNumber((Rstlist!poi_primreqdqty & ""), 4)
                lblUnitPrice = FormatNumber((Rstlist!pld_unitpric & ""), 4)
                lblAmount = FormatNumber((Rstlist!poi_totaprice & ""), 4)
                lblQtyDelv = FormatNumber((Rstlist!poi_qtydlvd & ""), 4)
                lblQtyInv = FormatNumber((Rstlist!poi_qtyship & ""), 4)
                LblTobeInven = FormatNumber((Rstlist!pld_reqdqty & ""), 4)
                TxtBeShipped = FormatNumber((Rstlist!pld_reqdqty & ""), 4)
                
If Err Then Call LogErr(Name & "::LoadValues", Err.Description, Err.number, True)
End Sub
'set values to form text box and lable
Private Sub LoadValuesrst()
On Error Resume Next
        
        
'        Private FMainfestNumber As String
'        Private FNamespace As String
'        Private FManiFestSerialNumb As Integer
'        Private FPoNumber As String
'        CoBLineitem = pl.LineNumber
'        TxtBoxNumber = pl.BoxNumber
''        Private FRequestedQty As Double
'        lblUnitprice = pl.UnitPrice
'        lblAmount = pl.TotalPrice
      
        CoBLineitem = Rstitem!poi_liitnumb
        TxtDescription = Rstitem!poi_desc
        lblReqQty = CDbl(Rstitem!poi_primreqdqty)
        lblUnitPrice = CDbl(Rstitem!poi_unitprice)
        lblAmount = CDbl(Rstitem!poi_totaprice)
        lblQtyDelv = CDbl(Rstitem!poi_qtydlvd)
        lblQtyInv = CDbl(Rstitem!poi_qtyship)
        LblTobeInven = (Rstitem!poi_qtydlvd - Rstitem!poi_qtyship)
        TxtBeShipped = LblTobeInven
        TxtBoxNumber = pld.Count + 1
                 
'                cboPackingNumber = Rstitem!pld_manfnumb & ""
'                ssdcboPoNumb = Rstitem!pld_ponum & ""
'                CoBLineitem = Rstitem!pld_liitnumb & ""
'                TxtBoxNumber = Rstitem!pld_boxnumb & ""
'                TxtDescription = Rstitem!poi_desc & ""
'                lblReqQty = Rstitem!poi_primreqdqty & ""
'                lblUnitPrice = Rstitem!pld_unitpric & ""
'                lblAmount = Rstitem!poi_totaprice & ""
'                lblQtyDelv = Rstitem!poi_qtydlvd & ""
'                lblQtyInv = Rstitem!poi_qtyship & ""
'                LblTobeInven = Rstitem!pld_reqdqty & ""
'                TxtBeShipped = Rstitem!pld_reqdqty & ""
                
If Err Then Call LogErr(Name & "::LoadValuesrst", Err.Description, Err.number, True)
End Sub
'function check quatity being over shipped
Public Function Checkovershiped() As Boolean
On Error Resume Next
Dim Msg, Style, Title
Dim Num1 As Double
Dim Num2 As Double
Dim Num3 As Double

    Checkovershiped = False

'Modified by Juan (9/13/2000) for Multilingual
msg1 = translator.Trans("L00186") 'J added
msg2 = translator.Trans("M00305") 'J added
Msg = IIf(msg1 = "", " Lineitem# ", msg1 + " ") & CoBLineitem & IIf(msg2 = "", " is being over shipped, Do you want to continue ?", " " + msg2) 'J modified
'---------------------------------------------

Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Imswin"

'    lblQtyDelv = Rstitem!poi_qtydlvd

    Num1 = CDbl(lblQtyInv.Caption)
    Num2 = CDbl(TxtBeShipped.text)
    Num3 = Num1 + Num2
    If Not Len(Trim(TxtBeShipped)) = 0 Then
         If Num3 > CDbl(lblReqQty) Then
            If MsgBox(Msg, Style, Title) = vbNo Then
                Exit Function
            End If
        Else
        End If
    End If
    
    Checkovershiped = True
    
    If Err Then Call LogErr(Name & "::Checkovershiped", Err.Description, Err.number, True)
End Function
'check line item already exist on recordset
Private Function Lineitemcheck() As Boolean
On Error Resume Next
Dim pl As Object
Dim Msg, Style, Title
Dim Num1, Num2 As Integer
Dim Ponumb As String


'Modified by Juan (9/13/2000) for Multilingual
msg1 = translator.Trans("L00186") 'J added
msg2 = translator.Trans("M00307") 'J added
Msg = IIf(msg1 = "", " Lineitem# ", msg1 + " ") & CoBLineitem & IIf(msg2 = "", " is already exists, Do you want to add ?", " " + msg2) 'J modified
'---------------------------------------------

Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Imswin"
     
     Ponumb = ssdcboPoNumb
     Num1 = CoBLineitem
'     Num = pl.LineNumber
     Lineitemcheck = False
     
     For Each pl In pld
        If Not Len(Trim$(CoBLineitem)) = 0 Then
            If pl.LineNumber = Num1 And pl.PoNumber = Ponumb Then
                If MsgBox(Msg, Style, Title) = vbNo Then
'                    Exit For
                    Exit Function
                End If
            Else
            End If
        End If
    Next
    
    Lineitemcheck = True
    
    If Err Then Call LogErr(Name & "::Checkovershiped", Err.Description, Err.number, True)
End Function
'disable navigetor button
Private Sub DisableNav(BackWard As Boolean, Forward As Boolean)
On Error Resume Next

    Forward = Not Forward
    BackWard = Not BackWard
    Navbar1.LastEnabled = Forward
    Navbar1.NextEnabled = Forward
    Navbar1.FirstEnabled = BackWard
    Navbar1.PreviousEnabled = BackWard
End Sub

Private Sub Txttotavolu_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub

Private Sub Txtviacarr_KeyPress(KeyAscii As Integer)
    'Added by Juan 11/17/2000
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If
    '------------------------
End Sub


