VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWarehouse 
   ClientHeight    =   9885
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   14415
   MaxButton       =   0   'False
   ScaleHeight     =   13717.92
   ScaleMode       =   0  'User
   ScaleWidth      =   14415
   Tag             =   "02050700"
   Begin VB.PictureBox imsMsgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4560
      ScaleHeight     =   2025
      ScaleWidth      =   5145
      TabIndex        =   107
      Top             =   3600
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton noButton 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   110
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton yesButton 
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   109
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "If YES it will be received with PO value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   111
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "No supplier invoice has been entered, do you want to continue?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   108
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.PictureBox savingLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4320
      ScaleHeight     =   945
      ScaleWidth      =   3105
      TabIndex        =   63
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "SAVING..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   64
         Top             =   360
         Width           =   3135
      End
   End
   Begin MSComctlLib.TreeView treeNothing 
      Height          =   735
      Left            =   10320
      TabIndex        =   106
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox baseFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   1800
      ScaleHeight     =   1935
      ScaleWidth      =   6975
      TabIndex        =   102
      Top             =   3960
      Width           =   6975
      Begin VB.PictureBox treeFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         ScaleHeight     =   2895
         ScaleWidth      =   4815
         TabIndex        =   103
         Top             =   0
         Width           =   4815
         Begin VB.PictureBox linesH 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10650
            TabIndex        =   104
            Top             =   0
            Visible         =   0   'False
            Width           =   10650
         End
      End
   End
   Begin VB.TextBox invoiceLineBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   7920
      MousePointer    =   1  'Arrow
      TabIndex        =   100
      Text            =   "invoiceListBOX"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox invoiceBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   6480
      MousePointer    =   1  'Arrow
      TabIndex        =   97
      Text            =   "invoiceBOX"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox emailRecepient 
      Height          =   375
      Left            =   3000
      TabIndex        =   93
      Top             =   9405
      Width           =   3255
   End
   Begin VB.CommandButton searchButton 
      Caption         =   "Search"
      Height          =   255
      Left            =   2040
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1800
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid summaryValues 
      Height          =   1815
      Left            =   7680
      TabIndex        =   90
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox price2BOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   840
      MousePointer    =   1  'Arrow
      TabIndex        =   89
      TabStop         =   0   'False
      Text            =   "price2BOX"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox positionBox 
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   600
      MousePointer    =   1  'Arrow
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox poItemBox 
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   87
      Text            =   "poItemBox"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid unitCombo 
      Height          =   495
      Left            =   10800
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo TxtCompany 
      Height          =   375
      Left            =   1080
      TabIndex        =   78
      Top             =   1320
      Width           =   615
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBCamChart 
      Height          =   735
      Left            =   8520
      TabIndex        =   76
      Top             =   7320
      Width           =   1455
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBStockType 
      Height          =   735
      Left            =   8520
      TabIndex        =   75
      Top             =   7200
      Width           =   975
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1720
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBUsChart 
      Height          =   735
      Left            =   8520
      TabIndex        =   74
      Top             =   7200
      Width           =   1455
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBLocation 
      Height          =   735
      Left            =   8520
      TabIndex        =   73
      Top             =   6960
      Width           =   975
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1720
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleCompany 
      Height          =   735
      Left            =   8520
      TabIndex        =   72
      Top             =   6720
      Width           =   855
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1508
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.TextBox remarks 
      Height          =   3015
      Left            =   120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Top             =   6240
      Width           =   11775
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   240
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12840
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton newBUTTON 
      Caption         =   "&New Transaction"
      Height          =   375
      Left            =   9240
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   5
      Left            =   1560
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox NEWconditionBOX 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5640
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   59
      Text            =   "NEWconditionBOX"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1560
      TabIndex        =   57
      Top             =   4290
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox newDESCRIPTION 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   460
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox userNAMEbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox repairBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5880
      MousePointer    =   1  'Arrow
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "repairBOX"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox unitBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   220
      Index           =   0
      Left            =   5880
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "unitBOX"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show &Remarks, FQA"
      Height          =   375
      Left            =   120
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1815
   End
   Begin VB.CommandButton emailButton 
      Caption         =   "E-Mail to"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   9405
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6960
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1575
   End
   Begin VB.CommandButton removeDETAIL 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   12360
      TabIndex        =   45
      Top             =   3870
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox priceBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5880
      MousePointer    =   1  'Arrow
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "priceBOX"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton hideDETAIL 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   11400
      TabIndex        =   41
      Top             =   3870
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton submitDETAIL 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   13320
      TabIndex        =   40
      Top             =   3870
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox quantityBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5880
      MousePointer    =   1  'Arrow
      TabIndex        =   38
      Text            =   "quantityBOX"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox balanceBOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   220
      Index           =   0
      Left            =   5880
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "balanceBOX"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox sublocaBOX 
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5880
      MousePointer    =   1  'Arrow
      TabIndex        =   36
      Text            =   "sublocaBOX"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox logicBOX 
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   5880
      MousePointer    =   1  'Arrow
      TabIndex        =   35
      Text            =   "logicBOX"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox quantity 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   220
      Index           =   0
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "quantity"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWarehouse.frx":0000
            Key             =   "thing"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWarehouse.frx":0142
            Key             =   "thing 0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWarehouse.frx":0284
            Key             =   "thing 1"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox searchFIELD 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   3020
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1770
      Width           =   6210
   End
   Begin VB.TextBox searchFIELD 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   620
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1410
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton saveBUTTON 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1575
   End
   Begin VB.TextBox dateBOX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TextLINE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   11040
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   12735
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1575
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   10440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   16777215
      CustomFormat    =   "MMMM/dd/yyyy"
      Format          =   54001667
      CurrentDate     =   36867
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid STOCKlist 
      Height          =   1620
      Left            =   120
      TabIndex        =   12
      Top             =   2085
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   2858
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid matrix 
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   6840
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1296
      _Version        =   393216
      BackColor       =   16776960
      Rows            =   11
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   2
      Left            =   4080
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1455
      Index           =   0
      Left            =   2520
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.PictureBox linesV 
      Height          =   975
      Index           =   0
      Left            =   2880
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   34
      Top             =   4680
      Visible         =   0   'False
      Width           =   15
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid detailHEADER 
      Height          =   300
      Left            =   120
      TabIndex        =   33
      Top             =   4320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   529
      _Version        =   393216
      Cols            =   6
      RowHeightMin    =   240
      Enabled         =   0   'False
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   4
      Left            =   8040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   4725
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4560
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8334
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid SUMMARYlist 
      Height          =   3900
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6879
      _Version        =   393216
      Cols            =   8
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColorFixed  =   0
      GridColorUnpopulated=   16777215
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo2 
      Height          =   375
      Left            =   2280
      TabIndex        =   79
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   93
      Text            =   "SSOleDBCombo1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo TxtStockType 
      Height          =   375
      Left            =   7680
      TabIndex        =   80
      Top             =   1320
      Width           =   1095
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo TxtUSChart 
      Height          =   375
      Left            =   5280
      TabIndex        =   81
      Top             =   1320
      Width           =   1455
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo TxtLocation 
      Height          =   375
      Left            =   2760
      TabIndex        =   82
      Top             =   1320
      Width           =   1215
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo TxtCamChart 
      Height          =   375
      Left            =   10080
      TabIndex        =   83
      Top             =   1320
      Width           =   1455
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   3
      Left            =   4080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12632064
      BackColorBkg    =   12648447
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox quantity2BOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   220
      Index           =   0
      Left            =   2880
      MousePointer    =   1  'Arrow
      TabIndex        =   84
      Text            =   "quantityBOX"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox unit2BOX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   220
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   85
      TabStop         =   0   'False
      Text            =   "unitBOX"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBFQA 
      Height          =   2220
      Left            =   120
      TabIndex        =   71
      Top             =   3840
      Visible         =   0   'False
      Width           =   12615
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   10
      Columns(0).Width=   2566
      Columns(0).Caption=   "StockNumber"
      Columns(0).Name =   "StockNumber"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3122
      Columns(1).Caption=   "Company"
      Columns(1).Name =   "Company"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4207
      Columns(2).Caption=   "Location"
      Columns(2).Name =   "Location"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2566
      Columns(3).Caption=   "USChart#"
      Columns(3).Name =   "USChart#"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   4419
      Columns(4).Caption=   "StockType"
      Columns(4).Name =   "StockType"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2619
      Columns(5).Caption=   "CamChart#"
      Columns(5).Name =   "CamChart#"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "Ponumb"
      Columns(6).Name =   "Ponumb"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "LineNo"
      Columns(7).Name =   "LineNo"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   2381
      Columns(8).Caption=   "Condition"
      Columns(8).Name =   "ToCond"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1693
      Columns(9).Caption=   "Quantity"
      Columns(9).Name =   "Quantity"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      _ExtentX        =   22251
      _ExtentY        =   3916
      _StockProps     =   79
      Caption         =   "FQA"
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
   Begin VB.Label nothing 
      Caption         =   "nothing"
      Height          =   135
      Left            =   12720
      TabIndex        =   105
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label invoiceLineLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "invoiceLine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7845
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label invoiceNumberLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "invoice:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9390
      TabIndex        =   99
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label invoiceLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "invoice:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7845
      TabIndex        =   98
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label logLabel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label5"
      Height          =   255
      Left            =   10560
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label serialLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   95
      Top             =   4035
      Width           =   1215
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial:"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   94
      Top             =   4035
      Width           =   1335
   End
   Begin VB.Label poItemLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblUSChart 
      Caption         =   "US Chart#"
      Height          =   255
      Left            =   4200
      TabIndex        =   70
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LBLCompany 
      Caption         =   "Company"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LblLocation 
      Caption         =   "Location"
      Height          =   255
      Left            =   1920
      TabIndex        =   69
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LblType 
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   68
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LblCamChart 
      Caption         =   "Cam. Chart #"
      Height          =   255
      Left            =   9000
      TabIndex        =   67
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Search Field"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   66
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Search Field"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   65
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label unitLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   61
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   60
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   1500
      X2              =   1500
      Y1              =   5034.761
      Y2              =   4028.642
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "New Commodity:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   55
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label summaryLABEL 
      Caption         =   "Summary"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   44
      Top             =   4035
      Width           =   1335
   End
   Begin VB.Label unitLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   43
      Top             =   4035
      Width           =   1335
   End
   Begin VB.Label descriptionLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3000
      TabIndex        =   30
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11880
      Y1              =   1783.26
      Y2              =   1783.26
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "Commodity:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label transactionACTIVE 
      Caption         =   "transactionACTIVE"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label label 
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label label 
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   10920
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label remarksLABEL 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label label 
      Caption         =   "Transaction #"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label label 
      Caption         =   "Company"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label label 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label commodityLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Menu treeMENU 
      Caption         =   "Edit"
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu addITEM 
         Caption         =   "Add Serial"
      End
      Begin VB.Menu deleteITEM 
         Caption         =   "Delete Item"
      End
   End
End
Attribute VB_Name = "frmWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim thisFORM As FormMode
Dim usingARROWS As Boolean
Public POrowguid As String
Dim locked As Boolean
Dim STOCKlocked As Boolean
Dim dbtablename As String, grid1 As Boolean, grid2 As Boolean
Dim POValue
Dim doChanges As Boolean
'Juan 2010-7-17
Dim inProgress As Boolean
Dim isReset As Boolean

Dim ctt As New cTreeTips
Dim ctt1 As New cTreeTips
Dim ctt2 As New cTreeTips
Dim ctt3 As New cTreeTips


Sub arrowKEYS(direction As String, Index As Integer)
Dim grid As MSHFlexGrid
    With cell(Index)
        Set grid = combo(Index)
            grid.Visible = True
            Call gridCOLORnormal(grid, Val(grid.tag))
            Select Case direction
                Case "down"
                    If grid.row < (grid.Rows - 1) Then
                        If grid.row = 0 And .text = "" Then
                            .text = grid.text
                        Else
                            grid.row = grid.row + 1
                        End If
                    Else
                        grid.row = grid.Rows - 1
                    End If
                Case "up"
                    If grid.row > 0 Then
                        grid.row = grid.row - 1
                    Else
                        grid.row = 1
                    End If
            End Select
            
            grid.tag = grid.row
            If Not grid.Visible Then
                grid.Visible = True
            End If
            grid.ZOrder
            grid.TopRow = IIf(grid.row = 0, 1, grid.row)
            usingARROWS = True
            Call gridCOLORdark(grid, grid.row)
            grid.SetFocus
    End With
End Sub

Sub editMarkRow(StockNumber As String, isSerial As Boolean)
    Dim i As Integer
    Dim markIt As Boolean
    Dim currentformname, currentformname1
    Dim imsLock As imsLock.Lock
    markIt = False
    With STOCKlist
        For i = 1 To STOCKlist.Rows - 1
            If RTrim(.TextMatrix(i, 1)) = RTrim(StockNumber) Then
                markIt = True
                .row = i
                Exit For
            End If
        Next
                
        Set imsLock = New imsLock.Lock
        currentformname = frmWarehouse.tag + "stock"
        currentformname1 = currentformname
        
        If markIt Then
            .col = 0
            .CellFontName = "Wingdings 3"
            .CellFontSize = 10
            If isSerial Then
                .text = "?"
            Else
                .text = "?"
            End If
        End If
    End With
End Sub

Sub editSummaryList()
Dim r, QTYpo, cond, condName, serial, key, lastLine, thick, total, poItem, loca, subloca, logicname, sublocaname
Dim isSerial As Boolean
Dim moreSerial As Boolean
Dim StockNumber As String
Dim qtyColumn
Dim datax As ADODB.Recordset
Dim sql As String
Dim isEditionActive  As Boolean
Dim translationAdjustmentEntry
isSerial = False
isEditionActive = False
Screen.MousePointer = 11
Set datax = New ADODB.Recordset
sql = "SELECT * FROM form_settings WHERE form = 'warehouse' AND mode = 'edition'"
datax.Open sql, cn, adOpenStatic
If datax.RecordCount > 0 Then
    isEditionActive = datax!active
End If
datax.Close
If isEditionActive = False Then Exit Sub
    Call workBOXESlist("clean")
    isFirstSubmit = False
    qtyColumn = 5
    With SUMMARYlist
        r = .row
        '1 "Commodity"
        '2 "Serial"
        '3 "Condition"
        '4 "Unit Price"
        '5 "Description"
        '6 "Unit"
        '7 "Qty"
        '8 "node"
        '9 "From Logical"
        '10 "From Subloca"
        '11 "To Logical"
        '12 "To Subloca"
        '13 "New Condition Code"
        '14 "New Condition Description"
        '15 "Unit Code"
        '16 "Computer Factor"
        '17 "PO qty"
        '20 "Original Condition Code"
        '21 "Secundary Qty"
        '25 "ratio"
       
        Tree.width = frmWarehouse.detailHEADER.width
         'Juan 2014-01-12, resizing treeFrame ------
'        treeFrame.Left = detailHEADER.ColWidth(0) + 5200
'        treeFrame.width = Tree.width - frmWarehouse.detailHEADER.ColWidth(0)
'        treeFrame.Top = width = frmWarehouse.detailHEADER.Top + frmWarehouse.detailHEADER.Height
        
        baseFrame.Left = detailHEADER.ColWidth(0) + 5200
        baseFrame.width = Tree.width - frmWarehouse.detailHEADER.ColWidth(0)
        baseFrame.Top = width = frmWarehouse.detailHEADER.Top + frmWarehouse.detailHEADER.Height
        baseFrame.Height = 5000
        treeFrame.Top = 0
        treeFrame.width = baseFrame.width
        
        ' ------------
        
        Tree.Nodes.Clear
        StockNumber = .TextMatrix(r, 1)
        QTYpo = 0
        Dim translationCondition, translationLogical, translationSublocation, translationSerial, translationPool, translationTotal, translationTot
        translationCondition = translator.getIt("translationCondition") + ": "
        translationLogical = translator.getIt("translationLogical") + ": "
        translationSublocation = translator.getIt("translationSubLocation") + ": "
        translationSerial = translator.getIt("translationSerial") + ": "
        translationPool = translator.getIt("translationPool") + ": "
        translationTotal = translator.getIt("translationTotal") + ": "
        translationTot = translator.getIt("translationTot")
        
        Select Case frmWarehouse.tag
            'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales,ReturnFromWell
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                cond = .TextMatrix(r, 13)
                condName = .TextMatrix(r, 14)
                loca = .TextMatrix(r, 9)
                subloca = .TextMatrix(r, 10)
                serial = .TextMatrix(r, 2)
                QTYpo = .TextMatrix(r, 17)
                
                Tree.Nodes.Add , tvwChild, "@" + cond, translationCondition + cond + " - " + condName, "thing"
                Tree.Nodes("@" + cond).Bold = True
                Tree.Nodes("@" + cond).backcolor = &HE0E0E0
                Tree.Nodes.Add "@" + cond, tvwChild, "@" + cond + "{{" + loca, translationLogical + loca, "thing 0"
                'If (frmWarehouse.tag = "02040400" Or frmWarehouse.tag = "02040500" Or frmWarehouse.tag = "02040300") Then
                    logicname = .TextMatrix(r, 9)
                    sublocaname = subloca
                    key = cond + "-" + condName + "{{" + loca + "{{" + subloca
                    If UCase(serial) = "POOL" Then
                        Tree.Nodes.Add "@" + cond + "{{" + loca, tvwChild, key, translationSublocation + sublocaname, "thing 1"
                        Call setupBoxes2(Tree.Nodes.Count, r, False, QTYpo)
                    Else
                        isSerial = True
                        Tree.Nodes.Add "@" + cond + "{{" + loca, tvwChild, key, translationSublocation + sublocaname, "thing 0"
                    End If
                'End If
                If isSerial Then
                    Tree.Nodes.Add key, tvwChild, key + "#" + serial, translationSerial + serial, "thing 1"
                    Call setupBoxes2(Tree.Nodes.Count, r, True, .TextMatrix(r, 17))
                End If
                Tree.Nodes.Add , , translationTot, Space(53) + translationTotal
                Tree.Nodes("Total").Bold = True
                Tree.Nodes("Total").backcolor = &HC0C0C0
                Tree.Nodes(Tree.Nodes.Count - 1).Selected = True
                Call imsWarehouse.bottomLine(totalNode, total, Not isSerial, StockNumber, True, lastLine, ctt)
            Case "02050200" 'AdjustmentEntry
                serial = .TextMatrix(r, 2)
                commodityLABEL = StockNumber
                unitLABEL(0) = .TextMatrix(r, 6)
                unitLABEL(1) = .TextMatrix(r, 15)
                descriptionLABEL = .TextMatrix(r, 5)
                translationAdjustmentEntry = translator.getIt("translationAdjustmentEntry") + " "
                Tree.Nodes.Add , tvwChild, "@AE", translationAdjustmentEntry, "thing"
                Tree.Nodes("@AE").Bold = True
                Tree.Nodes("@AE").backcolor = &HE0E0E0
                key = "AE-NEW{"
                If UCase(serial) = "POOL" Then
                    Tree.Nodes.Add "@AE", tvwChild, key + "{{Pool", translationPool, "thing 1"
                    Call setupBoxes2(Tree.Nodes.Count, r, False)
                Else
                    isSerial = True
                    Tree.Nodes.Add "@AE", tvwChild, key + "{{Serial", translationSerial, "thing 1"
                    Tree.Nodes(Tree.Nodes.Count).text = translationSerial + serial
                    Call setupBoxes2(Tree.Nodes.Count, r, True)
                End If
                Call setupBoxes2(Tree.Nodes.Count, r, isSerial)
                
                Tree.Nodes.Add , , translationTot, Space(53) + translationTotal
                Tree.Nodes("Total").Bold = True
                Tree.Nodes("Total").backcolor = &HC0C0C0
                Tree.Nodes(Tree.Nodes.Count - 1).Selected = True
                Call imsWarehouse.bottomLine(totalNode, total, Not isSerial, StockNumber, True, lastLine, ctt)

            Case "02040100" 'WarehouseReceipt
                '17 "Qty po"
                '21 "PO" --- changed to 24
                '22 "lineitem"
                QTYpo = .TextMatrix(r, 17)
                cond = .TextMatrix(r, 13)
                condName = .TextMatrix(r, 14)
                serial = .TextMatrix(r, 2)
                commodityLABEL = StockNumber
                poItemLabel = .TextMatrix(r, 8)
                unitLABEL(0) = .TextMatrix(r, 6)
                unitLABEL(1) = .TextMatrix(r, 15)
                descriptionLABEL = .TextMatrix(r, 5)
                poItem = .TextMatrix(r, 22)
                poItemLabel = poItem
                'Juan 2010-9-5 To add ratio functionallity
                If IsNumeric(.TextMatrix(r, 25)) Then
                    ratioValue = Val(.TextMatrix(r, 25))
                Else
                    ratioValue = 1
                End If
                '-------------------
                    'Juan 2014-8-27 new version of calculation based on invoice if exists
                    Select Case frmWarehouse.tag
                        Case "02040100" 'WarehouseReceipt
                            Set datax = New ADODB.Recordset
                            sql = "select * from invoicedetl where invd_npecode = '" + nameSP + "'  " + _
                                "and invd_invcnumb = '" + summaryValues.TextMatrix(r, 2) + "' " + _
                                "and invd_liitnumb = '" + Trim(summaryValues.TextMatrix(r, 3)) + "'"
                            datax.Open sql, cn, adOpenStatic
                            If datax.RecordCount > 0 Then
                                ratioValue = datax!invd_secoreqdqty / datax!invd_primreqdqty
                            End If
                        Case Else
                    End Select
                    datax.Close
                    '----------------------
            
                Tree.Nodes.Add , tvwChild, "@" + cond, cond + "-" + condName, "thing"
                Tree.Nodes("@" + cond).Bold = True
                Tree.Nodes("@" + cond).backcolor = &HE0E0E0
                key = cond + "-" + condName + "{{"
                If serial = "Pool" Then
                    Tree.Nodes.Add "@" + cond, tvwChild, key + "{{Pool", translationPool, "thing 1"
                    Call setupBoxes2(Tree.Nodes.Count, r, False, QTYpo)
                Else
                    isSerial = True
                    Tree.Nodes.Add "@" + cond, tvwChild, key + "{{Serial", translationSerial, "thing 1"
                    Tree.Nodes(Tree.Nodes.Count).text = translationSerial + serial
                    Call setupBoxes2(Tree.Nodes.Count, r, True, QTYpo)
                End If
            
                Select Case frmWarehouse.tag
                    Case "02040100" 'WarehouseReceipt
                        Tree.Nodes.Add , , translationTot, Space(90) + IIf(frmWarehouse.newBUTTON.Enabled, Space(24), "")
                    Case Else
                        Tree.Nodes.Add , , translationTot, Space(53) + IIf(frmWarehouse.newBUTTON.Enabled, Space(24), translationTotal)
                End Select
                Tree.Nodes("Total").Bold = True
                Tree.Nodes("Total").backcolor = &HC0C0C0
                
                totalNode = Tree.Nodes.Count
                total = .TextMatrix(r, 7)
                Call imsWarehouse.bottomLine(totalNode, total, Not isSerial, StockNumber, True, lastLine, ctt)
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        Dim stockListRow  As Integer
        stockListRow = findSTUFF(StockNumber, STOCKlist, 1)
        Select Case frmWarehouse.tag
            Case "02050200" 'AdjustmentEntry
                latestStockNumberQty = "n/a"
            Case Else
                latestStockNumberQty = STOCKlist.TextMatrix(stockListRow, qtyColumn)
        End Select
    End With
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    SUMMARYlist.Visible = False
    Call SHOWdetails
    Screen.MousePointer = 0
End Sub

Sub getEmail()
    Dim datax As New ADODB.Recordset
    Dim sql, emailText As String
    emailRecepient.text = ""
    sql = "select email from xuserprofile where " _
        & "usr_npecode = '" + nameSP + "' and  usr_userid = '" + CurrentUser + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If Not datax.EOF Then
        datax.MoveFirst
        If IsNull(datax!Email) Then
            emailText = ""
        Else
            emailText = LTrim(datax!Email)
        End If
        emailRecepient.text = emailText
    End If
    datax.Close
End Sub

Sub limitQty(Index As Integer)
    Select Case frmWarehouse.tag
        Case "02040100" 'WarehouseReceipt
            If (invoiceBOX(Index) <> "" And invoiceBOX(Index) <> "invoiceBOX") Then
                Dim originalQty, sumQty As Double
                Dim i As Integer
                For i = 1 To STOCKlist.Rows - 1
                    If commodityLABEL = STOCKlist.TextMatrix(i, 1) Then
                        If Left(STOCKlist.TextMatrix(i, 7), 1) <> "@" Then
                            originalQty = CDbl(STOCKlist.TextMatrix(i, 9)) 'gets original qty to receieve
                            Exit For
                        End If
                    End If
                Next
                For i = 1 To SUMMARYlist.Rows - 1
                    If commodityLABEL = STOCKlist.TextMatrix(i, 1) Then
                        If IsNumeric(SUMMARYlist.TextMatrix(i, 7)) Then
                            sumQty = sumQty + CDbl(SUMMARYlist.TextMatrix(i, 7))
                        End If
                    End If
                Next
                If (originalQty - sumQty) = 0 Then
                    MsgBox "No more items to receive"
                    quantityBOX(Index) = "0.00"
                Else
                    If CDbl(quantityBOX(Index)) > (originalQty - sumQty) Then
                        If Left(STOCKlist.TextMatrix(i, 7), 1) = "@" Then
                            If CDbl(quantityBOX(Index)) > (quantity(Index)) Then
                                quantityBOX(Index) = quantity(Index)
                            Else
                                quantityBOX(Index).text = Format((originalQty - sumQty), "0.00")
                                MsgBox "Total remaining to be received vs PO: " + Format((originalQty - sumQty), "0.00")
                            End If
                        Else
                            quantityBOX(Index).text = Format((originalQty - sumQty), "0.00")
                            MsgBox "Total remaining to be received vs PO: " + Format((originalQty - sumQty), "0.00")
                        End If
                    Else
                        If Left(STOCKlist.TextMatrix(i, 7), 1) = "@" Then
                            If CDbl(quantityBOX(Index)) > (quantity(Index)) Then
                                quantityBOX(Index) = quantity(Index)
                            End If
                        End If
                    End If
                End If
            End If
        'Case "02050300" 'AdjustmentIssue
        'Case "02040200" 'WarehouseIssue
        'Case "02040500" 'WellToWell
        'Case "02040300" 'Return from Well
        'Case "02040600" 'WarehouseToWarehouse
        Case "02040400" 'ReturnFromRepair
        Case "02050200" 'AdjustmentEntry
        Case "02040700" 'InternalTransfer
        Case "02050400" 'Sales
        Case Else
            If CDbl(quantityBOX(Index)) > CDbl(quantity(Index)) Then quantityBOX(Index).text = quantity(Index)
    End Select
End Sub

Function locateLine(StockNumber As String, searchValue As String, Optional col As Integer) As Integer
    Dim i
    If col = 0 Then col = 10
    With frmWarehouse.SUMMARYlist
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = StockNumber Then
                If .TextMatrix(i, col) = searchValue Then
                    locateLine = i
                    Exit For
                Else
                    locateLine = 0
                End If
            Else
                locateLine = 0
            End If
        Next
    End With
End Function

Public Sub searchStockNumber(Index As Integer)
Dim datax As New ADODB.Recordset
Dim sql, list, i, ii, t
Screen.MousePointer = 11
      
            If Index = 0 Then
                If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                    sql = "SELECT stk_stcknumb, stk_desc, uni_desc " _
                        & "FROM STOCKMASTER LEFT OUTER JOIN UNIT ON " _
                        & "stk_npecode = uni_npecode AND " _
                        & "stk_primuon = uni_code WHERE " _
                        & "(stk_npecode = '" + nameSP + "') AND " _
                        & "(stk_stcknumb like '" + searchFIELD(Index).text + "%')"
                    datax.Open sql, cn, adOpenStatic
                    With STOCKlist
                        .Rows = 2
                        .TextMatrix(1, 0) = ""
                        .TextMatrix(1, 1) = ""
                        .TextMatrix(1, 2) = ""
                        .TextMatrix(1, 3) = ""
                        If datax.RecordCount > 0 Then
                            STOCKlist.Rows = datax.RecordCount + 1
                            Dim r As Integer
                            r = 1
                            Do While Not datax.EOF
                                If findSTUFF(datax!stk_stcknumb, STOCKlist, 1) = 0 Then
                                    .TextMatrix(r, 1) = datax!stk_stcknumb
                                    .TextMatrix(r, 2) = datax!stk_desc
                                    .TextMatrix(r, 3) = datax!uni_desc
                                    r = r + 1
                                    'STOCKlist.addITEM "" + vbTab + datax!stk_stcknumb + vbTab + datax!stk_desc + vbTab + datax!uni_desc & "", 1
                                End If
                                datax.MoveNext
                                Loop
                                STOCKlist.RowHeight(1) = 240
                            If STOCKlist.Rows > 2 And STOCKlist.TextMatrix(1, 1) = "" Then STOCKlist.RemoveItem 1
                            Call reNUMBER(STOCKlist)
                        End If
                    End With
                End If
            Else
                If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                    If searchFIELD(Index) <> "" Then
                    sql = "SELECT stk_stcknumb, stk_desc, uni_desc " _
                        & "FROM STOCKMASTER LEFT OUTER JOIN UNIT ON " _
                        & "stk_npecode = uni_npecode AND " _
                        & "stk_primuon = uni_code WHERE " _
                        & "(stk_npecode = '" + nameSP + "') "
                        Call doARRAYS("s", searchFIELD(1), list)
                        If UBound(list) >= 0 Then
                            For i = 0 To UBound(list)
                                sql = sql + "AND stk_desc LIKE '%" + list(i) + "%' "
                            Next
                            datax.Open sql, cn, adOpenStatic
                            If datax.RecordCount > 0 Then
                                Do While Not datax.EOF
                                    If findSTUFF(datax!stk_stcknumb, STOCKlist, 1) = 0 Then
                                        t = "" + vbTab
                                        t = t + datax!stk_stcknumb + vbTab
                                        t = t + IIf(IsNull(datax!stk_desc), "", datax!stk_desc) + vbTab
                                        t = t + IIf(IsNull(datax!uni_desc), "", datax!uni_desc)
                                        STOCKlist.addITEM t, 1
                                    End If
                                    datax.MoveNext
                                Loop
                                STOCKlist.RowHeight(1) = 240
                                If STOCKlist.Rows > 2 And STOCKlist.TextMatrix(1, 1) = "" Then STOCKlist.RemoveItem 1
                                Call reNUMBER(STOCKlist)
                            End If
                        End If
                    End If
                Else
                    Call search(searchFIELD(1), STOCKlist, 3)
                    searchFIELD(1).SelStart = 0
                    searchFIELD(1).SelLength = Len(searchFIELD(1))
                End If
            End If
        
    STOCKlist.TopRow = 1
    Screen.MousePointer = 0
End Sub

Sub sendEMail(fileName As String, reportCaption As String, parameters() As String, recipents As String, subject As String, reportName As String, path As String)
Dim Params(1) As String
Dim i As Integer
Dim Attachments() As String
Dim str As String
Dim attention As String
On Error GoTo errMESSAGE
    Dim size As Integer
    attention = "Please find here attached report  "
    Attachments = generateattachmentswithCR11(fileName, reportCaption, parameters, reportName, path)
    Dim sender As String
    sender = ""
    Call WriteParameterFiles(recipents, sender, Attachments, subject, attention)
errMESSAGE:
    If Err.Number <> 0 Then
        MsgBox "Process sendEMail " + Err.description
    End If
End Sub

Private Sub StockListDuplicate_Click()

End Sub


Private Sub Command2_Click()
    Call searchFIELD_KeyPress(0, 13)
End Sub

Sub updateEmail()
    Dim sql, emailText As String
    sql = "update  xuserprofile set email = ' " + emailRecepient + "' where " _
        & "usr_npecode = '" + nameSP + "' and  usr_userid = '" + CurrentUser + "'"
    cn.Execute sql
End Sub

Private Sub emailButton_Click()
Dim reportName As String
Dim reportPATH As String
Dim parameters(2) As String
Dim subject As String
Dim reportCaption As String
reportPATH = repoPATH + "\"
If treeFrame.Visible = True Then
    Screen.MousePointer = 0
    MsgBox "There is a pending item to submit"
    Exit Sub
End If
Select Case frmWarehouse.tag
    Case "02040400" 'ReturnFromRepair
        reportCaption = "Return From Repair"
        reportName = "wareRR.rpt"
    Case "02050200" 'AdjustmentEntry
        reportCaption = "Adjustment Entry"
        reportName = "wareAEIA.rpt"
    Case "02040200" 'WarehouseIssue
        reportCaption = "Warehouse Issue"
        reportName = "wareI.rpt"
    Case "02040500" 'WellToWell
        reportCaption = "Well To Well"
        reportName = "wareI.rpt"
    Case "02040700" 'InternalTransfer
        reportCaption = "Internal Transfer"
        reportName = "wareI.rpt"
    Case "02050300" 'AdjustmentIssue
        reportCaption = "Adjustment Issue"
        reportName = "wareI.rpt"
    Case "02040600" 'WarehouseToWarehouse
        reportCaption = "Warehouse To Warehouse"
        reportName = "wareI.rpt"
    Case "02040100" 'WarehouseReceipt
        reportCaption = "Warehouse Receipt"
        reportName = "wareR.rpt"
    Case "02050400" 'Sales
        reportCaption = "Sales"
        reportName = "wareSL.rpt"
    Case "02040300" 'Return from Well
        reportCaption = "Return from Well"
        reportName = "wareRT.rpt"
    End Select
    parameters(1) = nameSP
    parameters(0) = cell(0)
    subject = "Copy of  " + reportName
    reportCaption = reportCaption + "Report"
    Call sendEMail(reportPATH & reportName, reportCaption, parameters, emailRecepient.text, subject, reportName, reportPATH)
    Call updateEmail
End Sub

Private Sub logicBOX_Validate(Index As Integer, Cancel As Boolean)
If skipExistance Then
    skipExistance = False
    Exit Sub
End If
If UCase(logicBOX(Index)) <> "GENERAL" Then
    If DoesItemExist(logicBOX(Index), grid(1), 1) = False Then
        Cancel = True
        MsgBox "Logic Warehouse does not exist, please select a valid one from the list.", vbInformation
        skipExistance = True
    End If
End If
With logicBOX(Index)
    If .text = "" Then
        .backcolor = &HC0C0FF
    Else
        .backcolor = vbWhite
    End If
End With
End Sub

Private Sub noButton_Click()
    
    msgBoxResponse = False
End Sub

Private Sub quantity2BOX_Change(Index As Integer)
    'If doChanges Then
        'Call quantity2BOX_Validate(Index, True)
    'Else
    '    doChanges = True
    'End If
End Sub


Private Sub quantity2BOX_Click(Index As Integer)
    With quantity2BOX(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub


Private Sub quantity2BOX_GotFocus(Index As Integer)
    If Index <> totalNode Then
        Call whitening
        quantity2BOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantity2BOX_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call quantity2BOX_Validate(Index, True)
    End If
End Sub

Private Sub quantity2BOX_LostFocus(Index As Integer)
    Call quantity2BOX_Validate(Index, True)
    If Index <> totalNode Then quantity2BOX(Index).backcolor = vbWhite
End Sub

Private Sub quantity2BOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        quantity2BOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantity2BOX_Validate(Index As Integer, Cancel As Boolean)
Dim qty, qty2
On Error Resume Next
    With quantity2BOX(Index)
        If Index <> totalNode Then
            If IsNumeric(.text) Then
                If CDbl(.text) > 0 Then
                    'Juan 2010-6-5
                    '.text = Format(.text, 0)
                    .text = Format(.text, "0.00")
                    'doChanges = False
                    
                    'Juan 2010-9-4 implementing ratio rather than computer factor
                    If ratioValue > 1 Then
                        If IsNumeric(.text) Then
                            qty2 = CDbl(.text)
                            If qty2 > 0 Then
                                qty = qty2 / ratioValue
                                quantityBOX(Index).text = Format(qty, "0.00")
                            Else
                                quantityBOX(Index).text = .text
                            End If
                        End If
                    Else
                        quantityBOX(Index).text = .text
                    End If
'                    If computerFactorValue > 0 Then
'                        If IsNumeric(.text) Then
'                            qty2 = CDbl(.text)
'                            If qty2 > 0 Then
'                                If Round(computerFactorValue) > 0 Then
'                                    qty = qty2 * computerFactorValue / 10000
'                                Else
'                                    qty = qty2 * (10000 * computerFactorValue)
'                                End If
'                                quantityBOX(Index).text = Format(qty, "0.00")
'                            Else
'                                quantityBOX(Index).text = .text
'                            End If
'                        End If
'                    Else
'                        quantityBOX(Index).text = .text
'                    End If
                    '--------------------------
                    Select Case frmWarehouse.tag
                        Case "02050200" 'AdjustmentEntry
                        Case Else
                            'If CDbl(.text) > CDbl(quantity(Index)) Then .text = quantity(Index)
                    End Select
                Else
                    'Juan 2010-6-5
                    '.text = "0"
                    .text = "0.00"
                    '-----------------
                End If
                If Err.Number = 0 Then

                    Select Case .tag
                        Case "02040100" 'WarehouseReceipt
                            Call calculations(True, True)
                        Case Else
                            Call calculations(True)
                    End Select
                        
                End If

            Else
                    'Juan 2010-6-5
                    '.text = "0"
                    .text = "0.00"
                    '-----------------
            End If
        End If
        .SelStart = Len(.text)
    End With

End Sub

Private Sub quantityBOX_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    submitted = False
End Sub

Private Sub searchButton_Click()
    Call searchStockNumber(0)
End Sub

Private Sub sublocaBOX_Validate(Index As Integer, Cancel As Boolean)
'juan 2012-1-14 to avoid t he problem when logical warehouse shows up with no reason
If SUMMARYlist.Visible Then Exit Sub
If Tree.Visible = False Then Exit Sub
'------------
If skipExistance Then
    skipExistance = False
    Exit Sub
End If
If UCase(sublocaBOX(Index)) <> "GENERAL" Then
    If DoesItemExist(sublocaBOX(Index), grid(2), 0) = False Then
        Cancel = True
        MsgBox "Sub Location does not exist, please select a valid one from the list.", vbInformation
    End If
End If
End Sub

Sub fillTRANSACTION(datax As ADODB.Recordset)
Dim i, n, rec, condition, key, conditionCODE, fromlogic
Dim fromSubLoca, unitCODE, unit, StockNumber, unitPRICE
Dim shot
    Call cleanDETAILS
    Call hideDETAILS
    STOCKlist.Visible = False
    searchFIELD(0).Visible = False
    searchFIELD(1).Visible = False
    searchButton.Visible = False
    'treeFrame.Visible = False
    baseFrame.Visible = False
    
    Tree.Height = 2000
    SUMMARYlist.Top = searchFIELD(0).Top
    'SUMMARYlist.Height = (Tree.Top + Tree.Height) - SUMMARYlist.Top M
    
    SUMMARYlist.Height = 1980 + 2340 + 1740 'M
    
    SUMMARYlist.ZOrder
    summaryLABEL.Top = SUMMARYlist.Top - 240
    'summaryLABEL.Visible = True M
    summaryLABEL.Visible = False
    
    If newBUTTON.Enabled Then

        remarks.Top = SSOleDBFQA.Top + SSOleDBFQA.Height + 200   'detailHEADER.Top
        remarks.Height = Tree.Top - detailHEADER.Top + Tree.Height '- SSOleDBFQA.Height
        remarks.width = detailHEADER.width
    Else
        remarks.Top = Tree.Top + 2000 + 600
        If Me.Height > (remarks.Top + 990) Then
            remarks.Height = Me.Height - remarks.Top - 790
        End If
    End If
    remarks.Visible = True
    'remarks.ZOrder M
    'remarksLABEL.Top = remarks.Top - 240 M
    'remarksLABEL.Visible = True M
    
    
    remarks.locked = True
    Me.Refresh
    
    dateBOX = Format(datax!Date, "Short Date")
    userNAMEbox = getUSERname(datax!userCODE)
    remarks = IIf(IsNull(datax!remarks), "", datax!remarks)
    With SUMMARYlist
        .Rows = 2
        i = 0
        
        cell(1).tag = datax!Company
        directCLICK = True
        cell(1) = getCOMPANYdescription(cell(1).tag)
        Select Case frmWarehouse.tag
            Case "02040400", "02040100", "02040300" 'ReturnFromRepair, 'WarehouseReceipt, 'Return from Well
                cell(2).tag = datax!FromPlace
                directCLICK = True
                cell(2) = getLOCATIONdescription(cell(2).tag)
                cell(3).tag = datax!Warehouse
                directCLICK = True
                cell(3) = getLOCATIONdescription(cell(3).tag)
                If frmWarehouse.tag = "02040100" Then
                    cell(4).tag = datax!PO
                    directCLICK = True
                    cell(4) = cell(4).tag
                End If
            Case "02050200" 'AdjustmentEntry
                cell(2).tag = datax!FromPlace
                directCLICK = True
                cell(2) = getLOCATIONdescription(cell(2).tag)
            Case "02040200", "02040500", "02040600", "02050400"  'WarehouseIssue, 'WellToWell, 'WarehouseToWarehouse, 'Sales
                cell(2).tag = datax!Warehouse
                directCLICK = True
                cell(2) = getLOCATIONdescription(cell(2).tag)
                cell(3).tag = datax!IssueToPlace
                directCLICK = True
                cell(3) = getLOCATIONdescription(cell(3).tag)
            Case "02040700", "02050300" 'InternalTransfer, 'AdjustmentIssue
                cell(2).tag = datax!Warehouse
                directCLICK = True
                cell(2) = getLOCATIONdescription(cell(2).tag)
        End Select
        Do While Not datax.EOF
            shot = ImsDataX.GetConditions(nameSP, IIf(IsNull(datax!originalcondition), "", datax!originalcondition), True, cn)
            condition = shot(0)
            conditionCODE = shot(1)
            StockNumber = datax!StockNumber
            rec = Format(datax!TransactionLine) + vbTab
            rec = rec + StockNumber + vbTab
            If datax!serialNumber <> "" Then
                If newBUTTON.Enabled Then
                    rec = rec + Trim(datax!serialNumber) + vbTab
                Else
                    rec = rec + Trim(datax!serialNumber) + vbTab
                End If
            Else
                rec = rec + "Pool" + vbTab
            End If
            rec = rec + condition + vbTab
            rec = rec + Format(datax!unitPRICE, "0.00") + vbTab
            rec = rec + IIf(IsNull(datax!StockDescription), "", datax!StockDescription) + vbTab
            unitCODE = getUNIT(StockNumber)
            unit = getUNITdescription(unitCODE)
            rec = rec + unit + vbTab
            rec = rec + Format(datax!qty1) + vbTab
            rec = rec + Format(i) + vbTab
            rec = rec + Trim(IIf(IsNull(datax!fromlogic), "", datax!fromlogic)) + vbTab
            rec = rec + Trim(IIf(IsNull(datax!fromSubLoca), "", datax!fromSubLoca)) + vbTab
            rec = rec + IIf(IsNull(datax!toLOGIC), "", Trim(datax!toLOGIC)) + vbTab
            rec = rec + IIf(IsNull(datax!toSUBLOCA), "", Trim(datax!toSUBLOCA)) + vbTab
            rec = rec + IIf(IsNull(datax!originalcondition), "", datax!originalcondition) + vbTab
            rec = rec + unit
            .addITEM rec
            .TextMatrix(.Rows - 1, 20) = conditionCODE
            datax.MoveNext
            i = i + 1
        Loop
        If .Rows > 2 Then .RemoveItem 1
        Call reNUMBER(SUMMARYlist)
    End With
    directCLICK = False
End Sub

Sub fillGRID(ByRef grid As MSHFlexGrid, box As textBOX, Index)
'On Error Resume Next
Dim paraVECTOR
Dim i, n, rec, list, size, totalwidth, cols, wide(), title(), extraW, sql, clue, Flag
Dim datax As New ADODB.Recordset
    Err.Clear
Dim translationLogical, translationCode, translationDescription, translationSublocation, translationCondition
    
    translationLogical = translator.getIt("translationLogical")
    translationCode = translator.getIt("translationCode")
    translationDescription = translator.getIt("translationDescription")
    translationSublocation = translator.getIt("translationSubLocation")
    translationCondition = translator.getIt("translationCondition")
    Select Case box.name
        Case "logicBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 3000
            wide(1) = 1200
            ReDim title(2)
            'title(0) = "Logical Warehouse"
            title(0) = translationLogical
            title(1) = translationCode
            sql = "select lw_code Code , lw_desc Description from LOGWAR" _
                & " where lw_actvflag = 1 AND lw_npecode = '" & nameSP & "' order by lw_desc "
            Set datax = New ADODB.Recordset
            list = Array("Description", "Code")
        Case "sublocaBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 3000
            wide(1) = 1200
            ReDim title(2)
            title(0) = translationSublocation
            title(1) = translationCode
            Set datax = getDATA("getSUBLOCA", nameSP)
            list = Array("Description", "Code")
        Case "NEWconditionBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 500
            wide(1) = 2400
            ReDim title(2)
            title(0) = translationCode
            title(1) = translationCondition
            sql = "SELECT cond_condcode as Code, cond_desc as Condition FROM condition WHERE " _
                & "cond_npecode = '" + nameSP + "' " _
                & "ORDER BY cond_condcode"
            list = Array("code", "Condition")
    End Select
    If datax.State <> 1 Then datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount < 1 Then Exit Sub
    If Err.Number = 3704 Then
        Err.Clear
        Exit Sub
    End If
    
    With grid
        totalwidth = 0
        .Rows = 2
        .cols = cols
        .ColAlignment(0) = 1
        Select Case box.name
            Case "NEWconditionBOX"
                .ColAlignment(0) = 3
        End Select
        For i = 0 To cols - 1
            .TextMatrix(0, i) = title(i)
            .ColWidth(i) = wide(i)
            totalwidth = totalwidth + wide(i)
        Next
        
        .Height = 2340
        extraW = 270
        .ScrollBars = flexScrollBarVertical
        If (box.width) > (totalwidth + extraW) Then
            .width = box.width
            .ColWidth(0) = .ColWidth(0) + (.width - totalwidth) - extraW
        Else
            .width = totalwidth + extraW
        End If
        .tag = Format(Index, "00") + box.name
        
        n = 1
        Do While Not datax.EOF
            rec = ""
            For i = 0 To cols - 1
                rec = rec + Trim(Format(datax(list(i))))
                If i < (datax.Fields.Count - 1) Then
                    rec = rec + vbTab
                End If
            Next
            .addITEM rec
            If datax(clue) = box.tag Then
                Flag = .Rows - 1
            End If
            If n = 6 And datax.RecordCount > 10 Then
                Call showGRID(grid, Index, box, True)
                Screen.MousePointer = 11
                .RemoveItem (1)
                grid.Refresh
            End If
            datax.MoveNext
            n = n + 1
        Loop
        If datax.RecordCount <= 10 Then
            .RemoveItem (1)
            If Flag > 1 Then Flag = Flag - 1
        End If
        .row = Flag
        .RowHeightMin = 240
        If .Rows < 6 Then
            .Height = 300 * (.Rows + 1)
            .width = .width - extraW
            .ScrollBars = flexScrollBarNone
        End If
    End With
    Screen.MousePointer = 0
End Sub

Sub fillCOMBO(ByRef grid As MSHFlexGrid, Index)
On Error Resume Next
Dim paraVECTOR, sql
Dim i, n, Params, shot, x, spot, rec, list, list2, size, totalwidth, extraW, align, clue
Dim datax As New ADODB.Recordset
Dim addCOMBO As Boolean
    Err.Clear
    With combo(Index)
        totalwidth = 0
        .Rows = 2
        .cols = matrix.TextMatrix(1, Index)
        Call doARRAYS("s", matrix.TextMatrix(8, Index), list)
        Call doARRAYS("n", matrix.TextMatrix(9, Index), size)
        Call doARRAYS("n", matrix.TextMatrix(5, Index), align)
        n = 0
        For i = 0 To matrix.TextMatrix(1, Index) - 1
            .TextMatrix(0, i) = list(i)
            .TextMatrix(1, i) = ""
            .ColWidth(i) = size(i)
            .ColAlignment(i) = align(i)
            totalwidth = totalwidth + size(i)
        Next
        list = ""
    End With
    
    Err.Clear
    clue = matrix.TextMatrix(0, Index)
    Select Case clue
        Case "WarehouseIssue"
            
        Case "Get_Location2"
            Params = matrix.TextMatrix(6, Index)
            Call doARRAYS("s", Params, list)
            Call doARRAYS("s", matrix.TextMatrix(2, Index), list2)
            n = UBound(list)
            
            For i = 0 To n
                If Params = "" Then
                    Set datax = GetLocation(nameSP, "BASE", cell(1).tag, cn, False)
                    addCOMBO = True
                Else
                    If list(n) = "TRUE" Then
                        If i < n Then
                            Set datax = GetLocation(nameSP, Format(list(i)), cell(1).tag, cn, True)
                            addCOMBO = True
                        Else
                            addCOMBO = False
                        End If
                    Else
                        Set datax = GetLocation(nameSP, Format(list(i)), cell(1).tag, cn, False)
                        addCOMBO = True
                    End If
                End If
                If addCOMBO Then
                    If datax.RecordCount > 0 Then
                        datax.Sort = "loc_name"
                        Call doCOMBO(Index, datax, list2, totalwidth)
                    End If
                End If
            Next
            Exit Sub
        Case "query"
            sql = "SELECT po_ponumb, PO_Date, po_buyr, po_sendby, po_apprby, po_stas, po_freigforwr FROM PO INNER JOIN STATUS ON " _
                & "po_stas = sts_code AND po_npecode = sts_npecode WHERE " _
                & "po_stas IN ('OP') AND " _
                & "po_npecode = '" + nameSP + "' AND " _
                & "po_compcode = '" + cell(1).tag + "' AND " _
                & "po_invloca = '" + cell(3).tag + "' AND " _
                & "po_docutype IN ('P', 'O', 'L', 'W', 'S', 'PO', 'C', 'E') AND " _
                & "((po_freigforwr=1 and  po_stasdelv in('RP','RC')) or (po_freigforwr=0) and po_stasinvt <> 'IC') " _
                + "order by po_creadate desc"   'Juan 2014-09-09
            datax.Open sql, cn, adOpenForwardOnly
        Case "suppliers"
            sql = "SELECT sup_code, sup_name FROM supplier WHERE " _
                & "sup_npecode = '" & nameSP & "' AND " _
                & "sup_actvflag = 1 " _
                & "ORDER BY sup_name, sup_npecode"
            datax.Open sql, cn, adOpenForwardOnly
        Case "AdjustmentEntry"
            sql = "SELECT loc_name, loc_locacode FROM Location " _
                & "WHERE loc_actvflag = 1 AND (UPPER(loc_gender) = 'BASE') AND " _
                & "(UPPER(loc_npecode) = '" & nameSP & "') AND " _
                & "loc_compcode = '" & cell(1).tag & "' " _
                & "ORDER BY loc_name "
            datax.Open sql, cn, adOpenForwardOnly
        Case Else
            Params = matrix.TextMatrix(6, Index)
            If Params <> "" Then If Len(Params) = 0 Then Exit Sub
            If Err.Number = 0 Then
                n = howMANY(matrix.TextMatrix(6, Index), ",")
                ReDim paraVECTOR(n)
                paraVECTOR(0) = ""
                For i = 0 To n
                    x = InStr(Params, ",") - 1
                    If x < 0 Then x = Len(Params)
                    spot = Trim(Left(Params, x))
                    If Left(spot, 1) = "@" Then
                        If UCase(Left(spot, 5)) = "@CELL" Then
                            spot = cell(Val(Mid(spot, 7, 1))).tag
                        Else
                            spot = cell(Val(Mid(spot, 2, 1)))
                        End If
                    End If
                    paraVECTOR(i) = Trim(spot)
                    If InStr(Params, ",") > 0 Then
                        Params = Mid(Params, x + 2)
                    End If
                Next
                Set datax = getDATA(clue, paraVECTOR)
                Err.Clear
            End If
    End Select
            
    If datax.RecordCount < 1 Then Exit Sub
    Call doARRAYS("s", matrix.TextMatrix(2, Index), list)
    Call doCOMBO(Index, datax, list, totalwidth)
    Set datax = New ADODB.Recordset
End Sub

Sub getLINEitems(transaction As String)
Dim dataPO As New ADODB.Recordset
Dim sql, rowTEXT, stock As String
Dim i As Integer
Dim qty As Double

    On Error Resume Next
    Screen.MousePointer = 11
    Call makeLists
    If transaction = "*" Then
        sql = "SELECT * from PO_Details_For_transaction WHERE NAMESPACE = '" + nameSP + "' " _
            & "AND PO = '" + cell(0) + "' ORDER BY PO, CONVERT(integer, LineItem)"
    Else
        transaction = Trim(transaction)
        sql = "SELECT * from transaction_Details WHERE NAMESPACE = '" + nameSP + "' " _
            & "AND PO = '" + cell(0) + "' AND transaction = '" + transaction + "' ORDER BY PO, CONVERT(integer, LineItem)"
    End If
    STOCKlist.RowHeightMin = 0
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, cn, adOpenForwardOnly
    If Err.Number <> 0 Then Exit Sub
    With dataPO
        If .RecordCount > 0 Then
            Do While Not .EOF
                rowTEXT = "" + vbTab
                rowTEXT = rowTEXT + IIf(IsNull(!LineItem), "", !LineItem) + vbTab 'PO Line Item
                stock = IIf(IsNull(!StockNumber), "", Trim(!StockNumber)) + " - " + IIf(IsNull(!description), "", !description)
                rowTEXT = rowTEXT + stock + vbTab 'Stock Number + Description
                rowTEXT = rowTEXT + "" + vbTab 'Line
                
                'Purchase
                rowTEXT = rowTEXT + FormatNumber(!qty1, 2) + vbTab 'Primary Quantity
                rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice1), 0, !UnitPrice1), 2) + vbTab 'Primary Unit Price
                
                'transaction
                rowTEXT = rowTEXT + "" + vbTab 'Line
                If transaction = "*" Then
                    If IsNumeric(!SumQty1) Then
                        qty = !SumQty1
                    Else
                        qty = 0
                    End If
                    rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Sumary Primary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!SumUnitPrice1), "", !SumUnitPrice1), 2) + vbTab 'Sumary Primary Unit Price
                Else
                    If IsNumeric(!QuantityI1) Then
                        qty = !QuantityI1
                    Else
                        qty = 0
                    End If
                    rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Primary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPriceI1), 0, !UnitPriceI1), 2) + vbTab 'Primary Unit Price
                End If
                
                STOCKlist.addITEM rowTEXT
                STOCKlist.row = STOCKlist.Rows - 1
                STOCKlist.TextMatrix(STOCKlist.row, 16) = !Unit1Code
                STOCKlist.TextMatrix(STOCKlist.row, 17) = IIf(IsNull(!transactions), 0, !transactions)
                Call colorCOLS
                Call differences(STOCKlist.row)
                If !unit1 = !unit2 Then
                    STOCKlist.TextMatrix(STOCKlist.row, 15) = ""
                Else
                    STOCKlist.TextMatrix(STOCKlist.row, 15) = !UnitSwitch
                    STOCKlist.RowHeight(STOCKlist.row) = 240
                    rowTEXT = "" + vbTab + "" + vbTab + "" + vbTab
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    
                    'Purchase
                    rowTEXT = rowTEXT + FormatNumber(!qty2, 2) + vbTab 'Secundary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Secundary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!unitPRICE2), 0, !unitPRICE2), 2) + vbTab 'Secundary Unit Price
                    
                    'transaction
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    If transaction = "*" Then
                        If IsNumeric(!SumQty2) Then
                            qty = !SumQty2
                        Else
                            qty = 0
                        End If
                        rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Sumary Primary Quantity
                        rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!SumUnitPrice2), "", !SumUnitPrice2), 2) + vbTab 'Sumary Primary Unit Price
                    Else
                        If IsNumeric(!QuantityI2) Then
                            qty = !QuantityI2
                        Else
                            qty = 0
                        End If
                        rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Primary Quantity
                        rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPriceI2), 0, !UnitPriceI2), 2) + vbTab 'Primary Unit Price
                    End If
                    
                    STOCKlist.addITEM rowTEXT
                    STOCKlist.row = STOCKlist.Rows - 1
                    STOCKlist.TextMatrix(STOCKlist.row, 15) = !UnitSwitch
                    STOCKlist.TextMatrix(STOCKlist.row, 16) = !Unit2Code
                    STOCKlist.TextMatrix(STOCKlist.row, 17) = IIf(IsNull(!transactions), 0, !transactions)
                    Call colorCOLS
                    STOCKlist.col = 1
                    STOCKlist = "?"
                    STOCKlist.CellFontName = "Wingdings"
                    'stocklist.CellFontSize = 8
                    Call differences(STOCKlist.row)
                    If UCase(Trim(!UnitSwitch)) = "P" Or IsNull(!UnitSwitch) Then STOCKlist.row = STOCKlist.Rows - 2
                    For i = 4 To 6
                        STOCKlist.col = i
                        STOCKlist.CellBackColor = &HC0C0FF
                    Next
                    
                    STOCKlist.row = STOCKlist.Rows - 1
                End If
                
                STOCKlist.RowHeight(STOCKlist.row) = 240
                STOCKlist.addITEM ""
                STOCKlist.row = STOCKlist.Rows - 1
                For i = 0 To STOCKlist.cols - 1
                    STOCKlist.col = i
                    If i = 0 Then
                        STOCKlist.CellBackColor = &H808080
                    Else
                        STOCKlist.CellBackColor = &HE0E0E0
                    End If
                Next
                STOCKlist.RowHeight(STOCKlist.row) = 50
                STOCKlist.TextMatrix(STOCKlist.row, 13) = 50
                .MoveNext
            Loop
            STOCKlist.RemoveItem (1)
            STOCKlist.RemoveItem (STOCKlist.Rows - 1)
            STOCKlist.row = 0
        End If
    End With
    Screen.MousePointer = 0
End Sub


Sub gridLIST(ByVal mainGRID As MSHFlexGrid, ByVal childGRID As MSHFlexGrid)
Dim h, i As Integer
    
    With childGRID
        .Left = mainGRID.Left + mainGRID.ColWidth(0)
        h = 20
        For i = 0 To mainGRID.row
            h = h + mainGRID.RowHeight(i)
        Next
        .Top = h + mainGRID.Top - 30
        .Visible = True
        .SetFocus
    End With
End Sub

Sub gridONfocus(ByRef grid As MSHFlexGrid)
Dim i, x As Integer
    With grid
        x = .col
        For i = 0 To .cols - 1
            .col = i
            .CellBackColor = &H800000   'Blue
            .CellForeColor = &HFFFFFF   'White
        Next
        .col = x
        .tag = .row
    End With
End Sub

Sub lockDOCUMENT(locked As Boolean)
Dim i As Integer
    For i = 1 To 5
        If locked Then
            cell(i).locked = True
            cell(0).locked = False
        Else
            cell(i).locked = False
            cell(0).locked = True
        End If
    Next
End Sub

Sub makeLists()
Dim i, col, c, dark As Integer
Dim translationCommodity, translationUnit, translationUnitPrice, translationDescription
Dim translationQty, translationSerial, translationPurchaseQty, translationQtyToRec
Dim translationPrimaryUnit, translationSecondaryUnit, translationOriginal, translationItem
Dim translationCondition, translationLogicalWarehouse, translationBalance, translationSublocation
Dim translationFrom, translationTo, translationLogical, translationNewCond, translationNewConditionDescription
Dim translationSecondaryQty

    For i = 0 To 4
        If cell(i).Visible Then cell(i).tabindex = i
    Next
    STOCKlist.tabindex = 5
    Tree.tabindex = 6
    translationCommodity = translator.getIt("translationCommodity") + ": "
    translationUnit = translator.getIt("translationUnit")
    translationUnitPrice = translator.getIt("translationUnitPrice")
    translationDescription = translator.getIt("translationDescription")
    translationQty = translator.getIt("translationQty")
    translationSerial = translator.getIt("translationSerial")
    translationPurchaseQty = translator.getIt("translationPurchaseQty")
    translationQtyToRec = translator.getIt("translationQtyToRec")
    translationPrimaryUnit = translator.getIt("translationPrimaryUnit")
    translationSecondaryUnit = translator.getIt("translationSecondaryUnit")
    translationOriginal = translator.getIt("translationOriginal")
    translationItem = translator.getIt("translationItem")
    translationCondition = translator.getIt("translationCondition")
    translationLogicalWarehouse = translator.getIt("translationLogicalWarehouse")
    translationBalance = translator.getIt("translationBalance")
    translationSublocation = translator.getIt("translationSublocation")
    translationFrom = translator.getIt("translationFrom")
    translationTo = translator.getIt("translationTo")
    translationLogical = translator.getIt("translationLogical")
    translationNewCond = translator.getIt("translationNewCond")
    translationNewConditionDescription = translator.getIt("translationNewConditionDescription")
    translationSecondaryQty = translator.getIt("translationSecondaryQty")
    
    dark = 1
    With STOCKlist
        .width = 12615 + 1500 'Juan 2010-5-9
        .Clear
        .Rows = 2
        .ColWidth(0) = 485
        .row = 0
        .col = 0
        .TextMatrix(0, 0) = "#"
        '.TextMatrix(0, 1) = "Commodity"
        .TextMatrix(0, 1) = translationCommodity '2015/03/24
        .ColWidth(1) = 1400
        For i = 1 To .cols - 1
            .ColAlignment(i) = 0
            .ColAlignmentFixed(i) = 4
        Next
        .ColAlignment(2) = 6
        Select Case frmWarehouse.tag
            'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales
            Case "02040400", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                .cols = 7
                dark = 1
                '.TextMatrix(0, 2) = "Unit Price"
                .TextMatrix(0, 2) = translationUnitPrice
                .ColWidth(2) = 1000
                '.TextMatrix(0, 3) = "Description"
                .TextMatrix(0, 3) = translationDescription
                .ColWidth(3) = 6200
                '.TextMatrix(0, 4) = "Unit"
                .TextMatrix(0, 4) = translationUnit
                .ColWidth(4) = 1200
                .ColAlignment(5) = 6
                '.TextMatrix(0, 5) = "Qty"
                .TextMatrix(0, 5) = translationQty
                .ColWidth(5) = 1200
                .ColWidth(6) = 0
            'WarehouseIssue 2012-3-23 to add serial
            Case "02040200"
                .cols = 8
                dark = 1
                '.TextMatrix(0, 2) = "Serial"
                .TextMatrix(0, 2) = translationSerial
                .ColWidth(2) = 1000
                '.TextMatrix(0, 3) = "Unit Price"
                .TextMatrix(0, 3) = translationUnitPrice
                .ColWidth(3) = 1000
                '.TextMatrix(0, 4) = "Description"
                .TextMatrix(0, 4) = translationDescription
                .ColWidth(4) = 6200
                '.TextMatrix(0, 5) = "Unit"
                .TextMatrix(0, 5) = translationUnit
                .ColWidth(5) = 1200
                .ColAlignment(6) = 6
                '.TextMatrix(0, 6) = "Qty"
                .TextMatrix(0, 6) = translationQty
                .ColWidth(6) = 1200
                .TextMatrix(0, 7) = "original qty"
                .ColWidth(7) = 0
            Case "02050200" 'AdjustmentEntry
                dark = 0
                .cols = 4
                '.TextMatrix(0, 2) = "Description"
                .TextMatrix(0, 2) = translationDescription
                .ColAlignment(2) = 0
                .ColWidth(2) = 8400
                '.TextMatrix(0, 3) = "Unit"
                .TextMatrix(0, 3) = translationUnit
                .ColWidth(3) = 1200
            Case "02040100" 'WarehouseReceipt
                dark = 1
                .cols = 14
                For i = 8 To .cols - 1
                    .ColWidth(i) = 0
                Next
                'Juan 2010-6-6
                For i = 1 To .cols - 1
                    .col = i
                    .CellFontName = "Arial"
                    .CellFontSize = 7
                    .ColAlignment(i) = 0
                    '.ColAlignmentFixed(i) = 4
                Next
                '-----------------------

                .ColAlignment(2) = 6
                .ColAlignment(3) = 6
                .ColAlignment(4) = 4
                .ColAlignment(6) = 4
                '.TextMatrix(0, 2) = "Purchase Qty"
                .TextMatrix(0, 2) = translationPurchaseQty
                .ColWidth(2) = 1100
                '.TextMatrix(0, 3) = "1Qty to Rec"
                .TextMatrix(0, 3) = translationQtyToRec
                .ColWidth(3) = 1100
                '.TextMatrix(0, 4) = "Primary Unit"
                .TextMatrix(0, 4) = translationPrimaryUnit
                .ColWidth(4) = 1100
                
                
                'Juan 2010-6-6
                '.TextMatrix(0, 5) = "Description"
                '.ColWidth(5) = 6200
                '.TextMatrix(0, 6) = "Item #"
                '.ColWidth(6) = 0
                '.TextMatrix(0, 5) = "2Qty to Rec"
                .TextMatrix(0, 5) = "Sec " + translationQtyToRec
                .ColWidth(5) = 1300
                '.TextMatrix(0, 6) = "Secundary Unit"
                .TextMatrix(0, 6) = translationSecondaryUnit
                .ColWidth(6) = 1100
                '.TextMatrix(0, 7) = "Description"
                .TextMatrix(0, 7) = translationDescription
                .ColWidth(7) = 6200
                '.TextMatrix(0, 8) = "Item #"
                .TextMatrix(0, 8) = translationItem + "#"
                .ColWidth(8) = 0
                '.TextMatrix(0, 9) = "original Qty to Rec"
                .TextMatrix(0, 9) = translationOriginal + " " + translationQtyToRec
                .ColWidth(9) = 0
                '.TextMatrix(0, 10) = "original 2Qty to Rec"
                .TextMatrix(0, 10) = translationOriginal + " 2" + translationQtyToRec
                .ColWidth(10) = 0
                '---------------------
        End Select
        .RowHeight(0) = 240
        .RowHeightMin = 0
        .RowHeight(1) = 0
        '.WordWrap = True
        .tag = ""
    End With
    
    With detailHEADER
        .width = STOCKlist.width 'Juan 2010-5-9
        .cols = 7
        c = 7
        .ColWidth(0) = 4800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1900
        .ColWidth(3) = 1900
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 260
        Select Case frmWarehouse.tag
            Case "02040100" 'WarehouseReceipt
                .cols = 9
                For i = 1 To .cols - 1
                    .col = i
                    .CellFontName = "Arial"
                    .CellFontSize = 7
                Next
                c = 9
                '.TextMatrix(0, 0) = "Condition / Serial-Pool"
                .TextMatrix(0, 0) = translationCondition + " / " + translationSerial + "-Pool"
                '.TextMatrix(0, 1) = "Qty po"
                .TextMatrix(0, 1) = translationQty + " PO"
                '.TextMatrix(0, 2) = "Logical Warehouse"
                .TextMatrix(0, 2) = translationLogicalWarehouse
                '.TextMatrix(0, 3) = "Sublocation"
                .TextMatrix(0, 3) = translationSublocation
                'Juan 2010-6-6
                '.TextMatrix(0, 4) = "Prim Unit"
                .TextMatrix(0, 4) = translationPrimaryUnit
                '.TextMatrix(0, 5) = "Qty"
                .TextMatrix(0, 5) = translationQty
                '.TextMatrix(0, 6) = "Sec Unit"
                .TextMatrix(0, 6) = translationSecondaryUnit
                '.TextMatrix(0, 7) = "Qty"
                .TextMatrix(0, 7) = translationSecondaryQty
                '.TextMatrix(0, 8) = "Balance"
                .TextMatrix(0, 8) = translationBalance
                '---------------------
            .ColWidth(0) = 3970
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
            .ColWidth(8) = 1000
            .ColWidth(9) = 260
            Case Else
                '.TextMatrix(0, 0) = "Condition / Logical Warehouse / Sublocation"
                .TextMatrix(0, 0) = translationCondition + " / " + translationLogicalWarehouse + " / " + translationSublocation
                '.TextMatrix(0, 1) = "Qty"
                .TextMatrix(0, 1) = translationQty
                '.TextMatrix(0, 2) = "Logical Warehouse"
                .TextMatrix(0, 2) = translationLogicalWarehouse
                '.TextMatrix(0, 3) = "Sublocation"
                .TextMatrix(0, 3) = translationSublocation
                '.TextMatrix(0, 4) = "Qty"
                .TextMatrix(0, 4) = translationQty
                '.TextMatrix(0, 5) = "Balance"
                .TextMatrix(0, 5) = translationBalance
                .ColWidth(2) = 2500 + 480 'Juan 2014-09-13 it was 2230 but made it longer
                .ColWidth(3) = 2500 + 480 'Juan 2014-09-13 it was 2230 but made it longer
        End Select

        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                .cols = 9
                c = 9
                '.TextMatrix(0, 2) = "Logical Ware."
                .TextMatrix(0, 2) = translationLogicalWarehouse
                '.TextMatrix(0, 4) = "Condition"
                .TextMatrix(0, 4) = translationCondition
                .TextMatrix(0, 5) = "Repair Cost"
                '.TextMatrix(0, 6) = "Qty"
                .TextMatrix(0, 6) = translationQty
                '.TextMatrix(0, 7) = "Balance"
                .TextMatrix(0, 7) = translationBalance
                .ColWidth(1) = 950
                .ColWidth(2) = 1040
                .ColWidth(3) = 1040
                .ColWidth(4) = 800
                .ColWidth(5) = 950
                .ColWidth(6) = 950
                .ColWidth(7) = 950
                .ColWidth(8) = 260
            Case "02050200" 'AdjustmentEntry
                .cols = 7
                c = 6
                '.TextMatrix(0, 0) = "Condition"
                .TextMatrix(0, 0) = translationCondition
                '.TextMatrix(0, 1) = "Logical Warehouse"
                .TextMatrix(0, 1) = translationLogicalWarehouse
                '.TextMatrix(0, 2) = "Sublocation"
                .TextMatrix(0, 2) = translationSublocation
                '.TextMatrix(0, 3) = "Unit Price"
                .TextMatrix(0, 3) = translationUnitPrice
                '.TextMatrix(0, 4) = "Condition"
                .TextMatrix(0, 4) = translationCondition
                '.TextMatrix(0, 5) = "Qty"
                .TextMatrix(0, 5) = translationQty
                .TextMatrix(0, 6) = ""
                .ColWidth(0) = 5000
                .ColWidth(1) = 2800 'Juan 2014-01-02 it was 2230 but made it longer
                .ColWidth(2) = 2800 'Juan 2014-01-02 it was 2230 but made it longer
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 1000
                .ColWidth(6) = 260
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
                .cols = 7
                '.TextMatrix(0, 4) = "Condition"
                .TextMatrix(0, 4) = translationCondition
                '.TextMatrix(0, 5) = "Qty"
                .TextMatrix(0, 5) = translationQty
                '.TextMatrix(0, 6) = "Balance"
                .TextMatrix(0, 6) = translationBalance
                .ColWidth(0) = 4400
                .ColWidth(2) = 2630 'Juan 2014-09-13 it was 2230 but made it longer
                .ColWidth(3) = 2630 'Juan 2014-09-13 it was 2230 but made it longer
                .ColWidth(4) = 1000
                .ColWidth(5) = 1000
                .ColWidth(6) = 1000
                .ColWidth(7) = 260
        End Select
        .row = 0
        For i = 1 To c - 1
            .ColAlignmentFixed(i) = 3
            If i > dark Then
                .col = i
                .CellBackColor = &H808080
                .CellForeColor = vbWhite
            End If
        Next
    End With
    
    With SUMMARYlist
        .width = STOCKlist.width 'Juan 2010-5-9
        .Height = newBUTTON.Top - .Top - 100 'Juan 2011-5-7
        .cols = 26
        .Clear
        .Rows = 2
        .ColWidth(0) = 285
        .row = 0
        .col = 0
        .TextMatrix(0, 0) = "#"
        For i = 1 To .cols - 1
            .ColAlignment(i) = 0
            .ColAlignmentFixed(i) = 4
        Next
        .ColAlignment(4) = 6
        .ColAlignment(7) = 6
        .ColAlignment(23) = 6
        '.TextMatrix(0, 1) = "Commodity"
        .TextMatrix(0, 1) = translationCommodity
        .ColWidth(1) = 1400
        '.TextMatrix(0, 2) = "Serial"
        .TextMatrix(0, 2) = translationSerial
        .ColWidth(2) = 800
        '.TextMatrix(0, 3) = "Condition"
        .TextMatrix(0, 3) = translationCondition
        .ColWidth(3) = 1000
        '.TextMatrix(0, 4) = "Prim. Unit Price"
        .TextMatrix(0, 4) = translationUnitPrice
        .ColWidth(4) = 1200
        '.TextMatrix(0, 5) = "Description"
        .TextMatrix(0, 5) = translationDescription
        .ColWidth(5) = 4400
        '.TextMatrix(0, 6) = "Unit"
        .TextMatrix(0, 6) = translationUnit
        .ColWidth(6) = 1100
        '.TextMatrix(0, 7) = "Qty"
        .TextMatrix(0, 7) = translationQty
        .ColWidth(7) = 1200
        .TextMatrix(0, 8) = "node"
        '.TextMatrix(0, 9) = "From Logical"
        .TextMatrix(0, 9) = translationFrom + " " + translationLogical
        '.TextMatrix(0, 10) = "From Subloca"
        .TextMatrix(0, 10) = translationFrom + " " + translationSublocation

        '.TextMatrix(0, 11) = "To Logical"
        .TextMatrix(0, 11) = translationTo + " " + translationLogical
        '.TextMatrix(0, 12) = "To Subloca"
        .TextMatrix(0, 12) = translationTo + " " + translationSublocation

        '.TextMatrix(0, 13) = "New Cond."
        .TextMatrix(0, 13) = translationNewCond
        
        '.TextMatrix(0, 14) = "New Condition Description"
        .TextMatrix(0, 14) = translationNewConditionDescription
        
        '.TextMatrix(0, 15) = "Unit Code"
        .TextMatrix(0, 15) = translationUnit
        
        .TextMatrix(0, 16) = "Computer Factor"
        .TextMatrix(0, 20) = "Original Condition Code"
        '.TextMatrix(0, 21) = "Secundary Qty"  'It will be just secondary unit and not 'qty'
        .TextMatrix(0, 21) = translationSecondaryUnit
        '.TextMatrix(0, 22) = "Po Item"
        .TextMatrix(0, 22) = "PO " + translationItem
        '.TextMatrix(0, 23) = "Unit 2"
        '.TextMatrix(0, 23) = "Secundary Qty"
        .TextMatrix(0, 23) = translationSecondaryQty
        .TextMatrix(0, 25) = "ratio"
        c = 8
        For i = c To .cols
            .ColWidth(i) = 0
        Next
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                .TextMatrix(0, 17) = "repaircost"
                .TextMatrix(0, 18) = "newcomodity"
                .TextMatrix(0, 19) = "newdescription"
                'Juan 2010-11-1 new arrangement to show the new condition
                .ColWidth(14) = 2000
                .TextMatrix(0, 3) = "Old Cond."
                '---------------------------
            Case "02050200" 'AdjustmentEntry
                'Juan 2010-11-20 new arrangement to show the new condition
                .ColWidth(14) = 2000
                .ColWidth(3) = 0
                '---------------------------
            Case "02040200" 'WarehouseIssue
                .ColWidth(12) = 1200
            Case "02040500" 'WellToWell
                .TextMatrix(0, 17) = "originalcondition"
            Case "02040700" 'InternalTransfer
                .ColWidth(11) = 1200
                '.TextMatrix(0, 11) = "Logical WH"
                .TextMatrix(0, 11) = translationLogicalWarehouse
                .ColWidth(12) = 1200
                '.TextMatrix(0, 12) = "Sub Location"
                .TextMatrix(0, 12) = translationSublocation
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                .cols = .cols + 2
                .TextMatrix(0, 17) = "Qty po"
                .TextMatrix(0, 24) = "PO"
                'Juan 2010-8-19
                .ColWidth(21) = 1100
                .ColWidth(23) = 1200
                '------------
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
                'Juan 2010-10-31 new arrangement to show the new condition
                .ColWidth(14) = 2000
                .TextMatrix(0, 3) = "Old Cond."
                '---------------------------
        End Select
        .RowHeight(0) = 240
        .RowHeightMin = 0
        .RowHeight(1) = 0
        .WordWrap = True
        .tag = ""
        .ZOrder
    End With
    
    'This grid is used to store values related with the SUMMARYlist grid in case needed
    With summaryValues
        .cols = 4
        .TextMatrix(0, 0) = "quantities array"
        .TextMatrix(0, 1) = "from sublocations array"
        .TextMatrix(0, 2) = "invoice"
        .TextMatrix(0, 2) = "invoice line item"
    End With
    
    
    'Juan 2014-01-12, resizing treeFrame ------
'    With treeFrame
'        .Left = detailHEADER.ColWidth(0) + Tree.Left
'        .width = detailHEADER.width - .Left - 800
'        .Top = detailHEADER.Top + detailHEADER.Height + 300
'        .Height = Tree.Height - 420
'    End With
    With baseFrame
        .Left = detailHEADER.ColWidth(0) + Tree.Left
        .width = detailHEADER.width - .Left - 800
        .Top = detailHEADER.Top + detailHEADER.Height + 300
        .Height = Tree.Height - 420
    End With
    ' ------------
End Sub


Function Iexists() As Boolean
Dim sql, transaction As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    Iexists = True
    transaction = Trim(cell(0))
    sql = "SELECT inv_invcnumb from transaction WHERE inv_npecode = '" + nameSP + "' " _
        & "AND inv_ponumb = '" + cell(0) + "' AND inv_invcnumb = '" + cell(1) + "'"
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, cn, adOpenForwardOnly
    If Err.Number <> 0 Then
        Iexists = False
        Exit Function
    End If
    If dataPO.RecordCount < 1 Then
        Iexists = False
    End If
End Function


Sub moveBOXES(start As Integer, direction As Integer)
Dim i, n As Integer
On Error Resume Next
    n = Tree.Nodes.Count
    For i = start To n
        logicBOX(i).Top = logicBOX(i).Top + (240 * direction)
        If Err.Number = 0 Then
            sublocaBOX(i).Top = sublocaBOX(i).Top + (240 * direction)
            quantity(i).Top = quantity(i).Top + (240 * direction)
            balanceBOX(i).Top = balanceBOX(i).Top + (240 * direction)
            NEWconditionBOX(i).Top = NEWconditionBOX(i).Top + (240 * direction)
            quantityBOX(i).Top = quantityBOX(i).Top + (240 * direction)
            priceBOX(i).Top = priceBOX(i).Top + (240 * direction)
            unitBOX(i).Top = unitBOX(i).Top + (240 * direction)
            repairBOX(i).Top = repairBOX(i).Top + (240 * direction)
        End If
        Err.Clear
    Next
    quantity(n).Top = quantity(n).Top + (240 * direction)
    quantityBOX(n).Top = quantityBOX(n).Top + (240 * direction)
    balanceBOX(n).Top = balanceBOX(n).Top + (240 * direction)
Err.Clear
End Sub


'
'
'
Sub hideDETAILS(Optional unmark As Boolean, Optional resetStockList As Boolean, Optional isSubmit As Boolean)
    Dim stockListRow  As String
    Dim i As Integer
    Dim selectedStockNumber As String
    'Call ctt.enable(False)
    'Set ctt = Nothing
    selectedStockNumber = commodityLABEL
    stockListRow = findSTUFF(commodityLABEL, STOCKlist, 1)
    'Juan 2010-6-4
    If IsMissing(unmark) Then unmark = True
    
    If unmark Then
        Dim stock
        'TODO.... stocknumber search
        stock = STOCKlist.TextMatrix(STOCKlist.MouseRow, 1)
        Call unmarkRow(stock, True, ctt)
    End If
    inProgress = False
    '---------------------

    SUMMARYlist.Visible = True
    SUMMARYlist.ZOrder
    hideDETAIL.Visible = False
    submitDETAIL.Visible = False
    removeDETAIL.Visible = False
    Label4(0).Visible = False
    Label4(1).Visible = False
    'treeFrame.Visible = False
    baseFrame.Visible = False
    If isReset Then
        isReset = False
    Else
        If isFirstSubmit Then
            If resetStockList Then Call calculationsFlat(selectedStockNumber)
            isFirstSubmit = False
        Else
            
            If isSubmit Then
            Else
                Select Case frmWarehouse.tag
                    Case "02050200" 'AdjustmentEntry
                    Case Else
                        STOCKlist.TextMatrix(STOCKlist.row, 5) = latestStockNumberQty
                End Select
            End If
        End If
    End If
End Sub
Sub reNUMBER(grid As MSHFlexGrid)
Dim i
    With grid
        For i = 1 To .Rows - 1
''            If IsNumeric(.TextMatrix(i, 0)) Or .TextMatrix(i, 0) = "" Then
''                .TextMatrix(i, 0) = Format(i)
''            End If
            If (IsNumeric(.TextMatrix(i, 0)) Or .TextMatrix(i, 0) = "") And Len(Trim(.TextMatrix(i, 1))) > 0 Then 'Code modified by Muzammil
                .TextMatrix(i, 0) = Format(i)
            End If
        Next
    End With
End Sub

Sub search(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word, bigKEY, key
    bigKEY = Trim(UCase(cellACTIVE))
    With gridACTIVE
        If cellACTIVE <> "" Then
            If Not .Visible Then .Visible = True
            .col = column
            .tag = ""
            .RowHeightMin = 0
            .RowHeight(-1) = 0
            .RowHeight(0) = 240
            Do While True
                If InStr(bigKEY, ",") = 0 Then
                    key = Trim(bigKEY)
                    bigKEY = ""
                Else
                    key = Trim(Left(bigKEY, InStr(bigKEY, ",") - 1))
                    bigKEY = Trim(Mid(bigKEY, InStr(bigKEY, ",") + 1))
                End If
                For i = 1 To .Rows - 1
                    If .RowHeight(i) = 0 Then
                        word = Trim(UCase(.TextMatrix(i, column)))
                        If InStr(word, key) > 0 Then
                            .RowHeight(i) = 240
                        End If
                    End If
                Next
                If bigKEY = "" Then Exit Do
            Loop
        Else
            .RowHeight(-1) = 240
        End If
    End With
End Sub





Public Sub setUSER(user As String)
    CurrentUser = user
End Sub

Sub showGRID(ByRef grid As MSHFlexGrid, Index, box As textBOX, Optional noFILLING As Boolean)
Dim n
    With grid
        'juan 2012-1-14 to avoid t he problem when logical warehouse shows up with no reason
        If SUMMARYlist.Visible Then Exit Sub
        If remarks.Visible Then Exit Sub
        If Tree.Visible = False Then Exit Sub
        '------------
        If Not noFILLING Then Call fillGRID(grid, box, Index)
        If .Rows > 0 And .text <> "" Then
            n = box.Left + .width
            If n >= frmWarehouse.width Then
                .Left = box.Left - (n - frmWarehouse.width) - 100
            Else
                .Left = box.Left
            End If
            '.Left = .Left + treeFrame.Left 'Juan 2014-02-04, to move cell
            .Left = .Left + baseFrame.Left 'Juan 2014-02-04, to move cell
            'If (box.Top) < (treeFrame.Height - .Height - 800) Then
            If (box.Top) < (baseFrame.Height - .Height - 800) Then
                .Top = box.Top + box.Height + 10
            Else
                .Top = box.Top - .Height - 10
            End If
            '.Top = .Top + treeFrame.Top + (80 * Index) 'Juan 2014-02-04, to move cell
            .Top = .Top + baseFrame.Top + (80 * Index) 'Juan 2014-02-04, to move cell
            .ZOrder
            .Visible = True
        End If
    End With
End Sub
Sub showCOMBO(ByRef grid As MSHFlexGrid, Index)
    With grid
        Call fillCOMBO(grid, Index)
        If .Rows > 0 And .text <> "" Then
            .Visible = True
            .ZOrder
            If Index < 5 Then .Top = cell(Index).Top + 370
        End If
        .MousePointer = 0
    End With
End Sub

Sub hideREMARKS()
    SSOleDBFQA.Visible = False
    otherLABEL(0).Visible = True
    otherLABEL(1).Visible = True
    Line2.Visible = True
    unitLABEL(0).Visible = True
    commodityLABEL.Visible = True
    descriptionLABEL.Visible = True
    remarksLABEL.Visible = False
    remarks.Visible = False
    SUMMARYlist.Visible = True
    SUMMARYlist.ZOrder
    hideDETAIL.Visible = True
    'juan 2012-1-8 commented the line until edition mode works well
    'removeDETAIL.Visible = True
    submitDETAIL.Visible = True
    Tree.Visible = True 'M
    'sublocaBOX(0).Visible = True ' M 'Juan 2010-6-2
    'grid(2).Visible = True 'M 'Juan 2015-10-15
End Sub

Sub showREMARKS()
    Dim h
    SUMMARYlist.Visible = False
    otherLABEL(0).Visible = False
    otherLABEL(1).Visible = False
    unitLABEL(0).Visible = False
    Line2.Visible = False
    commodityLABEL.Visible = False
    descriptionLABEL.Visible = False
    Command5.Caption = "&Hide Remarks"
    remarks.locked = False
    Tree.Visible = False 'M
    'treeFrame.Visible = False
    baseFrame.Visible = False
    remarks.Top = SSOleDBFQA.Top + SSOleDBFQA.Height + 200   'detailHEADER.Top
    h = Tree.Top - detailHEADER.Top + Tree.Height - SSOleDBFQA.Height
    If h < 0 Then h = Tree.Top - detailHEADER.Top + Tree.Height '- SSOleDBFQA.Height
    remarks.Height = h
    remarksLABEL.Visible = True
    remarks.Visible = True
    remarks.ZOrder
    
    'Juan 2010-6-20
    SSOleDBFQA.width = STOCKlist.width
    remarks.width = STOCKlist.width
    '----------------------
    
    sublocaBOX(0).Visible = False ' M
    summaryLABEL.Top = SUMMARYlist.Top - 240
    hideDETAIL.Visible = False
    removeDETAIL.Visible = False
    submitDETAIL.Visible = False
    SSOleDBFQA.Visible = True 'M
    SSOleDBFQA.ZOrder 'M
    remarks.SetFocus
    
    grid(2).Visible = False 'M
    
End Sub

Sub showTEXTline()
Dim positionX, positionY, i, currentCOL, currentROW As Integer
    With STOCKlist
        currentCOL = .col
        currentROW = .row
        If .TextMatrix(.row, 0) <> "" Then
            If Trim(.TextMatrix(.row, 15)) = "P" Then
                If .TextMatrix(.row, 1) = "?" Then
                    If .col = 10 Then Exit Sub
                End If
            Else
                If .TextMatrix(.row, 1) <> "?" Then
                    If .col = 10 Then Exit Sub
                End If
            End If
                positionX = .Left + 20
                For i = 0 To .col - 1
                    positionX = positionX + .ColWidth(i)
                Next
                positionY = .Top + 20
                For i = .TopRow - 1 To .row - IIf(.TopRow = 1, 1, 0)
                    positionY = positionY + .RowHeight(i)
                Next
                TextLINE.text = .text
                TextLINE.Left = positionX
                TextLINE.width = .ColWidth(.col) + 10
                TextLINE.Top = positionY
                TextLINE.Height = .RowHeight(.row) + 10
                TextLINE.tag = .row
                TextLINE.SelStart = 0
                TextLINE.SelLength = Len(TextLINE.text)
                TextLINE.Visible = True
                TextLINE.SetFocus
        End If
        .col = currentCOL
        .row = currentROW
    End With
End Sub



Sub whitening()
On Error Resume Next
    With logicBOX(currentBOX)
        If .text = "" Then
            .backcolor = &HC0C0FF
        Else
            .backcolor = vbWhite
        End If
    End With
    With sublocaBOX(currentBOX)
        If .text = "" Then
            .backcolor = &HC0C0FF
        Else
            .backcolor = vbWhite
        End If
    End With
    quantityBOX(currentBOX).backcolor = vbWhite
    quantity2BOX(currentBOX).backcolor = vbWhite 'Juan 2010-6-14
    NEWconditionBOX(currentBOX).backcolor = vbWhite
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
            repairBOX(currentBOX).backcolor = vbWhite
        Case "02050200" 'AdjustmentEntry
            priceBOX(currentBOX).backcolor = vbWhite
    End Select
    Err.Clear
End Sub

Private Sub addITEM_Click()
Dim n As Integer
Dim nody As node
    With Tree
        n = .SelectedItem.Index + .SelectedItem.Children
        Call moveBOXES(n, 1)
        .Nodes.Add .SelectedItem.key, tvwChild, .SelectedItem.key + "{{Serial", "Serial ", "thing 1"
        .Nodes(.SelectedItem.key + "{{Serial").Selected = True
        .StartLabelEdit
        'Added by Juan 2010-4-27
        quantityBOX(quantityBOX.Count - 1).text = "1"
        quantityBOX(quantityBOX.Count - 1).Enabled = False
        '----------------------------------------
    End With
End Sub
Private Sub cell_Change(Index As Integer)
Dim n As Integer
    If Not directCLICK Then
        If Index = 4 Or Index = 0 Then
            n = 0
        Else
            n = 1
        End If
        Call alphaSEARCH(cell(Index), combo(Index), n)
    Else
        directCLICK = False
    End If
End Sub

Private Sub combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    justCLICK = False
    With cell(Index)
        If Not .locked Then
            Select Case KeyCode
                Case 27
                    combo(Index).Visible = False
                Case 40
                    Call arrowKEYS("down", Index)
                Case 38
                    Call arrowKEYS("up", Index)
                Case Else
                Dim col
            End Select
        End If
    End With
End Sub

Private Sub combo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call combo_Click(Index)
        Case 27
            combo(Index).Visible = False
            Exit Sub
    End Select
    combo(Index).Visible = False
    If Index > 0 Then
        If Index < 4 Then
            cell(Index + 1).SetFocus
            Call cell_Click(Index + 1)
        Else
            cell(Index).SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub deleteITEM_Click()
    Tree.Nodes.Remove (Tree.SelectedItem.Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim imsLock As imsLock.Lock
    'Unlock
    Call unlockBUNCH
        
    grid1 = True
    grid2 = False
    Set imsLock = New imsLock.Lock
    Call imsLock.Unlock_Row(locked, cn, CurrentUser, frmWarehouse.POrowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
    '------

    Unload frmWarehouse
    GFQAComboFilled = False
    GDefaultValue = False
End Sub

Private Sub logicBOX_Change(Index As Integer)
    Call alphaSEARCH(logicBOX(Index), grid(1), 0)
End Sub

Private Sub NEWconditionBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            grid(0).Visible = False
        Case 27
            grid(0).Visible = False
    End Select
End Sub

Private Sub priceBOX_Change(Index As Integer)
    If noRETURN Then
        noRETURN = False
    Else
        Call priceBOX_Validate(Index, True)
    End If
End Sub

Private Sub priceBOX_Click(Index As Integer)
    With priceBOX(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub


Private Sub priceBOX_GotFocus(Index As Integer)
    Call whitening
    priceBOX(Index).backcolor = &H80FFFF
End Sub

Private Sub priceBOX_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If Err.Number = 6 Then Exit Sub
        Call priceBOX_Validate(Index, True)
        If IsNumeric(priceBOX(Index)) Then
            priceBOX(Index) = Format(priceBOX(Index), "0.00")
        End If
    End If
End Sub

Private Sub priceBOX_LostFocus(Index As Integer)
    priceBOX(Index).backcolor = vbWhite
    If IsNumeric(priceBOX(Index)) Then
        priceBOX(Index) = Format(priceBOX(Index), "0.00")
    End If
End Sub

Private Sub priceBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        priceBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub priceBOX_Validate(Index As Integer, Cancel As Boolean)
    Call validateQTY(priceBOX(Index), Index)
End Sub

Private Sub PrintButton_Click()

End Sub

Private Sub saveBUTTON_Click()
Dim i
Dim retval As Boolean
Dim PrimUnit As Double
Dim NP As String
Dim CompCode As String
Dim stocknumb As String
Dim stockDESC As String
Dim FromWH As String
Dim ToWH As String
Dim fromlogic As String
Dim fromSubLoca As String
Dim toLOGIC As String
Dim toSUBLOCA As String
Dim condition As String
Dim NEWcondition As String
Dim unitPRICE As Double
Dim newUNITprice As Double
Dim serial As String
Dim computerFactor
Dim imsLock As imsLock.Lock
Dim TranType As String
Dim data As New ADODB.Recordset
Dim datax As New ADODB.Recordset
Dim datay As New ADODB.Recordset
Screen.MousePointer = 11
    If treeFrame.Visible = True Then
        Screen.MousePointer = 0
        MsgBox "There is a pending item to submit"
        Exit Sub
    End If

    'MDI_IMS.StatusBar1.Panels(1).text = "Checking fields"
    
    If SUMMARYlist.Rows = 2 And SUMMARYlist.TextMatrix(1, 1) = "" Then
        Screen.MousePointer = 0
        MsgBox "No line Items selected"
        Exit Sub
    End If
    
    'Header
    For i = 1 To 3
        If cell(i) = "" And cell(i).Visible Then
            Screen.MousePointer = 0
            MsgBox "Missing " + label(i)
            cell(i).SetFocus
            Exit Sub
        End If
    Next
        
    If SSOleDBFQA.Rows = 0 Then
    
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please fill out the FQA values for this transaction.", vbCritical, "Ims"
        'remarks.backcolor = &HC0FFFF
        'remarks.SetFocus
        Exit Sub
    End If
    
    If ValidateFromFqa = False Then Screen.MousePointer = 0: Exit Sub
    If ValidateTOFqa = False Then Screen.MousePointer = 0: Exit Sub
    
    If remarks = "" Then
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please include the remarks for this transaction"
        remarks.backcolor = &HC0FFFF
        remarks.SetFocus
        Exit Sub
    End If
    
    
    Call hideREMARKS
        
    'MDI_IMS.StatusBar1.Panels(1).text = "Beginning the process"
    Screen.MousePointer = 11
    savingLABEL.Visible = True
    savingLABEL.ZOrder
    Me.Enabled = False
    Me.Refresh
    
    'cn.BeginTrans
    Set cn = cn
    NP = nameSP
    FromWH = cell(2).tag
    ToWH = cell(3).tag
    CompCode = cell(1).tag
    
    Call BeginTransaction(cn)
    
    
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
            retval = PutReturnData("RR")
            TranType = "RR"
            'Call PutReceiptRemarks
        Case "02050200" 'AdjustmentEntry
            retval = PutReturnData2
            Call InvtReceiptRem_Insert(nameSP, cell(1).tag, cell(2).tag, Format(Transnumb), remarks, CurrentUser, cn)
            TranType = "IA"
        Case "02040200" 'WarehouseIssue
            retval = PutInvtIssue("I")
            Call PutIssueRemarks
            TranType = "I"
        Case "02040500" 'WellToWell
            retval = PutInvtIssue("TI")
            Call PutIssueRemarks
            TranType = "TI"
        Case "02040700" 'InternalTransfer
            retval = PutInvtIssue("IT")
            Call PutIssueRemarks
            TranType = "IT"
        Case "02050300" 'AdjustmentIssue
            retval = PutInvtIssue("AI")
            Call PutIssueRemarks
            TranType = "AI"
        Case "02040600" 'WarehouseToWarehouse
            retval = PutInvtIssue("TI")
            Call PutIssueRemarks
            TranType = "TI"
        Case "02040100" 'WarehouseReceipt
            Transnumb = "R-" & GetTransNumb(nameSP, cn)
            If Err Then GoTo RollBack
            retval = InvtReceipt_Insert(NP, cell(4).tag, "R", cell(1).tag, ToWH, CurrentUser, cn, , FromWH, Format(Transnumb))
            retval = InvtReceiptRem_Insert(NP, CompCode, ToWH, Format(Transnumb), remarks, CurrentUser, cn)
            TranType = "R"
        Case "02050400" 'Sales
            retval = PutInvtIssue("SL")
            Call PutIssueRemarks
            TranType = "SL"
        Case "02040300" 'Return from Well
            retval = PutReturnData("RT")
            Call InvtReceiptRem_Insert(NP, CompCode, ToWH, Format(Transnumb), remarks, CurrentUser, cn)
            TranType = "RT"
    End Select
    If Not retval Then Call RollbackTransaction(cn)
        Screen.MousePointer = 11
        'MDI_IMS.StatusBar1.Panels(1).text = "Saving Line Items"
        For i = 1 To SUMMARYlist.Rows - 1
            stocknumb = SUMMARYlist.TextMatrix(i, 1)
            stockDESC = SUMMARYlist.TextMatrix(i, 5)
            PrimUnit = CDbl(IIf(SUMMARYlist.TextMatrix(i, 7) = "", 0, SUMMARYlist.TextMatrix(i, 7)))
            'Juan 2010-11-24 to obtain original price for AE
            Select Case frmWarehouse.tag
                Case "02040300" 'Return from Well
                    Dim r
                    r = findSTUFF(stocknumb, STOCKlist, 1)
                    If r > 0 Then
                        unitPRICE = CDbl(STOCKlist.TextMatrix(r, 2))
                    Else
                        unitPRICE = 0
                    End If
                Case Else
                    unitPRICE = CDbl(IIf(SUMMARYlist.TextMatrix(i, 4) = "", 0, SUMMARYlist.TextMatrix(i, 4)))
            End Select
            condition = SUMMARYlist.TextMatrix(i, 20)
            fromlogic = SUMMARYlist.TextMatrix(i, 9)
            fromSubLoca = SUMMARYlist.TextMatrix(i, 10)
            toLOGIC = SUMMARYlist.TextMatrix(i, 11)
            toSUBLOCA = SUMMARYlist.TextMatrix(i, 12)
            serial = SUMMARYlist.TextMatrix(i, 2)
            'Juan 2010-9-4 implementing ratio rather than computer factor
            'computerFactor = ImsDataX.ComputingFactor(nameSP, stocknumb, cn)
            'Set datax = getDATA("getStockRatio", Array(nameSP, stocknumb, cell(3).tag))
            'If datax.RecordCount > 0 Then
                'If IsNull(datax!realratio) Or datax!realratio = 0 Then
                    ratioValue = getStockRatioFromStockMaster(nameSP, stocknumb)
                'Else
                    'ratioValue = datax!realratio
                'End If
            'Else
                'ratioValue = getStockRatioFromStockMaster(nameSP, stocknumb)
            'End If
            Dim sql As String
            'Juan 2014-8-27 new version of calculation based on invoice if exists
            Select Case frmWarehouse.tag
                Case "02040100" 'WarehouseReceipt
                    Set datax = New ADODB.Recordset
                    sql = "select * from invoicedetl where invd_npecode = '" + nameSP + "'  " + _
                        "and invd_invcnumb = '" + summaryValues.TextMatrix(i, 2) + "' " + _
                        "and invd_liitnumb = '" + Trim(summaryValues.TextMatrix(i, 3)) + "'"
                    datax.Open sql, cn, adOpenStatic
                    If datax.RecordCount > 0 Then
                        ratioValue = datax!invd_secoreqdqty / datax!invd_primreqdqty
                    End If
                Case Else
            End Select
            datax.Close
            '----------------------
            
            SecUnit = PrimUnit * ratioValue

'            If computerFactor = 0 Then
'                SecUnit = PrimUnit
'            Else
'                If IsNumeric(SUMMARYlist.TextMatrix(i, 21)) Then
'                    SecUnit = CDbl(SUMMARYlist.TextMatrix(i, 21))
'                Else
'                    SecUnit = PrimUnit * 10000 / computerFactor
'                End If
'            End If
            '-----------------
            NEWcondition = SUMMARYlist.TextMatrix(i, 13)
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    retval = PutDataInsert2(i, unitPRICE)
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    Dim repairCOST As Double
                    repairCOST = CDbl(SUMMARYlist.TextMatrix(i, 17))
                    
                    retval = Update_Sap_With_repair_Cost(NP, CompCode, stocknumb, ToWH, PrimUnit, CDbl(1), repairCOST, unitPRICE, NEWcondition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), CDbl(i), ToWH, "RR", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), FromWH, CDbl(i), ToWH, "RR", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), FromWH, "RT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "RT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                Case "02050200" 'AdjustmentEntry
                    ToWH = cell(2).tag
                    retval = PutDataInsert2(i, unitPRICE)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If

                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, CDbl(1), unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "AE", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, Val(serial), ToWH, "AE", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                Case "02040200" 'WarehouseIssue
                    retval = PutDataInsert(i)
                    If retval = False Then Call RollbackTransaction(cn)
                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)
                        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "I", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, CDbl(i), ToWH, "I", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                        
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
            
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), ToWH, "I", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), ToWH, "I", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                Case "02040500" 'WellToWell
                    retval = PutDataInsert(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    condition = SUMMARYlist.TextMatrix(i, 13)
                
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "TI", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), FromWH, "TI", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                Case "02040700" 'InternalTransfer
                    ToWH = FromWH
                    retval = PutDataInsert(i)
                    If retval = False Then Call RollbackTransaction(cn)

                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)

                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "IT", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, CDbl(i), FromWH, "IT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), FromWH, "IT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "IT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If

                Case "02050300" 'AdjustmentIssue
                    retval = PutDataInsert(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                                        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), ToWH, "AI", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), ToWH, "AI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                
                Case "02040600" 'WarehouseToWarehouse
                    retval = PutDataInsert(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "TI", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), FromWH, "TI", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                                        
                Case "02040100" 'WarehouseReceipt
                    retval = PutDataInsert2(i, unitPRICE)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        newBUTTON.Enabled = True
                        saveBUTTON.Enabled = False
                        savingLABEL.Visible = False
                        Command3.Enabled = True
                        Call lockDOCUMENT(True)
                        Me.Enabled = True
                        'Unlock
                        Call unlockBUNCH
                            
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, CDbl(1), unitPRICE, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, CurrentUser, cn)
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), CDbl(i), ToWH, "R", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), FromWH, CDbl(i), ToWH, "R", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    'Unlock
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(locked, cn, CurrentUser, frmWarehouse.POrowguid)  'jawdat
                    '------
                
                Case "02050400" 'Sales
                    retval = PutDataInsert(i)
                    If retval = False Then Call RollbackTransaction(cn)
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                                        
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                     
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), ToWH, "SL", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), FromWH, CDbl(i), ToWH, "SL", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                
                Case "02040300" 'Return from Well
                    If condition = NEWcondition Then
                        newUNITprice = unitPRICE
                    Else
                        Set data = getDATA("conditionVALUE", Array(NP, unitPRICE, NEWcondition))
                        If data.RecordCount = 0 Then
                            Call RollbackTransaction(cn)
                            MsgBox "Error in Transaction"
                            Exit Sub
                        Else
                            newUNITprice = CDbl(data(0))
                        End If
                    End If
                    retval = PutDataInsert2(i, newUNITprice)
                    If retval = False Then Call RollbackTransaction(cn)
                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, newUNITprice, NEWcondition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, CurrentUser, cn)
                                            
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), CDbl(i), ToWH, "RT", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), FromWH, CDbl(i), FromWH, "RT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    SecUnit = SecUnit * -1
                    PrimUnit = PrimUnit * -1
                    
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), FromWH, "RT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "RT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
            End Select
            If retval = False Then
                Call RollbackTransaction(cn)
                MsgBox "Error in Transaction"
                Exit Sub
            End If
        Next
        
    If retval = True Then retval = SaveFQA(Transnumb, TranType)
        
    If retval Then
        Call CommitTransaction(cn)
        If frmWarehouse.tag = "02040100" Then  'WarehouseReceipt
            Dim poSTATUS As ADODB.Command
            Set poSTATUS = getCOMMAND("UPDATE_PO_INVSTATES")
            poSTATUS.parameters(1) = nameSP
            poSTATUS.parameters(2) = cell(4).tag
            poSTATUS.Execute
        End If
        
        'cn.CommitTrans
        cell(0) = Transnumb
        cell(0).tag = cell(0)
        cell(0).Visible = True
        combo(0).Visible = False
        combo(1).Visible = False
        combo(0).TextMatrix(1, 0) = Transnumb
    End If
    Screen.MousePointer = 11
        
    If Err Then Err.Clear
    newBUTTON.Enabled = True
    saveBUTTON.Enabled = False
    savingLABEL.Visible = False
    Command3.Enabled = True
    Call lockDOCUMENT(True)
    Me.Enabled = True
    'Unlock
    Call unlockBUNCH
        
    Screen.MousePointer = 0
    Exit Sub
RollBack:
    Call RollbackTransaction(cn)
    Screen.MousePointer = 0
    Exit Sub
    'MDI_IMS.StatusBar1.Panels(1).text = ""
End Sub

Function getStockRatioFromStockMaster(NameSpace, StockNumber) As Double
Dim data As New ADODB.Recordset
Dim ratio As Double
    Set data = getDATA("getStockRatioFromStockMaster", Array(nameSP, StockNumber))
    If data.RecordCount > 0 Then
        If IsNull(data!realratio) Then
            ratio = 1
        Else
            ratio = data!realratio
        End If
    Else
        ratio = 1
    End If
getStockRatioFromStockMaster = ratio
End Function


Private Function PutDataInsert2(Item, price) As Boolean
    Dim psVALUE, serial
    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert2 = False

    Set cmd = getCOMMAND("INVTRECEIPTDETL_INSERT")

    'Set the parameter values for the command to be executed.
    cmd.parameters("@ird_curr") = "USD"
    cmd.parameters("@ird_currvalu") = 1
    cmd.parameters("@ird_ponumb") = Null
    cmd.parameters("@ird_lirtnumb") = Null
    cmd.parameters("@ird_compcode") = cell(1).tag
    cmd.parameters("@ird_trannumb") = Transnumb
    cmd.parameters("@ird_npecode") = nameSP
    With SUMMARYlist
        Select Case frmWarehouse.tag
            Case "02050200" 'AdjustmentEntry
                cmd.parameters("@ird_ware") = cell(2).tag
            Case "02040100" 'WarehouseReceipt
                cmd.parameters("@ird_ponumb") = .TextMatrix(Item, 24)
                cmd.parameters("@ird_lirtnumb") = Val(.TextMatrix(Item, 22))
                cmd.parameters("@ird_ware") = cell(3).tag
            Case Else
                cmd.parameters("@ird_ware") = cell(3).tag
        End Select
        cmd.parameters("@ird_transerl") = Item
        cmd.parameters("@ird_stcknumb") = .TextMatrix(Item, 1)
        If UCase(.TextMatrix(Item, 2)) = "POOL" Or .TextMatrix(Item, 2) = "" Then
            psVALUE = 1
            serial = Null
        Else
            psVALUE = 0
            serial = .TextMatrix(Item, 2)
        End If
        cmd.parameters("@ird_ps") = psVALUE
        cmd.parameters("@ird_serl") = serial
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                cmd.parameters("@ird_reprcost") = CDbl(.TextMatrix(Item, 17))
                cmd.parameters("@ird_newcond") = .TextMatrix(Item, 13)
                cmd.parameters("@ird_newstcknumb") = .TextMatrix(Item, 18)
                cmd.parameters("@ird_newdesc") = .TextMatrix(Item, 19)
            Case "02050200" 'AdjustmentEntry
                'cmd.Parameters("@ird_newcond") = .TextMatrix(Item, 13)
                cmd.parameters("@ird_newcond") = .TextMatrix(Item, 20) 'M
                'Modified by Muzammil
                'Reason : would not save the condition only to the ird_origcond.
                'The above value begin passed is just an empty string
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                cmd.parameters("@ird_newcond") = "01"
                cmd.parameters("@ird_invoice") = summaryValues.TextMatrix(Item, 2) 'Juan 2014-8-29
                cmd.parameters("@ird_invoiceLine") = summaryValues.TextMatrix(Item, 3) 'Juan 2014-8-30
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
                cmd.parameters("@ird_newcond") = .TextMatrix(Item, 13)
        End Select
        cmd.parameters("@ird_stcktype") = ""
        cmd.parameters("@ird_ctry") = "US"
        cmd.parameters("@ird_tosubloca") = .TextMatrix(Item, 12)
        cmd.parameters("@ird_tologiware") = .TextMatrix(Item, 11)
        cmd.parameters("@ird_owle") = 1
        cmd.parameters("@ird_leasecomp") = Null
        cmd.parameters("@ird_primqty") = CDbl(.TextMatrix(Item, 7))
        cmd.parameters("@ird_secoqty") = SecUnit
        cmd.parameters("@ird_unitpric") = CDbl(price)
        cmd.parameters("@ird_stckdesc") = .TextMatrix(Item, 5)
        cmd.parameters("@ird_fromlogiware") = .TextMatrix(Item, 9)
        cmd.parameters("@ird_fromsubloca") = .TextMatrix(Item, 10)
        
        If Me.tag <> "02050200" Then cmd.parameters("@ird_origcond") = .TextMatrix(Item, 20) 'M
        
        cmd.parameters("@user") = CurrentUser
    End With
    'Execute the command.
    cmd.Execute

    PutDataInsert2 = True

    Exit Function

errPutDataInsert:
    MsgBox Err.description: Err.Clear
End Function


Private Function PutDataInsert(row) As Boolean
Dim cmd As Command
On Error Resume Next
    PutDataInsert = False
    Set cmd = getCOMMAND("InvtIssueDetl_INSERT")

    With SUMMARYlist
        'Set the parameter values for the command to be executed.
        cmd.parameters("@iid_trannumb") = Transnumb
        cmd.parameters("@iid_compcode") = cell(1).tag
        cmd.parameters("@iid_npecode") = nameSP
        cmd.parameters("@iid_ware") = cell(2).tag
        cmd.parameters("@iid_transerl") = .TextMatrix(row, 0)
        cmd.parameters("@iid_stcknumb") = .TextMatrix(row, 1)
        cmd.parameters("@iid_ps") = IIf(.TextMatrix(row, 2) = "", 1, 0)
        cmd.parameters("@iid_serl") = IIf(.TextMatrix(row, 2) = "", Null, .TextMatrix(row, 2))
        
        'cmd.Parameters("@iid_newcond") = .TextMatrix(row, 13) 'M
        'Modified by Muz
        'Reason :  In the older version this Field was NULL only the Orig cond is being Populated.
        'this is in case of AI
        If Me.tag <> "02050300" Then cmd.parameters("@iid_newcond") = .TextMatrix(row, 13) 'M
        
        cmd.parameters("@iid_stcktype") = "I"
        cmd.parameters("@iid_ctry") = "US"
        cmd.parameters("@iid_tosubloca") = .TextMatrix(row, 12)
        cmd.parameters("@iid_tologiware") = .TextMatrix(row, 11)
        cmd.parameters("@iid_owle") = 1
        cmd.parameters("@iid_leasecomp") = Null
        cmd.parameters("@iid_primqty") = CDbl(.TextMatrix(row, 7))
        cmd.parameters("@iid_secoqty") = SecUnit
        cmd.parameters("@iid_unitpric") = CDbl(.TextMatrix(row, 4))
        cmd.parameters("@iid_curr") = "USD"
        cmd.parameters("@iid_currvalu") = 1
        cmd.parameters("@iid_stckdesc") = .TextMatrix(row, 5)
        cmd.parameters("@iid_fromlogiware") = .TextMatrix(row, 9)
        If frmWarehouse.tag = "02040700" Then 'InternalTransfer
            cmd.parameters("@iid_fromsubloca") = .TextMatrix(row, 10)
        Else
            cmd.parameters("@iid_fromsubloca") = .TextMatrix(row, 12)
        End If
        
        cmd.parameters("@iid_origcond") = .TextMatrix(row, 13)
        cmd.parameters("@user") = CurrentUser
    End With
    'Execute the command.
    Call cmd.Execute(Options:=adExecuteNoRecords)
    PutDataInsert = True
End Function

Private Function PutReceiptRemarks() As Boolean
Dim cmd As New ADODB.Command

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceiptRem_Insert"
        .parameters.Append .CreateParameter("@CompanyCode", adChar, adParamInput, 10, cell(1).tag)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, nameSP)
        .parameters.Append .CreateParameter("@WhareHouse", adChar, adParamInput, 10, cell(2).tag)
        .parameters.Append .CreateParameter("@TranNumb", adVarChar, adParamInput, 15, Transnumb)
        .parameters.Append .CreateParameter("@LINENUMB", adInteger, adParamInput, , 1)
        .parameters.Append .CreateParameter("@REMARKS", adVarChar, adParamInput, 7000, remarks)
        .parameters.Append .CreateParameter("@USER", adChar, adParamInput, 20, CurrentUser)
        Call .Execute(, , adExecuteNoRecords)
    End With
    PutReceiptRemarks = cmd.parameters(0).Value = 0
End Function
Private Function PutIssueRemarks() As Boolean
Dim cmd As ADODB.Command

    Set cmd = getCOMMAND("InvtIssuetRem_Insert")
    
    cmd.parameters("@LineNumb") = 1
    cmd.parameters("@REMARKS") = remarks
    cmd.parameters("@TranNumb") = Transnumb
    cmd.parameters("@CompanyCode") = cell(1).tag
    cmd.parameters("@NAMESPACE") = nameSP
    cmd.parameters("@WhareHouse") = cell(2).tag
    cmd.parameters("@USER") = CurrentUser
    
    Call cmd.Execute(0, , adExecuteNoRecords)
    PutIssueRemarks = cmd.parameters(0).Value = 0
End Function

Private Function PutInvtIssue(prefix) As Boolean
Dim NP As String
Dim cmd As Command
On Error GoTo errPutInvtIssue

    PutInvtIssue = False
    Set cmd = getCOMMAND("InvtIssue_Insert")
    NP = nameSP
    Transnumb = prefix + "-" & GetTransNumb(NP, cn)
    cmd.parameters("@NAMESPACE") = NP
    cmd.parameters("@TRANTYPE") = prefix
    cmd.parameters("@COMPANYCODE") = cell(1).tag
    cmd.parameters("@TRANSNUMB") = Transnumb
    cmd.parameters("@ISSUTO") = cell(3).tag
    cmd.parameters("@SUPPLIERCODE") = Null
    Select Case frmWarehouse.tag
        Case "02040500" 'WellToWell
        'Juan 2010-11-30 Modified becasue internal transfer was wrong
        Case "02040700", "02050300" 'InternalTransfer, AdjustmentIssue
            cmd.parameters("@ISSUTO") = cell(2).tag
        Case "02040600" 'WarehouseToWarehouse
        Case "02050400" 'Sales
            cmd.parameters("@SUPPLIERCODE") = cell(4).tag
    End Select
    cmd.parameters("@WHAREHOUSE") = cell(2).tag
    cmd.parameters("@STCKNUMB") = Null
    cmd.parameters("@COND") = Null
    cmd.parameters("@SAP") = Null
    cmd.parameters("@NEWSAP") = Null
    cmd.parameters("@ENTYNUMB") = Null
    cmd.parameters("@USER") = CurrentUser
    cmd.Execute
    PutInvtIssue = cmd.parameters(0).Value = 0
    Exit Function

errPutInvtIssue:
    MsgBox Err.description
    Err.Clear
End Function

Private Sub Command3_Click()
Dim reportPATH, cnSTRING, text
If treeFrame.Visible = True Then
    Screen.MousePointer = 0
    MsgBox "There is a pending item to submit"
    Exit Sub
End If
Screen.MousePointer = 11

    With CrystalReport1
        .Reset
        reportPATH = repoPATH + "\"
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                .ReportFileName = reportPATH & "wareRR.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02050200" 'AdjustmentEntry
                .ReportFileName = reportPATH & "wareAEIA.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040200" 'WarehouseIssue
                .ReportFileName = reportPATH & "wareI.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040500" 'WellToWell
                .ReportFileName = reportPATH & "wareI.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040700" 'InternalTransfer
                .ReportFileName = reportPATH & "wareI.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02050300" 'AdjustmentIssue
                .ReportFileName = reportPATH & "wareI.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040600" 'WarehouseToWarehouse
                .ReportFileName = reportPATH & "wareI.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040100" 'WarehouseReceipt
                .ReportFileName = reportPATH + "wareR.rpt"
                .ParameterFields(1) = "NAMESPACE;" + nameSP + ";TRUE"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
            Case "02050400" 'Sales
                .ReportFileName = reportPATH + "wareSL.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
            Case "02040300" 'Return from Well
                .ReportFileName = reportPATH & "wareRT.rpt"
                .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
                .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
        End Select
        cnSTRING = Split(cn.ConnectionString, ";")
        For Each text In cnSTRING
            Select Case Left(UCase(text), InStr(text, "="))
                Case "PASSWORD="
                    dsnPWD = Mid(text, InStr(text, "=") + 1)
                Case "USER ID="
                    dsnUID = Mid(text, InStr(text, "=") + 1)
                Case "INITIAL CATALOG="
                    dsnDSQ = Mid(text, InStr(text, "=") + 1)
            End Select
        Next
        
        .LogOnServer "pdsodbc.dll", dsnF, dsnDSQ, dsnUID, dsnPWD
        Set thisrepo = CrystalReport1
        mainREPORT = True
        Call Translate_Reports(CrystalReport1.ReportFileName)
        Call Translate_SubReports
        .Action = 1
        .Reset
    End With
Screen.MousePointer = 0
End Sub

Public Sub Command5_Click()
    With Command5
    
''        If newBUTTON.Enabled = True And Len(Trim(cell(0).text)) > 0 Then
''
''            SSOleDBFQA.Top = Tree.Top
''            SSOleDBFQA.Height = 1740 + 2340
''
''         End If
    
        If treeFrame.Visible = True Then
            Screen.MousePointer = 0
            MsgBox "There is a pending item to submit"
            Exit Sub
        End If
        If .Caption = "Show &Remarks, FQA" Then
            .Caption = "Hide &Remarks, FQA"
            showREMARKS
            If GFQAComboFilled = False Then GFQAComboFilled = PopulateCombosWithFQA(cell(1).tag, cell(3).tag)
        Else
            .Caption = "Show &Remarks, FQA"
            SSOleDBFQA.Update
            hideREMARKS
        End If
    End With
    
    
    
End Sub

Private Sub newBUTTON_Click()
Dim i
    nodeONtop = 0
    treeFrame.Top = 0
    treeFrame.Refresh
    baseFrame.Refresh
    isReset = True
    label(0).Visible = False
    cell(0).Visible = False
    Command3.Enabled = False
    emailButton.Enabled = False
    Call cleanSTOCKlist
    Call cleanSUMMARYlist
    Call hideDETAILS
    Line2.Visible = False
    'STOCKlist.Top = 1920
    'STOCKlist.Top = 2080
    STOCKlist.Visible = True
    searchFIELD(0).Visible = True
    searchFIELD(1).Visible = True
    searchButton.Visible = True
    'detailHEADER.Top = 4320
    'Tree.Top = 4560
    'Tree.Height = 3660
    cell(1).SetFocus
    saveBUTTON.Enabled = True
    newBUTTON.Enabled = False
    cell(0).backcolor = &HFFFFC0
    cell(0) = ""
    SUMMARYlist.Rows = 2
    summaryValues.Rows = 2
    For i = 0 To SUMMARYlist.Rows - 1
        SUMMARYlist.TextMatrix(1, i) = ""
    Next
    'SUMMARYlist.Top = 3870
    'SUMMARYlist.Height = 4375
    For i = 1 To 4
        cell(i) = ""
        cell(i).backcolor = vbWhite
        cell(i).locked = False
    Next
    Call hideREMARKS
    Call CleanFQA
    Call ChangeMode(False)
    remarks = ""
    
    'frmWarehouse.Height = 8910
    Call cell_Click(1)
End Sub

Private Sub commodityLABEL_Change()
    Call whitening
End Sub
Private Sub grid_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call grid_Click(Index)
        Case 27
    End Select
End Sub

Private Sub hideDETAIL_Click()
Dim answer, i
    With STOCKlist
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = commodityLABEL Then
                .row = i
                Exit For
            End If
        Next
        Select Case frmWarehouse.tag
            'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales
            Case "02040400", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                .TextMatrix(i, 5) = .TextMatrix(i, 6)
            Case "02050200" 'AdjustmentEntry
            'WarehouseIssue
            Case "02040200"
                .TextMatrix(i, 6) = .TextMatrix(i, 7)
            Case "02040100" 'WarehouseReceipt
        End Select
    End With
    If Tree.Nodes.Count > 0 Then
        If IsNumeric(quantityBOX(totalNode)) Then
            If CDbl(quantityBOX(totalNode)) > 0 Then
                answer = MsgBox("Are you sure you want to lose last changes?", vbYesNo)
                If answer = vbYes Then
                    Call hideDETAILS(True, True, False)
                End If
            Else
                hideDETAILS (True)
            End If
        Else
            hideDETAILS (True)
        End If
    Else
        hideDETAILS (True)
    End If
    For i = 0 To 2
        grid(i).Visible = False
    Next
'
End Sub

Private Sub cell_Click(Index As Integer)
Dim datax As New ADODB.Recordset
Dim sql As String
Dim i
Screen.MousePointer = 11
    With cell(Index)
        Select Case Index
            Case 5
                If saveBUTTON.Enabled Then
                    If Not combo(5).Visible Then
                        Set datax = GetSpecificStockInfo("", nameSP, cn)
                        If datax.RecordCount > 0 Then
                            combo(5).Rows = 2
                            Do While Not datax.EOF
                                combo(5).addITEM Trim(datax!stk_stcknumb)
                                datax.MoveNext
                            Loop
                            Screen.MousePointer = 0
                            combo(5).Visible = True
                            combo(5).ZOrder
                            combo(5).RemoveItem 1
                            combo(5).ColWidth(0) = combo(5).width - 270
                            combo(5).ColAlignment(0) = 0
                            combo(5).TextMatrix(0, 0) = "Stock Number"
                            combo(5).ColAlignmentFixed(0) = 3
                            .tag = .text
                            .text = ""
                            .text = .tag
                            .SelLength = 0
                            .SelStart = Len(.text)
                        End If
                    End If
                End If
                Screen.MousePointer = 0
            Case Else
                If saveBUTTON.Enabled Or Index = 0 Then
                    If Index > 1 Then
                        If combo(Index - 1) = "" Then
                            MsgBox "Please select " + label(Index - 1) + " first"
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                End If
                If Not (saveBUTTON.Enabled And Index = 0) Then
                        Call showCOMBO(combo(Index), Index)
                    End If
                End If
                Screen.MousePointer = 0
        End Select
        .SelStart = 0
        .SelLength = Len(.text)
    End With
Screen.MousePointer = 0
End Sub

Private Sub cell_GotFocus(Index As Integer)
    If saveBUTTON.Enabled Or Index = 0 Then
        If Not (saveBUTTON.Enabled And Index = 0) Then
            With cell(Index)
                .backcolor = &H80FFFF
                .Appearance = 1
                .Refresh
                activeCELL = Index
                .SelLength = Len(.text)
                .SelStart = 0
            End With
        End If
    End If
End Sub

Private Sub cell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    justCLICK = False
    With cell(Index)
        If Not .locked Then
                Select Case KeyCode
                    Case 27
                        combo(Index).Visible = False
                    Case 40
                        Call arrowKEYS("down", Index)
                    Case 38
                        Call arrowKEYS("up", Index)
                    Case Else
                    Dim col
                End Select
        End If
    End With
End Sub
Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i, t, n
Dim gotIT As Boolean
    With cell(Index)
        Select Case KeyAscii
            Case 13
                KeyAscii = 0
                If Not .locked Then
                    justCLICK = False
                    gotIT = False
                    If Index = 4 Or Index = 0 Then
                        n = 0
                    Else
                        n = 1
                    End If
                    t = UCase(combo(Index).TextMatrix(combo(Index).row, n))
                    
                    If UCase(cell(Index)) = Left(t, Len(cell(Index))) Then
                        gotIT = True
                        i = combo(Index).row
                    Else
                        For i = 1 To combo(Index).Rows - 1
                            If UCase(cell(Index)) = UCase(combo(Index).TextMatrix(i, n)) Then
                                gotIT = True
                                Exit For
                            End If
                        Next
                    End If
                    If gotIT Then
                        Call combo_Click(Index)
                    Else
                        cell(Index) = ""
                    End If
                End If
            Case 27
                combo(Index).Visible = False
                Select Case Index
                    Case 1, 5
                        cell(Index) = cell(Index).tag
                End Select
        End Select
    End With
End Sub

Private Sub cell_LostFocus(Index As Integer)
Dim continue As Boolean
    If usingARROWS Then
        usingARROWS = False
    Else
        If saveBUTTON.Enabled Or Index = 0 Then
            If Not (saveBUTTON.Enabled And Index = 0) Then
                If Index < 6 Then
                    combo(activeCELL).Visible = False
                End If
            End If
        End If
    End If
    If saveBUTTON.Enabled Or Index = 0 Then
        With cell(Index)
            .backcolor = vbWhite
        End With
    End If
    Screen.MousePointer = 0
End Sub



Public Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If findSTUFF(cell(Index), combo(Index), 0) = 0 Then cell(Index) = ""
End Sub

Private Sub combo_Click(Index As Integer)
Dim i, sql, t
Dim cleanDETAILS As Boolean
Dim datax As New ADODB.Recordset
Dim currentformname, currentformname1
Dim MSGBOXReply As VbMsgBoxResult
Dim labelname As String
Dim computerFactor As Double
Dim ratio As Integer
    combo(Index).Visible = False
    DoEvents
    Screen.MousePointer = 11
    DoEvents
    directCLICK = True
    Set datax = New ADODB.Recordset
    DoEvents
    With combo(Index)
        STOCKlist.Enabled = True
        If Index = 5 Then
            Set datax = New ADODB.Recordset
            sql = "SELECT stk_desc,stk_ratio2 FROM STOCKMASTER WHERE " _
                & "stk_npecode = '" + nameSP + "' and " _
                & "stk_stcknumb = '" + .text + "'"
            datax.Open sql, cn, adOpenStatic
            cell(5) = .text
            If datax.RecordCount > 0 Then
                newDESCRIPTION = IIf(IsNull(datax!stk_desc), "", datax!stk_desc)
                'Juan 2010-9-4 implementing ratio rather than computer factor
                computerFactor = datax!stk_compfctr
                ratio = datax!stk_ratio1
                '----------------
            Else
                newDESCRIPTION = ""
                'Juan 2010-9-4 implementing ratio rather than computer factor
                computerFactor = 0
                ratio = 1
                '----------------
            End If
        Else
            If Not savingLABEL.Visible Then
                DoEvents
                '------------------------------
                'Added by Muzammil, this code check if stocks have already been selected, if yes then does not let the
                'user change the FROM Location
                If (Index = 1 Or Index = 2 Or Index = 3) And Len(cell(Index).text) > 0 And HasUserSelectedAnyStocks = True Then
                  If Index = 2 Then
                    labelname = label(2).Caption
                  ElseIf Index = 3 Then
                    labelname = label(3).Caption
                  End If
                  Call MsgBox("Please select and remove each selected Line items before changing the " & labelname & " .", vbInformation, "Imswin")
                  Screen.MousePointer = 0
                  Exit Sub
                End If
                '-------------------------------
                cell(Index) = .TextMatrix(.row, 0)
                DoEvents
                .Refresh
                cell(Index).tag = .TextMatrix(.row, matrix.TextMatrix(10, Index))
            End If
            If Index < 2 Then
                For i = 2 To 4
                    cell(i) = ""
                    cell(i).tag = ""
                Next
            End If
            
            currentformname = frmWarehouse.tag
            currentformname1 = currentformname
            
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Receptions WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                        Case 3
                            .Visible = False
                            Screen.MousePointer = 0
                            Exit Sub
                    End Select
                Case "02050200" 'AdjustmentEntry
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Receptions WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                    End Select
                Case "02040200" 'WarehouseIssue
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then 'Juan 2014-01-05, had to use the new StockInfoForIssue
                                sql = "SELECT  * " _
                                    & "FROM StockInfoForIssue WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                        Case 3
                            Screen.MousePointer = 0
                            Exit Sub
                    End Select
                Case "02040500" 'WellToWell
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
                Case "02040700" 'InternalTransfer
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
                Case "02050300" 'AdjustmentIssue
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
                Case "02040600" 'WarehouseToWarehouse
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
                Case "02040100" 'WarehouseReceipt
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Receptions WHERE QTY1 > 0 AND " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1
                            cleanDETAILS = True
                        Case 2
                            cleanDETAILS = True
                        Case 3
                            cleanDETAILS = True
                        Case 4
                            'Unlock
                            Dim imsLock As imsLock.Lock
                            Set imsLock = New imsLock.Lock
                            If locked = True Then 'sets locked = true because another user has this record open in edit mode
                                Call imsLock.Unlock_Row(locked, cn, CurrentUser, frmWarehouse.POrowguid)  'jawdat
                            End If
                            
                            'Lock
                            Dim ListOfPrimaryControls() As String
                            Call imsLock.Check_Lock(locked, cn, CurrentUser, Array("", frmWarehouse.cell(4), nameSP, "", "", "", "", ""), currentformname1, POrowguid)
                            If locked = True Then 'sets locked = true because another user has this record open in edit mode
                                Screen.MousePointer = 0
                                Exit Sub 'Exit Edit sub because theres nothing the user can do
                            Else
                                locked = True
                            End If
                            '----
                            
                            sql = "StoredProcedure"
                            cleanDETAILS = True
                    End Select
                Case "02050400" 'Sales
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
                Case "02040300" 'Return from Well
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Receptions WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo WHERE " _
                                    & "NAMESPACE = '" + nameSP + "' AND " _
                                    & "Company = '" + cell(1).tag + "' AND " _
                                    & "Location = '" + cell(2).tag + "' " _
                                    & "ORDER BY Stocknumber"
                            End If
                            cleanDETAILS = True
                    End Select
            End Select
            If sql = "" Then
            Else
                If Index = 0 Then
                    datax.Open sql, cn, adOpenForwardOnly
                    If datax.RecordCount > 0 Then
                        Call fillTRANSACTION(datax)
                    End If
                Else
                    Call cleanSTOCKlist
                    If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                        If sql = "StoredProcedure" Then
                            t = cell(4)
                            Set datax = getDATA("getStockInfoPO", Array(nameSP, t))
                        Else
                            datax.Open sql, cn, adOpenForwardOnly
                        End If
                        If datax.RecordCount > 0 Then
                            
                            If datax.RecordCount > 100 Then
                                Label3 = "Loading " + Format(datax.RecordCount) + " records..."
                                savingLABEL.Visible = True
                                DoEvents
                                savingLABEL.ZOrder
                                DoEvents
                            End If
                            DoEvents
                            .MousePointer = 11
                            DoEvents
                            Me.Refresh
                            DoEvents
                         
                         If savingLABEL.Visible Then
                         
                            If frmWarehouse.tag = "02040200" And Index = 2 Then
                            
                                    'StockListDuplicate.Visible = True
                                    
                            End If
                        End If
                            For i = 1 To 4
                                If cell(i).Visible And cell(i) = "" Then STOCKlist.Enabled = False
                            Next
                            Call fillSTOCKlist(datax)
                            'detailHEADER.ZOrder 0
                            If savingLABEL.Visible Then
                                Label3 = "SAVING..."
                                savingLABEL.Visible = False
                                If frmWarehouse.tag = "02040200" And Index = 2 Then
                                    'StockListDuplicate.Visible = False
                                 End If
                            End If
                            
                        End If
                    End If
                End If
            End If
            .Visible = False
        End If
    End With
    If cleanDETAILS Then
        inProgress = False 'Juan 2010-7-22
        Call fillDETAILlist("", "", "", , , , , ctt)
        Call unlockBUNCH
    End If
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
        Case "02050200" 'AdjustmentEntry
        Case "02040200" 'WarehouseIssue
        Case "02040500" 'WellToWell
            If cell(2).tag + cell(3).tag <> "" Then
                If cell(2).tag = cell(3).tag Then
                    cell(Index) = ""
                    cell(Index).tag = ""
                    If Index = 2 Then Call cleanSTOCKlist
                    MsgBox label(2) + " and " + label(Index) + " can not be the same"
                    cell(Index).SetFocus
                End If
            End If
        Case "02040700" 'InternalTransfer
        Case "02050300" 'AdjustmentIssue
        Case "02040600" 'WarehouseToWarehouse
        Case "02040100" 'WarehouseReceipt
            If Index < 4 Then Call cleanSTOCKlist
        Case "02050400" 'Sales
    End Select
    Dim x As String
    'Loads the FQA Details of the saved Transaction ( Only in the case of a modification)
    If Index = 0 Then Call PopulateFQAOftheTransaction(combo(0))
    'Gets the FQA code for the selected Location ( only in the case of a creation)
    'only for WarehouseReceipt,Well to Well, Return From Well
    'If Index = 2 And (Me.tag = "02040100" Or Me.tag = "02040500" Or Me.tag = "02040300") Then
    If Index = 2 And (Me.tag = "02040100") Then
            Call LoadFromFQA(Trim(cell(1).tag), Trim(cell(2).tag))
    End If
        
    If newBUTTON.Enabled = True Then
        Call ChangeMode(True)
    ElseIf newBUTTON.Enabled = False Then
        Call ChangeMode(False)
    End If
    emailButton.Enabled = True
    Screen.MousePointer = 0
End Sub


Private Sub combo_LostFocus(Index As Integer)
    combo(Index).Visible = False
End Sub


Public Sub DTPicker1_DropDown()
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    With DTPicker1
        Select Case KeyCode
            Case 13
                cell(Val(.tag)).text = Format(.Value, "MMMM/dd/yyyy")
                cell(Val(.tag) + 1).SetFocus
        End Select
    End With
End Sub

Private Sub DTPicker1_LostFocus()
Dim indexCELL As Integer
    With DTPicker1
        If IsNumeric(.tag) Then
            cell(Val(.tag)).text = Format(.Value, "MMMM/dd/yyyy")
            indexCELL = Val(.tag)
            If Me.ActiveControl.name = "cell" Then
                If Me.ActiveControl.Index <> Val(.tag) Then .Visible = False
                indexCELL = Me.ActiveControl.Index
            End If
            If Me.ActiveControl.name = "cell" Then
                cell(indexCELL).SetFocus
            Else
                .Visible = False
            End If
        End If
        .Value = Now
    End With
End Sub


Private Sub Form_Activate()
Dim rights As Boolean
    SSOleDBFQA.Visible = False
    inProgress = False
    Screen.MousePointer = 0
    rights = Getmenuuser(nameSP, CurrentUser, Me.tag, cn)
    newBUTTON.Enabled = rights
    
    'Added by Juan (2015/02/13) for Multilingual
    Call translator.Translate_Forms("frmWarehouse")
    '------------------------------------------
    
    Me.Visible = True
    If newBUTTON.Enabled Then newBUTTON.SetFocus
    Me.Refresh
    userNAMEbox = CurrentUser
    dateBOX = Format(Now, "mm/dd/yyyy")
    hideDETAILS
    Call makeLists
    Load grid(1)
    Load grid(2)
    DoEvents
    Call fillGRID(grid(1), logicBOX(0), 0)
    DoEvents
    Call fillGRID(grid(2), sublocaBOX(0), 0)
    Call cleanDETAILS
    Call getEmail
End Sub

Public Sub setCN(conn As ADODB.Connection)
    Set cn = conn
    If Not IsConnectionOpen(conn) Then Exit Sub
End Sub
Private Sub Form_Load()
On Error Resume Next
    'Call translator.Translate_Forms("frmWarehouse")
    Screen.MousePointer = 11
    Call lockDOCUMENT(True)
    frmWarehouse.Caption = frmWarehouse.Caption + " - " + frmWarehouse.tag
    Screen.MousePointer = 0
    If Err Then MsgBox "Error: " + Err.description
    'StockListDuplicate.Visible = False
    
    SSOleCompany.columns(0).width = 855
    SSOleDBLocation.columns(0).width = 975
    SSOleDBUsChart.columns(0).width = 1455
    SSOleDBCamChart.columns(0).width = 1455
    With frmWarehouse
        .Left = Round((Screen.width - .width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub


Sub SAVE()
Dim header As New ADODB.Recordset
Dim details As New ADODB.Recordset
Dim remarks As New ADODB.Recordset

Dim INVitem As New ADODB.Recordset

Dim i, row As Integer
Dim sql As String
Dim q, quantity, price As Double
On Error Resume Next
    
    If readyFORsave Then
        'Header routine
        'msg1 = translator.Trans("M00708")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Header", msg1)
        cn.BeginTrans
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM transaction WHERE inv_ponumb = ''"
        header.Open sql, cn, adOpenDynamic, adLockPessimistic
        With header
            .AddNew
            !inv_creauser = CurrentUser
            !inv_npecode = nameSP
            
            !inv_ponumb = cell(0)
            !inv_invcnumb = cell(1)
            !inv_invcdate = CDate(cell(3))
            !inv_creadate = CDate(cell(4))
            .Update
        End With
        
        'Remarks routine
        'msg1 = translator.Trans("M00719")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Remarks", msg1)
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM transactionREM WHERE invr_ponumb = ''"
        remarks.Open sql, cn, adOpenDynamic, adLockPessimistic
        With remarks
            .AddNew
            !invr_creauser = CurrentUser
            !invr_npecode = nameSP
            !invr_creadate = CDate(cell(4))
            
            !invr_ponumb = cell(0)
            !invr_invcnumb = cell(1)
            !invr_rem = remarks
            !invr_linenumb = 1
            .Update
        End With
                
        'Details routine
        'msg1 = translator.Trans("M00710")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Details", msg1)
        Set details = New ADODB.Recordset
        sql = "SELECT * FROM transactionDETL WHERE invd_ponumb = ''"
        details.Open sql, cn, adOpenKeyset, adLockPessimistic
        With details
            For i = 1 To STOCKlist.Rows - 1
                If STOCKlist.TextMatrix(i, 0) <> "" Then
                    If IsNumeric(STOCKlist.TextMatrix(i, 1)) Then
                        .AddNew
                        !invd_npecode = nameSP
                        !invd_creauser = CurrentUser
                        !invd_creadate = CDate(cell(4))
                        
                        !invd_ponumb = cell(0)
                        !invd_invcnumb = cell(1)
                        !invd_liitnumb = STOCKlist.TextMatrix(i, 1)
                        
                        quantity = IIf(IsNumeric(STOCKlist.TextMatrix(i, 8)), CDbl(STOCKlist.TextMatrix(i, 8)), 0)
                        !invd_primreqdqty = quantity
                        !invd_primuom = STOCKlist.TextMatrix(i, 16)
                        price = IIf(IsNumeric(STOCKlist.TextMatrix(i, 10)), CDbl(STOCKlist.TextMatrix(i, 10)), 0)
                        !invd_unitpric = price
                        !invd_totapric = quantity * price
                                                
                        If Trim(STOCKlist.TextMatrix(i, 15)) = "" Then
                            row = i
                        Else
                            row = i + 1
                        End If
                        quantity = IIf(IsNumeric(STOCKlist.TextMatrix(row, 8)), CDbl(STOCKlist.TextMatrix(row, 8)), 0)
                        !invd_secoreqdqty = quantity
                        !invd_secouom = STOCKlist.TextMatrix(row, 16)
                        price = IIf(IsNumeric(STOCKlist.TextMatrix(row, 10)), CDbl(STOCKlist.TextMatrix(row, 10)), 0)
                        !invd_secounitprice = price
                        !invd_secototaprice = quantity * price
                    End If
                End If
            Next
            'msg1 = translator.Trans("M00714")
            'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Transaction", msg1)
            .UpdateBatch
        End With
        'msg1 = translator.Trans("M00715")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Commiting Transaction", msg1)
        cn.CommitTrans
        'MDI_IMS.StatusBar1.Panels(1).Text = ""
        Screen.MousePointer = 0
        Screen.MousePointer = 11
        Call lockDOCUMENT(True)
        Call clearDOCUMENT
'        Call getPOComboList
    End If
End Sub

Public Sub grid_Click(Index As Integer)
Dim i, name
Dim data As New ADODB.Recordset
skipAlphaSearch = True
skipExistance = True
    With grid(Index)
        justCLICK = True
        If Index = 0 Then
            i = Val(Left(.tag, 2))
            name = Mid(.tag, 3)
        Else
            i = Val(.ToolTipText)
            Select Case Index
                Case 1
                    name = "logicBOX"
                Case 2
                    name = "sublocaBOX"
            End Select
        End If
        
        Dim tempRow As Integer
        If .row = 0 Then Exit Sub
        tempRow = .row
        Select Case name
            Case "logicBOX"
                logicBOX(i) = .TextMatrix(.row, 0) 'Juan  2014-01-02 it was col=1
                .row = tempRow
                logicBOX(i).tag = .TextMatrix(.row, 1)
                logicBOX(i).ToolTipText = .TextMatrix(.row, 0)
                logicBOX(i).SetFocus
            Case "sublocaBOX"
                sublocaBOX(i) = .TextMatrix(.row, 0) 'Juan  2014-01-02 it was col=1
                .row = tempRow
                sublocaBOX(i).tag = .TextMatrix(.row, 1)
                sublocaBOX(i).ToolTipText = .TextMatrix(.row, 0)
                sublocaBOX(i).SetFocus
            Case "NEWconditionBOX"
                NEWconditionBOX(i) = "0" + .TextMatrix(.row, 0)
                .row = tempRow
                NEWconditionBOX(i).tag = .TextMatrix(.row, 0)
                NEWconditionBOX(i).ToolTipText = .TextMatrix(.row, 1)
                NEWconditionBOX(i).SetFocus
                'Juan 2010-10-31 to get new price after changing condition
                Set data = getDATA("conditionVALUE", Array(nameSP, priceBOX(i), NEWconditionBOX(i)))
                If data.RecordCount = 0 Then
                Else
                    Select Case frmWarehouse.tag
                        Case "02050200" 'AdjustmentEntry
                        Case Else
                            priceBOX(i) = Format(CDbl(data(0)), "0.00")
                    End Select
                End If
        End Select
        .Visible = False
    End With
End Sub

Private Sub logicBOX_Click(Index As Integer)
    grid(1).ToolTipText = Format(Index, "00") + "logicBOX"
    Call showGRID(grid(1), Index, logicBOX(Index), True)
End Sub

Private Sub logicBOX_GotFocus(Index As Integer)
    Call whitening
    With logicBOX(Index)
        .backcolor = &H80FFFF
        .SelStart = 0
        .SelLength = Len(.text)
        If justCLICK Then
            grid(1).Visible = False
            justCLICK = False
        Else
            grid(1).ToolTipText = Format(Index, "00") + "logicBOX"
            Call showGRID(grid(1), Index, logicBOX(Index), True)
        End If
   End With
End Sub


Private Sub logicBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call grid_Click(1)
            grid(1).Visible = False
        Case 27
            grid(1).Visible = False
    End Select
End Sub

Private Sub logicBOX_LostFocus(Index As Integer)
    With logicBOX(Index)
        If .text = "" Then
            .backcolor = &HC0C0FF
        Else
            .backcolor = vbWhite
        End If
    End With
    grid(1).Visible = False
End Sub


Private Sub logicBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        With logicBOX(Index)
            .backcolor = &H80FFFF
        End With
    End If
End Sub

Private Sub NEWconditionBOX_Click(Index As Integer)
    Call showGRID(grid(0), Index, NEWconditionBOX(Index))
End Sub


Private Sub NEWconditionBOX_GotFocus(Index As Integer)
    Call whitening
    NEWconditionBOX(Index).backcolor = &H80FFFF
End Sub


Private Sub NEWconditionBOX_LostFocus(Index As Integer)
    NEWconditionBOX(Index).backcolor = vbWhite
    grid(0).Visible = False
End Sub

Private Sub NEWconditionBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        NEWconditionBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantityBOX_Change(Index As Integer)
    If doChanges Then
        'Call quantityBOX_Validate(Index, True)
    Else
        doChanges = True
    End If
End Sub

Private Sub quantityBOX_Click(Index As Integer)
    With quantityBOX(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub quantityBOX_GotFocus(Index As Integer)
    If Index <> totalNode Then
        Call whitening
        quantityBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantityBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        submitted = False
        Call quantityBOX_Validate(Index, True)
    End If
End Sub

Private Sub quantityBOX_LostFocus(Index As Integer)
    If submitted Then Exit Sub
    Call quantityBOX_Validate(Index, True)
    If Index <> totalNode Then quantityBOX(Index).backcolor = vbWhite
End Sub


Private Sub quantityBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        quantityBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Public Sub quantityBOX_Validate(Index As Integer, Cancel As Boolean)
Dim qty, qty2
On Error Resume Next
    If submitted Then Exit Sub
    With quantityBOX(Index)
        If Index <> totalNode Then
            If IsNumeric(.text) Then
                If CDbl(.text) > 0 Then
                    If serialStockNumber Then .text = 1
                    'Juan 2010-6-5
                    '.text = Format(.text, 0)
                    .text = Format(.text, "0.00")
                    'doChanges = False
                    
                    'Juan 2010-9-4 implementing ratio rather than computer factor
                    If ratioValue > 1 Then
                        If IsNumeric(.text) Then
                            qty = CDbl(.text)
                            If qty > 0 Then
                                Call limitQty(Index)
                                'If (invoiceBOX(index) <> "" And invoiceBOX(index) <> "invoiceBOX") And qty > CDbl(quantity(index)) Then .text = quantity(index)
                                qty2 = qty * ratioValue
                                quantity2BOX(Index).text = Format(qty2, "0.00")
                            Else
                                quantity2BOX(Index).text = .text
                            End If
                        End If
                    Else
                        Call limitQty(Index)
                        'If (invoiceBOX(index) <> "" And invoiceBOX(index) <> "invoiceBOX") And CDbl(.text) > CDbl(quantity(index)) Then .text = quantity(index)
                        quantity2BOX(Index).text = .text
                    End If
                    If Err.Number = 340 Then Err.Clear 'if error is about element n doesn't exist it clear errors variables
                    
'                    If computerFactorValue > 0 Then
'                        If IsNumeric(.text) Then
'                            qty = CDbl(.text)
'                            If qty > 0 Then
'                                If Round(computerFactorValue) > 0 Then
'                                    qty2 = qty * 10000 / computerFactorValue
'                                Else
'                                    qty2 = qty / (10000 * computerFactorValue)
'                                End If
'                                quantity2BOX(Index).text = Format(qty2, "0.00")
'                            Else
'                                quantity2BOX(Index).text = .text
'                            End If
'                        End If
'                    Else
'                        quantity2BOX(Index).text = .text
'                    End If
                    '--------------------
                    Select Case frmWarehouse.tag
                        Case "02040100" 'WarehouseReceipt
                            
                        Case "02040200" 'WarehouseIssue
                            
                        Case "02050200" 'AdjustmentEntry
                        Case Else
                            'If CDbl(.text) > CDbl(quantity(Index)) Then .text = quantity(Index)
                    End Select
                Else
                    'Juan 2010-6-5
                    '.text = "0"
                    .text = "0.00"
                    '-----------------
                End If
                    If Err.Number = 0 Then
                        If isFirstSubmit Then
                            If frmWarehouse.Tree.Nodes(Index).text = "Pool" Then
                                Call calculations(True, True, True)
                            Else
                                Call calculations(True, True)
                            End If
                        Else
                            Select Case .tag
                                Case "02040100" 'WarehouseReceipt
                                    Call calculations(True, True)
                                Case Else
                                    Call calculations2(SUMMARYlist.row, Tree.Nodes(Index).text, Index)
                            End Select
                        End If
                    End If
            Else
                    'Juan 2010-6-5
                    '.text = "0"
                    .text = "0.00"
                    '-----------------
            End If
        End If
        .SelStart = Len(.text)
    End With
End Sub

Private Sub remarks_GotFocus()
    remarks.backcolor = &HC0FFFF
End Sub


Private Sub remarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command5_Click
    End If
End Sub

Private Sub remarks_LostFocus()
    remarks.backcolor = vbWhite
End Sub


Private Sub removeDETAIL_Click()
Dim i, ii
Dim RowPosition As Integer
    With SUMMARYlist
        If .Rows > 2 Then
            .RemoveItem .row
        Else
            For ii = 0 To .cols - 2
                .TextMatrix(1, ii) = ""
            Next
            RowPosition = 1
        End If
        Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "delete", "", "", "", "", RowPosition)
        Call reNUMBER(SUMMARYlist)
        Call fillDETAILlist("", "", "", , , , , ctt)
        'Call updateStockListBalance
'Juan 30-10-2010 this part is not useful any more
'        With STOCKlist
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, 1) = commodityLABEL Then
'                    .row = i
'                    Exit For
'                End If
'            Next
'        End With
        Call hideDETAILS(True)
        .Visible = True
        .ZOrder
    End With
End Sub

Private Sub repairBOX_Change(Index As Integer)
    If repairBOX(Index).Visible Then Call repairBOX_Validate(Index, True)
End Sub

Private Sub repairBOX_Click(Index As Integer)
    With repairBOX(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub


Private Sub repairBOX_GotFocus(Index As Integer)
    Call whitening
    repairBOX(Index).backcolor = &H80FFFF
End Sub

Private Sub repairBOX_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        Call repairBOX_Validate(repairBOX(Index), True)
        If Err.Number = 6 Then Exit Sub
        If IsNumeric(repairBOX) Then
            repairBOX(Index) = Format(repairBOX(Index), "0.00")
        End If
    End If
End Sub

Private Sub repairBOX_LostFocus(Index As Integer)
    repairBOX(Index).backcolor = vbWhite
    If IsNumeric(repairBOX(Index)) Then
        repairBOX(Index) = Format(repairBOX(Index), "0.00")
    End If
End Sub

Private Sub repairBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        repairBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub repairBOX_Validate(Index As Integer, Cancel As Boolean)
    Call validateQTY(repairBOX(Index), Index)
End Sub

Private Sub searchFIELD_Change(Index As Integer)
    With STOCKlist
        If Index = 0 Then
            If .row <> 1 Or .RowSel <> .Rows - 1 Then
                .row = 1
                .RowSel = .Rows - 1
            End If
            If .ColSel <> 1 Then
                .col = 1
                .ColSel = 1
                .Sort = 1
            End If
            If STOCKlist.Rows > 2 Then Call alphaSEARCH(searchFIELD(0), STOCKlist, 1)
        End If
    End With
End Sub

Private Sub searchFIELD_GotFocus(Index As Integer)
    searchFIELD(Index).backcolor = &H80FFFF
End Sub


Public Sub searchFIELD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call searchStockNumber(Index)
    End If
End Sub
Private Sub searchFIELD_LostFocus(Index As Integer)
    searchFIELD(Index).backcolor = &HC0E0FF
End Sub

Private Sub SSOleCompany_Click()
SSOleDBFQA.columns("company").Value = SSOleCompany.columns(0).text
End Sub

Private Sub SSOleDBCamChart_Click()
SSOleDBFQA.columns("Camchart#").Value = SSOleDBCamChart.columns(0).text
End Sub

Private Sub SSOleDBFQA_BeforeRowColChange(Cancel As Integer)
Dim Location As String
Dim stockprefix As String

On Error GoTo ErrHand

''''''    SSOleDBStockType.ListAutoValidate = False
''''''    SSOleDBLocation.ListAutoValidate = False
''''''    SSOleDBUsChart.ListAutoValidate = False
''''''    SSOleDBCamChart.ListAutoValidate = False
''''''
''''''    If Len(Trim(SSOleDBFQA.columns(2).text)) > 0 Then LOCATION = UCase(Trim(SSOleDBFQA.columns(2).text))
''''''
''''''    stockprefix = Mid(Trim(SSOleDBFQA.columns(0).text), 1, 2)
''''''
''''''If (LOCATION = "K69871" Or LOCATION = "K69023" Or LOCATION = "K69022") And (stockprefix = "55" Or stockprefix = "66") Then
''''''
''''''    SSOleDBStockType.ListAutoValidate = True
''''''    SSOleDBLocation.ListAutoValidate = True
''''''    SSOleDBUsChart.ListAutoValidate = True
''''''    SSOleDBCamChart.ListAutoValidate = True
''''''
''''''''    LOCKLOCATION = True
''''''''    LOCKUSCHART = True
''''''''    LockType = True
''''''''    LOCKCAMCHART = True
''''''
''''''End If
''''''
''''''If (LOCATION = "K69871" Or LOCATION = "K69023" Or LOCATION = "K69022") And (stockprefix = "33") Then
''''''
''''''    SSOleDBLocation.ListAutoValidate = True
''''''    SSOleDBUsChart.ListAutoValidate = True
''''''
''''''End If

Select Case SSOleDBFQA.col

    Case 0
        If Len(Trim(SSOleDBFQA.columns(0).text & "")) > 20 Then
            MsgBox "Stocknumber is too long. Please make sure it is not larger than 20 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    'company
    Case 1
    
        If Len(Trim(SSOleDBFQA.columns(1).text & "")) > 2 Then
            
            MsgBox "Company is too long. Please make sure it is not larger than 2 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    'location
    Case 2
    
       If Len(Trim(SSOleDBFQA.columns(2).text & "")) > 11 Then
            
            MsgBox "Location is too long. Please make sure it is not larger than 11 characters.", vbInformation, "Imswin"
            Cancel = 1
        End If
    'Us chart
    Case 3
    
        
       If Len(Trim(SSOleDBFQA.columns(3).text)) > 9 Then
            
            MsgBox "US Chart is too long. Please make sure it is not larger than 9 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    'Stocktype
    Case 4
    
        
       If Len(Trim(SSOleDBFQA.columns(4).text)) > 4 Then
            
            MsgBox "Stocktype is too long. Please make sure it is not larger than 4 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    'Can chart
    Case 5
    
        
       If Len(Trim(SSOleDBFQA.columns(5).text)) > 8 Then
            
            MsgBox "Cam Chart is too long. Please make sure it is not larger than 8 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    
End Select
Exit Sub
ErrHand:
MsgBox "Errors Occurred. error description : " & Err.description
Err.Clear
End Sub

Private Sub SSOleDBFQA_InitColumnProps()
SSOleDBFQA.columns("company").DropDownHwnd = SSOleCompany.hWnd
SSOleDBFQA.columns("location").DropDownHwnd = SSOleDBLocation.hWnd
SSOleDBFQA.columns("uschart#").DropDownHwnd = SSOleDBUsChart.hWnd
'SSOleDBFQA.columns("stocktype").DropDownHwnd = SSOleDBStockType.hWnd
SSOleDBFQA.columns("camchart#").DropDownHwnd = SSOleDBCamChart.hWnd

End Sub
Private Sub SSOleDBFQA_KeyPress(KeyAscii As Integer)

If SSOleDBFQA.col = 0 Then
'stockno
    

ElseIf SSOleDBFQA.col = 1 Then

        

ElseIf SSOleDBFQA.col = 2 Then

        SSOleDBLocation.DroppedDown = True

ElseIf SSOleDBFQA.col = 3 Then

    SSOleDBUsChart.DroppedDown = True
    

    
ElseIf SSOleDBFQA.col = 4 Then

        

ElseIf SSOleDBFQA.col = 5 Then

        SSOleDBCamChart.DroppedDown = True

End If

'If SSOleDBFQA.col <> 4 Then KeyAscii = 0



End Sub

Private Sub SSOleDBFQA_KeyUp(KeyCode As Integer, Shift As Integer)
SSOleDBUsChart.DroppedDown = True
End Sub

Private Sub SSOleDBLocation_Click()
SSOleDBFQA.columns("location").Value = SSOleDBLocation.columns(0).text
End Sub

Private Sub SSOleDBStockType_Click()
SSOleDBFQA.columns("stocktype").Value = SSOleDBStockType.columns(0).text
End Sub

Private Sub SSOleDBUsChart_Click()
SSOleDBFQA.columns("UsChart#").Value = SSOleDBUsChart.columns(0).text
End Sub

Private Sub STOCKlist_Click()
Dim i, pointerCOL As Integer
Screen.MousePointer = 11
    doChanges = False
    With STOCKlist
        rowMark = .row
        If Not inProgress Then
            If .MouseCol = 0 Then
                .col = 0
                If .row > 0 Then
                    pointerCOL = 0
                    Call markROW(STOCKlist, , ctt)
                    hideDETAIL.Visible = True
                    submitDETAIL.Visible = True
                    'juan 2012-1-8 commented the line until edition mode works well
                    'removeDETAIL.Visible = True
                End If
            End If
        End If
    End With
Screen.MousePointer = 0
frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub STOCKlist_DblClick()
'    With STOCKlist
'        If Not inProgress Then
'            Me.MousePointer = vbHourglass
'            .col = 0
'            If .text = "?" Then 'Juan 2010-5-25
'            Else '----
'                inProgress = True
'                Call markROW(STOCKlist)
'            End If '-----
'            hideDETAIL.Visible = True
'            submitDETAIL.Visible = True
'            removeDETAIL.Visible = True
'            Call PREdetails
'            Me.MousePointer = 0
'        End If
'    End With
'frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub stocklist_EnterCell()
Screen.MousePointer = 11
    Call selectROW(STOCKlist)
Screen.MousePointer = 0
frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub STOCKlist_GotFocus()
    With STOCKlist
        If STOCKlist.Rows > 2 Or Not (STOCKlist.Rows = 2 And STOCKlist.TextMatrix(1, 0) = "") Then
            If .row = 0 And .col = 1 Then .row = 1
            Call selectROW(STOCKlist)
        End If
    End With
End Sub

Private Sub STOCKlist_LostFocus()
    If STOCKlist.Rows > 2 Or Not (STOCKlist.Rows = 2 And STOCKlist.TextMatrix(1, 0) = "") Then
        Call selectROW(STOCKlist, True)
        STOCKlist.tag = 0
    End If
End Sub

Private Sub STOCKlist_RowColChange()
'    With STOCKlist
'        If IsNumeric(.TextMatrix(.row, 0)) Then
'            Call fillDETAILlist("", "", "")
'        Else
'            Call PREdetails
'        End If
'    End With
End Sub

Private Sub sublocaBOX_Change(Index As Integer)
    Call alphaSEARCH(sublocaBOX(Index), grid(2), 0)
End Sub

Private Sub sublocaBOX_Click(Index As Integer)
    grid(2).ToolTipText = Format(Index, "00") + "sublocaBOX"
    Call showGRID(grid(2), Index, sublocaBOX(Index), True)
End Sub

Private Sub sublocaBOX_GotFocus(Index As Integer)
If ("Sublocation: " + sublocaBOX(Index)) = RTrim(Tree.Nodes(Tree.Nodes.Count - 1).text) Then
    'sublocaBOX(Index).text = ""
Else
    Call whitening
    With sublocaBOX(Index)
        .backcolor = &H80FFFF
        .SelStart = 0
        .SelLength = Len(.text)
        If justCLICK Then
            grid(2).Visible = False
            justCLICK = False
        Else
            grid(2).ToolTipText = Format(Index, "00") + "sublocaBOX"
            Call showGRID(grid(2), Index, sublocaBOX(Index), True)
        End If
    End With
End If
End Sub


Private Sub sublocaBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call grid_Click(2)
            grid(2).Visible = False
        Case 27
            grid(2).Visible = False
    End Select
End Sub

Private Sub sublocaBOX_LostFocus(Index As Integer)
    With sublocaBOX(Index)
        If .text = "" Then
            .backcolor = &HC0C0FF
        Else
            .backcolor = vbWhite
        End If
    End With
    grid(2).Visible = False
End Sub


Private Sub sublocaBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 0 And Index <> totalNode Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        With sublocaBOX(currentBOX)
            .backcolor = &H80FFFF
        End With
    End If
End Sub

Private Sub SUMMARYlist_Click()
If newBUTTON.Enabled Then Exit Sub
Screen.MousePointer = 11
Dim i, pointerCOL As Integer
    With SUMMARYlist
        If .MouseCol = 0 Then
            .row = .MouseRow
            Call SUMMARYlist_DblClick
        End If
    End With
Screen.MousePointer = 0
End Sub

Public Sub SUMMARYlist_DblClick()
If newBUTTON.Enabled Then Exit Sub
Screen.MousePointer = 11
    With SUMMARYlist
        Select Case frmWarehouse.tag
            'ReturnFromRepair, AdjustmentEntry,WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales
            Case "02040400", "02050200", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                'Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6))
                'Call editSummaryList 'Juan new procedure to edit items: juan 2013-12-28 not working at this point
            Case "02040100" 'WarehouseReceipt
                'Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6), .TextMatrix(.row, 17))
                'Call editSummaryList 'Juan new procedure to edit items juan 2013-12-28 not working at this point
            Case "02050200" 'AdjustmentEntry
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 2), .TextMatrix(.row, 3), , .row, , , ctt)
        End Select
    End With
Screen.MousePointer = 0
End Sub


Private Sub SUMMARYlist_EnterCell()
    If newBUTTON.Enabled Then Exit Sub
    Call selectROW(SUMMARYlist)
End Sub

Private Function fxLogicFilledIn() As Boolean
Dim i, n

fxLogicFilledIn = True

'-----> (gib 10/04) Itirate through all of the node items and for each node which has a
'       quantity value, make sure that the sub-location(sublocaBOX) is filled in; if it
'       is not then return FALSE.
'
On Error Resume Next
    
n = frmWarehouse.Tree.Nodes.Count

For i = 1 To n
    Err.Clear
    If CDbl(quantityBOX(i)) > 0 Then
        If Err.Number = 0 Then  'we need to check the error number because this control may not exist for this 'i' value.
            Err.Clear
            If logicBOX(i) = "" Or IsNull(logicBOX(i)) = True Then
                If Err.Number = 0 Then
                    fxLogicFilledIn = False
                    Exit Function
                End If
            End If
        End If
    End If
Next

End Function
Private Function fxSubLocFilledIn() As Boolean
Dim i, n

fxSubLocFilledIn = True

'-----> (gib 10/04) Itirate through all of the node items and for each node which has a
'       quantity value, make sure that the sub-location(sublocaBOX) is filled in; if it
'       is not then return FALSE.
'
On Error Resume Next
    
n = frmWarehouse.Tree.Nodes.Count

For i = 1 To n
    Err.Clear
    If CDbl(quantityBOX(i)) > 0 Then
        If Err.Number = 0 Then  'we need to check the error number because this control may not exist for this 'i' value.
            Err.Clear
            If sublocaBOX(i) = "" Or IsNull(sublocaBOX(i)) = True Then
                If Err.Number = 0 Then
                    fxSubLocFilledIn = False
                    Exit Function
                End If
            End If
        End If
    End If
Next

End Function
Private Function isQtyEntered() As Boolean
'Juan function to check across the tree if a qty has to be entered
Dim i
isQtyEntered = True
On Error GoTo errorHandler
For i = 1 To frmWarehouse.Tree.Nodes.Count
    Err.Clear
        If Err.Number = 0 Then  'we need to check the error number because this control may not exist for this 'i' value.
            If sublocaBOX(i) = "" Or IsNull(sublocaBOX(i)) = True Then
            Else
                If Err.Number = 0 Then
                    If CDbl(quantityBOX(i)) <= 0 Then
                        isQtyEntered = False
                        Exit Function
                    End If
                End If
            End If
        End If
Next
errorHandler:
If Err.Number > 0 Then
    If Err.Number = 340 Then
        Err.Clear
        Resume Next
    Else
        MsgBox Err.description
    End If
End If
End Function
Private Sub submitDETAIL_Click()
Dim aproved As Boolean
On Error Resume Next
Dim i, n, rec, condition, key, conditionCODE, fromlogic, unitCODE
Dim fromSubLoca As String
Dim str As String
Dim PONumb As String
Dim lineno As String
Dim quant As String
'Juan 2010-04-21
Dim nodeText As String
Dim startingPoint As Integer
Dim differenceWithTable As Integer
Dim pieceText, serialText As String
Dim rowKey As String
Dim datax As New ADODB.Recordset
Dim summaryValueFirstTime As Boolean
'-----------------------

'-----> (gib 10/04) If no sub-location has been entered, exit this Sub(do not continue until user enters one).
'
Dim askForSubLocation As Boolean
Dim askForLogic As Boolean
askForSubLocation = False
askForLogic = False
Dim askForQTy As Boolean
askForQTy = False
serialText = ""


Select Case frmWarehouse.tag
    Case "02040400" 'ReturnFromRepair
    Case "02050200" 'AdjustmentEntry
        If UCase(Tree.Nodes.Item(2).text) = "SERIAL:" Or UCase(Tree.Nodes.Item(2).text) = "SERIAL" Or RTrim(Tree.Nodes.Item(2).text) = "" Then
            MsgBox "Please enter a valid serial #"
            Call Tree_Click
            Exit Sub
        End If
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02040200" 'WarehouseIssue
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02040500" 'WellToWell
    Case "02040700" 'InternalTransfer
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02050300" 'AdjustmentIssue
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02040600" 'WarehouseToWarehouse
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02040100" 'WarehouseReceipt
        If Tree.Nodes.Item(2).text = "Serial:" Or RTrim(Tree.Nodes.Item(2).text) = "" Then
            MsgBox "Please enter a valid serial #"
            Call Tree_Click
            Exit Sub
        End If
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
    Case "02050400" 'Sales
    Case "02040300" 'Return from Well
        askForSubLocation = True
        askForLogic = True
        askForQTy = True
End Select

If askForLogic Then
    If Not fxLogicFilledIn() Then
        MsgBox "Logical Warehouse must be entered."
        Exit Sub
    End If
End If
If askForSubLocation Then
    If Not fxSubLocFilledIn() Then
        MsgBox "Sub-Location must be entered."
        Exit Sub
    End If
End If

If isFirstSubmit Then
Else
    If askForQTy Then
        If Not isQtyEntered Then
            MsgBox "A qty bigger than zero must be entered."
            Exit Sub
        End If
    End If
End If

summaryValueFirstTime = True
    If IsNumeric(quantityBOX(totalNode)) Then
        If CDbl(quantityBOX(totalNode)) > 0 Or frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    aproved = True
                    For i = 1 To Tree.Nodes.Count
                        If i <> totalNode Then
                            If IsNumeric(repairBOX(i)) Then
                                If Err.Number = 0 Then
                                    If CDbl(repairBOX(i)) = 0 Then
                                        aproved = False
                                        Exit For
                                    Else
                                        Exit For
                                    End If
                                Else
                                    Err.Clear
                                End If
                            Else
                                If repairBOX(i) = "" Then
                                    If Err.Number = 0 Then
                                        aproved = False
                                        Exit For
                                    Else
                                        Err.Clear
                                    End If
                                Else
                                    aproved = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If Not aproved Then
                        MsgBox "Invalid Repair Cost"
                        repairBOX(i).SelStart = 0
                        repairBOX(i).SelLength = Len(repairBOX(i))
                        repairBOX(i).SetFocus
                        Exit Sub
                    End If
                    If cell(5) = "" Then
                        MsgBox "Invalid New Commodity"
                        cell(5).SetFocus
                        Exit Sub
                    End If
                Case "02050200" 'AdjustmentEntry
                Case "02040200" 'WarehouseIssue
                Case "02040500" 'WellToWell
                Case "02040700" 'InternalTransfer
                Case "02050300" 'AdjustmentIssue
                Case "02040600" 'WarehouseToWarehouse
                Case "02040100" 'WarehouseReceipt
                Case "02050400" 'Sales
                Case "02040300" 'Return from Well
            End Select
            
            With SUMMARYlist
'Juan 2010-7-4
'                    For i = .Rows - 1 To 1 Step -1
'                        If .TextMatrix(i, 1) = commodityLABEL Then
'                            If STOCKlist.text = "?" Then
'                                If .TextMatrix(i, 2) <> "Pool" Then
'                                    serialText = Tree.Nodes(Tree.Nodes.Count - 1).text
'                                    If InStr(serialText, "#") > 0 Then
'                                        serialText = Mid(nodeText, InStrRev(nodeText, "#") + 3)
'                                        If .TextMatrix(i, 2) = serialText Then
'                                            If .Rows > 2 Then
'                                                .RemoveItem i
'                                            Else
'                                                .TextMatrix(1, 0) = ""
'                                                .TextMatrix(1, 1) = ""
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            Else
'                                If .Rows > 2 Then
'                                    .RemoveItem i
'                                Else
'                                    .TextMatrix(1, 0) = ""
'                                    .TextMatrix(1, 1) = ""
'                                End If
'                            End If
'                        End If
'                    Next
'-----------------------

                If Tree.Nodes.Count > 3 Then
                    'differenceWithTable = 1 'Juan 2010-10-3 this was provocating a bug
                    differenceWithTable = 0
                    startingPoint = 3
                    Select Case frmWarehouse.tag
                        Case "02040400" 'ReturnFromRepair
                        Case "02050200" 'AdjustmentEntry
                            startingPoint = 2
                        Case "02040200" 'WarehouseIssue
                        Case "02040500" 'WellToWell
                        Case "02040700" 'InternalTransfer
                        Case "02050300" 'AdjustmentIssue
                        Case "02040600" 'WarehouseToWarehouse
                        Case "02040100" 'WarehouseReceipt
                        Case "02050400" 'Sales
                        Case "02040300" 'Return from Well
                    End Select
                Else
                    startingPoint = 2
                    differenceWithTable = 0
                    Select Case frmWarehouse.tag
                        Case "02040400" 'ReturnFromRepair
                        Case "02050200" 'AdjustmentEntry
                        Case "02040200" 'WarehouseIssue
                        Case "02040500" 'WellToWell
                        Case "02040700" 'InternalTransfer
                        Case "02050300" 'AdjustmentIssue
                        Case "02040600" 'WarehouseToWarehouse
                        Case "02040100" 'WarehouseReceipt
                        Case "02050400" 'Sales
                        Case "02040300" 'Return from Well
                    End Select
                End If
                For i = startingPoint To Tree.Nodes.Count - 1
                    If i <> totalNode + differenceWithTable Then
                        key = Tree.Nodes(i - IIf(differenceWithTable > 0, 1, 0)).key
                        rowKey = Tree.Nodes(i).key
                        nodeText = Tree.Nodes(i + differenceWithTable).text
                        condition = Mid(key, InStr(key, "-") + 1, InStr(key, "{{") - InStr(key, "-") - 1)
                        conditionCODE = Left(key, 2)
                        If InStr(conditionCODE, "{") > 0 Or InStr(conditionCODE, "-") Then conditionCODE = Left(conditionCODE, 1)
                        Err.Clear
                        Dim alreadyPosition As Integer
                        If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                            fromSubLoca = "GENERAL"
                            condition = NEWconditionBOX(i - differenceWithTable)
                            conditionCODE = condition
                        Else
                            fromSubLoca = Mid(key, InStr(key, "{{") + 2)
                        End If

                        If isFirstSubmit Then
                        Else
                            Dim serialNumber
                            If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                serialNumber = pieceText
                                alreadyPosition = locateLine(commodityLABEL, fromSubLoca, 2) 'serial
                            Else
                                alreadyPosition = locateLine(commodityLABEL, fromSubLoca)
                            End If
                        End If
                        If positionBox(i - differenceWithTable) = "" And alreadyPosition = 0 Then
                            If Val(quantityBOX(i - differenceWithTable)) > 0 Then
                                Dim Message
                                If Val(quantity(i - differenceWithTable)) = Val(quantityBOX(i - differenceWithTable)) Then
                                    Message = "This Receipt PO Item valuation is done with Supplier invoice item values"
                                End If
'                                If Val(quantity(i - differenceWithTable)) = Val(quantityBOX(i - differenceWithTable)) Then
'                                    message = "This Receipt valuation is done with Supplier invoice Item values"
'                                End If
'                            "This PO Item has already been received.Please continue"
                            
                            
                            
                            
                                If Err.Number = 0 Then
                                    rec = "" + vbTab
                                    rec = rec + commodityLABEL + vbTab
                                    'If InStrRev(key, "#") > 0 Then  ' Juan 2010-04-21 change for the below one
                                    If InStrRev(nodeText, "#") > 0 Then
                                        'Juan 2010-4-21
                                        'rec = rec + Mid(key, InStrRev(key, "#") + 1) + vbTab
                                        'pieceText = Mid(rowKey, 2)
                                        pieceText = Mid(nodeText, InStrRev(nodeText, "#") + 3)
                                        '---------------------
                                    Else
                                        pieceText = "Pool"
                                    End If
                                    serialText = pieceText
                                    'rec = rec + pieceText + vbTab
                                    rec = rec + serialText + vbTab
                                    rec = rec + condition + vbTab
                                    rec = rec + Format(priceBOX(i - differenceWithTable), "0.00") + vbTab
                                    
                                    
                                    
                                    rec = rec + descriptionLABEL + vbTab
                                    rec = rec + unitLABEL(0) + vbTab
                                    rec = rec + quantityBOX(i - differenceWithTable) + vbTab
                                    rec = rec + Format(i) + vbTab
                                    fromlogic = Mid(key, InStr(key, "{{") + 2)
                                    If InStr(fromlogic, "Serial") > 0 Then
                                        If frmWarehouse.tag = "02040300" Or frmWarehouse.tag = "02040200" Then  'Return from Well, 'WarehouseIssue
                                            fromlogic = Left(fromlogic, InStr(fromlogic, "{{") - 1)
                                        Else
                                            fromlogic = "GENERAL"
                                        End If
                                    Else
                                        If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                            fromlogic = "GENERAL"
                                        Else
                                            fromlogic = Left(fromlogic, InStr(fromlogic, "{{") - 1)
                                        End If
                                    End If
                                    rec = rec + fromlogic + vbTab
                                    If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                        fromSubLoca = "GENERAL"
                                    Else
                                        fromSubLoca = Mid(key, InStr(key, "{{") + 2)
                                        If InStr(fromSubLoca, "{{") > 0 Then
                                            fromSubLoca = Mid(fromSubLoca, InStr(fromSubLoca, "{{") + 2)
                                            If InStr(fromSubLoca, "{{#") > 0 Then
                                                fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "#") - 3)
                                            Else
                                                If InStr(fromSubLoca, "Serial") > 0 Then
                                                    If frmWarehouse.tag = "02040300" Or frmWarehouse.tag = "02040200" Then  'Return from Well, 'WarehouseIssue
                                                        fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "{{") - 1)
                                                    Else
                                                        fromSubLoca = "GENERAL"
                                                    End If
                                                Else
                                                    fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "{{") - 2)
                                                End If
                                            End If
                                        End If
                                    End If
                                    rec = rec + fromSubLoca + vbTab
                                    rec = rec + logicBOX(i - differenceWithTable).tag + vbTab 'Juan 2014-12-04, it was text but changed to tag
                                    rec = rec + sublocaBOX(i - differenceWithTable).tag + vbTab 'Juan 2014-12-04, it was text but changed to tag
                                    rec = rec + NEWconditionBOX(i - differenceWithTable) + vbTab
                                    rec = rec + NEWconditionBOX(i - differenceWithTable).ToolTipText + vbTab
                                    rec = rec + unitBOX(i)
                                    .addITEM rec
                                    Dim qtyArrayTxt, subLocationArrayTxt As String
                                    qtyArrayTxt = ""
                                    Dim counter
                                    For counter = 0 To UBound(qtyArray())
                                        qtyArrayTxt = qtyArrayTxt + Format(qtyArray(counter))
                                    Next
                                    subLocationArrayTxt = Join(subLocationArray())
                                    frmWarehouse.summaryValues.addITEM qtyArrayTxt + vbTab + subLocationArrayTxt + vbTab + invoiceBOX(i - differenceWithTable) + vbTab + invoiceLineBOX(i - differenceWithTable)
                                    If summaryValueFirstTime And summaryValues.TextMatrix(1, 0) + summaryValues.TextMatrix(1, 1) + summaryValues.TextMatrix(1, 2) = "" Then
                                        summaryValues.RemoveItem (1)
                                        summaryValueFirstTime = False
                                    End If
                                    'Juan 2010-9-4 implementing ratio rather than computer
                                    computerFactorValue = ImsDataX.ComputingFactor(nameSP, commodityLABEL, cn) 'deprecated
                                    .TextMatrix(.Rows - 1 - differenceWithTable, 16) = Format(computerFactorValue) 'deprecated
                                    .TextMatrix(.Rows - 1 - differenceWithTable, 23) = Format(quantity2BOX(i - differenceWithTable)) 'deprecated
                                    .TextMatrix(.Rows - 1 - differenceWithTable, 22) = poItemBox(i - differenceWithTable)
                                    Set datax = getDATA("getStockRatio", Array(nameSP, commodityLABEL))
                                    If datax.RecordCount > 0 Then
                                        ratioValue = datax!realratio
                                    Else
                                        ratioValue = 1
                                    End If
                                    
                                    'Juan 2014-8-28 fixing new ratio method
                                    Select Case frmWarehouse.tag
                                        Case "02040100" 'WarehouseReceipt
                                            Dim sql As String
                                            Set datax = New ADODB.Recordset
                                            sql = "select * from invoicedetl where invd_npecode = '" + nameSP + "'  " + _
                                                "and invd_invcnumb = '" + invoiceBOX(i) + "' " + _
                                                "and invd_liitnumb = '" + Trim(invoiceLineBOX(i)) + "'"
                                            datax.Open sql, cn, adOpenStatic
                                            If datax.RecordCount > 0 Then
                                                ratioValue = datax!invd_secoreqdqty / datax!invd_primreqdqty
                                            End If
                                        Case Else
                                    End Select
                                    
                                    .TextMatrix(.Rows - 1 - differenceWithTable, 25) = Format(ratioValue)
                                    '----------------------------------
                                    
                                    'SSOleDBFQA.addITEM commodityLABEL
                                    
                                    If frmWarehouse.tag = "02040100" Then
                                        PONumb = cell(4).tag
                                        lineno = SUMMARYlist.TextMatrix(.Rows - 1, 8)
                                    Else
                                        PONumb = ""
                                        lineno = ""
                                    End If
                                    
                                    quant = SUMMARYlist.TextMatrix(.Rows - 1, 7)
                                    
                                    Call LoadFromFQA(Trim(cell(1).tag), Trim(cell(2).tag), Trim(commodityLABEL))
                                    
                                    'Juan 2011-11-13 to validate fqa header and details
                                    Select Case frmWarehouse.tag
                                        'Case "02040400" 'ReturnFromRepair
                                        'Case "02050200" 'AdjustmentEntry
                                        Case "02040200" 'WarehouseIssue
                                            Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "insert", NEWconditionBOX(i - differenceWithTable).text, PONumb, lineno, quant)
                                        'Case "02040500" 'WellToWell
                                        'Case "02040700" 'InternalTransfer
                                        'Case "02050300" 'AdjustmentIssue
                                        'Case "02040600" 'WarehouseToWarehouse
                                        Case "02040100" 'WarehouseReceipt
                                            Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "insert", NEWconditionBOX(i - differenceWithTable).text, PONumb, lineno, quant, , True)
                                        'Case "02050400" 'Sales
                                        'Case "02040300" 'Return from Well
                                        Case Else
                                            Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "insert", NEWconditionBOX(i - differenceWithTable).text, PONumb, lineno, quant)
                                    End Select


                                    
                                    
                                    Select Case frmWarehouse.tag
                                        Case "02040400" 'ReturnFromRepair
                                            .TextMatrix(.Rows - 1, 17) = repairBOX(i)
                                            .TextMatrix(.Rows - 1, 18) = cell(5)
                                            .TextMatrix(.Rows - 1, 19) = newDESCRIPTION
                                        Case "02050200" 'AdjustmentEntry
                                            .TextMatrix(.Rows - 1, 19) = condition
                                        Case "02040200" 'WarehouseIssue
                                            .TextMatrix(.Rows - 1, 17) = quantity(i)
                                        Case "02040500" 'WellToWell
                                            .TextMatrix(.Rows - 1, 17) = Left(Tree.Nodes(i).key, 2)
                                        Case "02040700" 'InternalTransfer
                                            .TextMatrix(.Rows - 1, 17) = quantity(i)
                                        Case "02050300" 'AdjustmentIssue
                                        Case "02040600" 'WarehouseToWarehouse
                                        Case "02040100" 'WarehouseReceipt
                                            .TextMatrix(.Rows - 1, 17) = quantity(i - differenceWithTable)
                                            'Juan 2010-7-1
                                            '.TextMatrix(.Rows - 1, 22) = repairBOX(i)
                                            'It is necessary to switch between 21&23 to leave first the unit than the qty
                                            .TextMatrix(.Rows - 1, 23) = Format(quantity2BOX(i - differenceWithTable), "0.00")
                                            .TextMatrix(.Rows - 1, 21) = unit2BOX(i - differenceWithTable)
                                            .ColWidth(21) = 1200
                                            .ColWidth(23) = 1200
                                            .TextMatrix(.Rows - 1, 24) = cell(4).tag 'PO num
                                            '--------------------
                                        Case "02050400" 'Sales
                                        Case "02040300" 'Return from Well
                                            .TextMatrix(.Rows - 1, 17) = quantity(i - differenceWithTable)
                                    End Select
                                    'Juan 2010 - 4 - 28
                                    .TextMatrix(.Rows - 1, 20) = Left(Tree.Nodes(i).key, 2)
                                    .TextMatrix(.Rows - 1, 20) = conditionCODE
                                    '---------------------------
                                Else
                                    Err.Clear
                                End If
                            End If
                        Else 'if position <> ""... means editing
                                If Err.Number = 0 Then
                                    Dim position
                                    If alreadyPosition > 0 Then
                                        position = alreadyPosition
                                    Else
                                        position = Val(positionBox(i - differenceWithTable))
                                    End If
                                    .TextMatrix(position, 1) = commodityLABEL
                                    If InStrRev(nodeText, "#") > 0 Then
                                        pieceText = Mid(nodeText, InStrRev(nodeText, "#") + 3)
                                    Else
                                        pieceText = "Pool"
                                    End If
                                    serialText = pieceText
                                    .TextMatrix(position, 2) = serialText
                                    .TextMatrix(position, 3) = condition
                                    .TextMatrix(position, 4) = Format(priceBOX(i - differenceWithTable), "0.00")
                                    .TextMatrix(position, 5) = descriptionLABEL
                                    .TextMatrix(position, 6) = unitBOX(i).text
                                    .TextMatrix(position, 7) = quantityBOX(i - differenceWithTable)
                                    .TextMatrix(position, 8) = Format(i)
                                    fromlogic = Mid(key, InStr(key, "{{") + 2)
                                    If InStr(fromlogic, "Serial") > 0 Then
                                        If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                            fromlogic = Left(fromlogic, InStr(fromlogic, "{{") - 1)
                                        Else
                                            fromlogic = "GENERAL"
                                        End If
                                    Else
                                        If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                            fromlogic = "GENERAL"
                                        Else
                                            fromlogic = Left(fromlogic, InStr(fromlogic, "{{") - 1)
                                        End If
                                    End If
                                    .TextMatrix(position, 9) = fromlogic
                                    If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                                        fromSubLoca = "GENERAL"
                                    Else
                                        fromSubLoca = Mid(key, InStr(key, "{{") + 2)
                                        If InStr(fromSubLoca, "{{") > 0 Then
                                            fromSubLoca = Mid(fromSubLoca, InStr(fromSubLoca, "{{") + 2)
                                            If InStr(fromSubLoca, "#") > 0 Then
                                                fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "#") - 1)
                                            Else
                                                If InStr(fromSubLoca, "Serial") > 0 Then
                                                    If frmWarehouse.tag = "02040300" Or frmWarehouse.tag = "02040200" Then  'Return from Well, 'WarehouseIssue
                                                        fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "{{") - 2)
                                                    Else
                                                        fromSubLoca = "GENERAL"
                                                    End If
                                                Else
                                                    fromSubLoca = Left(fromSubLoca, InStr(fromSubLoca, "{{") - 2)
                                                End If
                                            End If
                                        End If
                                    End If
                                    .TextMatrix(position, 10) = fromSubLoca
                                    .TextMatrix(position, 11) = logicBOX(i - differenceWithTable)
                                    .TextMatrix(position, 12) = sublocaBOX(i - differenceWithTable)
                                    .TextMatrix(position, 13) = NEWconditionBOX(i - differenceWithTable)
                                    .TextMatrix(position, 14) = NEWconditionBOX(i - differenceWithTable).ToolTipText
                                    .TextMatrix(position, 15) = unitLABEL(i - differenceWithTable)
                                    computerFactorValue = ImsDataX.ComputingFactor(nameSP, commodityLABEL, cn)
                                    .TextMatrix(position, 16) = Format(computerFactorValue)
                                    .TextMatrix(position, 23) = Format(quantity2BOX(i - differenceWithTable), "0.00")
                                    '.TextMatrix(position, 21) = Format(quantity2BOX(i - differenceWithTable))
                                    
                                    .TextMatrix(position, 22) = poItemBox(i - differenceWithTable)
                                    .TextMatrix(position, 21) = unit2BOX(i)
                                    '.TextMatrix(position, 23) = unit2BOX(i)
                                    If frmWarehouse.tag = "02040100" Then
                                        PONumb = cell(4).tag
                                        lineno = SUMMARYlist.TextMatrix(.Rows - 1, 8)
                                    Else
                                        PONumb = ""
                                        lineno = ""
                                    End If
                                    'Juan 2010-9-5 Added to have new ratio functionality
                                    Set datax = getDATA("getStockRatio", Array(nameSP, commodityLABEL))
                                    If datax.RecordCount > 0 Then
                                        ratioValue = datax!realratio
                                    Else
                                        ratioValue = 1
                                    End If
                                    'Juan 2014-8-28 new version of calculation based on invoice if exists
                                    Select Case frmWarehouse.tag
                                        Case "02040100" 'WarehouseReceipt
                                            Set datax = New ADODB.Recordset
                                            sql = "select * from invoicedetl where invd_npecode = '" + nameSP + "'  " + _
                                                "and invd_invcnumb = '" + invoiceBOX(i) + "' " + _
                                                "and invd_liitnumb = '" + Trim(invoiceLineBOX(i)) + "'"
                                            datax.Open sql, cn, adOpenStatic
                                            If datax.RecordCount > 0 Then
                                                ratioValue = datax!invd_secoreqdqty / datax!invd_primreqdqty
                                            End If
                                        Case Else
                                    End Select
                                    datax.Close
                    '----------------------
                                    .TextMatrix(position, 25) = Format(ratioValue)
                                    '---------------------
                                    
                                    quant = SUMMARYlist.TextMatrix(.Rows - 1, 7)
                                    Select Case frmWarehouse.tag
                                        Case "02040400" 'ReturnFromRepair
                                            .TextMatrix(position, 17) = repairBOX(i)
                                            .TextMatrix(position, 18) = cell(5)
                                            .TextMatrix(position, 19) = newDESCRIPTION
                                        Case "02050200" 'AdjustmentEntry
                                        Case "02040200" 'WarehouseIssue
                                        Case "02040500" 'WellToWell
                                            .TextMatrix(position, 17) = Left(Tree.Nodes(i).key, 2)
                                        Case "02040700" 'InternalTransfer
                                        Case "02050300" 'AdjustmentIssue
                                        Case "02040600" 'WarehouseToWarehouse
                                        Case "02040100" 'WarehouseReceipt
                                            .TextMatrix(position, 17) = quantity(i)
                                            .TextMatrix(position, 24) = PONumb
                                            'Juan 2010-8-19
                                            ' It is necessary to switch between 21 & 23
                                            .TextMatrix(position, 23) = Format(quantity2BOX(i - differenceWithTable), "0.00")
                                            .TextMatrix(position, 21) = unit2BOX(i)
                                            '---------------------
                                        Case "02050400" 'Sales
                                        Case "02040300" 'Return from Well
                                    End Select
                                    .TextMatrix(position, 20) = conditionCODE
                                Else
                                    Err.Clear
                                End If
                        End If
                    End If
                Next
                If .Rows > 2 And .TextMatrix(1, 0) = "" Then .RemoveItem 1
                Call reNUMBER(SUMMARYlist)
                .RowHeight(.Rows - 1) = .RowHeight(1)
            End With
            If serialText = "" Or UCase(serialText) = "POOL" Then
                Call hideDETAILS(False, , True)
            Else
                'Call hideDETAILS(False, True, True) 'juan 2012-3-10
                Call hideDETAILS(False, , True)
            End If
            Exit Sub
        End If
    End If
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
            MsgBox "Please enter the quantity you want to return from repair"
        Case "02050200" 'AdjustmentEntry
        Case "02040200" 'WarehouseIssue
            MsgBox "Please enter the quantity you want to issue of this commodity"
        Case "02040500" 'WellToWell
        Case "02040700" 'InternalTransfer
        Case "02050300" 'AdjustmentIssue
        Case "02040600" 'WarehouseToWarehouse
        Case "02040100" 'WarehouseReceipt
        Case "02050400" 'Sales
        Case "02040300" 'Return from Well
    End Select
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
'    grid(3).Visible = False
End Sub

Private Sub TextLINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call TextLINE_Validate(True)
        Case 27
            TextLINE.Visible = False
    End Select
End Sub


Private Sub TextLINE_LostFocus()
    With TextLINE
        If .Visible Then
            .Visible = False
            Call TextLINE_Validate(True)
        End If
    End With
End Sub

Public Sub TextLINE_Validate(Cancel As Boolean)
Dim i, col, row As Integer
Dim qty, switch As String
Dim newPRICE, qty1, qty2, uPRICE1, uPRICE2 As Double
Dim newPRICEok As Boolean
    With TextLINE
        If STOCKlist.col = 8 Or STOCKlist.col = 10 Then
            col = STOCKlist.col
            If IsNumeric(.text) Then
                If Val(.text) > 0 Then
                     STOCKlist.TextMatrix(STOCKlist.row, col) = FormatNumber(.text, 2)
                    switch = Trim(STOCKlist.TextMatrix(STOCKlist.row, 15))
                    Select Case switch
                        Case ""
                            Call differences(STOCKlist.row)
                        Case "P", "S"
                            If STOCKlist.TextMatrix(STOCKlist.row, 1) = "?" Then
                                row = STOCKlist.row - 1
                            Else
                                row = STOCKlist.row
                            End If
                            newPRICEok = True
                            If IsNumeric(STOCKlist.TextMatrix(row, 8)) Then
                                qty1 = CDbl(STOCKlist.TextMatrix(row, 8))
                            Else
                                qty1 = 0
                                newPRICEok = False
                            End If
                            If IsNumeric(STOCKlist.TextMatrix(row + 1, 8)) Then
                                qty2 = CDbl(STOCKlist.TextMatrix(row + 1, 8))
                            Else
                                qty2 = 0
                                newPRICEok = False
                            End If
                            If switch = "P" Then
                                If IsNumeric(STOCKlist.TextMatrix(row, 10)) Then
                                    uPRICE1 = CDbl(STOCKlist.TextMatrix(row, 10))
                                Else
                                    uPRICE1 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE2 = (qty1 * uPRICE1) / qty2
                                    STOCKlist.TextMatrix(row + 1, 10) = FormatNumber(uPRICE2, 2)
                                End If
                            Else
                                If IsNumeric(STOCKlist.TextMatrix(row + 1, 10)) Then
                                    uPRICE2 = CDbl(STOCKlist.TextMatrix(row + 1, 10))
                                Else
                                    uPRICE2 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE1 = (qty2 * uPRICE2) / qty1
                                    STOCKlist.TextMatrix(row, 10) = FormatNumber(uPRICE1, 2)
                                End If
                            End If
                            Call differences(row)
                            Call differences(row + 1)
                    End Select
                    
                    .tag = ""
                    .text = ""
                    .Visible = False
                    Exit Sub
                End If
            End If
            If .text <> "" Then
                'msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Missing value in field", msg1)
                TextLINE = ""
            End If
        End If
    End With
End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim nody As node
Dim sql
Dim datax As New ADODB.Recordset
Dim n As Integer
    For Each nody In Tree.Nodes
        If nody.text = NewString Then
            Tree.Nodes.Remove (Tree.SelectedItem.Index)
            Exit For
        End If
    Next
    If NewString = "Serial:" Then
        MsgBox "Please enter a valid serial #"
        Exit Sub
    End If
    sql = "SELECT * From QTYST6 WHERE " _
        & "qs6_npecode = '" + nameSP + "' AND " _
        & "qs6_stcknumb = '" + commodityLABEL + "' AND " _
        & "qs6_serl = '" + NewString + "' AND " _
        & "qs6_primqty > 0"
    If sql = "" Then
        Cancel = True
        Tree.Nodes.Remove (Tree.SelectedItem.Index)
        Exit Sub
    Else
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenForwardOnly
        If datax.RecordCount > 0 Then
            'Tree.Nodes.Remove (Tree.SelectedItem.Index) 'Juan 2014-01-06, commented to fix bug when found it on db
            MsgBox "That serial number is already registered in the warehouse.  Please enter a different one"
            NewString = "Serial" 'Juan 2014-01-06, added to reset serial
            Exit Sub
        End If
    End If
    'Juan 2013-12-29, to prevent enter a duplicated serial on the transaction
    Dim i
    With SUMMARYlist
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 2) = NewString Then
                MsgBox "That serial number is already taken within the transaction"
                NewString = "Serial:"
                Exit Sub
            End If
        Next
    End With
    '--------------------------
    n = InStr(Tree.SelectedItem.key, "{{Serial")
    If n > 0 Then
        Tree.SelectedItem.key = Left(Tree.SelectedItem.key, n + 7)
    End If
    Tree.SelectedItem.key = Tree.SelectedItem.key + "@" + NewString
    NewString = "Serial #: " + NewString
End Sub

Public Sub Tree_Click()
On Error Resume Next
Dim n
    With Tree
        n = .SelectedItem.Index
        If n = totalNode Then
            If nodeSEL <> totalNode Then
                quantity(totalNode).backcolor = &H800000
                quantity(totalNode).ForeColor = vbWhite
            End If
        Else
            'Juan 2010-5-29
            Select Case frmWarehouse.tag
                Case "02040100", "02050200"  'WarehouseReceipt, AdjustmentEntry
                    n = InStr(.SelectedItem.key, "{{Serial")
                    If n > 0 Then
                        .SelectedItem.text = "Serial:"
                        .StartLabelEdit
                    End If
            End Select
            '---------------------
        End If
    End With
End Sub

Private Sub Tree_Collapse(ByVal node As MSComctlLib.node)
    node.Expanded = True
End Sub


Private Sub Tree_LostFocus()
    Tree.SelectedItem = Nothing
    'Call Tree_Click
End Sub


Private Sub Tree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    With Tree
        nodeSEL = .SelectedItem.Index
        If nodeSEL > 0 Then
            quantity(totalNode).backcolor = &HC0C0C0
            quantity(totalNode).ForeColor = vbBlack
            If nodeSEL <> totalNode Then
                quantity(nodeSEL).backcolor = vbWhite
                quantity(nodeSEL).ForeColor = vbBlack
            End If
        End If
    End With
End Sub

Private Sub Tree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If currentBOX > 0 Then
        Call whitening
        currentBOX = 0
    End If
End Sub

Private Sub Tree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo getOUT
Dim nody As node
    If newBUTTON.Enabled Then Exit Sub
    If Button = 2 Then
        Set nody = Tree.HitTest(x, y)
        If nody.Image = "thing" Then
            deleteITEM.Enabled = False
        Else
            deleteITEM.Enabled = True
        End If
        PopupMenu treeMENU
    End If
    Exit Sub
    
getOUT:
    Exit Sub
    If Err.Number = 91 Then
        addITEM.Enabled = False
        deleteITEM.Enabled = True
        PopupMenu treeMENU
    End If
End Sub

Public Sub setNAMESPACE(NP As String)
    nameSP = NP
End Sub
Public Sub setNAMESPACE_name(NP_name As String)
    nameSPname = NP_name
End Sub
Public Function SaveFQA(Transnumb, TransactionType As String) As Boolean
Dim TranNo As String
Dim i As Integer
TranNo = CStr(Transnumb)

SaveFQA = False

Dim RsTOFQA As New ADODB.Recordset
Dim RsInventoryFQA As New ADODB.Recordset
Dim RsFROMFQA As New ADODB.Recordset
Dim RsUnitPrice As ADODB.Recordset
Dim Location As String
Dim Company As String
RsInventoryFQA.source = " select * from inventoryfqa where 1=2"
RsInventoryFQA.Open , cn, adOpenDynamic, adLockOptimistic

If Me.tag = "02050300" Or Me.tag = "02050200" Then

    Location = cell(2).tag
    Company = cell(1).tag
 Else
 
    Location = Trim(cell(3).tag & "")
    Company = Trim(cell(1).tag & "")
    
End If

   'The is the header(FROM) of the FQA data being stored in the Inventory Table

SSOleDBFQA.MoveFirst

For i = 1 To SSOleDBFQA.Rows
        
    RsInventoryFQA.AddNew
    
    RsInventoryFQA("Npce_code") = nameSP
    RsInventoryFQA("TransactionNo") = Trim(TranNo)
    RsInventoryFQA("ItemNo") = Trim(i)
    RsInventoryFQA("TransactionType") = Trim(TransactionType)
    RsInventoryFQA("transactiondate") = Now
    RsInventoryFQA("Ponumb") = SSOleDBFQA.columns("PONUMB").text
    RsInventoryFQA("PoItemNo") = IIf(Len(Trim(SSOleDBFQA.columns("LINENO").text)) = 0, Null, SSOleDBFQA.columns("LINENO").text)
    RsInventoryFQA("StockNo") = IIf(Len(Trim(SSOleDBFQA.columns("stocknumber").text)) = 0, Null, SSOleDBFQA.columns("stocknumber").text)
    RsInventoryFQA("ToCondition") = IIf(Len(Trim(SSOleDBFQA.columns("tocond").text)) = 0, Null, SSOleDBFQA.columns("tocond").text)
    
    Set RsUnitPrice = New ADODB.Recordset
    
    RsUnitPrice.source = " select sap_value, sap_value * (select top 1 curd_value from currencydetl where curd_code ='" & GExtendedCurrency & "' and"
    RsUnitPrice.source = RsUnitPrice.source & " getdate() > curd_from and getdate() < curd_to) Extnsapvalue from sap where sap_compcode ='" & Company & "' and sap_npecode ='" & nameSP & "'"
    RsUnitPrice.source = RsUnitPrice.source & " and sap_loca='" & Location & "' and sap_stcknumb='" & SSOleDBFQA.columns("STOCKNUMBER").Value & "' and sap_cond ='" & SSOleDBFQA.columns("tocond").Value & "'"
    RsUnitPrice.Open , cn
        
    RsInventoryFQA("BaseCURUnitPrice") = "0"
        
    If Len(GExtendedCurrency) > 0 And RsUnitPrice.RecordCount > 0 Then
        
            RsInventoryFQA("ExtendedUnitPrice") = Round(RsUnitPrice("Extnsapvalue"), 4)
            RsInventoryFQA("BaseCURUnitPrice") = RsUnitPrice("sap_value")
            
    End If
    
    
    RsInventoryFQA("BaseCurrency") = "USD"
    RsInventoryFQA("ExtendedCurrency") = GExtendedCurrency
    
    
    RsInventoryFQA("Quantity") = SSOleDBFQA.columns("quantity").text
    RsInventoryFQA("FromCompany") = Trim(TxtCompany.text)
    RsInventoryFQA("FromLocation") = Trim(TxtLocation)
    RsInventoryFQA("FromUsChar") = Trim(TxtUSChart)
    RsInventoryFQA("FromStockType") = Trim(TxtStockType)
    RsInventoryFQA("FromCamChar") = Trim(TxtCamChart)
    RsInventoryFQA("ToCompany") = Trim(SSOleDBFQA.columns("company").Value)
    RsInventoryFQA("ToLocation") = Trim(SSOleDBFQA.columns("location").Value)
    RsInventoryFQA("ToUsChar") = Trim(SSOleDBFQA.columns("USChart#").Value)
    RsInventoryFQA("ToStockType") = Trim(SSOleDBFQA.columns("stocktype").Value)
    RsInventoryFQA("ToCamChar") = Trim(SSOleDBFQA.columns("CamChart#").Value)
    RsInventoryFQA("TBS") = 1
    RsInventoryFQA("CreaUser") = CurrentUser
    RsInventoryFQA("CreaDate") = Now()
    RsInventoryFQA("ModiUser") = CurrentUser
    RsInventoryFQA("ModiDate") = Now()
    
    SSOleDBFQA.MoveNext
    
    Set RsUnitPrice = Nothing
    
Next
    
    RsInventoryFQA.UpdateBatch
    

''''''''    RsTOFQA.source = "select * from TOFQA where 1=2 "
''''''''    RsTOFQA.ActiveConnection = cn
''''''''    RsTOFQA.Open , , adOpenStatic, adLockBatchOptimistic
''''''''
''''''''    RsFROMFQA.source = "select * from FROMFQA where 1=2 "
''''''''    RsFROMFQA.ActiveConnection = cn
''''''''    RsFROMFQA.Open , , adOpenStatic, adLockOptimistic
''''''''    RsFROMFQA.AddNew
''''''''
''''''''    RsFROMFQA("npce_code") = nameSP
''''''''    RsFROMFQA("TransactionNo") = Trim(TranNo)
''''''''    RsFROMFQA("TransactionType") = Trim(TransactionType)
''''''''    RsFROMFQA("FromCompany") = Trim(TxtCompany.text)
''''''''    RsFROMFQA("FromLocation") = Trim(TxtLocation)
''''''''    RsFROMFQA("FromUsChar") = Trim(TxtUSChart)
''''''''    RsFROMFQA("FromStockType") = Trim(TxtStockType)
''''''''    RsFROMFQA("FromCamChar") = Trim(TxtCamChart)
''''''''    RsFROMFQA("creadate") = Now()
''''''''    RsFROMFQA("tbs") = 1
''''''''    RsFROMFQA("Creauser") = CurrentUser
''''''''    RsFROMFQA.Update
''''''''
''''''''
''''''''
''''''''
''''        RsTOFQA("npce_code") = nameSP
''''        RsTOFQA("TransactionNo") = Trim(TranNo)
''''        RsTOFQA("ItemNo") = Trim(i)
''''        RsTOFQA("StockNo") = Trim(SSOleDBFQA.columns("stocknumber").Value)
''''        RsTOFQA("Company") = Trim(SSOleDBFQA.columns("company").Value)
''''        RsTOFQA("Location") = Trim(SSOleDBFQA.columns("location").Value)
''''        RsTOFQA("UsChar") = Trim(SSOleDBFQA.columns("USChart#").Value)
''''        RsTOFQA("StockType") = Trim(SSOleDBFQA.columns("stocktype").Value)
''''        RsTOFQA("CamChar") = Trim(SSOleDBFQA.columns("CamChart#").Value)
''''        RsTOFQA("creadate") = Now()
''''        RsTOFQA("tbs") = 1
''''        RsTOFQA("Creauser") = CurrentUser
''''        RsTOFQA("TransactionType") = Trim(TransactionType)
''''        RsTOFQA.Update
''''
''''        SSOleDBFQA.MoveNext
''''''''
''''''''    Next

        

SaveFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes." & Err.description, vbCritical, "Ims"

Err.Clear

End Function

Public Function CleanFQA()
SSOleDBFQA.RemoveAll
TxtCamChart.text = ""
TxtCompany.text = ""
TxtLocation.text = ""
TxtUSChart.text = ""
TxtStockType.text = ""
End Function

Public Function GetFROMFQAForTransaction(TranNo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

On Error GoTo ErrHand
Set rs = New ADODB.Recordset
rs.source = "SELECT * from FromFQA where TransactionNo ='" & TranNo & "' and Npce_code ='" & nameSP & "'"
rs.Open , cn

Set GetFROMFQAForTransaction = rs

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to get the FQA details for the transaction. " & Err.description, vbCritical, "Ims"
Err.Clear
End Function

Public Function GetTOFQAForTransaction(TranNo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

On Error GoTo ErrHand

Set rs = New ADODB.Recordset

rs.source = "SELECT * from TOFQA where TransactionNo ='" & TranNo & "' and Npce_code ='" & nameSP & "'"
rs.Open , cn

Set GetTOFQAForTransaction = rs

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to get the FQA details for the transaction." & Err.description, vbCritical, "Ims"
Err.Clear
End Function
Public Function GetFQAForTransaction(TranNo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

On Error GoTo ErrHand

Set rs = New ADODB.Recordset

rs.source = "SELECT * from INVENTORYFQA where TransactionNo ='" & TranNo & "' and Npce_code ='" & nameSP & "'"
rs.Open , cn

Set GetFQAForTransaction = rs

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to get the FQA details for the transaction." & Err.description, vbCritical, "Ims"
Err.Clear
End Function
Public Function PopulateFQAOftheTransaction(TranNo As String) As Boolean
Dim rsFrom As ADODB.Recordset
Dim RsTo As ADODB.Recordset

On Error GoTo ErrHand

   ' Set RsTo = GetTOFQAForTransaction(TranNo)
    
    Set rs = GetFQAForTransaction(TranNo)
    
    If rs.EOF = False Then
    
            TxtCompany = rs("FromCompany")
            TxtLocation = rs("FromLocation")
            TxtUSChart = rs("FromUsChar")
            TxtStockType = rs("FromStockType")
            TxtCamChart = rs("FromCamChar")
            
    End If

    SSOleDBFQA.RemoveAll

    Do While Not rs.EOF
            
        SSOleDBFQA.addITEM rs("StockNo") & vbTab & rs("toCompany") & vbTab & rs("toLocation") & vbTab & rs("toUsChar") & vbTab & rs("toStockType") & vbTab & rs("toCamChar") & vbTab & rs("ponumb") & "" & vbTab & rs("PoItemNo") & "" & vbTab & rs("ToCondition") & vbTab & rs("Quantity")
        'SSOleDBFQA.addITEM No & vbTab & GDefaultFQA.Company & vbTab & GDefaultFQA.Location & vbTab & GDefaultFQA.UsChart & vbTab & StockType & vbTab & GDefaultFQA.CamChart & vbTab & PONumb & vbTab & lineno & vbTab & Tocondition & vbTab & quantity
        rs.MoveNext
    Loop

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to populate the grid with FQA for the transaction." & Err.description, vbCritical, "Ims"
Err.Clear
End Function

Public Function PopulateCombosWithFQA(Companycode As String, Optional LocationCode As String) As Boolean

On Error GoTo ErrHand
PopulateCombosWithFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

LocationCode = Trim(LocationCode)

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Level ='C' order by FQA"

RsCompany.Open , cn

Do While Not RsCompany.EOF

    SSOleCompany.addITEM RsCompany("FQA")
    RsCompany.MoveNext
    
Loop

'RsLocation.source = "select distinct(FQA) from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS'"
RsLocation.source = "select distinct(FQA) from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Level ='LB' OR LEVEL ='LS' order by FQA"

RsLocation.Open , cn

If RsLocation.RecordCount = 0 Then SSOleDBLocation.addITEM LocationCode
Do While Not RsLocation.EOF

    SSOleDBLocation.addITEM RsLocation("FQA")
    RsLocation.MoveNext
    
Loop


'Get US Chart FQA

RsUC.source = "select distinct(FQA) from  FQA where Namespace ='" & nameSP & "'  and Level ='UC'  order by FQA" ' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC'"

RsUC.Open , cn


Do While Not RsUC.EOF

    SSOleDBUsChart.addITEM RsUC("FQA")
    RsUC.MoveNext
    
Loop

'Get Cam Chart FQA

RsCC.source = "select  distinct(FQA) from FQA where Namespace ='" & nameSP & "'  and Level ='CC'  order by FQA" ' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='CC'"

RsCC.Open , cn


Do While Not RsCC.EOF

    SSOleDBCamChart.addITEM RsCC("FQA")
    RsCC.MoveNext
    
Loop

Set RsCompany = Nothing
Set RsLocation = Nothing
Set RsUC = Nothing
Set RsCC = Nothing

PopulateCombosWithFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes." & Err.description, vbCritical, "Ims"

Err.Clear

End Function





Public Function LoadFromFQA(Companycode As String, LocationCode As String, Optional stockno As String)

'Receipt, Return From Well, Well to Well
'-----------------------------------------
'IN this case this function is called when the user selects the FROM Location and the FROM Fqas are populated right away.
'In case of receipt the FROM FQA are hard coded value , in case of Return from well and Well to well ,
'there is only one FQA account for each well


' Issues, warehouse to Warehouse, Adjustment Entry\ Issue
'-----------------------------------------
'In this case this function is called when the user selects the first stock no and depending on if it is a
'controlled or expense stock the FROM FQA is populated.


Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

Dim RsCompanyDefault As New ADODB.Recordset
Dim RsLocationDefault As New ADODB.Recordset
Dim RsUCDefault As New ADODB.Recordset
Dim RsCCDefault As New ADODB.Recordset

Dim companyFQA As String
Dim LocationFQA As String
Dim USChartFQA As String
Dim CamChartFQA As String
Dim StockType As String

On Error GoTo ErrHand

'Which mean that it should be executed only once wihch would be the first time a stock no is selected.

If (SUMMARYlist.Rows > 3 Or SUMMARYlist.Rows = 3) And TxtCompany.Rows > 0 Then Exit Function


        RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Level ='C'"
        RsCompany.Open , cn
        
        RsLocation.source = "select distinct(FQA) from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Level ='LB' OR LEVEL ='LS' order by FQA"
        RsLocation.Open , cn
        
        RsUC.source = "select distinct(FQA) from  FQA where Namespace ='" & nameSP & "'  and Level ='UC'  order by FQA"
        RsUC.Open , cn
        
        RsCC.source = "select  distinct(FQA) from FQA where Namespace ='" & nameSP & "'  and Level ='CC'  order by FQA"
        RsCC.Open , cn


    If Me.tag = "02040100" Then ' Or Me.tag = "02040500" Or Me.tag = "02040300" Then    'Receipt, Return From Well, Well to Well
    
            RsLocationDefault.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='LB' or  Level ='LS'"
            RsLocationDefault.Open , cn
            
            RsUCDefault.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='UC'"
            RsUCDefault.Open , cn
            
            RsCCDefault.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='CC'"
            RsCCDefault.Open , cn


            If RsCompany.RecordCount > 0 Then companyFQA = RsCompany("FQA")


            If RsLocation.RecordCount > 0 Then
            
                If RsLocationDefault.RecordCount > 0 Then
                    LocationFQA = RsLocationDefault("FQA")
                Else
                    LocationFQA = ""
                End If
                
            End If


            If RsUC.RecordCount > 0 Then
            
                If RsUCDefault.RecordCount > 0 Then
                    USChartFQA = RsUCDefault("FQA")
                Else
                    USChartFQA = ""
                End If
                
            End If


            If RsCC.RecordCount > 0 Then
            
                If RsCCDefault.RecordCount > 0 Then
                    CamChartFQA = RsCCDefault("FQA")
                Else
                    CamChartFQA = ""
                End If

            End If
            
            'If this is a Receipt then ...
            If Me.tag = "02040100" Then StockType = "0000"
  ' Return From Well, Well to Well, Issues, warehouse to Warehouse, Adjustment Entry\ Issue
ElseIf Me.tag = "02040500" Or Me.tag = "02040300" Or Me.tag = "02040200" Or Me.tag = "02040600" Or Me.tag = "02050300" Or Me.tag = "02050200" Then
        
        'Companycode As String, LocationCode As String, stockno As String
        Call LoadDefaultValuesForFROMFQA(Companycode, LocationCode, stockno)
        companyFQA = GDefaultFQA.Company
        LocationFQA = GDefaultFQA.Location
        USChartFQA = GDefaultFQA.UsChart
        CamChartFQA = GDefaultFQA.CamChart
        StockType = GDefaultFQA.StockType
        
Else
    'this would cover Internal Tansfer as well
        'Companycode As String, LocationCode As String, stockno As String
        Call LoadDefaultValuesForFROMFQA(Companycode, LocationCode, stockno)
        companyFQA = GDefaultFQA.Company
        LocationFQA = GDefaultFQA.Location
        USChartFQA = GDefaultFQA.UsChart
        CamChartFQA = GDefaultFQA.CamChart
        StockType = GDefaultFQA.StockType
End If
            

            If RsCompany.RecordCount > 0 Then
            
                'companyFQA = RsCompany("FQA")
                TxtCompany.RemoveAll
                Do While Not RsCompany.EOF
                    
                    TxtCompany.addITEM RsCompany("FQA")
                    RsCompany.MoveNext
                    
                Loop
                
                TxtCompany.text = companyFQA
            
            End If
            
            If RsLocation.RecordCount > 0 Then
            
                'If RsLocationDefault.RecordCount > 0 Then
                '    LocationFQA = RsLocationDefault("FQA")
                'Else
                '    LocationFQA = ""
                'End If
                TxtLocation.RemoveAll
                Do While Not RsLocation.EOF
                                
                    TxtLocation.addITEM RsLocation("FQA")
                    RsLocation.MoveNext
                    
                Loop
                
                TxtLocation.text = LocationFQA
                
            End If
            
            If RsUC.RecordCount > 0 Then
            
                'If RsUCDefault.RecordCount > 0 Then
                '   USChartFQA = RsUCDefault("FQA")
                'Else
                '   USChartFQA = ""
                'End If
                TxtUSChart.RemoveAll
                Do While Not RsUC.EOF
                
                    'USChartFQA = RsUC("FQA")
                    TxtUSChart.addITEM RsUC("FQA")
                    RsUC.MoveNext
                    
                Loop
                
                    TxtUSChart.text = USChartFQA
                    
            End If
            
            If RsCC.RecordCount > 0 Then
            
                'If RsCCDefault.RecordCount > 0 Then
                '    CamChartFQA = RsCCDefault("FQA")
                'Else
                '    CamChartFQA = ""
                'End If
                TxtCamChart.RemoveAll
                Do While Not RsCC.EOF
                    
                    TxtCamChart.addITEM RsCC("FQA")
                    RsCC.MoveNext
                    
                Loop
                 
                TxtCamChart.text = CamChartFQA
                 
            End If
            
            TxtStockType = StockType
            TxtStockType.RemoveAll
            TxtStockType.addITEM "0000"

Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the combo boxes.", vbCritical, "Ims"
End Function


Public Function VerifyAddDeleteFQAFromGrid(stockno As String, Insert_delete As String, Tocondition As String, PONumb As String, lineno As String, quantity As String, Optional RowPositionToBedeleted As Integer, Optional onlyDetail As Boolean) As Boolean
Dim i As Integer
Dim Flag As Integer
Dim datax As ADODB.Recordset
Dim sql As String

On Error GoTo ErrHand

Insert_delete = UCase(Insert_delete)

    Select Case Insert_delete
    
    Case "INSERT"
    
        'If GDefaultValue = False Then
        'Juan 2010-11-3 To populate default values
        Set datax = New ADODB.Recordset
        sql = "SELECT * FROM PESYS WHERE psys_npecode='" + nameSP + "'"
        datax.Open sql, cn, adOpenStatic
        If datax.RecordCount = 0 Then
            GDefaultValue = LoadDefaultValuesForTOFQA(cell(1).tag, cell(3).tag, stockno)
        Else
            If Null = datax!EnableinventoryFQA Then
                GDefaultValue = LoadDefaultValuesForTOFQA(cell(1).tag, cell(3).tag, stockno)
            Else
                If datax!EnableinventoryFQA = True Then
                    GDefaultValue = LoadDefaultValuesForTOFQA(cell(1).tag, cell(3).tag, stockno)
                Else
                    Dim doHeader As Boolean
                    ' Juan 2011-11-13 to don't blank the header
                    If IsMissing(onlyDetail) Then
                        doHeader = True
                    Else
                        If onlyDetail = True Then
                            doHeader = False
                        Else
                            doHeader = True
                        End If
                    End If
                    If doHeader Then
                        'This variable list was here already, it just got validated- Juan
                        'Header values
                        TxtCompany.text = "0"
                        TxtLocation = "0"
                        TxtUSChart = "0"
                        TxtStockType = "0"
                        TxtCamChart = "0"
                    End If
                    '------------------------------------------
                    'Detail values
                    GDefaultFQA.Company = "0"
                    GDefaultFQA.CamChart = "0"
                    GDefaultFQA.Location = "0"
                    GDefaultFQA.StockType = "0"
                    GDefaultFQA.UsChart = "0"
                End If
            End If
        End If
        datax.Close
        '----------------------------
    
        Flag = 1
        
        'This is to check if the stockno is not repeatedly added again
    
''          For i = 0 To SSOleDBFQA.Rows
''
''            If STOCKNo = SSOleDBFQA.columns(0).Value And Tocondition = SSOleDBFQA.columns("tocond").Value Then
''
''                Flag = 0
''
''
''            End If
''
''            SSOleDBFQA.MoveNext
''
''          Next i
        
        If Flag = 1 Then
        
'''                Dim RsStockMaster As New ADODB.Recordset
                Dim StockType As String
'''                RsStockMaster.source = "select  isnull(stk_stcktype,'0000') stcktype from stockmaster where  stk_stcknumb ='" & stockno & "' and  stk_npecode ='" & nameSP & "'"
'''                RsStockMaster.Open , cn
'''
'''                If RsStockMaster.EOF = True Then
'''
'''
'''                    StockType = "0000"
'''
'''                ElseIf Len(Trim(RsStockMaster("stcktype"))) = 0 Then
'''
'''                        StockType = "0000"
'''                Else
'''                        StockType = Trim(RsStockMaster("stcktype"))
'''
'''                End If
            
            
            
                StockType = GDefaultFQA.StockType
                
                SSOleDBFQA.addITEM stockno & vbTab & GDefaultFQA.Company & vbTab & GDefaultFQA.Location & vbTab & GDefaultFQA.UsChart & vbTab & StockType & vbTab & GDefaultFQA.CamChart & vbTab & PONumb & vbTab & lineno & vbTab & Tocondition & vbTab & quantity
                
        End If
        
    Case "DELETE"
    SSOleDBFQA.MoveFirst
          For i = 0 To SSOleDBFQA.Rows - 1
          
                If stockno = SSOleDBFQA.columns(0).Value And i = RowPositionToBedeleted - 1 Then
                      
                      SSOleDBFQA.RemoveItem i
                      Exit Function
                
                End If
                
                SSOleDBFQA.MoveNext
          
          Next i
    
    End Select

Exit Function
ErrHand:


MsgBox "Errors occurred while trying to Insert\ Delete a record in the FQA grid.", vbCritical, "Ims"


End Function

Public Function LoadDefaultValuesForTOFQA(Companycode As String, LocationCode As String, stockno As String) As Boolean
' This function is used only for loading the TO FQA values
On Error GoTo ErrHand
LoadDefaultValuesForTOFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset
Dim stockprefix As String


LocationCode = UCase(Trim(LocationCode))

stockno = Trim(stockno)

stockprefix = Mid(stockno, 1, 2)

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Level ='C' and ""default"" =1"

RsCompany.Open , cn


If RsCompany.EOF = False Then
    
    GDefaultFQA.Company = RsCompany("FQA").Value
    
    Else
    
    GDefaultFQA.Company = ""
    
End If

' If this is an Issue then the To FQAs are free entry and should be set to "" and exit
If Me.tag = "02040200" Then

    GDefaultFQA.CamChart = ""
    GDefaultFQA.Location = ""
    GDefaultFQA.StockType = ""
    GDefaultFQA.UsChart = ""
    
    LoadDefaultValuesForTOFQA = True
    Exit Function
    
End If

'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS' and ""default"" =1"

RsLocation.Open , cn

If RsLocation.EOF = False Then
    
    GDefaultFQA.Location = RsLocation("FQA").Value
    
    Else
    
    GDefaultFQA.Location = ""
    
End If

If Me.tag = "02040100" Then 'Return from well and Warehouse to warehouse

'If it is PRD, DRL , CHM then get the defaults
If LocationCode = "PRD" Or LocationCode = "DRL" Or LocationCode = "CHM" Then

    'Get US Chart FQA
    RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC' and ""default"" =1"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "'  and Locationcode='" & Trim(LocationCode) & "' and Level ='CC' and ""default"" =1"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

'In case this is M&T or SUR then just give them free entry with no defaults
Else

    RsUC.source = "select FQA from FQA where 1=2"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where 1=2"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

End If

ElseIf Me.tag = "02040300" Or Me.tag = "02040600" Then 'Return from well and Warehouse to warehouse



    'Get US Chart FQA
    RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC' and ""default"" =1"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "'  and Locationcode='" & Trim(LocationCode) & "' and Level ='CC' and ""default"" =1"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

ElseIf Me.tag = "02040200" Or Me.tag = "02040500" Or Me.tag = "02050200" Or Me.tag = "02050300" Then  'Warehouse Issue , "well to well" , "Write on", "write off"

    RsUC.source = "select FQA from FQA where 1=2"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where 1=2"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

Else

    RsUC.source = "select FQA from FQA where 1=2"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where 1=2"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

End If

 
                        If RsCC.EOF = False Then
                            
                            GDefaultFQA.CamChart = RsCC("FQA").Value
                            
                            Else
                            
                            GDefaultFQA.CamChart = "00000"
                            
                        End If
                        
                        If RsUC.EOF = False Then
                            
                            GDefaultFQA.UsChart = RsUC("FQA").Value
                            
                            Else
                            
                            GDefaultFQA.UsChart = "00000"
                            
                        End If
        
        LocationCode = Trim(UCase(LocationCode))
        
''        If (LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL") And _
''        (stockprefix = "55" Or stockprefix = "66") And Me.tag = "02040100" Then
''
''            GDefaultFQA.stocktype = "0000"
''
''        ElseIf (LocationCode = "SUR") And (stockprefix = "44" Or stockprefix = "88") And Me.tag = "02040100" Then
''
''            GDefaultFQA.stocktype = "0000"
''
''        Else
''
''            GDefaultFQA.stocktype = ""
''
''        End If
        
    ' 'WellToWell,
    If (stockprefix = "55" Or stockprefix = "66" Or stockprefix = "88" Or stockprefix = "44") And (Me.tag <> "02040500") Then
    
        GDefaultFQA.StockType = "0000"
     
    Else
    
        GDefaultFQA.StockType = "0000"
    
    End If
        
        'If LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL" Then GDefaultFQA.stocktype = "0000"
        
        'If LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL" Then GDefaultFQA.stocktype = "0000"
        


Set RsCompany = Nothing
Set RsLocation = Nothing
Set RsUC = Nothing
Set RsCC = Nothing

LoadDefaultValuesForTOFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to get the default values." & Err.description, vbCritical, "Ims"

Err.Clear

End Function

Public Function LoadDefaultValuesForFROMFQA(Companycode As String, LocationCode As String, stockno As String) As Boolean
' This function is used only for loading the FROM FQA values for all the transaction except RECEIPTS
On Error GoTo ErrHand
LoadDefaultValuesForFROMFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset
Dim stockprefix As String


stockno = Trim(stockno)

stockprefix = Mid(stockno, 1, 2)

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Level ='C' and ""default"" =1"

RsCompany.Open , cn

If RsCompany.EOF = False Then
    
    GDefaultFQA.Company = RsCompany("FQA").Value
    
    Else
    
    GDefaultFQA.Company = ""
    
End If

'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS' and ""default"" =1"

RsLocation.Open , cn

If RsLocation.EOF = False Then
    
    GDefaultFQA.Location = RsLocation("FQA").Value
    
    Else
    
    GDefaultFQA.Location = ""
    
End If

If Me.tag = "02040200" Or Me.tag = "02040600" Or Me.tag = "02050300" Then  'Warehouse Issue, Warehouse to warehouse,  "write off"

    'Get US Chart FQA
    RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC' and ""default"" =1"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "'  and Locationcode='" & Trim(LocationCode) & "' and Level ='CC' and ""default"" =1"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn

ElseIf Me.tag = "02040300" Or Me.tag = "02040500" Or Me.tag = "02050200" Then   'Return from well, "well to well" , "Write on"

    'Have to be left blank, user will entery the US and Cameroon CC.
    'Since RSCC and RSUC are EOF , the steps below will set US and Cam Charts to ""
    
        'Get US Chart FQA
    RsUC.source = "select FQA from FQA where 1=2"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where 1=2"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn
    
    ' this might be the case of Internal Transfer and ... and left over other if there are any
Else

    'Have to be left blank, user will entery the US and Cameroon CC.
    'Since RSCC and RSUC are EOF , the steps below will set US and Cam Charts to ""
    
        'Get US Chart FQA
    RsUC.source = "select FQA from FQA where 1=2"
    'RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='UC' and ""default"" =1"
    RsUC.Open , cn
    
    'Get Cam Chart FQA
    RsCC.source = "select FQA from FQA where 1=2"
    'RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and stockprefix ='" & stockprefix & "' and Level ='CC' and ""default"" =1"
    RsCC.Open , cn


End If



    If RsCC.EOF = False Then
        
        GDefaultFQA.CamChart = RsCC("FQA").Value
        
        Else
        
        GDefaultFQA.CamChart = ""
        
    End If
    
    If RsUC.EOF = False Then
        
        GDefaultFQA.UsChart = RsUC("FQA").Value
        
        Else
        
        GDefaultFQA.UsChart = ""
        
    End If
    
    LocationCode = Trim(UCase(LocationCode))
    
    'Getting the StockTypes
    
    If (stockprefix = "55" Or stockprefix = "66" Or stockprefix = "88" Or stockprefix = "44") And (Me.tag <> "02040500" Or Me.tag <> "02040300") Then
    
        GDefaultFQA.StockType = "0000"
     
    Else
    
        GDefaultFQA.StockType = "0000"
    
    End If
        
        'If it is a receipt then
        'If (LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL") And _
        '(stockprefix = "55" Or stockprefix = "66") And Me.tag = "02040100" Then
       '
       '     GDefaultFQA.stocktype = "0000"
            
       ' ElseIf (LocationCode = "SUR") And (stockprefix = "44" Or stockprefix = "88") And Me.tag = "02040100" Then
        
       '     GDefaultFQA.stocktype = "0000"
        
       ' Else
        
        '    GDefaultFQA.stocktype = ""
       '
       ' End If
        
        'If LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL" Then GDefaultFQA.stocktype = "0000"
        
        'If LocationCode = "PRD" Or LocationCode = "CHM" Or LocationCode = "DRL" Then GDefaultFQA.stocktype = "0000"
        


Set RsCompany = Nothing
Set RsLocation = Nothing
Set RsUC = Nothing
Set RsCC = Nothing

LoadDefaultValuesForFROMFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to get the default values." & Err.description, vbCritical, "Ims"

Err.Clear

End Function

Public Function ChangeMode(ReadOnly As Boolean)

SSOleDBFQA.Enabled = Not ReadOnly

TxtCompany.Enabled = Not ReadOnly
TxtLocation.Enabled = Not ReadOnly
TxtUSChart.Enabled = Not ReadOnly
TxtCamChart.Enabled = Not ReadOnly
TxtStockType.Enabled = Not ReadOnly

End Function

Public Function PopulateFROMCombosWithFQA(RsCompany As ADODB.Recordset, RsLocation As ADODB.Recordset, RsUC As ADODB.Recordset, RsCC As ADODB.Recordset) As Boolean

On Error GoTo ErrHand

PopulateFROMCombosWithFQA = False

'Get Company FQA

If RsCompany.RecordCount > 0 Then
Do While Not RsCompany.EOF

    TxtCompany.addITEM RsCompany("FQA")
    RsCompany.MoveNext
    
Loop
End If

If RsLocation.RecordCount > 0 Then
Do While Not RsLocation.EOF

    TxtLocation.addITEM RsLocation("FQA")
    RsLocation.MoveNext
    
Loop
End If
'Get US Chart FQA
If RsUC.RecordCount > 0 Then
Do While Not RsUC.EOF

    TxtUSChart.addITEM RsUC("FQA")
    RsUC.MoveNext
    
Loop
End If
'Get Cam Chart FQA

If RsCC.RecordCount > 0 Then
Do While Not RsCC.EOF

    TxtCamChart.addITEM RsCC("FQA")
    RsCC.MoveNext
    
Loop
End If

PopulateFROMCombosWithFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes." & Err.description, vbCritical, "Ims"

Err.Clear

End Function

Private Sub TxtCompany_Validate(Cancel As Boolean)
If Len(Trim$(TxtCompany.text)) > 0 And Not TxtCompany.IsItemInList Then
  Cancel = True
   TxtCompany.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub TxtCompany_LostFocus()
'Call NormalBackground(TxtCompany)
End Sub

Private Sub TxtCompany_KeyDown(KeyCode As Integer, Shift As Integer)


If Not TxtCompany.DroppedDown Then TxtCompany.DroppedDown = True
End Sub

Private Sub TxtCompany_GotFocus()
TxtCompany.SelStart = 0
TxtCompany.SelLength = 0
' Call HighlightBackground(TxtCompany)
End Sub

'-----------
Private Sub TxtLocation_Validate(Cancel As Boolean)
If Len(Trim$(TxtLocation.text)) > 0 And Not TxtLocation.IsItemInList Then
  Cancel = True
   TxtLocation.SetFocus
   
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub TxtLocation_LostFocus()
'Call NormalBackground(TxtLocation)
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)

If Not TxtLocation.DroppedDown Then TxtLocation.DroppedDown = True
End Sub

Private Sub TxtLocation_GotFocus()
TxtLocation.SelStart = 0
TxtLocation.SelLength = 0
' Call HighlightBackground(TxtLocation)
End Sub
'---------------
Private Sub TxtUSChart_Validate(Cancel As Boolean)
If Len(Trim$(TxtUSChart.text)) > 0 And Not TxtUSChart.IsItemInList Then
  Cancel = True
   TxtUSChart.SetFocus
   
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub TxtUSChart_LostFocus()
'Call NormalBackground(TxtUSChart)
End Sub

Private Sub TxtUSChart_KeyDown(KeyCode As Integer, Shift As Integer)


If Not TxtUSChart.DroppedDown Then TxtUSChart.DroppedDown = True
End Sub

Private Sub TxtUSChart_GotFocus()
TxtUSChart.SelStart = 0
TxtUSChart.SelLength = 0
' Call HighlightBackground(TxtUSChart)
End Sub

'---------------
Private Sub TxtCamChart_Validate(Cancel As Boolean)
If Len(Trim$(TxtCamChart.text)) > 0 And Not TxtCamChart.IsItemInList Then
  Cancel = True
   TxtCamChart.SetFocus
   
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub TxtCamChart_LostFocus()
'Call NormalBackground(TxtCamChart)
End Sub

Private Sub TxtCamChart_KeyDown(KeyCode As Integer, Shift As Integer)


If Not TxtCamChart.DroppedDown Then TxtCamChart.DroppedDown = True
End Sub

Private Sub TxtCamChart_GotFocus()
TxtCamChart.SelStart = 0
TxtCamChart.SelLength = 0
' Call HighlightBackground(TxtCamChart)
End Sub

'---------------
'Private Sub TxtStockType_Validate(Cancel As Boolean)
'If Len(Trim$(TxtStockType.text)) > 0 And Not TxtStockType.IsItemInList Then
'  Cancel = True
'   TxtStockType.SetFocus
'
' MsgBox "Invalid Value", , "Imswin"
'End If
'End Sub

Private Sub TxtStockType_LostFocus()
'Call NormalBackground(TxtStockType)
End Sub

Private Sub TxtStockType_KeyDown(KeyCode As Integer, Shift As Integer)
If Not TxtStockType.DroppedDown Then TxtStockType.DroppedDown = True
End Sub

Private Sub TxtStockType_GotFocus()
TxtStockType.SelStart = 0
TxtStockType.SelLength = 0
 'Call HighlightBackground(TxtStockType)
End Sub

Public Function ValidateFromFqa() As Boolean
On Error GoTo ErrHand

If Me.tag = "02040700" Then ValidateFromFqa = True: Exit Function 'InternalTransfer

If Len(Trim(TxtCompany)) = 0 Then
        
        Screen.MousePointer = 0
        MsgBox "Please fill out the FROM FQA Company for this transaction.", vbCritical, "Ims"
        If TxtCompany.Enabled = True Then TxtCompany.backcolor = &HC0FFFF
        If TxtCompany.Enabled = True Then TxtCompany.SetFocus
        Exit Function
        
    End If
    
    If Len(Trim(TxtLocation.text)) = 0 Then
    
        Screen.MousePointer = 0
        MsgBox "Please fill out the FROM FQA Location for this transaction.", vbCritical, "Ims"
        If TxtLocation.Enabled = True Then TxtLocation.backcolor = &HC0FFFF
        If TxtLocation.Enabled = True Then TxtLocation.SetFocus
        Exit Function
    
    End If
    
    If Len(Trim(TxtStockType.text)) = 0 Then
    
            Screen.MousePointer = 0
        MsgBox "Please fill out the FROM FQA StockType for this transaction.", vbCritical, "Ims"
        If TxtStockType.Enabled = True Then TxtStockType.backcolor = &HC0FFFF
        If TxtStockType.Enabled = True Then TxtStockType.SetFocus
        Exit Function
    
    End If
    
    If Len(Trim(TxtUSChart.text)) = 0 Then
    
            Screen.MousePointer = 0
        MsgBox "Please fill out the FROM FQA US Chart for this transaction.", vbCritical, "Ims"
        If TxtUSChart.Enabled = True Then TxtUSChart.backcolor = &HC0FFFF
        If TxtUSChart.Enabled = True Then TxtUSChart.SetFocus
        Exit Function
    
    End If
    
    If Len(Trim(TxtCamChart.text)) = 0 Then
    
            Screen.MousePointer = 0
        MsgBox "Please fill out the FROM FQA Cam Chart for this transaction.", vbCritical, "Ims"
        If TxtCamChart.Enabled = True Then TxtCamChart.backcolor = &HC0FFFF
        If TxtCamChart.Enabled = True Then TxtCamChart.SetFocus
        Exit Function
    
    End If
    
ValidateFromFqa = True

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to validate FQA. Err Desc : " & Err.description
Err.Clear

End Function
Public Function ValidateTOFqa() As Boolean
Dim i As Integer
On Error GoTo ErrHand

If Me.tag = "02040700" Then ValidateTOFqa = True: Exit Function 'InternalTransfer

SSOleDBFQA.MoveFirst

For i = 0 To SSOleDBFQA.Rows - 1

    If Len(SSOleDBFQA.columns(1).text & "") = 0 Then
        
        
        MsgBox "Please make sure that all the TO FQA Company codes are entered.", vbCritical, "Ims"
        Exit Function
        
    End If
    
    If Len(SSOleDBFQA.columns(2).text & "") = 0 Then
    
        
        MsgBox "Please make sure that all the TO FQA Location codes are entered.", vbCritical, "Ims"
        
        Exit Function
    
    End If
    
    If Len(SSOleDBFQA.columns(3).text & "") = 0 Then
    
        MsgBox "Please make sure that all the TO FQA StockType codes are entered.", vbCritical, "Ims"
    
        Exit Function
    
    End If
    
    If Len(SSOleDBFQA.columns(4).text & "") = 0 Then
    
    
        MsgBox "Please make sure that all the TO FQA US Chart codes are entered.", vbCritical, "Ims"
    
        Exit Function
    
    End If
    
    If Len(SSOleDBFQA.columns(5).text & "") = 0 Then
    
    
        MsgBox "Please make sure that all the TO FQA Cameroon chart codes are entered.", vbCritical, "Ims"
    
        Exit Function
    
    End If
    
    SSOleDBFQA.MoveNext
    
Next i

ValidateTOFqa = True

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to validate FQA. Err Desc : " & Err.description
Err.Clear

End Function

'Created by Muzammil
Public Function HasUserSelectedAnyStocks() As Boolean
On Error GoTo ErrHand

If SUMMARYlist.Rows > 2 Then
 HasUserSelectedAnyStocks = True

ElseIf SUMMARYlist.Rows = 2 And SUMMARYlist.TextMatrix(1, 0) <> "" Then
 HasUserSelectedAnyStocks = True

ElseIf SUMMARYlist.Rows = 2 And SUMMARYlist.TextMatrix(1, 0) = "" Then
 HasUserSelectedAnyStocks = False


End If
Exit Function
ErrHand:
    
    MsgBox "Errors occurred. Err Description :" & Err.description
    Err.Clear

End Function


Private Sub yesButton_Click()
    imsMsgBox.Visible = False
    msgBoxResponse = True
End Sub


