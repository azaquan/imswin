VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmWarehouse 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Tag             =   "02050700"
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBCamChart 
      Height          =   735
      Left            =   8520
      TabIndex        =   83
      Top             =   7680
      Width           =   1455
      DataFieldList   =   "Column 0"
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
      TabIndex        =   82
      Top             =   7560
      Width           =   975
      DataFieldList   =   "Column 0"
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
      TabIndex        =   81
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
      TabIndex        =   80
      Top             =   6960
      Width           =   975
      DataFieldList   =   "Column 0"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBFQA 
      Height          =   2340
      Left            =   120
      TabIndex        =   78
      Top             =   3840
      Width           =   11775
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   2566
      Columns(0).Caption=   "StockNumber"
      Columns(0).Name =   "StockNumber"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Company"
      Columns(1).Name =   "Company"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1852
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
      Columns(4).Width=   1640
      Columns(4).Caption=   "StockType"
      Columns(4).Name =   "StockType"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2566
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
      Columns(8).Width=   1746
      Columns(8).Caption=   "Condition"
      Columns(8).Name =   "ToCond"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1402
      Columns(9).Caption=   "Quantity"
      Columns(9).Name =   "Quantity"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   20770
      _ExtentY        =   4128
      _StockProps     =   79
      Caption         =   "FQA"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleCompany 
      Height          =   735
      Left            =   8520
      TabIndex        =   79
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
      Height          =   1980
      Left            =   120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   6240
      Width           =   11775
   End
   Begin VB.TextBox TxtCompany 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1080
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   1320
      Width           =   450
   End
   Begin VB.TextBox TxtLocation 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1320
      Width           =   810
   End
   Begin VB.TextBox TxtUSChart 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1170
   End
   Begin VB.TextBox TxtStockType 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   1320
      Width           =   570
   End
   Begin VB.TextBox TxtCamChart 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid StockListDuplicate 
      Height          =   1740
      Left            =   120
      TabIndex        =   68
      Top             =   2080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3069
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   0
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
   Begin VB.PictureBox savingLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4320
      ScaleHeight     =   945
      ScaleWidth      =   3105
      TabIndex        =   64
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
         TabIndex        =   65
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   240
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5520
      Top             =   8280
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
      Left            =   6000
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   8320
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   5
      Left            =   1560
      TabIndex        =   59
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
      TabIndex        =   60
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   54
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
      TabIndex        =   52
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
      TabIndex        =   51
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
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   8320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E-Mail"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   8320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3480
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8320
      Width           =   1575
   End
   Begin VB.CommandButton removeDETAIL 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   10080
      TabIndex        =   46
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
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "priceBOX"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton hideDETAIL 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   42
      Top             =   3870
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton submitDETAIL 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   11040
      TabIndex        =   41
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
      Text            =   "logicBOX"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox linesH 
      Height          =   15
      Index           =   0
      Left            =   960
      ScaleHeight     =   15
      ScaleWidth      =   10650
      TabIndex        =   33
      Top             =   4920
      Visible         =   0   'False
      Width           =   10650
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
      Left            =   8640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8320
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
      Left            =   10320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8320
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
      Format          =   58785795
      CurrentDate     =   36867
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid STOCKlist 
      Height          =   1740
      Left            =   120
      TabIndex        =   12
      Top             =   2080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3069
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid matrix 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   450
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
      TabIndex        =   40
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
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   15
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid detailHEADER 
      Height          =   300
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   11775
      _ExtentX        =   20770
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
      Height          =   3660
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6456
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid SUMMARYlist 
      Height          =   4380
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7726
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
   Begin VB.Label LBLCompany 
      Caption         =   "Company"
      Height          =   255
      Left            =   120
      TabIndex        =   84
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LblUSChart 
      Caption         =   "US Chart#"
      Height          =   255
      Left            =   3255
      TabIndex        =   77
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LblLocation 
      Caption         =   "Location"
      Height          =   255
      Left            =   1560
      TabIndex        =   76
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LblType 
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   75
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LblCamChart 
      Caption         =   "Cam. Chart #"
      Height          =   255
      Left            =   6840
      TabIndex        =   74
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Search Field"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   67
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Search Field"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   66
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
      TabIndex        =   62
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   61
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   1500
      X2              =   1500
      Y1              =   4800
      Y2              =   3840
   End
   Begin VB.Label otherLABEL 
      Alignment       =   1  'Right Justify
      Caption         =   "New Commodity:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   56
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   8040
      TabIndex        =   55
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label summaryLABEL 
      Caption         =   "Summary"
      Height          =   255
      Left            =   120
      TabIndex        =   53
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      Y1              =   1700
      Y2              =   1700
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
      Left            =   -120
      TabIndex        =   23
      Top             =   7680
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
      Width           =   975
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
      Caption         =   "Search Transaction #"
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


Sub fillGRID(ByRef grid As MSHFlexGrid, box As textBOX, Index)
'On Error Resume Next
Dim paraVECTOR
Dim i, n, list, size, totalwidth, cols, wide(), title(), extraW, sql, clue, Flag
Dim datax As New ADODB.Recordset
    Err.Clear
    Screen.MousePointer = 11
    Select Case box.name
        Case "logicBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 3000
            wide(1) = 1200
            ReDim title(2)
            title(0) = "Logical Warehouse"
            title(1) = "Code"
            sql = "select lw_code Code , lw_desc Description from LOGWAR" _
                & " where lw_actvflag = 1 AND lw_npecode = '" & nameSP & "' order by lw_desc "
            Set datax = New ADODB.Recordset
            list = Array("description", "code")
        Case "sublocaBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 3000
            wide(1) = 1200
            ReDim title(2)
            title(0) = "Sub Location"
            title(1) = "Code"
            Set datax = getDATA("getSUBLOCA", nameSP)
            list = Array("description", "code")
        Case "NEWconditionBOX"
            clue = "Code"
            cols = 2
            ReDim wide(2)
            wide(0) = 500
            wide(1) = 2100
            ReDim title(2)
            title(0) = "Code"
            title(1) = "Condition"
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
        
        .Height = 1455
        extraW = 270
        .ScrollBars = flexScrollBarVertical
        If box.width > (totalwidth + extraW) Then
            .width = box.width
            .ColWidth(0) = .ColWidth(0) + (.width - totalwidth) - extraW
        Else
            .width = totalwidth + extraW
        End If
        .tag = Format(Index, "00") + box.name
        
        n = 1
        .Rows = datax.RecordCount + 1
        Do While Not datax.EOF
            .row = n
            For i = 0 To cols - 1
                .TextMatrix(n, i) = Trim(Format(datax(list(i))))
            Next
            If datax(clue) = box.tag Then
                Flag = .Rows - 1
            End If
            If Tree.Nodes.Count > 0 Then
                If n = 6 And datax.RecordCount > 10 Then
                    Call showGRID(grid, Index, box, True)
                    Screen.MousePointer = 11
                    grid.Refresh
                End If
            End If
            datax.MoveNext
            n = n + 1
        Loop
        .row = Flag
        If .Rows < 6 Then
            .Height = 240 * .Rows
            extraW = 0
            .ScrollBars = flexScrollBarNone
        End If
    End With
    Screen.MousePointer = 0
End Sub

Sub fillCOMBO(ByRef grid As MSHFlexGrid, Index)
'On Error Resume Next
Dim paraVECTOR, sql
Dim i, n, params, shot, x, spot, rec, list, list2, size, totalwidth, extraW, align, clue
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
            params = matrix.TextMatrix(6, Index)
            Call doARRAYS("s", params, list)
            Call doARRAYS("s", matrix.TextMatrix(2, Index), list2)
            n = UBound(list)
            
            For i = 0 To n
                If params = "" Then
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
                If i = 0 Then
                    combo(Index).Rows = 2
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
                & "po_docutype IN ('P', 'O', 'L', 'W', 'S') AND " _
                & "((po_freigforwr=1 and  po_stasdelv in('RP','RC')) or (po_freigforwr=0) and po_stasinvt <> 'IC')"
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
            params = matrix.TextMatrix(6, Index)
            If params <> "" Then If Len(params) = 0 Then Exit Sub
            If Err.Number = 0 Then
                n = howMANY(matrix.TextMatrix(6, Index), ",")
                ReDim paraVECTOR(n)
                paraVECTOR(0) = ""
                For i = 0 To n
                    x = InStr(params, ",") - 1
                    If x < 0 Then x = Len(params)
                    spot = Trim(Left(params, x))
                    If Left(spot, 1) = "@" Then
                        If UCase(Left(spot, 5)) = "@CELL" Then
                            spot = cell(Val(Mid(spot, 7, 1))).tag
                        Else
                            spot = cell(Val(Mid(spot, 2, 1)))
                        End If
                    End If
                    paraVECTOR(i) = Trim(spot)
                    If InStr(params, ",") > 0 Then
                        params = Mid(params, x + 2)
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
    Call makeLISTS
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
                rowTEXT = rowTEXT + FormatNumber(!QTY1, 2) + vbTab 'Primary Quantity
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
                    rowTEXT = rowTEXT + FormatNumber(!QTY2, 2) + vbTab 'Secundary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Secundary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice2), 0, !UnitPrice2), 2) + vbTab 'Secundary Unit Price
                    
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
                    STOCKlist = "§"
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

Sub hideGRIDS()
Dim i
    For i = 0 To 2
        grid(i).Visible = False
    Next
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

Sub makeLISTS()
Dim i, col, c, dark As Integer
    For i = 0 To 4
        If cell(i).Visible Then cell(i).tabindex = i
    Next
    STOCKlist.tabindex = 5
    Tree.tabindex = 6
    
    dark = 1
    With STOCKlist
        .Clear
        .Rows = 2
        .cols = 8
        .ColWidth(0) = 485
        .row = 0
        .col = 0
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Commodity"
        .ColWidth(1) = 1400
        For i = 1 To .cols - 1
            .ColAlignment(i) = 0
            .ColAlignmentFixed(i) = 4
        Next
        'cc
        .ColAlignment(2) = 6
        Select Case frmWarehouse.tag
            'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                dark = 1
                .TextMatrix(0, 2) = "Unit Price"
                .ColWidth(2) = 1000
                .TextMatrix(0, 3) = "Description"
                .ColWidth(3) = 6200
                .TextMatrix(0, 4) = "Unit"
                .ColWidth(4) = 1200
                .ColAlignment(5) = 6
                .TextMatrix(0, 5) = "Qty"
                .ColWidth(5) = 1200
                .ColWidth(6) = 0
                .ColWidth(7) = 0
            Case "02050200" 'AdjustmentEntry
                dark = 0
                '.cols = 4
                .TextMatrix(0, 2) = "Description"
                .ColAlignment(2) = 0
                .ColWidth(2) = 8400
                .TextMatrix(0, 3) = "Unit"
                .ColWidth(3) = 1200
            Case "02040100" 'WarehouseReceipt
                dark = 1
                .ColAlignment(2) = 6
                .ColAlignment(3) = 6
                .ColAlignment(4) = 4
                .TextMatrix(0, 2) = "Purchase QTY"
                .ColWidth(2) = 1100
                .TextMatrix(0, 3) = "QTY to Rec."
                .ColWidth(3) = 1100
                .TextMatrix(0, 4) = "Unit"
                .ColWidth(4) = 1200
                .TextMatrix(0, 5) = "Description"
                .ColWidth(5) = 6200
             '   .TextMatrix(0, 6) = "Item #"
                .ColWidth(6) = 0
                .ColWidth(7) = 0
        End Select
        .TextMatrix(0, 6) = "Initial Value"
        .ColWidth(7) = 0
        .RowHeight(0) = 240
        .RowHeightMin = 0
        .RowHeight(1) = 0
        .WordWrap = True
        .tag = ""
    End With
    
    With detailHEADER
        .cols = 7
        c = 7
        .TextMatrix(0, 0) = "Condition / Logical Warehouse / Sublocation"
        .TextMatrix(0, 1) = "QTY"
        .TextMatrix(0, 2) = "Logical Warehouse"
        .TextMatrix(0, 3) = "Sublocation"
        .TextMatrix(0, 4) = "QTY"
        .TextMatrix(0, 5) = "Balance"
        .ColWidth(0) = 4800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1830
        .ColWidth(3) = 1830
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 260
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                .cols = 9
                c = 9
                .TextMatrix(0, 2) = "Logical Ware."
                .TextMatrix(0, 4) = "Condition"
                .TextMatrix(0, 5) = "Repair Cost"
                .TextMatrix(0, 6) = "QTY"
                .TextMatrix(0, 7) = "Balance"
                .ColWidth(1) = 950
                .ColWidth(2) = 1040
                .ColWidth(3) = 1040
                .ColWidth(4) = 800
                .ColWidth(5) = 950
                .ColWidth(6) = 950
                .ColWidth(7) = 950
                .ColWidth(8) = 260
            Case "02050200" 'AdjustmentEntry
                .cols = 6
                c = 6
                .TextMatrix(0, 0) = "Condition"
                .TextMatrix(0, 1) = "Logical Warehouse"
                .TextMatrix(0, 2) = "Sublocation"
                .TextMatrix(0, 3) = "Unit Price"
                .TextMatrix(0, 4) = "QTY"
                .TextMatrix(0, 5) = ""
                .ColWidth(0) = 5000
                .ColWidth(1) = 2230
                .ColWidth(2) = 2230
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 260
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                dark = 0
                .TextMatrix(0, 0) = "Condition / Serial"
                .TextMatrix(0, 1) = "Logical Warehouse"
                .TextMatrix(0, 2) = "Sublocation"
                .TextMatrix(0, 3) = "QTY"
                .TextMatrix(0, 4) = ""
                .ColWidth(0) = 5500
                .ColWidth(1) = 2470
                .ColWidth(2) = 2470
                .ColWidth(3) = 1000
                .ColWidth(4) = 260
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
                .cols = 7
                .TextMatrix(0, 4) = "Condition"
                .TextMatrix(0, 5) = "QTY"
                .TextMatrix(0, 6) = "Balance"
                .ColWidth(0) = 4400
                .ColWidth(2) = 1530
                .ColWidth(3) = 1530
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
        .cols = 21
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
        .TextMatrix(0, 1) = "Commodity"
        .ColWidth(1) = 1400
        .TextMatrix(0, 2) = "Serial"
        .ColWidth(2) = 800
        .TextMatrix(0, 3) = "Condition"
        .ColWidth(3) = 1000
        .TextMatrix(0, 4) = "Unit Price"
        .ColWidth(4) = 1200
        .TextMatrix(0, 5) = "Description"
        .ColWidth(5) = 4400
        .TextMatrix(0, 6) = "Unit"
        .ColWidth(6) = 1200
        .TextMatrix(0, 7) = "Qty"
        .ColWidth(7) = 1200
        .TextMatrix(0, 8) = "node"
        .TextMatrix(0, 9) = "From Logical"
        .TextMatrix(0, 10) = "From Subloca"
        .TextMatrix(0, 11) = "To Logical"
        .TextMatrix(0, 12) = "To Subloca"
        .TextMatrix(0, 13) = "New Condition Code"
        .TextMatrix(0, 14) = "New Condition Description"
        .TextMatrix(0, 15) = "Unit Code"
        .TextMatrix(0, 16) = "Computer Factor"
        .TextMatrix(0, 20) = "Original Condition Code"
        c = 8
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                .TextMatrix(0, 17) = "repaircost"
                .TextMatrix(0, 18) = "newcomodity"
                .TextMatrix(0, 19) = "newdescription"
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
                .TextMatrix(0, 17) = "originalcondition"
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                .cols = .cols + 2
                .TextMatrix(0, 17) = "QTYpo"
                .TextMatrix(0, 21) = "PO"
                .TextMatrix(0, 22) = "lineitem"
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        For i = c To .cols
            .ColWidth(i) = 0
        Next
        .RowHeight(0) = 240
        .RowHeightMin = 0
        .RowHeight(1) = 0
        .WordWrap = True
        .tag = ""
        .ZOrder
    End With
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




Sub searchIN(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
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

Sub showCOMBO(ByRef grid As MSHFlexGrid, Index)
    With grid
        Call fillCOMBO(grid, Index)
        If .Rows > 0 And .text <> "" Then
            .Visible = True
            .ZOrder
            If Index < 5 Then .Top = cell(Index).Top + 270
        End If
    End With
End Sub

Sub hideREMARKS()
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
    removeDETAIL.Visible = True
    submitDETAIL.Visible = True
    Tree.Visible = True 'M
    SSOleDBFQA.Visible = False
    sublocaBOX(0).Visible = True ' M
    grid(2).Visible = True 'M
End Sub

Sub showREMARKS()
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
    remarks.Top = SSOleDBFQA.Top + SSOleDBFQA.Height + 70    'detailHEADER.Top
    remarks.Height = 1980 'Tree.Top - detailHEADER.Top + Tree.Height - SSOleDBFQA.Height
    remarksLABEL.Visible = True
    remarks.Visible = True
    remarks.ZOrder
    
    SSOleDBFQA.Visible = True 'M
    SSOleDBFQA.ZOrder 'M
    sublocaBOX(0).Visible = False ' M
    summaryLABEL.Top = SUMMARYlist.Top - 240
    hideDETAIL.Visible = False
    removeDETAIL.Visible = False
    submitDETAIL.Visible = False
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
                If .TextMatrix(.row, 1) = "§" Then
                    If .col = 10 Then Exit Sub
                End If
            Else
                If .TextMatrix(.row, 1) <> "§" Then
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
    logicBOX(currentBOX).backcolor = vbWhite
    sublocaBOX(currentBOX).backcolor = vbWhite
    quantityBOX(currentBOX).backcolor = vbWhite
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
Dim nody As Node
    With Tree
        n = .SelectedItem.Index + .SelectedItem.Children
        Call moveBOXES(n, 1)
        .Nodes.Add .SelectedItem.key, tvwChild, .SelectedItem.key + "{{Serial", "Serial ", "thing 1"
        .Nodes(.SelectedItem.key + "{{Serial").Selected = True
        .StartLabelEdit
    End With
End Sub

Private Sub balanceBOX_GotFocus(Index As Integer)
    activeBOX = "balanceBOX"
End Sub


Private Sub cell_Change(Index As Integer)
Dim n As Integer
    If Not directCLICK Then
        n = Val(matrix.TextMatrix(10, 0))
        Call alphaSEARCH(cell(Index), combo(Index), n)
    Else
        directCLICK = False
    End If
    combo(Index).MousePointer = 0
End Sub

Private Sub combo_EnterCell(Index As Integer)
    cell(Index) = combo(Index).TextMatrix(combo(Index).row, 0)
    If Index = 5 Then
        cell(Index).tag = combo(Index).TextMatrix(combo(Index).row, 0)
    Else
        cell(Index).tag = combo(Index).TextMatrix(combo(Index).row, Val(matrix.TextMatrix(10, Index)))
    End If
End Sub

Private Sub combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    justCLICK = False
    With cell(Index)
        If Not .locked Then
            Select Case KeyCode
                Case 40
                    direction = "down"
                Case 38
                    direction = "up"
            End Select
            'cell(Index) = combo(Index).TextMatrix(combo(Index).row, 0)
        End If
    End With
End Sub

Private Sub combo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call combo_Click(Index)
        Case 27
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

Private Sub dateBOX_GotFocus()
    activeBOX = "dateBOX"
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

Private Sub grid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    justCLICK = False
    Select Case KeyCode
        Case 40
            direction = "down"
        Case 38
            direction = "up"
    End Select
End Sub

Private Sub logicBOX_Change(Index As Integer)
    Call alphaSEARCH(logicBOX(Index), grid(1), 0)
End Sub

Private Sub logicBOX_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40
            direction = "down"
            Call showGRID(grid(1), Index, logicBOX(Index), True)
            Call arrowKEYS(Index, logicBOX(Index), grid(1))
        Case 38
            direction = "up"
            Call showGRID(grid(1), Index, logicBOX(Index), True)
            Call arrowKEYS(Index, logicBOX(Index), grid(1))
    End Select
End Sub

Private Sub NEWconditionBOX_Change(Index As Integer)
    Call alphaSEARCH(NEWconditionBOX(Index), grid(0), 0)
End Sub

Private Sub NEWconditionBOX_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40
            direction = "down"
            Call showGRID(grid(0), Index, NEWconditionBOX(Index), True)
            Call arrowKEYS(Index, NEWconditionBOX(Index), grid(0))
        Case 38
            direction = "up"
            Call showGRID(grid(0), Index, NEWconditionBOX(Index), True)
            Call arrowKEYS(Index, NEWconditionBOX(Index), grid(0))
    End Select
End Sub

Private Sub NEWconditionBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call grid_Click(0)
            grid(0).Visible = False
            If quantityBOX(Index).Visible Then
                quantityBOX(Index).SetFocus
                sublocaBOX(Index).backcolor = vbWhite
                Exit Sub
            End If
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
    activeBOX = "priceBOX"
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

Private Sub priceBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index > 0 And Index <> totalNODE Then
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

Private Sub saveBUTTONold_Click()
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
Dim fromSUBLOCA As String
Dim toLOGIC As String
Dim toSUBLOCA As String
Dim condition As String
Dim NEWcondition As String
Dim unitprice As Double
Dim newUNITprice As Double
Dim serial As String
Dim ComputerFactor
Dim imsLock As imsLock.Lock
Dim TranType As String
Screen.MousePointer = 11
Dim data As New ADODB.Recordset
    
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
        
    If remarks = "" Then
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please include the remarks for this transaction"
        remarks.backcolor = &HC0FFFF
        remarks.SetFocus
        Exit Sub
    End If
    
    If SSOleDBFQA.Rows = 0 Then
    
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please fill out the FQA values for this transaction.", vbCritical, "Ims"
        'remarks.backcolor = &HC0FFFF
        'remarks.SetFocus
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
            unitprice = CDbl(IIf(SUMMARYlist.TextMatrix(i, 4) = "", 0, SUMMARYlist.TextMatrix(i, 4)))
            condition = SUMMARYlist.TextMatrix(i, 20)
            fromlogic = SUMMARYlist.TextMatrix(i, 9)
            fromSUBLOCA = SUMMARYlist.TextMatrix(i, 10)
            toLOGIC = SUMMARYlist.TextMatrix(i, 11)
            toSUBLOCA = SUMMARYlist.TextMatrix(i, 12)
            serial = SUMMARYlist.TextMatrix(i, 2)
            ComputerFactor = ImsDataX.ComputingFactor(nameSP, stocknumb, cn)
            If ComputerFactor = 0 Then
                SecUnit = PrimUnit
            Else
                SecUnit = PrimUnit * 10000 / ComputerFactor
            End If
            NEWcondition = SUMMARYlist.TextMatrix(i, 13)
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    retval = PutDataInsert2(i, unitprice)
                    
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    Dim repairCOST As Double
                    Dim newSTOCKNUMB As String
                    If SUMMARYlist.TextMatrix(i, 18) = "" Then
                        newSTOCKNUMB = stocknumb
                    Else
                        If SUMMARYlist.TextMatrix(i, 18) = stocknumb Then
                        Else
                            newSTOCKNUMB = SUMMARYlist.TextMatrix(i, 18)
                        End If
                    End If
                    repairCOST = CDbl(SUMMARYlist.TextMatrix(i, 17))
                    
                    
                    retval = Update_Sap_With_repair_Cost(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, CDbl(1), repairCOST, unitprice, NEWcondition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), CDbl(i), ToWH, "RR", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, newSTOCKNUMB, ToWH, PrimUnit, SecUnit, toLOGIC, toSUBLOCA, NEWcondition, Format(Transnumb), FromWH, CDbl(i), ToWH, "RR", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "RT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "RT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                Case "02050200" 'AdjustmentEntry
                    ToWH = cell(2).tag
                    retval = PutDataInsert2(i, unitprice)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If

                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, CDbl(1), unitprice, condition, CurrentUser, cn)
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
                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitprice, condition, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
            
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "I", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), ToWH, "I", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                Case "02040500" 'WellToWell
                    retval = PutDataInsert(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    condition = SUMMARYlist.TextMatrix(i, 13)
                
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitprice, condition, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "TI", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    
                Case "02040700" 'InternalTransfer
                    ToWH = FromWH
                    retval = PutDataInsert(i)
                    If retval = False Then Call RollbackTransaction(cn)

                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitprice, condition, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "IT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "IT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                                        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "AI", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), ToWH, "AI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                
                Case "02040600" 'WarehouseToWarehouse
                    retval = PutDataInsert(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, 1, unitprice, condition, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
        
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "TI", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "TI", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                                        
                Case "02040100" 'WarehouseReceipt
                    retval = PutDataInsert2(i, unitprice)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Exit Sub
                    End If
                                    
                    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, PrimUnit, CDbl(1), unitprice, condition, CurrentUser, cn)
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                     
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), ToWH, "SL", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), FromWH, CDbl(i), ToWH, "SL", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                
                Case "02040300" 'Return from Well
                    If condition = NEWcondition Then
                        newUNITprice = unitprice
                    Else
                        Set data = getDATA("conditionVALUE", Array(NP, unitprice, NEWcondition))
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
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, CurrentUser, cn)
                    
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), CDbl(i), FromWH, "RT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, PrimUnit, SecUnit, fromlogic, fromSUBLOCA, condition, Format(Transnumb), ToWH, CDbl(i), FromWH, "RT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
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
            poSTATUS.Parameters(1) = nameSP
            poSTATUS.Parameters(2) = cell(4).tag
            poSTATUS.Execute
        End If
        
        'cn.CommitTrans
        Tree.Visible = False
        cell(0) = Transnumb
        cell(0).tag = cell(0)
        combo(0).Visible = False
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
    label(0).Visible = True
    cell(0).Visible = True
    
    For i = 1 To 4
        cell(i).backcolor = &HFFFFC0
    Next
    Call unlockCELLS
    Screen.MousePointer = 0
    Exit Sub
RollBack:
    Call RollbackTransaction(cn)
    Screen.MousePointer = 0
    Exit Sub
    'MDI_IMS.StatusBar1.Panels(1).text = ""
End Sub
Private Sub saveBUTTON_Click()
Dim i
Dim retval As Boolean
Dim PO As String
Dim POitem As String
Dim transactionNO As String
Dim tranNUM As Integer
Dim NP As String
Dim CompanyCode As String
Dim TransactionLine
Dim stocknumberFROM As String
Dim stocknumberTO As String
Dim StockDescriptionFrom As String
Dim StockDescriptionTo As String
Dim LocationFrom As String
Dim LocationTO As String
Dim LogicalWarehouseFrom, LogicalWarehouseTo
Dim SubLocationFrom, SubLocationTo As String
Dim PrimaryQuantity As Double
Dim SecondaryQuantity As Double
Dim primaryUNIT, secondaryUNIT
Dim AdditionalCost As Double
Dim conditionFROM As String
Dim conditionTO As String
Dim StockType As String

Dim unitprice As Double
Dim serial As String
Dim PS As Integer
Dim ComputerFactor
Dim imsLock As imsLock.Lock
Dim TransactionType As String
Screen.MousePointer = 11
Dim data As New ADODB.Recordset
    
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
        
    If remarks = "" Then
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please include the remarks for this transaction"
        remarks.backcolor = &HC0FFFF
        remarks.SetFocus
        Exit Sub
    End If
    
    If SSOleDBFQA.Rows = 0 Then
    
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please fill out the FQA values for this transaction.", vbCritical, "Ims"
        Exit Sub
    End If
       
    Call hideREMARKS
    Screen.MousePointer = 11
    savingLABEL.Visible = True
    savingLABEL.ZOrder
    Me.Enabled = False
    Me.Refresh
                
    'Transaction Header
    CompanyCode = cell(1).tag
    PO = ""
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
            TransactionType = "RR"
        Case "02050200" 'AdjustmentEntry
            TransactionType = "AE"
        Case "02040200" 'WarehouseIssue
            TransactionType = "I"
        Case "02040500" 'WellToWell
            TransactionType = "TI"
        Case "02040700" 'InternalTransfer
            TransactionType = "IT"
        Case "02050300" 'AdjustmentIssue
            TransactionType = "AI"
        Case "02040600" 'WarehouseToWarehouse
            TransactionType = "TI"
        Case "02040100" 'WarehouseReceipt
            PO = cell(4).tag
            TransactionType = "R"
        Case "02050400" 'Sales
            TransactionType = "SL"
        Case "02040300" 'Return from Well
            TransactionType = "RT"
    End Select
    Call BeginTransaction(cn)
    tranNUM = GetTransNumb(nameSP, cn)
    If tranNUM > 0 Then
        transactionNO = TransactionType + "-" + Format(tranNUM)
        
        'Parameters for InsertInventoryTransaction stored procedure
        '@Transaction#      AS VARCHAR(15),
        '@TransactionType   AS CHAR(2),
        '@Namespace         AS VARCHAR(5),
        '@Company           AS CHAR(10),
        '@Remarks           AS TEXT,
        '@Currency          AS CHAR(3),
        '@CurrencyValue     AS DECIMAL(18,0),
        '@CreaUser          AS VARCHAR(20)
'                                                            Transaction,   TransactionType, Namesp, Company,     Remarks, Currency, CurrencyValue, CreaUser
        retval = putDATA("InsertInventoryTransaction", Array(transactionNO, TransactionType, nameSP, CompanyCode, remarks, "USD", CDbl(1), CurrentUser))
    Else
        retval = False
    End If
    'Transaction Line Items
    If retval Then
        Screen.MousePointer = 11
        'MDI_IMS.StatusBar1.Panels(1).text = "Saving Line Items"
        For i = 1 To SUMMARYlist.Rows - 1
            TransactionLine = i
            LogicalWarehouseFrom = SUMMARYlist.TextMatrix(i, 9)
            LogicalWarehouseTo = SUMMARYlist.TextMatrix(i, 11)
            LocationFrom = cell(2).tag
            LocationTO = cell(3).tag
            SubLocationFrom = SUMMARYlist.TextMatrix(i, 10)
            SubLocationTo = SUMMARYlist.TextMatrix(i, 12)
            stocknumberFROM = SUMMARYlist.TextMatrix(i, 1)
            stocknumberTO = stocknumberFROM
            conditionFROM = SUMMARYlist.TextMatrix(i, 20)
            conditionTO = conditionFROM
            StockDescriptionFrom = SUMMARYlist.TextMatrix(i, 5)
            StockDescriptionTo = StockDescriptionFrom
            serial = SUMMARYlist.TextMatrix(i, 2)
            If serial = "" Then serial = "POOL"
            If UCase(serial) = "POOL" Then
                PS = 0
            Else
                PS = 1
            End If
            POitem = ""
            PrimaryQuantity = CDbl(IIf(SUMMARYlist.TextMatrix(i, 7) = "", 0, SUMMARYlist.TextMatrix(i, 7)))
            ComputerFactor = ImsDataX.ComputingFactor(nameSP, stocknumberFROM, cn)
            If ComputerFactor = 0 Then
                SecondaryQuantity = PrimaryQuantity
            Else
                SecondaryQuantity = PrimaryQuantity * 10000 / ComputerFactor
            End If
            unitprice = CDbl(IIf(SUMMARYlist.TextMatrix(i, 4) = "", 0, SUMMARYlist.TextMatrix(i, 4)))
            AdditionalCost = 0
            Err.Clear
            Set data = getDATA("GetStocknumberValues", Array(nameSP, stocknumberTO))
            If data.RecordCount > 0 Then
                primaryUNIT = data!stk_primuon
                secondaryUNIT = data!stk_secouom
            Else
                primaryUNIT = "EA"
                secondaryUNIT = "EA"
            End If
            ''
            'NEWcondition = SUMMARYlist.TextMatrix(i, 13)
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    AdditionalCost = CDbl(SUMMARYlist.TextMatrix(i, 17))
                    'conditionTO = conditionbox(n)
                    If SUMMARYlist.TextMatrix(i, 18) = stocknumberFROM Then
                    Else
                        stocknumberTO = SUMMARYlist.TextMatrix(i, 18)
                        ''''''''stockdescriptionto = a
                    End If
                    retval = Update_Sap(nameSP, CompanyCode, stocknumberFROM, LocationTO, PrimaryQuantity, CDbl(1), unitprice, conditionTO, CurrentUser, cn)
                    
                Case "02050200" 'AdjustmentEntry
                    LocationTO = cell(2).tag
                    retval = Update_Sap(nameSP, CompanyCode, stocknumberFROM, LocationTO, PrimaryQuantity, CDbl(1), unitprice, conditionTO, CurrentUser, cn)
                
                Case "02040200" 'WarehouseIssue
                    
                Case "02040500" 'WellToWell
                    conditionTO = SUMMARYlist.TextMatrix(i, 13)
                    retval = Update_Sap(nameSP, CompanyCode, stocknumberTO, LocationTO, PrimaryQuantity, 1, unitprice, conditionTO, CurrentUser, cn)
                Case "02040700" 'InternalTransfer
                    LocationTO = LocationFrom
                
                Case "02050300" 'AdjustmentIssue
                
                Case "02040600" 'WarehouseToWarehouse
                    retval = Update_Sap(nameSP, CompanyCode, stocknumberFROM, LocationTO, PrimaryQuantity, 1, unitprice, conditionTO, CurrentUser, cn)
                                        
                Case "02040100" 'WarehouseReceipt
                    Err.Clear
                    Set data = getDATA("GetStockNumberPOValues", Array(nameSP, PO, stocknumberFROM))
                    If data.RecordCount > 0 Then
                        primaryUNIT = data!poi_primuom
                        secondaryUNIT = data!poi_secouom
                        POitem = data!poi_liitnumb
                    End If
                    LogicalWarehouseFrom = LogicalWarehouseTo
                    SubLocationFrom = SubLocationTo
                    retval = Update_Sap(nameSP, CompanyCode, stocknumberFROM, LocationTO, PrimaryQuantity, CDbl(1), unitprice, conditionTO, CurrentUser, cn)
                    
                    'Unlock
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(locked, cn, CurrentUser, frmWarehouse.POrowguid)  'jawdat
                    '------
                
                Case "02050400" 'Sales
                
                Case "02040300" 'Return from Well
                    
                    If conditionFROM = conditionTO Then
                        unitprice = unitprice
                    Else
                        Set data = getDATA("conditionVALUE", Array(nameSP, unitprice, conditionTO))
                        If data.RecordCount = 0 Then
                            Call RollbackTransaction(cn)
                            MsgBox "Error in Transaction"
                            Exit Sub
                        Else
                            unitprice = CDbl(data(0))
                        End If
                    End If
            End Select
            
            'Parameters for InsertInventoryTransactionItem stored procedure
            '@Transaction#          AS VARCHAR(15),
            '@TransactionLine       AS INT,
            '@Namespace             AS VARCHAR(5),
            '@TransactionType       AS CHAR(2),
            '@Company               AS CHAR(10),
            '@LogicalWarehouseFrom  AS CHAR(10),
            '@LogicalWarehouseTo    AS CHAR(10),
            '@LocationFrom          AS CHAR(10),
            '@LocationTo            AS CHAR(10),
            '@PO                    AS PONUMB,
            '@POITEM                AS VARCHAR(6),
            '@SubLocationFrom       AS CHAR(10),
            '@SubLocationTo         AS CHAR(10),
            '@StockNumberFrom       AS CHAR(20),
            '@StockNumberTo         AS CHAR(20),
            '@ConditionFrom         AS CHAR(2),
            '@ConditionTo           AS CHAR(2),
            '@StockDescriptionFrom  AS CHAR(1500),
            '@StockDescriptionTo    AS CHAR(1500),
            '@PS                    AS BIT,
            '@Serial                AS VARCHAR(15),
            '@PrimaryQuantity       AS DECIMAL(18,0),
            '@SecondaryQuantity     AS DECIMAL(18,0),
            '@PrimaryUnit           AS CHAR(4),
            '@SecondaryUnit         AS CHAR(4),
            '@UnitPrice             AS DECIMAL(18,0),
            '@AdditionalCost        AS DECIMAL(18,0),
            '@Remarks               AS TEXT,
            '@StockType             AS CHAR(4) ****I'm getting this value on fly from stockmaster,
            '@OWLE                  AS BIT,
            '@LeaseCompany          AS VARCHAR(20),
            '@CreaUser              AS VARCHAR(20)
            
            'Paramaters                                              Transaction#,  TransactionLine, Namespace,TransactionType,Company,    LogicalWarehouseFrom, LogicalWarehouseTo, LocationFrom, LocationTo, PO, POITEM, SubLocationFrom, SubLocationTo, StockNumberFrom, StockNumberTo, ConditionFrom, ConditionTo, StockDescriptionFrom, StockDescriptionTo, PS, Serial, PrimaryQuantity, SecondaryQuantity, PrimaryUnit, SecondaryUnit, UnitPrice, AdditionalCost, Remarks, StockType,OWLE,LeaseCompany, CreaUser
'            If PO = "" Then PO = Null
'            If POitem = "" Then POitem = Null
            retval = putDATA("InsertInventoryTransactionItem", Array(transactionNO, TransactionLine, nameSP, TransactionType, CompanyCode, LogicalWarehouseFrom, LogicalWarehouseTo, LocationFrom, LocationTO, PO, POitem, SubLocationFrom, SubLocationTo, stocknumberFROM, stocknumberTO, conditionFROM, conditionTO, StockDescriptionFrom, StockDescriptionTo, PS, serial, PrimaryQuantity, SecondaryQuantity, primaryUNIT, secondaryUNIT, unitprice, AdditionalCost, remarks, "", 1, Null, CurrentUser))
        Next
    End If
    If retval = True Then retval = SaveFQA(transactionNO, TransactionType)
    If retval Then
        Call CommitTransaction(cn)
        If frmWarehouse.tag = "02040100" Then  'WarehouseReceipt
            Dim poSTATUS As ADODB.Command
            Set poSTATUS = getCOMMAND("UPDATE_PO_INVSTATES")
            poSTATUS.Parameters(1) = nameSP
            poSTATUS.Parameters(2) = cell(4).tag
            poSTATUS.Execute
        End If
        
        Tree.Visible = False
        cell(0) = transactionNO
        cell(0).tag = cell(0)
        combo(0).Visible = False
        combo(0).TextMatrix(1, 0) = transactionNO
    Else
        Call RollbackTransaction(cn)
        MsgBox "The transaction can not be saved"
    End If
    Screen.MousePointer = 11
    If Err Then Err.Clear
    newBUTTON.Enabled = True
    saveBUTTON.Enabled = False
    savingLABEL.Visible = False
    Command3.Enabled = True
    Call lockDOCUMENT(True)
    Me.Enabled = True
    Call unlockBUNCH
    label(0).Visible = True
    cell(0).Visible = True
    
    For i = 1 To 4
        cell(i).backcolor = &HFFFFC0
    Next
    Call unlockCELLS
    Screen.MousePointer = 0
    Exit Sub
RollBack:
    Call RollbackTransaction(cn)
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Function PutDataInsert2(Item, price) As Boolean
    Dim psVALUE, serial
    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert2 = False

    Set cmd = getCOMMAND("INVTRECEIPTDETL_INSERT")

    'Set the parameter values for the command to be executed.
    cmd.Parameters("@ird_curr") = "USD"
    cmd.Parameters("@ird_currvalu") = 1
    cmd.Parameters("@ird_ponumb") = Null
    cmd.Parameters("@ird_lirtnumb") = Null
    cmd.Parameters("@ird_compcode") = cell(1).tag
    cmd.Parameters("@ird_trannumb") = Transnumb
    cmd.Parameters("@ird_npecode") = nameSP
    With SUMMARYlist
        Select Case frmWarehouse.tag
            Case "02050200" 'AdjustmentEntry
                cmd.Parameters("@ird_ware") = cell(2).tag
            Case "02040100" 'WarehouseReceipt
                cmd.Parameters("@ird_ponumb") = .TextMatrix(Item, 21)
                cmd.Parameters("@ird_lirtnumb") = Val(.TextMatrix(Item, 22))
                cmd.Parameters("@ird_ware") = cell(3).tag
            Case Else
                cmd.Parameters("@ird_ware") = cell(3).tag
        End Select
        cmd.Parameters("@ird_transerl") = Item
        cmd.Parameters("@ird_stcknumb") = .TextMatrix(Item, 1)
        If UCase(.TextMatrix(Item, 2)) = "POOL" Or .TextMatrix(Item, 2) = "" Then
            psVALUE = 1
            serial = Null
        Else
            psVALUE = 0
            serial = .TextMatrix(Item, 2)
        End If
        cmd.Parameters("@ird_ps") = psVALUE
        cmd.Parameters("@ird_serl") = serial
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
                cmd.Parameters("@ird_reprcost") = CDbl(.TextMatrix(Item, 17))
                cmd.Parameters("@ird_newcond") = .TextMatrix(Item, 13)
                cmd.Parameters("@ird_newstcknumb") = .TextMatrix(Item, 18)
                cmd.Parameters("@ird_newdesc") = .TextMatrix(Item, 19)
            Case "02050200" 'AdjustmentEntry
                'cmd.Parameters("@ird_newcond") = .TextMatrix(Item, 13)
                cmd.Parameters("@ird_newcond") = .TextMatrix(Item, 20) 'M
                'Modified by Muzammil
                'Reason : would not save the condition only to the ird_origcond.
                'The above value begin passed is just an empty string
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                cmd.Parameters("@ird_newcond") = "01"
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
                cmd.Parameters("@ird_newcond") = .TextMatrix(Item, 13)
        End Select
        cmd.Parameters("@ird_stcktype") = ""
        cmd.Parameters("@ird_ctry") = "US"
        cmd.Parameters("@ird_tosubloca") = .TextMatrix(Item, 12)
        cmd.Parameters("@ird_tologiware") = .TextMatrix(Item, 11)
        cmd.Parameters("@ird_owle") = 1
        cmd.Parameters("@ird_leasecomp") = Null
        cmd.Parameters("@ird_primqty") = CDbl(.TextMatrix(Item, 7))
        cmd.Parameters("@ird_secoqty") = SecUnit
        cmd.Parameters("@ird_unitpric") = CDbl(price)
        cmd.Parameters("@ird_stckdesc") = .TextMatrix(Item, 5)
        cmd.Parameters("@ird_fromlogiware") = .TextMatrix(Item, 9)
        cmd.Parameters("@ird_fromsubloca") = .TextMatrix(Item, 10)
        
        If Me.tag <> "02050200" Then cmd.Parameters("@ird_origcond") = .TextMatrix(Item, 20) 'M
        
        cmd.Parameters("@user") = CurrentUser
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
        cmd.Parameters("@iid_trannumb") = Transnumb
        cmd.Parameters("@iid_compcode") = cell(1).tag
        cmd.Parameters("@iid_npecode") = nameSP
        cmd.Parameters("@iid_ware") = cell(2).tag
        cmd.Parameters("@iid_transerl") = .TextMatrix(row, 0)
        cmd.Parameters("@iid_stcknumb") = .TextMatrix(row, 1)
        cmd.Parameters("@iid_ps") = IIf(.TextMatrix(row, 2) = "", 1, 0)
        cmd.Parameters("@iid_serl") = IIf(.TextMatrix(row, 2) = "", Null, .TextMatrix(row, 2))
        
        'cmd.Parameters("@iid_newcond") = .TextMatrix(row, 13) 'M
        'Modified by Muz
        'Reason :  In the older version this Field was NULL only the Orig cond is being Populated.
        'this is in case of AI
        If Me.tag <> "02050300" Then cmd.Parameters("@iid_newcond") = .TextMatrix(row, 13) 'M
        
        cmd.Parameters("@iid_stcktype") = "I"
        cmd.Parameters("@iid_ctry") = "US"
        cmd.Parameters("@iid_tosubloca") = .TextMatrix(row, 12)
        cmd.Parameters("@iid_tologiware") = .TextMatrix(row, 11)
        cmd.Parameters("@iid_owle") = 1
        cmd.Parameters("@iid_leasecomp") = Null
        cmd.Parameters("@iid_primqty") = CDbl(.TextMatrix(row, 7))
        cmd.Parameters("@iid_secoqty") = SecUnit
        cmd.Parameters("@iid_unitpric") = CDbl(.TextMatrix(row, 4))
        cmd.Parameters("@iid_curr") = "USD"
        cmd.Parameters("@iid_currvalu") = 1
        cmd.Parameters("@iid_stckdesc") = .TextMatrix(row, 5)
        cmd.Parameters("@iid_fromlogiware") = .TextMatrix(row, 9)
        cmd.Parameters("@iid_fromsubloca") = .TextMatrix(row, 12)
        cmd.Parameters("@iid_origcond") = .TextMatrix(row, 13)
        cmd.Parameters("@user") = CurrentUser
    End With
    'Execute the command.
    Call cmd.Execute(Options:=adExecuteNoRecords)
    PutDataInsert = True
End Function

Private Function PutReceiptRemarks() As Boolean
Dim cmd As New ADODB.Command

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceiptRem_Insert"
        .Parameters.Append .CreateParameter("@CompanyCode", adChar, adParamInput, 10, cell(1).tag)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, nameSP)
        .Parameters.Append .CreateParameter("@WhareHouse", adChar, adParamInput, 10, cell(2).tag)
        .Parameters.Append .CreateParameter("@TranNumb", adVarChar, adParamInput, 15, Transnumb)
        .Parameters.Append .CreateParameter("@LINENUMB", adInteger, adParamInput, , 1)
        .Parameters.Append .CreateParameter("@REMARKS", adVarChar, adParamInput, 7000, remarks)
        .Parameters.Append .CreateParameter("@USER", adChar, adParamInput, 20, CurrentUser)
        Call .Execute(, , adExecuteNoRecords)
    End With
    PutReceiptRemarks = cmd.Parameters(0).Value = 0
End Function
Private Function PutIssueRemarks() As Boolean
Dim cmd As ADODB.Command

    Set cmd = getCOMMAND("InvtIssuetRem_Insert")
    
    cmd.Parameters("@LineNumb") = 1
    cmd.Parameters("@REMARKS") = remarks
    cmd.Parameters("@TranNumb") = Transnumb
    cmd.Parameters("@CompanyCode") = cell(1).tag
    cmd.Parameters("@NAMESPACE") = nameSP
    cmd.Parameters("@WhareHouse") = cell(2).tag
    cmd.Parameters("@USER") = CurrentUser
    
    Call cmd.Execute(0, , adExecuteNoRecords)
    PutIssueRemarks = cmd.Parameters(0).Value = 0
End Function

Private Function PutInvtIssue(prefix) As Boolean
Dim NP As String
Dim cmd As Command
On Error GoTo errPutInvtIssue

    PutInvtIssue = False
    Set cmd = getCOMMAND("InvtIssue_Insert")
    NP = nameSP
    Transnumb = prefix + "-" & GetTransNumb(NP, cn)
    cmd.Parameters("@NAMESPACE") = NP
    cmd.Parameters("@TRANTYPE") = prefix
    cmd.Parameters("@COMPANYCODE") = cell(1).tag
    cmd.Parameters("@TRANSNUMB") = Transnumb
    cmd.Parameters("@ISSUTO") = cell(3).tag
    cmd.Parameters("@SUPPLIERCODE") = Null
    Select Case frmWarehouse.tag
        Case "02040500" 'WellToWell
        Case "02040700", "02050300" 'InternalTransfer, AdjustmentIssue
            cmd.Parameters("@ISSUTO") = cell(2).tag
        Case "02040600" 'WarehouseToWarehouse
        Case "02050400" 'Sales
            cmd.Parameters("@SUPPLIERCODE") = cell(4).tag
    End Select
    cmd.Parameters("@WHAREHOUSE") = cell(2).tag
    cmd.Parameters("@STCKNUMB") = Null
    cmd.Parameters("@COND") = Null
    cmd.Parameters("@SAP") = Null
    cmd.Parameters("@NEWSAP") = Null
    cmd.Parameters("@ENTYNUMB") = Null
    cmd.Parameters("@USER") = CurrentUser
    cmd.Execute
    PutInvtIssue = cmd.Parameters(0).Value = 0
    Exit Function

errPutInvtIssue:
    MsgBox Err.description
    Err.Clear
End Function

Private Sub Command3_Click()
Dim reportPATH, cnSTRING, text
Screen.MousePointer = 11
    With CrystalReport1
        .Reset
        reportPATH = repoPATH + "\"
        .ReportFileName = reportPATH & "InventoryTransaction.rpt"
        .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
        .ParameterFields(1) = "NAMESPACE;" & nameSP & ";TRUE"
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
    label(0).Visible = False
    cell(0).Visible = False
    Command3.Enabled = False
    Call cleanSTOCKlist
    Call cleanSUMMARYlist
    Call hideDETAILS
    Line2.Visible = False
    'STOCKlist.Top = 1920
    STOCKlist.Top = 2080
    STOCKlist.Visible = True
    searchFIELD(0).Visible = True
    searchFIELD(1).Visible = True
    detailHEADER.Top = 4320
    Tree.Top = 4560
    Tree.Height = 3660
    cell(1).SetFocus
    saveBUTTON.Enabled = True
    newBUTTON.Enabled = False
    cell(0).backcolor = &HFFFFC0
    cell(0) = ""
    SUMMARYlist.Rows = 2
    For i = 0 To SUMMARYlist.Rows - 1
        SUMMARYlist.TextMatrix(1, i) = ""
    Next
    SUMMARYlist.Top = 3870
    SUMMARYlist.Height = 4375
    For i = 1 To 4
        cell(i) = ""
        cell(i).backcolor = vbWhite
        cell(i).locked = False
    Next
    Call hideREMARKS
    Call CleanFQA
   Call ChangeMode(False)
    remarks = ""
    Call cell_Click(1)
End Sub

Private Sub commodityLABEL_Change()
    Call whitening
End Sub
Private Sub grid_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i, name
    Select Case KeyAscii
        Case 13
            Call grid_Click(Index)
        Case 27
    End Select
    grid(Index).Visible = False
        
    With grid(Index)
        i = Val(.ToolTipText)
        Select Case Index
            Case 0
                name = "NEWconditionBOX"
            Case 1
                name = "logicBOX"
            Case 2
                name = "sublocaBOX"
        End Select
        
        Select Case name
            Case "logicBOX"
                logicBOX(i).backcolor = vbWhite
                sublocaBOX(i).SetFocus
            Case "sublocaBOX"
                sublocaBOX(i).backcolor = vbWhite
                If NEWconditionBOX(i).Visible Then
                    NEWconditionBOX(i).SetFocus
                Else
                    quantityBOX(i).SetFocus
                End If
            Case "NEWconditionBOX"
                NEWconditionBOX(i).backcolor = vbWhite
                quantityBOX(i).SetFocus
        End Select
    End With
End Sub

Private Sub hideDETAIL_Click()
Dim answer, i
    If Tree.Nodes.Count > 0 Then
        If CDbl(quantityBOX(totalNODE)) > 0 Then
            answer = MsgBox("Are you sure you want to lose last changes?", vbYesNo)
            If answer = vbYes Then
                hideDETAILS
            End If
        Else
            hideDETAILS
        End If
    Else
        hideDETAILS
    End If
    With frmWarehouse
        If .STOCKlist.cols > 6 Then
            For i = 1 To .STOCKlist.Rows - 1
                Select Case .tag
                    'ReturnFromRepair, AdjustmentEntry,WarehouseIssue,WellToWell,InternalTransfer,
                    'AdjustmentIssue,WarehouseToWarehouse,Sales
                    Case "02040400", "02050200", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                        .STOCKlist.TextMatrix(i, 5) = .STOCKlist.TextMatrix(i, 7)
                    Case "02040100" 'WarehouseReceipt
                        .STOCKlist.TextMatrix(i, 3) = .STOCKlist.TextMatrix(i, 7)
                End Select
            Next
        End If
    End With
End Sub

Private Sub cell_Click(Index As Integer)
Dim datax As New ADODB.Recordset
Dim sql As String
Dim i
    With cell(Index)
        If .locked Then Exit Sub
        Screen.MousePointer = 11
        Select Case Index
            Case 5
                If saveBUTTON.Enabled Then
                    If Not combo(5).Visible Then
                        sql = "SELECT stk_stcknumb, stk_stcktype, stk_catecode, stk_poolspec, stk_desc " _
                            & "FROM StockMaster WHERE stk_npecode = '" & nameSP & "'" _
                            & "ORDER BY stk_stcknumb"
                        datax.Open sql, cn, adOpenForwardOnly
                        If datax.RecordCount > 0 Then
                            combo(5).Rows = 2
                            combo(5).Clear
                            combo(5).Rows = datax.RecordCount + 1
                            For i = 1 To datax.RecordCount
                                combo(5).TextMatrix(i, 0) = Trim(datax!stk_stcknumb)
                                datax.MoveNext
                            Next
                            combo(5).Visible = True
                            combo(5).ZOrder
                            combo(5).ColWidth(0) = combo(5).width - 270
                            combo(5).ColAlignment(0) = 0
                            combo(5).TextMatrix(0, 0) = "Stock Number"
                            combo(5).ColAlignmentFixed(0) = 3
                            .tag = .text
                            .text = ""
                            .text = .tag
                            .SelLength = 0
                            .SelStart = Len(.text)
                            Screen.MousePointer = 0
                        End If
                    End If
                End If
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
                        If cell(Index) <> "" Then Call alphaSEARCH(cell(Index), combo(Index), 0)
                        combo(Index).ColSel = combo(Index).cols - 1
                    End If
                End If
        End Select
        .SelStart = 0
        .SelLength = Len(.text)
        frmWarehouse.combo(Index).MousePointer = 0
    End With
    Screen.MousePointer = 0
End Sub

Private Sub cell_GotFocus(Index As Integer)
    If saveBUTTON.Enabled Or Index = 0 Then
        If Not (saveBUTTON.Enabled And Index = 0) Then
            If Index < 6 Then
                If Index <> activeCELL Then
                    combo(activeCELL).Visible = False
                    cell(activeCELL).backcolor = vbWhite
                End If
            End If
            With cell(Index)
                If .locked Then Exit Sub
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
        If .locked Then Exit Sub
        If Not .locked Then
                Select Case KeyCode
                    Case 40
                        If combo(Index).Visible Then
                            direction = "down"
                            Call arrowKEYS(Index, cell(Index), combo(Index))
                        Else
                            Call cell_Click(Index)
                        End If
                        If Index = 5 Then
                            cell(Index).tag = combo(Index).TextMatrix(combo(Index).row, 0)
                        Else
                            cell(Index).tag = combo(Index).TextMatrix(combo(Index).row, Val(matrix.TextMatrix(10, Index)))
                        End If
                    Case 38
                        If combo(Index).Visible Then
                            direction = "up"
                            Call arrowKEYS(Index, cell(Index), combo(Index))
                        Else
                            Call cell_Click(Index)
                        End If
                        'cell(Index).SetFocus
                End Select
        End If
    End With
End Sub
Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i, t, n
Dim gotIT As Boolean
    With cell(Index)
        If .locked Then Exit Sub
        Select Case KeyAscii
            Case 13
                KeyAscii = 0
                If Not .locked Then
                    justCLICK = False
                    gotIT = False
                    n = Val(matrix.TextMatrix(10, Index) = 0)
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
                        If Index = 4 Then
                            cell(1).SetFocus
                        Else
                            If cell(Index + 1).locked Then
                            Else
                                If cell(Index + 1).Visible Then cell(Index + 1).SetFocus
                            End If
                        End If
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
    If cell(Index).locked Then Exit Sub
    If usingARROWS Then
        usingARROWS = False
    Else
        If saveBUTTON.Enabled Or Index = 0 Then
            If Not (saveBUTTON.Enabled And Index = 0) Then
                If Index < 6 Then
                    If Index <> activeCELL Then combo(activeCELL).Visible = False
                End If
            End If
        End If
    End If
    If saveBUTTON.Enabled Or Index = 0 Then
        If combo(Index).Visible = False Then
            With cell(Index)
                .backcolor = vbWhite
            End With
        End If
    End If
End Sub



Public Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If cell(Index).locked Then Exit Sub
    If findSTUFF(cell(Index), combo(Index), 0) = 0 Then cell(Index) = ""
End Sub

Private Sub combo_Click(Index As Integer)
Dim i, sql, t
Dim cleanDETAILS As Boolean
Dim datax As New ADODB.Recordset
Dim currentformname, currentformname1
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
            sql = "SELECT stk_desc FROM STOCKMASTER WHERE " _
                & "stk_npecode = '" + nameSP + "' and " _
                & "stk_stcknumb = '" + .text + "'"
            datax.Open sql, cn, adOpenStatic
            cell(5) = .text
            If datax.RecordCount > 0 Then
                newDESCRIPTION = IIf(IsNull(datax!stk_desc), "", datax!stk_desc)
            Else
                newDESCRIPTION = ""
            End If
        Else
            If Not savingLABEL.Visible Then
                DoEvents
                cell(Index) = .TextMatrix(.row, 0)
                DoEvents
                .Refresh
                cell(Index).tag = .TextMatrix(.row, Val(matrix.TextMatrix(10, Index)))
            End If
            If Index < 2 Then
                For i = 2 To 4
                    cell(i) = ""
                    cell(i).tag = ""
                Next
                STOCKlist.Rows = 2
                STOCKlist.RowHeightMin = 0
                STOCKlist.RowHeight(1) = 0
                STOCKlist.TextMatrix(1, 0) = ""
                STOCKlist.TextMatrix(1, 1) = ""
            End If
            
            currentformname = frmWarehouse.tag
            currentformname1 = currentformname
            
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Receptions_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Receptions_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                    End Select
                Case "02040200" 'WarehouseIssue
                    Select Case Index
                        Case 0
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Receptions_New WHERE QTY1 > 0 AND " _
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
                            sql = "SELECT * FROM Issues_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                            sql = "SELECT * FROM Receptions_New WHERE " _
                                & "NAMESPACE = '" + nameSP + "' AND " _
                                & "Transaction# = '" + cell(0) + "' " _
                                & "ORDER BY TransactionLine"
                        Case 1, 2
                            If (Len(cell(1)) + Len(cell(2))) > Len(cell(1)) Then
                                sql = "SELECT * FROM StockInfo_New WHERE " _
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
                                StockListDuplicate.Visible = True
                            End If
                        End If
                            For i = 1 To 4
                                If cell(i).Visible And cell(i) = "" Then STOCKlist.Enabled = False
                            Next
                            Call fillSTOCKlist(datax)
                            If savingLABEL.Visible Then
                                Label3 = "SAVING..."
                                savingLABEL.Visible = False
                                If frmWarehouse.tag = "02040200" And Index = 2 Then
                                    StockListDuplicate.Visible = False
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
        Call fillDETAILlist("", "", "")
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
            If Index < 4 Then
                If Index > 0 Then
                    Call cleanSTOCKlist
                    For i = Index + 1 To 4
                        cell(i) = ""
                    Next
                End If
            End If
        Case "02050400" 'Sales
    End Select
    Dim x As String
    
    'Loads the FQA Details of the saved Transaction ( Only in the case of a modification)
    If Index = 0 Then Call PopulateFQAOftheTransaction(combo(0))
    
    'Gets the FQA code for the selected Location ( only in the case of a creation)
    If Index = 2 Then Call LoadFromFQA(cell(1).tag, cell(2).tag)
    
    If newBUTTON.Enabled = True Then
        Call ChangeMode(True)
    ElseIf newBUTTON.Enabled = False Then
        Call ChangeMode(False)
    End If
    If Index > 0 Then
        If Index < 4 Then
            If cell(Index + 1).Visible Then
                cell(Index + 1).Enabled = True
                cell(Index + 1).SetFocus
                Call cell_Click(Index + 1)
            End If
        Else
            cell(Index).SetFocus
        End If
    End If
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
    rights = Getmenuuser(nameSP, CurrentUser, Me.tag, cn)
    newBUTTON.Enabled = rights
    Me.Visible = True
    If newBUTTON.Enabled Then newBUTTON.SetFocus
    Me.Refresh
    userNAMEbox = CurrentUser
    dateBOX = Format(Now, "mm/dd/yyyy")
    hideDETAILS
    Call makeLISTS
    Load grid(1)
    Load grid(2)
    DoEvents
    Call fillGRID(grid(1), logicBOX(0), 0)
    DoEvents
    Call fillGRID(grid(2), sublocaBOX(0), 0)
    Call fillGRID(grid(0), NEWconditionBOX(0), 0)
End Sub

Public Sub setCN(conn As ADODB.Connection)
    Set cn = conn
    If Not IsConnectionOpen(conn) Then Exit Sub
End Sub
Private Sub Form_Load()
On Error Resume Next
    DoEvents
    Screen.MousePointer = 11
    
    'Call translator.Translate_Forms("frmWarehouse")
    Screen.MousePointer = 11
    Call lockDOCUMENT(True)
    frmWarehouse.Caption = frmWarehouse.Caption + " - " + frmWarehouse.tag
    Screen.MousePointer = 0
    If Err Then MsgBox "Error: " + Err.description
    StockListDuplicate.Visible = False
    
    SSOleCompany.columns(0).width = 855
    SSOleDBLocation.columns(0).width = 975
    SSOleDBUsChart.columns(0).width = 1455
    SSOleDBCamChart.columns(0).width = 1455

End Sub


Sub SAVE()
Dim header As New ADODB.Recordset
Dim details As New ADODB.Recordset
Dim remarks As New ADODB.Recordset

Dim INVitem As New ADODB.Recordset

Dim i, row As Integer
Dim sql As String
Dim Q, quantity, price As Double
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
    With grid(Index)
        justCLICK = True
        i = Val(.ToolTipText)
        Select Case Index
            Case 0
                name = "NEWconditionBOX"
            Case 1
                name = "logicBOX"
            Case 2
                name = "sublocaBOX"
        End Select
        
        Select Case name
            Case "logicBOX"
                logicBOX(i) = .TextMatrix(.row, 1)
                logicBOX(i).tag = .TextMatrix(.row, 1)
                logicBOX(i).ToolTipText = .TextMatrix(.row, 0)
                logicBOX(i).SetFocus
            Case "sublocaBOX"
                sublocaBOX(i) = .TextMatrix(.row, 1)
                sublocaBOX(i).tag = .TextMatrix(.row, 1)
                sublocaBOX(i).ToolTipText = .TextMatrix(.row, 0)
                sublocaBOX(i).SetFocus
            Case "NEWconditionBOX"
                NEWconditionBOX(i) = "0" + .TextMatrix(.row, 0)
                NEWconditionBOX(i).tag = .TextMatrix(.row, 0)
                NEWconditionBOX(i).ToolTipText = .TextMatrix(.row, 1)
                NEWconditionBOX(i).SetFocus
        End Select
        .Visible = False
    End With
End Sub

Public Sub logicBOX_Click(Index As Integer)
    grid(0).Visible = False
    grid(2).Visible = False
    grid(1).ToolTipText = Format(Index, "00") + "logicBOX"
    Call showGRID(grid(1), Index, logicBOX(Index), True)
End Sub

Public Sub logicBOX_GotFocus(Index As Integer)
    If Tree.Visible = True Then Exit Sub
        activeBOX = "logicBOX"
        Call whitening
        With logicBOX(Index)
            .backcolor = &H80FFFF
            .SelStart = 0
            .SelLength = Len(.text)
            If justCLICK Then
                grid(0).Visible = False
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
            grid(1).Visible = False
            If sublocaBOX(Index).Visible Then
                sublocaBOX(Index).SetFocus
                logicBOX(Index).backcolor = vbWhite
                Exit Sub
            End If
        Case 27
            grid(1).Visible = False
    End Select
End Sub

Private Sub logicBOX_LostFocus(Index As Integer)
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    If ActiveControl.name = "grid" Then
        If ActiveControl.Index = 1 Then
        Else
            Call hideGRIDS
            logicBOX(Index).backcolor = vbWhite
        End If
    Else
        Call hideGRIDS
        logicBOX(Index).backcolor = vbWhite
    End If
End Sub


Private Sub logicBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index > 0 And Index <> totalNODE Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        logicBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub NEWconditionBOX_Click(Index As Integer)
    grid(0).Visible = False
    grid(2).Visible = False
    grid(0).ToolTipText = Format(Index, "00") + "NEWconditionBOX"
    Call showGRID(grid(0), Index, NEWconditionBOX(Index), True)
End Sub


Private Sub NEWconditionBOX_GotFocus(Index As Integer)
    If Tree.Visible = True Then Exit Sub
    activeBOX = "NEWconditionBOX"
    Call whitening
    With NEWconditionBOX(Index)
        .backcolor = &H80FFFF
        .SelStart = 0
        .SelLength = Len(.text)
        If justCLICK Then
            grid(0).Visible = False
            justCLICK = False
        Else
            grid(0).ToolTipText = Format(Index, "00") + "NEWconditionBOX"
            Call showGRID(grid(0), Index, NEWconditionBOX(Index), True)
        End If
    End With
End Sub


Private Sub NEWconditionBOX_LostFocus(Index As Integer)
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    If activeBOX = "NEWconditionBOX" Then
    Else
        Call hideGRIDS
        NEWconditionBOX(Index).backcolor = vbWhite
    End If
End Sub

Private Sub NEWconditionBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index > 0 And Index <> totalNODE Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        NEWconditionBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantityBOX_Change(Index As Integer)
    Call quantityBOX_Validate(Index, True)
End Sub

Private Sub quantityBOX_Click(Index As Integer)
    With quantityBOX(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub quantityBOX_GotFocus(Index As Integer)
    activeBOX = "quantityBOX"
    If Index <> totalNODE Then
        Call whitening
        quantityBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Private Sub quantityBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call quantityBOX_Validate(Index, True)
    End If
End Sub

Private Sub quantityBOX_LostFocus(Index As Integer)
    If Index <> totalNODE Then quantityBOX(Index).backcolor = vbWhite
End Sub


Private Sub quantityBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index > 0 And Index <> totalNODE Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        quantityBOX(Index).backcolor = &H80FFFF
    End If
End Sub

Public Sub quantityBOX_Validate(Index As Integer, Cancel As Boolean)
Dim qty
Dim calculate As Boolean
On Error Resume Next
    With quantityBOX(Index)
        If Index <> totalNODE Then
            If IsNumeric(.text) Then
                If CDbl(.text) > 0 Then
                    '.text = Format(.text, "0.00")
                    Select Case frmWarehouse.tag
                        Case "02050200" 'AdjustmentEntry
                        Case "02040100" 'WarehouseReceipt
                        Case Else
                            If CDbl(.text) > CDbl(quantity(Index)) Then .text = quantity(Index)
                    End Select
                    calculate = True
                Else
                    .text = "0.00"
                End If
                If (Err.Number = 0) Then Call calculations
            Else
                If .text = "." Then
                Else
                    .text = "0.00"
                End If
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
Dim i
Dim RowPosition As Integer
    With SUMMARYlist
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, 1) = commodityLABEL Then
                If .Rows > 2 Then
                    .RemoveItem i
''                Else
''                    .TextMatrix(1, 0) = ""
''                    .TextMatrix(1, 1) = ""
                    RowPosition = i
            ElseIf .Rows = 2 Then
                    .RemoveItem 1
                    RowPosition = 1
                End If
            End If
        Next
        Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "delete", Null, Null, Null, Null, RowPosition)
        Call reNUMBER(SUMMARYlist)
        Call fillDETAILlist("", "", "")
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
    activeBOX = "repairBOX"
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

Private Sub repairBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index <> totalNODE Then
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


Private Sub searchFIELD_KeyPress(Index As Integer, KeyAscii As Integer)
Dim datax As New ADODB.Recordset
Dim sql, list, i, ii, t
Screen.MousePointer = 11
    If KeyAscii = 13 Then
        If Not newBUTTON.Enabled Then
            If Index = 0 Then
                If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                    sql = "SELECT stk_stcknumb, stk_desc, uni_desc " _
                        & "FROM STOCKMASTER LEFT OUTER JOIN UNIT ON " _
                        & "stk_npecode = uni_npecode AND " _
                        & "stk_primuon = uni_code WHERE " _
                        & "(stk_npecode = '" + nameSP + "') AND " _
                        & "(stk_stcknumb like '" + searchFIELD(Index).text + "%')"
                    datax.Open sql, cn, adOpenStatic
                    If datax.RecordCount > 0 Then
                        Do While Not datax.EOF
                            If findSTUFF(datax!stk_stcknumb, STOCKlist, 1) = 0 Then
                                If IsNull(datax!uni_desc) Then
                                    STOCKlist.addITEM "" + vbTab + datax!stk_stcknumb + vbTab + datax!stk_desc + vbTab + "", 1
                                Else
                                    STOCKlist.addITEM "" + vbTab + datax!stk_stcknumb + vbTab + datax!stk_desc + vbTab + datax!uni_desc, 1
                                End If
                            End If
                            datax.MoveNext
                            Loop
                        If STOCKlist.Rows > 2 And STOCKlist.TextMatrix(1, 1) = "" Then STOCKlist.RemoveItem 1
                        Call reNUMBER(STOCKlist)
                    End If
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
                                If STOCKlist.Rows > 2 And STOCKlist.TextMatrix(1, 1) = "" Then STOCKlist.RemoveItem 1
                                Call reNUMBER(STOCKlist)
                            End If
                        End If
                    End If
                Else
                    Call searchIN(searchFIELD(1), STOCKlist, 3)
                    searchFIELD(1).SelStart = 0
                    searchFIELD(1).SelLength = Len(searchFIELD(1))
                End If
            End If
        End If
    End If
    STOCKlist.TopRow = 1
    Screen.MousePointer = 0
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

Select Case SSOleDBFQA.col

    Case 0
        If Len(Trim(SSOleDBFQA.columns(0))) > 20 Then
            MsgBox "Stocknumber is too long. Please make sure it is not larger than 20 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    Case 1
    
        If Len(Trim(SSOleDBFQA.columns(1).text)) > 2 Then
            
            MsgBox "Company is too long. Please make sure it is not larger than 2 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    
    Case 2
    
       If Len(Trim(SSOleDBFQA.columns(2).text)) > 11 Then
            
            MsgBox "Location is too long. Please make sure it is not larger than 11 characters.", vbInformation, "Imswin"
            Cancel = 1
        End If
    
    Case 3
    
        
       If Len(Trim(SSOleDBFQA.columns(3).text)) > 9 Then
            
            MsgBox "US Chart is too long. Please make sure it is not larger than 9 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    
    Case 4
    
        
       If Len(Trim(SSOleDBFQA.columns(4).text)) > 4 Then
            
            MsgBox "Stocktype is too long. Please make sure it is not larger than 4 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    
    Case 5
    
        
       If Len(Trim(SSOleDBFQA.columns(5).text)) > 8 Then
            
            MsgBox "Cam Chart is too long. Please make sure it is not larger than 8 characters.", vbInformation, "Imswin"
            Cancel = 1
            
        End If
        
    
End Select

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
    Call lockCELLS
    With STOCKlist
        If .MouseCol = 0 Then
            If .row > 0 Then
                pointerCOL = 0
                Call markROW(STOCKlist)
            End If
        End If
        .col = 0
        .ColSel = .cols - 1
    End With
Screen.MousePointer = 0
frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub STOCKlist_DblClick()
    With STOCKlist
        Me.MousePointer = vbHourglass
        Call markROW(STOCKlist)
        Call PREdetails
        Me.MousePointer = 0
    End With
frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub stocklist_EnterCell()
    frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Private Sub STOCKlist_GotFocus()
    Call hideCOMBOS
End Sub

Private Sub STOCKlist_LostFocus()
    If STOCKlist.Rows > 2 Or Not (STOCKlist.Rows = 2 And STOCKlist.TextMatrix(1, 0) = "") Then
        STOCKlist.tag = 0
    End If
End Sub

Private Sub STOCKlist_RowColChange()
    With STOCKlist
        If IsNumeric(.TextMatrix(.row, 0)) Then
            Call fillDETAILlist("", "", "")
        Else
            Call PREdetails
        End If
    End With
End Sub

Private Sub sublocaBOX_Change(Index As Integer)
    Call alphaSEARCH(sublocaBOX(Index), grid(2), 0)
End Sub

Private Sub sublocaBOX_Click(Index As Integer)
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).ToolTipText = Format(Index, "00") + "sublocaBOX"
    Call showGRID(grid(2), Index, sublocaBOX(Index), True)
End Sub

Private Sub sublocaBOX_GotFocus(Index As Integer)
    activeBOX = "sublocaBOX"
    Call whitening
    With sublocaBOX(Index)
        .backcolor = &H80FFFF
        .SelStart = 0
        .SelLength = Len(.text)

        If justCLICK Then
            grid(0).Visible = False
            justCLICK = False
        Else
            grid(2).ToolTipText = Format(Index, "00") + "sublocaBOX"
            Call showGRID(grid(2), Index, sublocaBOX(Index), True)
        End If
    End With
End Sub


Private Sub sublocaBOX_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40
            direction = "down"
            Call showGRID(grid(2), Index, sublocaBOX(Index), True)
            Call arrowKEYS(Index, sublocaBOX(Index), grid(2))
        Case 38
            direction = "up"
            Call showGRID(grid(2), Index, sublocaBOX(Index), True)
            Call arrowKEYS(Index, sublocaBOX(Index), grid(2))
    End Select
End Sub

Private Sub sublocaBOX_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call grid_Click(2)
            grid(2).Visible = False
            If NEWconditionBOX(Index).Visible Then
                NEWconditionBOX(Index).SetFocus
                sublocaBOX(Index).backcolor = vbWhite
                Exit Sub
            End If
            If quantityBOX(Index).Visible Then
                quantityBOX(Index).SetFocus
                sublocaBOX(Index).backcolor = vbWhite
                Exit Sub
            End If
        Case 27
            grid(2).Visible = False
    End Select
End Sub

Private Sub sublocaBOX_LostFocus(Index As Integer)
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    If activeBOX = "sublocaBOX" Then
    Else
        Call hideGRIDS
        sublocaBOX(Index).backcolor = vbWhite
    End If
End Sub


Private Sub sublocaBOX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Index > 0 And Index <> totalNODE Then
        If currentBOX <> Index Then Call whitening
        currentBOX = Index
        sublocaBOX(Index).backcolor = &H80FFFF
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
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6))
            Case "02040100" 'WarehouseReceipt
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6), .TextMatrix(.row, 17))
        End Select
    End With
Screen.MousePointer = 0
End Sub


Private Sub SUMMARYlist_EnterCell()
    If newBUTTON.Enabled Then Exit Sub
End Sub

Private Sub submitDETAIL_Click()
Dim aproved As Boolean
On Error Resume Next
Dim i, n, rec, r, rcondition, key, conditionCODE, fromlogic, fromSUBLOCA, unitCODE, condition
Dim Str As String
Dim PONumb As String
Dim lineno As String
Dim quant As String

    If IsNumeric(quantityBOX(totalNODE)) Then
        If CDbl(quantityBOX(totalNODE)) > 0 Then
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    aproved = True
                    For i = 1 To Tree.Nodes.Count
                        If i <> totalNODE Then
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
                For i = .Rows - 1 To 1 Step -1
                    If .TextMatrix(i, 1) = commodityLABEL Then
                        If .Rows > 2 Then
                            .RemoveItem i
                        Else
                            .TextMatrix(1, 0) = ""
                            .TextMatrix(1, 1) = ""
                        End If
                    End If
                Next
                For i = 1 To Tree.Nodes.Count
                    If i <> totalNODE Then
                        key = Tree.Nodes(i).key
                        key = Replace(key, "@", "")
                        condition = Mid(key, InStr(key, "-") + 1, InStr(key, "{{") - InStr(key, "-") - 1)
                        conditionCODE = Left(key, 2)
                        If InStr(conditionCODE, "{") > 0 Or InStr(conditionCODE, "-") Then conditionCODE = Left(conditionCODE, 1)
                        Err.Clear
                        
'                        1 = "Commodity"
'                        2 = "Serial"
'                        3 = "Condition"
'                        4 = "Unit Price"
'                        5 = "Description"
'                        6 = "Unit"
'                        7 = "Qty"
'                        9 = "From Logical"
'                        10 = "From Subloca"
'                        11 = "To Logical"
'                        12 = "To Subloca"
'                        13 = "New Condition Code"
'                        14 = "New Condition Description"
'                        15 = "Unit Code"
'                        16 = "Computer Factor"
'                        20 = "Original Condition Code"
                        
                        If Val(quantityBOX(i)) > 0 Then
                            If Err.Number = 0 Then
                                .Rows = .Rows + 1
                                r = .Rows - 1
                                .TextMatrix(r, 1) = commodityLABEL
                                If InStrRev(key, "#") > 0 Then
                                    .TextMatrix(r, 2) = Mid(key, InStrRev(key, "#") + 1)
                                Else
                                    .TextMatrix(r, 2) = "Pool"
                                End If
                                .TextMatrix(r, 4) = Format(priceBOX(i))
                                .TextMatrix(r, 5) = descriptionLABEL
                                .TextMatrix(r, 6) = unitLABEL(0)
                                .TextMatrix(r, 7) = Format(quantityBOX(i), "0.00")
                                .TextMatrix(r, 8) = Format(i)
                                fromlogic = Mid(key, InStr(key, "{{") + 2)
                                fromlogic = Left(fromlogic, InStr(fromlogic, "{{") - 1)
                                .TextMatrix(r, 9) = fromlogic
                                fromSUBLOCA = Mid(key, InStr(key, "{{") + 2)
                                If InStr(fromSUBLOCA, "{{") > 0 Then
                                    fromSUBLOCA = Mid(fromSUBLOCA, InStr(fromSUBLOCA, "{{") + 2)
                                    If InStr(fromSUBLOCA, "#") > 0 Then
                                        fromSUBLOCA = Left(fromSUBLOCA, InStr(fromSUBLOCA, "#") - 1)
                                    Else
                                        fromSUBLOCA = Left(fromSUBLOCA, InStr(fromSUBLOCA, "{{") - 2)
                                    End If
                                End If
                                .TextMatrix(r, 10) = fromSUBLOCA
                                .TextMatrix(r, 11) = logicBOX(i)
                                .TextMatrix(r, 12) = sublocaBOX(i)
                                Err.Clear
                                If NEWconditionBOX(i) = "" Then NEWconditionBOX(i) = ""
                                If Err.Number = 0 Then
                                    .TextMatrix(r, 3) = NEWconditionBOX(i)
                                    
                                    .TextMatrix(r, 13) = conditionCODE
                                    .TextMatrix(r, 14) = condition
                                Else
                                    .TextMatrix(r, 3) = condition
                                    .TextMatrix(r, 13) = conditionCODE
                                    .TextMatrix(r, 14) = ""
                                End If
                                .TextMatrix(r, 15) = unitBOX(i)
                                
                                
                                'SSOleDBFQA.addITEM commodityLABEL
                                
                                If frmWarehouse.tag = "02040100" Then
                                
                                    PONumb = cell(4).tag
                                    'lineno = SUMMARYlist.TextMatrix(.Rows - 1, 22) ' Changed it from 8 to 22
                                    Dim stocknumberFROM As String
                                    Dim data As New ADODB.Recordset
                                    stocknumberFROM = SUMMARYlist.TextMatrix(i, 1)
                                    
                                    Set data = getDATA("GetStockNumberPOValues", Array(nameSP, PONumb, stocknumberFROM))
                                    If data.RecordCount > 0 Then
     
                                        lineno = data!poi_liitnumb
                                    End If
                                    stocknumberFROM = ""
                                Else
                                
                                    PONumb = ""
                                    lineno = ""
                                    
                                End If
                                
                                quant = SUMMARYlist.TextMatrix(.Rows - 1, 7)
                                
                                MsgBox SUMMARYlist.TextMatrix(0, 0) & vbTab & SUMMARYlist.TextMatrix(0, 1) & vbTab & SUMMARYlist.TextMatrix(0, 2) & vbTab & SUMMARYlist.TextMatrix(0, 3) & vbTab & SUMMARYlist.TextMatrix(0, 7) & vbCrLf & lineno & vbTab & commodityLABEL & vbTab & SUMMARYlist.TextMatrix(0, 2) & vbTab & .TextMatrix(.Rows - 1, 13) & vbTab & quant 'NEWconditionBOX(i).text
                                
                                Call VerifyAddDeleteFQAFromGrid(commodityLABEL, "insert", NEWconditionBOX(i).text, PONumb, lineno, quant)
                                
                                Select Case frmWarehouse.tag
                                    Case "02040400" 'ReturnFromRepair
                                        .TextMatrix(.Rows - 1, 17) = repairBOX(i)
                                        .TextMatrix(.Rows - 1, 18) = cell(5)
                                        .TextMatrix(.Rows - 1, 19) = newDESCRIPTION
                                    Case "02050200" 'AdjustmentEntry
                                    Case "02040200" 'WarehouseIssue
                                    Case "02040500" 'WellToWell
                                        .TextMatrix(.Rows - 1, 17) = Left(Tree.Nodes(i).key, 2)
                                    Case "02040700" 'InternalTransfer
                                    Case "02050300" 'AdjustmentIssue
                                    Case "02040600" 'WarehouseToWarehouse
                                    Case "02040100" 'WarehouseReceipt
                                        .TextMatrix(.Rows - 1, 17) = quantity(i)
                                        .TextMatrix(.Rows - 1, 21) = cell(4).tag
                                        .TextMatrix(.Rows - 1, 22) = repairBOX(i)
                                    Case "02050400" 'Sales
                                    Case "02040300" 'Return from Well
                                End Select
                                .TextMatrix(.Rows - 1, 20) = Left(Tree.Nodes(i).key, 2)
                            Else
                                Err.Clear
                            End If
                        End If
                    End If
                Next
                If .Rows > 2 And .TextMatrix(1, 0) = "" Then .RemoveItem 1
                Call reNUMBER(SUMMARYlist)
            End With
            Call hideDETAILS
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
Dim newPRICE, QTY1, QTY2, uPRICE1, uPRICE2 As Double
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
                            If STOCKlist.TextMatrix(STOCKlist.row, 1) = "§" Then
                                row = STOCKlist.row - 1
                            Else
                                row = STOCKlist.row
                            End If
                            newPRICEok = True
                            If IsNumeric(STOCKlist.TextMatrix(row, 8)) Then
                                QTY1 = CDbl(STOCKlist.TextMatrix(row, 8))
                            Else
                                QTY1 = 0
                                newPRICEok = False
                            End If
                            If IsNumeric(STOCKlist.TextMatrix(row + 1, 8)) Then
                                QTY2 = CDbl(STOCKlist.TextMatrix(row + 1, 8))
                            Else
                                QTY2 = 0
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
                                    uPRICE2 = (QTY1 * uPRICE1) / QTY2
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
                                    uPRICE1 = (QTY2 * uPRICE2) / QTY1
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
Dim nody As Node
Dim sql
Dim datax As New ADODB.Recordset
    For Each nody In Tree.Nodes
        If nody.text = NewString Then
            Tree.Nodes.Remove (Tree.SelectedItem.Index)
            Exit For
        End If
    Next
'    sql = "SELECT * From QTYST6 WHERE " _
'        & "qs6_npecode = '" + nameSP + "' AND " _
'        & "qs6_stcknumb = '" + commodityLABEL + "' AND " _
'        & "qs6_serl = '" + NewString + "' AND " _
'        & "qs6_primqty > 0"
    sql = "SELECT * FROM INVENTORY WHERE " _
        & "Namespace ='" + nameSP + "' AND " _
        & "StockNumber = '" + commodityLABEL + "' AND " _
        & "Serial = '" + NewString + "' AND " _
        & "PrimaryQuantity > 0"
        
    If sql = "" Then
        Cancel = True
        Tree.Nodes.Remove (Tree.SelectedItem.Index)
        Exit Sub
    Else
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenForwardOnly
        If datax.RecordCount > 0 Then
            Tree.Nodes.Remove (Tree.SelectedItem.Index)
            MsgBox "That serial number is already registered in the system"
            Exit Sub
        End If
    End If
    Tree.SelectedItem.key = "@" + NewString
    NewString = "Serial #: " + NewString
End Sub

Public Sub Tree_Click()
On Error Resume Next
Dim n
    With Tree
        n = .SelectedItem.Index
        If n = totalNODE Then
            If nodeSEL <> totalNODE Then
                quantity(totalNODE).backcolor = &H800000
                quantity(totalNODE).ForeColor = vbWhite
            End If
        End If
    End With
End Sub

Private Sub Tree_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Expanded = True
End Sub


Private Sub Tree_LostFocus()
    Tree.SelectedItem = Nothing
    Call Tree_Click
End Sub


Private Sub Tree_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    With Tree
        nodeSEL = .SelectedItem.Index
        If nodeSEL > 0 Then
            quantity(totalNODE).backcolor = &HC0C0C0
            quantity(totalNODE).ForeColor = vbBlack
            If nodeSEL <> totalNODE Then
                quantity(nodeSEL).backcolor = vbWhite
                quantity(nodeSEL).ForeColor = vbBlack
            End If
        End If
    End With
End Sub

Private Sub Tree_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If currentBOX > 0 Then
        Call whitening
        currentBOX = 0
    End If
End Sub

Private Sub Tree_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo getOUT
Dim nody As Node
    If newBUTTON.Enabled Then Exit Sub
    If Button = 2 Then
        Set nody = Tree.HitTest(x, Y)
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

Public Function PopulateCombosWithFQA(CompanyCode As String, Optional LocationCode As String) As Boolean

On Error GoTo ErrHand
PopulateCombosWithFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

LocationCode = Trim(LocationCode)

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Level ='C' order by FQA"

RsCompany.Open , cn

Do While Not RsCompany.EOF

    SSOleCompany.addITEM RsCompany("FQA")
    RsCompany.MoveNext
    
Loop

'RsLocation.source = "select distinct(FQA) from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(Companycode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS'"
RsLocation.source = "select distinct(FQA) from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Level ='LB' OR LEVEL ='LS' order by FQA"

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





Public Function LoadFromFQA(CompanyCode As String, LocationCode As String)

Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset



On Error GoTo ErrHand

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Level ='C'"

RsCompany.Open , cn


'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='LB' or  Level ='LS'"

RsLocation.Open , cn



'Get US Chart FQA

RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='UC'"

RsUC.Open , cn


'Get Cam Chart FQA

RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='CC'"

RsCC.Open , cn
            
            If RsCompany.EOF = False Then
                    TxtCompany = RsCompany("FQA")
            Else
                    TxtCompany = ""
            End If
            
            If RsLocation.EOF = False Then
                    TxtLocation = RsLocation("FQA")
            Else
                    TxtLocation = ""
            End If
            
            If RsUC.EOF = False Then
                    TxtUSChart = RsUC("FQA")
            Else
                        TxtUSChart = ""
            End If
            
            TxtStockType = "0000" 'rsFrom("FromStockType")
            
            If RsCC.EOF = False Then
            
                    TxtCamChart = RsCC("FQA")
            Else
                    TxtCamChart = ""
            End If
            







Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the combo boxes.", vbCritical, "Ims"
End Function

Public Function VerifyAddDeleteFQAFromGrid(STOCKNo As String, Insert_delete As String, Tocondition As String, PONumb As String, lineno As String, quantity As String, Optional RowPositionToBeDeleted As Integer) As Boolean
Dim i As Integer
Dim Flag As Integer

On Error GoTo ErrHand

Insert_delete = UCase(Insert_delete)

    Select Case Insert_delete
    
    Case "INSERT"
    
        If GDefaultValue = False Then GDefaultValue = LoadDefaultValuesForFQA(cell(1).tag, cell(3).tag)
    
    
        Flag = 1
        
        'This is to check if the stockno is no repeatedly added again
    
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
        
                Dim RsStockMaster As New ADODB.Recordset
                Dim StockType As String
                RsStockMaster.source = "select  isnull(stk_stcktype,'0000') stcktype from stockmaster where  stk_stcknumb ='" & STOCKNo & "' and  stk_npecode ='" & nameSP & "'"
                RsStockMaster.Open , cn
                
                If RsStockMaster.EOF = True Then
                    
                
                    StockType = "0000"
                    
                ElseIf Len(Trim(RsStockMaster("stcktype"))) = 0 Then
                        
                        StockType = "0000"
                Else
                        StockType = Trim(RsStockMaster("stcktype"))
                
                End If
                
                SSOleDBFQA.addITEM STOCKNo & vbTab & GDefaultFQA.Company & vbTab & GDefaultFQA.Location & vbTab & GDefaultFQA.UsChart & vbTab & StockType & vbTab & GDefaultFQA.UsChart & vbTab & PONumb & vbTab & lineno & vbTab & Tocondition & vbTab & quantity
                
        End If
        
    Case "DELETE"
    
          For i = 0 To SSOleDBFQA.Rows
          
                If STOCKNo = SSOleDBFQA.columns(0).Value And i = RowPositionToBeDeleted Then
                      
                      SSOleDBFQA.RemoveItem i
                      Exit Function
                
                End If
                
                SSOleDBFQA.MoveNext
          
          Next i
    
    End Select

Exit Function

ErrHand:


MsgBox "Errors occurred while trying to insert a record in the FQA grid.", vbCritical, "Ims"


End Function

Public Function LoadDefaultValuesForFQA(CompanyCode As String, LocationCode As String) As Boolean

On Error GoTo ErrHand
LoadDefaultValuesForFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Level ='C' and ""default"" =1"

RsCompany.Open , cn

'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS' and ""default"" =1"

RsLocation.Open , cn

'Get US Chart FQA

RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC' and ""default"" =1"

RsUC.Open , cn

'Get Cam Chart FQA

RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='CC' and ""default"" =1"

RsCC.Open , cn


                        If RsCompany.EOF = False Then
                            
                            GDefaultFQA.Company = RsCompany("FQA").Value
                            
                            Else
                            
                            GDefaultFQA.Company = ""
                            
                        End If

                        If RsLocation.EOF = False Then
                            
                            GDefaultFQA.Location = RsLocation("FQA").Value
                            
                            Else
                            
                            GDefaultFQA.Location = ""
                            
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
                        
                        

'GDefaultFQA.StockType  =


Set RsCompany = Nothing
Set RsLocation = Nothing
Set RsUC = Nothing
Set RsCC = Nothing

LoadDefaultValuesForFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to get the default values." & Err.description, vbCritical, "Ims"

Err.Clear

End Function

Public Function ChangeMode(ReadOnly As Boolean)

SSOleDBFQA.Enabled = Not ReadOnly

End Function

Private Sub unitBOX_GotFocus(Index As Integer)
    activeBOX = "unitBOX"
End Sub


