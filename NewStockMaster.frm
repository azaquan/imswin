VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form Frm_StockMaster 
   Caption         =   "Stock Master"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11610
   Tag             =   "02010100"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Stocks"
      TabPicture(0)   =   "NewStockMaster.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_ShortDescript"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Long(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSource"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEccn"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSoleEccnNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSDBHeader"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_ShortDescript"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_LongDescript"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtStockSearch"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FraStockHeader"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkLicense"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "SSoleSource"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Technical Specifications"
      TabPicture(1)   =   "NewStockMaster.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTechSpec"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Manufacturer"
      TabPicture(2)   =   "NewStockMaster.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtManuStock"
      Tab(2).Control(1)=   "txtTotal"
      Tab(2).Control(2)=   "TxtLineNumber"
      Tab(2).Control(3)=   "TxtPartnumb"
      Tab(2).Control(4)=   "txtManSpecs"
      Tab(2).Control(5)=   "TxtEstPrice"
      Tab(2).Control(6)=   "ChkManufActive"
      Tab(2).Control(7)=   "SSoleManufacturer"
      Tab(2).Control(8)=   "Label8"
      Tab(2).Control(9)=   "Label6"
      Tab(2).Control(10)=   "Label5"
      Tab(2).Control(11)=   "Label1"
      Tab(2).Control(12)=   "Label2"
      Tab(2).Control(13)=   "Label3"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Recepients"
      TabPicture(3)   =   "NewStockMaster.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lbl_search"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "SSOLEDBFax"
      Tab(3).Control(3)=   "SSOLEDBEmail"
      Tab(3).Control(4)=   "dgRecipientList"
      Tab(3).Control(5)=   "OptEmail"
      Tab(3).Control(6)=   "OptFax"
      Tab(3).Control(7)=   "Txt_search"
      Tab(3).Control(8)=   "cmdRemove"
      Tab(3).Control(9)=   "txt_Recipient"
      Tab(3).Control(10)=   "cmd_Add"
      Tab(3).Control(11)=   "fra_FaxSelect"
      Tab(3).Control(12)=   "TxtRecpStockNumb"
      Tab(3).ControlCount=   13
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSoleSource 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   5600
         Width           =   3750
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         Cols            =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   6615
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin VB.CheckBox chkLicense 
         Alignment       =   1  'Right Justify
         Caption         =   "License Required"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6000
         Width           =   3750
      End
      Begin VB.Frame FraStockHeader 
         Height          =   3015
         Left            =   4150
         TabIndex        =   39
         Top             =   420
         Width           =   7335
         Begin VB.Frame unitsRatio 
            Caption         =   "Units Ratio"
            Height          =   975
            Left            =   3840
            TabIndex        =   71
            Top             =   1560
            Width           =   3375
            Begin VB.TextBox qtySecondary 
               Alignment       =   1  'Right Justify
               DataField       =   "stk_estmprice"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               DataMember      =   "STOCKMASTER"
               Height          =   315
               Left            =   2400
               TabIndex        =   74
               Text            =   "1"
               Top             =   600
               Width           =   900
            End
            Begin VB.TextBox qtyPrimary 
               Alignment       =   1  'Right Justify
               DataField       =   "stk_estmprice"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               DataMember      =   "STOCKMASTER"
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   72
               Top             =   600
               Width           =   900
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "to make..."
               Height          =   225
               Left            =   1320
               TabIndex        =   76
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Secondary"
               Height          =   225
               Left            =   2400
               TabIndex        =   75
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Primary"
               Height          =   225
               Left            =   120
               TabIndex        =   73
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.TextBox TxtStockNumber 
            BackColor       =   &H00FFFFFF&
            DataField       =   "stm_partnumb"
            DataMember      =   "GetStockManufacturer"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox CheckActive 
            Caption         =   "Active"
            DataField       =   "stk_flag"
            DataMember      =   "STOCKMASTER"
            Height          =   195
            Left            =   1440
            TabIndex        =   7
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txt_Minimum 
            DataField       =   "stk_mini"
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   3000
            TabIndex        =   6
            Top             =   1320
            Width           =   540
         End
         Begin VB.TextBox txt_Estimate 
            DataField       =   "stk_estmprice"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   1920
            Width           =   2100
         End
         Begin VB.TextBox txt_Maximum 
            DataField       =   "stk_maxi"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   1320
            Width           =   540
         End
         Begin VB.TextBox txt_Standard 
            DataField       =   "stk_stdrcost"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1440
            TabIndex        =   12
            Top             =   2280
            Width           =   2100
         End
         Begin VB.OptionButton optPool 
            Caption         =   "Pool"
            Height          =   255
            Left            =   5145
            TabIndex        =   8
            Top             =   225
            Width           =   975
         End
         Begin VB.OptionButton optSpecific 
            Caption         =   "Specific"
            Height          =   255
            Left            =   6240
            TabIndex        =   40
            Top             =   225
            Width           =   1065
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SsoleSecUnit 
            Bindings        =   "NewStockMaster.frx":0070
            Height          =   315
            Left            =   5160
            TabIndex        =   4
            Top             =   840
            Width           =   2100
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   1561
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   2699
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            _ExtentX        =   3704
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SsoleStockType 
            Bindings        =   "NewStockMaster.frx":007B
            Height          =   315
            Left            =   1440
            TabIndex        =   9
            Top             =   960
            Width           =   2100
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
            ColumnHeaders   =   0   'False
            RowHeight       =   423
            Columns(0).Width=   3200
            _ExtentX        =   3704
            _ExtentY        =   556
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleCharge 
            Bindings        =   "NewStockMaster.frx":0086
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   2640
            Width           =   2100
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
            ColumnHeaders   =   0   'False
            RowHeight       =   423
            Columns(0).Width=   3200
            _ExtentX        =   3704
            _ExtentY        =   556
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOlePrimUnit 
            Bindings        =   "NewStockMaster.frx":0091
            Height          =   315
            Left            =   5160
            TabIndex        =   3
            Top             =   480
            Width           =   2100
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   1614
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   2858
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            _ExtentX        =   3704
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SsOleCategory 
            Bindings        =   "NewStockMaster.frx":009C
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   600
            Width           =   2100
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   1746
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   2884
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            _ExtentX        =   3704
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label lbl_Computed 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Computer Factor"
            Height          =   225
            Left            =   3720
            TabIndex        =   78
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbl_CompFactor 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "stk_compfctr"
            DataMember      =   "STOCKMASTER"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5160
            TabIndex        =   77
            Top             =   2640
            Width           =   2025
         End
         Begin VB.Label lblManufacNo 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5160
            TabIndex        =   67
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label LblManufno 
            Alignment       =   1  'Right Justify
            Caption         =   "Manufacturers"
            Height          =   255
            Left            =   3720
            TabIndex        =   66
            Top             =   1245
            Width           =   1335
         End
         Begin VB.Label lbl_Category 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbl_PrimaryUnit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Unit"
            Height          =   195
            Left            =   3720
            TabIndex        =   49
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lbl_SecondaryUnit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Unit"
            Height          =   195
            Left            =   3720
            TabIndex        =   48
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbl_StockNum 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Number"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   47
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lbl_Charge 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charge Account"
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbl_Minimum 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum"
            Height          =   195
            Left            =   2160
            TabIndex        =   45
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lbl_Maximum 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum"
            Height          =   225
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lbl_Standard 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Cost"
            Height          =   225
            Left            =   0
            TabIndex        =   43
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lbl_Estimate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Price"
            Height          =   225
            Left            =   0
            TabIndex        =   42
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lbl_StockNum 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Type"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtRecpStockNumb 
         BackColor       =   &H00FFFFC0&
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72840
         MaxLength       =   50
         TabIndex        =   63
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   690
         Left            =   -73545
         TabIndex        =   57
         Top             =   4035
         Width           =   2835
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   1680
            TabIndex        =   59
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -72285
         TabIndex        =   56
         Top             =   3270
         Width           =   1335
      End
      Begin VB.TextBox txt_Recipient 
         Height          =   288
         Left            =   -70320
         MaxLength       =   60
         TabIndex        =   55
         Top             =   3270
         Width           =   6150
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -72240
         TabIndex        =   54
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Txt_search 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   -70320
         MaxLength       =   60
         TabIndex        =   53
         Top             =   3720
         Width           =   3855
      End
      Begin VB.OptionButton OptFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   -70485
         TabIndex        =   52
         Top             =   2910
         Width           =   615
      End
      Begin VB.OptionButton OptEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -69645
         TabIndex        =   51
         Top             =   2910
         Width           =   735
      End
      Begin VB.TextBox TxtStockSearch 
         BackColor       =   &H00C0E0FF&
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   0
         Top             =   540
         Width           =   3735
      End
      Begin VB.TextBox txt_LongDescript 
         DataField       =   "stk_desc"
         DataMember      =   "STOCKMASTER"
         Height          =   1635
         Left            =   4200
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   4620
         Width           =   7260
      End
      Begin VB.TextBox txt_ShortDescript 
         DataField       =   "stk_hazmatclau"
         DataMember      =   "STOCKMASTER"
         Height          =   615
         Left            =   4200
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   3720
         Width           =   7260
      End
      Begin VB.TextBox TxtManuStock 
         BackColor       =   &H00FFFFC0&
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   35
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00FFFFC0&
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -64200
         MaxLength       =   50
         TabIndex        =   32
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox TxtLineNumber 
         BackColor       =   &H00FFFFC0&
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -64920
         MaxLength       =   50
         TabIndex        =   31
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox TxtPartnumb 
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Height          =   315
         Left            =   -71640
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1380
         Width           =   3495
      End
      Begin VB.TextBox txtTechSpec 
         DataField       =   "stk_techspec"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   5325
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   780
         Width           =   11115
      End
      Begin VB.TextBox txtManSpecs 
         DataField       =   "stm_techspec"
         DataMember      =   "GetStockManufacturer"
         Height          =   4095
         Left            =   -74760
         MaxLength       =   3500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2040
         Width           =   11055
      End
      Begin VB.TextBox TxtEstPrice 
         DataField       =   "stm_estmpric"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataMember      =   "GetStockManufacturer"
         Height          =   315
         Left            =   -66480
         MaxLength       =   9
         TabIndex        =   21
         Top             =   1380
         Width           =   1455
      End
      Begin VB.CheckBox ChkManufActive 
         Alignment       =   1  'Right Justify
         Caption         =   "Active"
         DataField       =   "stm_flag"
         DataMember      =   "GetStockManufacturer"
         Height          =   195
         Left            =   -74240
         TabIndex        =   24
         Top             =   1380
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBHeader 
         Height          =   3735
         Left            =   240
         TabIndex        =   23
         Top             =   900
         Width           =   3750
         _Version        =   196617
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   2143
         Columns(0).Caption=   "Stock Number"
         Columns(0).Name =   "Stock Number"
         Columns(0).DataField=   "stk_stcknumb"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "stk_desc"
         Columns(1).FieldLen=   256
         _ExtentX        =   6615
         _ExtentY        =   6588
         _StockProps     =   79
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSoleManufacturer 
         Bindings        =   "NewStockMaster.frx":00A7
         DataSource      =   "deIms"
         Height          =   315
         Left            =   -69360
         TabIndex        =   18
         Top             =   900
         Width           =   3180
         DataFieldList   =   "Column 0"
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1482
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).FieldLen=   256
         Columns(1).Width=   3016
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).FieldLen=   256
         _ExtentX        =   5609
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
         Height          =   2085
         Left            =   -70365
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   570
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         FieldSeparator  =   ";"
         stylesets.count =   2
         stylesets(0).Name=   "RowFont"
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "NewStockMaster.frx":00B2
         stylesets(0).AlignmentText=   0
         stylesets(1).Name=   "ColHeader"
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "NewStockMaster.frx":00CE
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowAddNew     =   -1  'True
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         SelectTypeCol   =   0
         SelectByCell    =   -1  'True
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   10239
         Columns(0).Caption=   "Recepients"
         Columns(0).Name =   "Column 0"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10927
         _ExtentY        =   3678
         _StockProps     =   79
         Caption         =   "Recepients"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBEmail 
         Height          =   2055
         Left            =   -70320
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4180
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   6165
         Columns(1).Caption=   "Email"
         Columns(1).Name =   "Email"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         _ExtentX        =   10927
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Email"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBFax 
         Height          =   2055
         Left            =   -70320
         TabIndex        =   64
         Top             =   4080
         Visible         =   0   'False
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4207
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   6138
         Columns(1).Caption=   "Fax"
         Columns(1).Name =   "Fax"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         _ExtentX        =   10927
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Fax"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSoleEccnNo 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4920
         Width           =   3750
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         Cols            =   3
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   6615
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin VB.Label lblEccn 
         Caption         =   "Eccn #"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   4700
         Width           =   3375
      End
      Begin VB.Label lblSource 
         Caption         =   "Source Of Information"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   5360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Part Specification"
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Number"
         Height          =   255
         Left            =   -74640
         TabIndex        =   65
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Lbl_search 
         Caption         =   "Search by name"
         Height          =   255
         Left            =   -72165
         TabIndex        =   62
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lbl_Long 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Description"
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   38
         Top             =   4380
         Width           =   1695
      End
      Begin VB.Label lbl_ShortDescript 
         BackStyle       =   0  'Transparent
         Caption         =   "Haz. Mat."
         Height          =   225
         Left            =   4200
         TabIndex        =   37
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "StockNumber"
         Height          =   255
         Left            =   -74400
         TabIndex        =   36
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "/"
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
         Left            =   -64440
         TabIndex        =   34
         Top             =   900
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ManuFacturer"
         Height          =   255
         Left            =   -70920
         TabIndex        =   27
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Part Number"
         Height          =   255
         Left            =   -72720
         TabIndex        =   26
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Estimated Price"
         Height          =   255
         Left            =   -68040
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
   End
   Begin LRNavigators.LROleDBNavBar LROleDBNavBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   29
      Top             =   6670
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      NewEnabled      =   -1  'True
      DeleteVisible   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   13680
      TabIndex        =   33
      Top             =   12240
      Width           =   1215
   End
   Begin VB.Label lblStatus 
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
      Left            =   7560
      TabIndex        =   30
      Top             =   6600
      Width           =   2700
   End
End
Attribute VB_Name = "Frm_StockMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Main As ImsStockMaster.Main
Dim lookups As ImsStockMaster.lookups
Dim StockHeader As ImsStockMaster.StockHeader
Dim manufacturer As ImsStockMaster.manufacturer
Dim GFormmode As FormMode
Dim mcheckStockHeader As Boolean
Dim mcheckManufac As Boolean
Dim RsStockNameDesc As ADODB.Recordset
Private Type InitiliazeParams

    StockHeaderCombosLoaded As Boolean
    ManufacturerCombosLoaded As Boolean
    StockAdded() As String
    StocksModified() As String
    NavBarNewEnabled As Boolean
    NavbarEditEnabled As Boolean
    NavbarSaveEnabled As Boolean
    
End Type
Dim GGridFilledWithEmails As Boolean
Dim GGridFilledWithFax As Boolean
Dim GInitiliazeParams As InitiliazeParams
Dim rowguid, locked As Boolean
Dim GRsEccnNo As ADODB.Recordset
Dim GRsSource As ADODB.Recordset
Function getRatio() As Double
    Dim ratio1, ratio2 As Double
    getRatio = 0
    If IsNumeric(qtyPrimary.Text) And IsNumeric(qtySecondary.Text) Then
        ratio1 = CDbl(qtyPrimary.Text)
        ratio2 = CDbl(qtySecondary.Text)
        
        If (ratio1 > 0) And (ratio2 > 0) Then
            If ratio1 > ratio2 Then
                getRatio = ratio1 / ratio2 * 10000
            Else
                getRatio = ratio1 / (ratio2 * 10000)
            End If
            If getRatio = 1 Then getRatio = 0
        End If
    End If
End Function

Private Sub Form_Load()

If deIms.cnIms.State = 0 Then deIms.cnIms.Open

Me.Height = 7530
Me.Width = 11730

Set Main = New ImsStockMaster.Main
Main.Configure deIms.NameSpace, deIms.cnIms

'SSoleEccnNo.DataFieldList = SSoleEccnNo.Columns(1)
'SSoleEccnNo.DataFieldToDisplay = SSoleEccnNo.Columns(1)

SSoleEccnno.Columns(0).Visible = False
SSoleEccnno.Columns(1).Caption = "Eccn#"
SSoleEccnno.Columns(2).Caption = "Desc"

SSoleSource.Columns(1).Width = SSoleSource.Width
SSoleSource.Columns(0).Visible = False

Call PopulateGrid

GFormmode = ChangeMode(mdVisualization)


    SSDBHeader.AllowDelete = False
    LROleDBNavBar1.EditVisible = True
    LROleDBNavBar1.DeleteVisible = False
    LROleDBNavBar1.CancelLastSepVisible = False
    LROleDBNavBar1.LastPrintSepVisible = False
    LROleDBNavBar1.EditEnabled = True
    LROleDBNavBar1.SaveEnabled = False
    Call DisableButtons(Me, LROleDBNavBar1)
    GInitiliazeParams.NavbarEditEnabled = LROleDBNavBar1.EditEnabled
    GInitiliazeParams.NavBarNewEnabled = LROleDBNavBar1.NewEnabled
    GInitiliazeParams.NavbarSaveEnabled = LROleDBNavBar1.SaveEnabled
    Call ToggleNavbar
    Call ToogleStockHeaderControls
    Call ToogleManufacturerControls
    

    mcheckManufac = True
    mcheckStockHeader = True
    
    Call GenerateStyleSheets

    Set StockHeader = InitializeStockHeader
    Call StockHeader.MoveToStocknumber(SSDBHeader.Columns(0).Text)

    Call ClearStockMasterDetails

    Call LoadFromStockheader(StockHeader)
    
    TxtStockSearch.Enabled = True
    'these are the controls for the Recepeints Tab
    '_______________________________
    dgRecipientList.Enabled = True
    OptFax.Enabled = True
    OptEmail.Enabled = True
    txt_Recipient.Enabled = True
    Txt_search.Enabled = True
    SSOLEDBEmail.Enabled = True
    opt_Email.Enabled = True
    opt_FaxNum.Enabled = True
    fra_FaxSelect.Enabled = True
    '__________________________________
    
    SSoleEccnno.Columns(2).Width = 6000
    SSoleEccnno.RowHeight = 500
End Sub

Private Sub SSOleDBGrid2_InitColumnProps()

End Sub

Public Function InitializeLookup() As ImsStockMaster.lookups

 If lookups Is Nothing Then Set lookups = Main.lookups
    
 Set InitializeLookup = lookups
    
End Function

Public Function InitializeStockHeader() As StockHeader

If StockHeader Is Nothing Then Set StockHeader = Main.StockHeader

Set InitializeStockHeader = StockHeader

End Function


Public Function InitializeManufacturer() As manufacturer

If manufacturer Is Nothing Then Set manufacturer = Main.manufacturer

Set InitializeManufacturer = manufacturer

End Function


Public Function PopulateGrid()

If deIms.rsActiveStockmasterLookup.State = 1 Then deIms.rsActiveStockmasterLookup.Close

Call deIms.ActiveStockMasterLooKUP(deIms.NameSpace)

Set lookups = InitializeLookup
Set RsStockNameDesc = lookups.GetStocknumbers

Set SSDBHeader.DataSource = RsStockNameDesc

End Function

Public Function ClearStockMasterDetails()
'TxtStockSearch = ""
TxtStockNumber = ""
SsOleCategory = ""
SsOleCategory.Tag = ""
SSOlePrimUnit = ""
SSOlePrimUnit.Tag = ""
SsoleSecUnit = ""
SsoleSecUnit.Tag = ""
txt_Maximum = ""
txt_Minimum = ""
CheckActive.value = 1
'Juan 2010/8/9
'optPool.value = False
'optSpecific.value = True
optPool.value = True
qtySecondary = "1"
'--------------------------

SsoleStockType = ""
SsoleStockType.Tag = ""
lbl_CompFactor = ""
txt_Estimate = ""
SSOleCharge = ""
SSOleCharge.Tag = ""
txt_Standard = ""
txt_ShortDescript = ""
txt_LongDescript = ""
txtTechSpec = ""
SSoleEccnno = ""
SSoleEccnno.Tag = 0
SSoleSource = ""
SSoleSource.Tag = 0
chkLicense.value = False

End Function

Public Function LoadFromStockheader(StockHeader As ImsStockMaster.StockHeader)

If GInitiliazeParams.StockHeaderCombosLoaded = False Then Call LoadCombos

If Len(Trim(SSDBHeader.Columns(0).Text)) = 0 Then Exit Function

Call ClearStockMasterDetails

If StockHeader.Count = 0 Then Exit Function
       
TxtStockNumber = StockHeader.StockNumber

SsOleCategory.Tag = StockHeader.CategoryCode
SsOleCategory = GetNameForTagFromCombo(SsOleCategory, StockHeader.CategoryCode)

SSOlePrimUnit.Tag = StockHeader.PrimUOfMeasure
SSOlePrimUnit = GetNameForTagFromCombo(SSOlePrimUnit, StockHeader.PrimUOfMeasure)

SsoleSecUnit.Tag = StockHeader.SecoUOfMeasure
SsoleSecUnit = GetNameForTagFromCombo(SsoleSecUnit, StockHeader.SecoUOfMeasure)

txt_Maximum = StockHeader.Maximum

txt_Minimum = StockHeader.Minimum

CheckActive.value = IIf(StockHeader.Activeflag = True, 1, 0)

optPool.value = StockHeader.PoolOrSpecific

optSpecific.value = Not optPool.value

SsoleStockType = StockHeader.stocktype

SsoleStockType.Tag = StockHeader.stocktype

lbl_CompFactor = StockHeader.ComputationFactor

'Juan 2010-8--9
qtyPrimary.Text = StockHeader.ratio1
qtySecondary.Text = StockHeader.ratio2
'-----------------

txt_Estimate = StockHeader.estmprice

SSOleCharge.Tag = StockHeader.characctcode

SSOleCharge = GetNameForTagFromCombo(SSOleCharge, StockHeader.characctcode)

txt_Standard = StockHeader.stdrcost

txt_ShortDescript = StockHeader.hazmatclau

txt_LongDescript = StockHeader.Description

txtTechSpec = StockHeader.Techspec

'SSoleEccnno = StockHeader.Eccnno
SSoleEccnno.Tag = StockHeader.Eccnid

'SSoleSource = StockHeader.Eccnsourcename
SSoleSource.Tag = StockHeader.Eccnsource

chkLicense.value = IIf(StockHeader.Eccnlicsreq = True, 1, 0)



Set lookups = InitializeLookup

If GRsEccnNo Is Nothing Then Set GRsEccnNo = lookups.GetListofEccns(0)

If Len(StockHeader.Eccnid) > 0 And GRsEccnNo.RecordCount > 0 Then

 GRsEccnNo.MoveFirst
 GRsEccnNo.Find "eccnid=" & StockHeader.Eccnid
 If GRsEccnNo.EOF = False Then SSoleEccnno.Text = GRsEccnNo!eccn_no
 
End If

If GRsSource Is Nothing Then Set GRsSource = lookups.GetListofEccnSource(0)

If Len(StockHeader.Eccnid) > 0 And GRsSource.RecordCount > 0 Then

    GRsSource.MoveFirst
    GRsSource.Find "SourceID=" & StockHeader.Eccnsource
    If GRsSource.EOF = False Then SSoleSource.Text = GRsSource!Source
 
End If


lblManufacNo = lookups.HowManyManufacturers(StockHeader.StockNumber) ' & " Manufacturers "

End Function

Public Function LoadCombos() As Boolean

LoadCombos = False
On Error GoTo ErrHandler

    If GInitiliazeParams.StockHeaderCombosLoaded = False Then
    
       Call populateCategory
       Call populatePriAndSecUnit
       Call populateStockType
       Call populateChargeAccount
       Call PopulateEccn
       Call PopulateEccnSource
       
       GInitiliazeParams.StockHeaderCombosLoaded = True
       
    End If
    
LoadCombos = True

Exit Function

ErrHandler:
Err.Clear



End Function

Public Function populateCategory() As Boolean

populateCategory = False

On Error GoTo ErrHandler

If deIms.rsCATEGORY.State = 1 Then deIms.rsCATEGORY.Close

Call deIms.Category(deIms.NameSpace)

SsOleCategory.RemoveAll

Do While Not deIms.rsCATEGORY.EOF

SsOleCategory.AddItem deIms.rsCATEGORY("cate_catecode") & vbTab & deIms.rsCATEGORY("cate_catename")

deIms.rsCATEGORY.MoveNext

Loop

populateCategory = True

Exit Function

ErrHandler:

Err.Clear

End Function

Public Function populatePriAndSecUnit() As Boolean

populatePriAndSecUnit = False

On Error GoTo ErrHandler

If deIms.rsGetUnit.State = 1 Then deIms.rsGetUnit.Close

Call deIms.Getunit(deIms.NameSpace)

SSOlePrimUnit.RemoveAll

SsoleSecUnit.RemoveAll

Do While Not deIms.rsGetUnit.EOF

    SSOlePrimUnit.AddItem deIms.rsGetUnit("uni_code") & vbTab & deIms.rsGetUnit("uni_desc")
    
    SsoleSecUnit.AddItem deIms.rsGetUnit("uni_code") & vbTab & deIms.rsGetUnit("uni_desc")
    
    deIms.rsGetUnit.MoveNext

Loop


populatePriAndSecUnit = True

Exit Function

ErrHandler:

Err.Clear
End Function

Public Function populateStockType() As Boolean

Dim RsGetStockType As ADODB.Recordset

populateStockType = False

On Error GoTo ErrHandler

Set lookups = InitializeLookup

Set RsGetStockType = lookups.GetStockTypes

SsoleStockType.RemoveAll

Do While Not RsGetStockType.EOF

    SsoleStockType.AddItem RsGetStockType("sty_stcktype") '& vbTab & deIms.RsGetStockType("sty_desc")
    
    RsGetStockType.MoveNext

Loop

populateStockType = True

Exit Function

ErrHandler:

Err.Clear


End Function

Public Function PopulateEccn() As Boolean

Dim RsGetEccn As ADODB.Recordset

PopulateEccn = False

On Error GoTo ErrHandler

Set lookups = InitializeLookup

Set RsGetEccn = lookups.GetListofEccns(1)

SSoleEccnno.RemoveAll

Do While Not RsGetEccn.EOF

    SSoleEccnno.AddItem RsGetEccn("eccnid") & vbTab & RsGetEccn("eccn_no") & vbTab & RsGetEccn("eccn_desc")
    
    RsGetEccn.MoveNext

Loop

PopulateEccn = True

Exit Function

ErrHandler:

Err.Clear


End Function

Public Function PopulateEccnSource() As Boolean

Dim RsSource As ADODB.Recordset

PopulateEccnSource = False

On Error GoTo ErrHandler

Set lookups = InitializeLookup

Set RsSource = lookups.GetListofEccnSource(1)

SSoleSource.RemoveAll

Do While Not RsSource.EOF

    SSoleSource.AddItem RsSource("SourceID") & vbTab & RsSource("source")
    
    RsSource.MoveNext

Loop

PopulateEccnSource = True

Exit Function

ErrHandler:

Err.Clear

End Function
Public Function populateChargeAccount() As Boolean

populateChargeAccount = False

On Error GoTo ErrHandler

If deIms.rsCHARGE.State = 1 Then deIms.rsCHARGE.Close

Call deIms.CHARGE(deIms.NameSpace)

SSOleCharge.RemoveAll
SSOleCharge.AddItem "N\A"
Do While Not deIms.rsCHARGE.EOF

    SsOleCategory.AddItem deIms.rsCHARGE("cha_acctcode") & vbTab & deIms.rsCHARGE("cha_acctname")
    
    deIms.rsCHARGE.MoveNext

Loop

populateChargeAccount = True

Exit Function

ErrHandler:

Err.Clear
End Function

Public Function GetNameForTagFromCombo(SsOleCombo As SSOleDBCombo, Code As String) As String

Dim desc As String

On Error GoTo ErrHandler

 SsOleCombo.MoveFirst
 
 Do While Not SsOleCombo.AddItemRowIndex(SsOleCombo.Bookmark) = SsOleCombo.Rows - 1
 'Debug.Print SsOleCombo.AddItemRowIndex(SsOleCombo.Bookmark)
      If Trim(UCase(SsOleCombo.Columns(0).Text)) = Trim(UCase(Code)) Then
    
            desc = SsOleCombo.Columns(1).Text
            
            Exit Do
            
      End If
      
      SsOleCombo.MoveNext
  
  Loop
  
  GetNameForTagFromCombo = desc
  
  Exit Function
  
ErrHandler:
  
End Function

Public Function LoadFromManufacturer(manufacturer As ImsStockMaster.manufacturer) As Boolean

Dim desc As String

LoadFromManufacturer = False

On Error GoTo ErrHandler


   Set manufacturer = InitializeManufacturer

'    If Trim(UCase(manufacturer.StockNumber)) <> Trim(UCase(TxtStockNumber)) Then
            
        If loadManufacturerCombos = False Then Exit Function
            
        If manufacturer.Count = 0 Then Exit Function
        
        Call ClearManufacturer
        
        SSoleManufacturer.Tag = manufacturer.ManufactCode
            
        SSoleManufacturer = GetNameForTagFromCombo(SSoleManufacturer, SSoleManufacturer.Tag)
        
        TxtPartnumb = manufacturer.PartNumb
        
        TxtEstPrice = manufacturer.Estmpric
        
        txtManSpecs = manufacturer.Techspec

        TxtLineNumber = manufacturer.AbsolutePosition
        
        txtTotal = manufacturer.Count

        TxtManuStock = manufacturer.StockNumber
        
'    End If

LoadFromManufacturer = True

Exit Function

ErrHandler:
  
End Function

Public Function ClearManufacturer()

TxtManuStock = ""

SSoleManufacturer = ""

SSoleManufacturer.Tag = ""

TxtPartnumb = ""

TxtEstPrice = ""

txtManSpecs = ""

TxtLineNumber = ""

txtTotal = ""

End Function

Private Sub Form_Unload(Cancel As Integer)
Dim x As InitiliazeParams
Set StockHeader = Nothing
Set manufacturer = Nothing
Set lookups = Nothing
Set Main = Nothing
Set RsStockNameDesc = Nothing
GGridFilledWithEmails = False
GGridFilledWithFax = False
GInitiliazeParams = x

End Sub

Private Sub LROleDBNavBar1_BeforeSaveClick()

Dim x As String
Dim EditMode As Integer

 Screen.MousePointer = vbHourglass

If ValidateStockHeaderValues = True Then

    Call SaveToStockHeader
    
    EditMode = StockHeader.EditMode
    
    x = Main.Save
    
    If Len(Trim(x)) > 0 Then
    
        MsgBox "Errors Occurred while trying to save the StockNumber. Please try again. Error Description :" & x, vbCritical, "Imswin"
        
        LROleDBNavBar1.SaveEnabled = True
        
    ElseIf Len(Trim(x)) = 0 Then
            
        GFormmode = ChangeMode(mdVisualization)
        Call ToggleNavbar
        Call ToogleManufacturerControls
        Call ToogleStockHeaderControls
        
        RsStockNameDesc.CancelUpdate
        RsStockNameDesc.Close
        RsStockNameDesc.Open , deIms.cnIms, 3, 3

        RsStockNameDesc.Requery
        
        Set lookups = InitializeLookup
        
        lblManufacNo = lookups.HowManyManufacturers(TxtStockNumber)
        
        Set manufacturer = Nothing
        
        Call StoreStocksPlayedwith(EditMode)
        
        Call MoveGridTo(TxtStockNumber)
        Call LoadStockMaster
        
         MsgBox "Saved Successfully"
         
         
         Dim imsLock As imsLock.Lock
         Set imsLock = New imsLock.Lock
         Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
            
          
        
    End If
    
Else

    LROleDBNavBar1.SaveEnabled = True
    
End If

 Screen.MousePointer = vbArrow

End Sub

Private Sub LROleDBNavBar1_OnCancelClick()

    If GFormmode = mdCreation Then
        
        Select Case SSTab1.Tab
        
            Case 0
            
                 StockHeader.CancelUpdate
                 Call ClearStockMasterDetails
                 Call LoadFromStockheader(StockHeader)
                 
                GFormmode = ChangeMode(mdVisualization)
                Call ToggleNavbar
                Call ToogleManufacturerControls
                Call ToogleStockHeaderControls
                
                RsStockNameDesc.CancelUpdate
                RsStockNameDesc.Requery
                
                Set manufacturer = Nothing
                Call MoveGridTo(TxtStockNumber)
                mcheckStockHeader = True
                
            Case 2
            
                 manufacturer.CancelUpdate
                 Call ClearManufacturer
                 Call LoadFromManufacturer(manufacturer)
                 mcheckManufac = True
            
         End Select
    
    ElseIf GFormmode = mdModification Then
    
        
        Select Case SSTab1.Tab
        
            Case 0
            
                 StockHeader.CancelUpdate
                 Call StockHeader.MoveToStocknumber(TxtStockNumber, True)
                 Call ClearStockMasterDetails
                 Call LoadFromStockheader(StockHeader)
                 
                GFormmode = ChangeMode(mdVisualization)
                Call ToggleNavbar
                Call ClearManufacturer
                Call ToogleManufacturerControls
                Call ToogleStockHeaderControls
                
                RsStockNameDesc.CancelUpdate
                RsStockNameDesc.Requery
                
                Set manufacturer = Nothing
                 
                Call MoveGridTo(TxtStockNumber)
                
                mcheckStockHeader = True
                
            Case 2
                
                 manufacturer.CancelUpdate
                 Call ClearManufacturer
                 Call LoadFromManufacturer(manufacturer)
                 Call ToogleManufacturerCombo
                 mcheckManufac = True
                 
         End Select
    
    End If
    
 Dim imsLock As imsLock.Lock
 Set imsLock = New imsLock.Lock
 Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode



End Sub

Private Sub LROleDBNavBar1_OnCloseClick()

 Dim imsLock As imsLock.Lock
 Set imsLock = New imsLock.Lock
 Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

Unload Me
End Sub

Private Sub LROleDBNavBar1_OnEditClick()

If Len(Trim(TxtStockNumber)) = 0 Then

    MsgBox "No Stock Number selected to modify. Please make a selection first.", vbInformation, "Ims"
    Exit Sub
End If
                          Dim currentformname, currentformname1
                        currentformname = Me.Name 'Forms(3).Name
                        currentformname1 = Me.Name 'Forms(3).Name
                         Dim imsLock As imsLock.Lock
                         Dim ListOfPrimaryControls() As String
                         
                         Set imsLock = New imsLock.Lock
                        
                         ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
                        
                         Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid)   'lock should be here, added by jawdat, 2.1.02
                        
                    If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
                    
                        Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
                        
                    End If
    
GFormmode = ChangeMode(mdModification)

Call ToggleNavbar
Call ToogleStockHeaderControls
Call ToogleManufacturerControls

End Sub

Private Sub LROleDBNavBar1_OnEMailClick()

Dim ParamsForRPTI(1) As String

Dim rptinf As RPTIFileInfo

Dim ParamsForCrystalReports(1) As String

Dim subject As String

Dim FieldName As String

Dim Message As String

Dim attention As String

Dim Recipients As New ADODB.Recordset
Dim i As Integer
On Error GoTo ErrHandler


If dgRecipientList.Rows = 0 Then Exit Sub
                   
    Recipients.Fields.Append "Recipients", adBSTR, 40
    'Recipients.Fields(0).Name = "Recipients"
    
    dgRecipientList.MoveFirst
    
    Recipients.Open
    
    For i = 0 To dgRecipientList.Rows - 1
    
    'Do While dgRecipientList.AddItemRowIndex(dgRecipientList.Bookmark) < dgRecipientList.Rows
    
        Recipients.AddNew
        
        Recipients("Recipients") = dgRecipientList.Columns(0).value
    
        dgRecipientList.MoveNext
    
    Next
    
                    
    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    
    ParamsForCrystalReports(1) = "stcknumb;" + Trim$(TxtStockNumber) + ";TRUE"
    
    ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
    
    ParamsForRPTI(1) = "stcknumb=" & Trim$(TxtStockNumber)
    
    FieldName = "Recipients"
    
    subject = "Stock Master Record " & Trim$(TxtStockNumber)
    
    If ConnInfo.EmailClient = Outlook Then
    
        'Call sendOutlookEmailandFax("stckmaster1.rpt", "StockMaster", MDI_IMS.CrystalReport1, ParamsForCrystalReports, Recipients, subject, attention) MM 030209 EFCR11
        Call sendOutlookEmailandFax(Report_EmailFax_Stockmaster_name, "StockMaster", MDI_IMS.CrystalReport1, ParamsForCrystalReports, Recipients, subject, attention)
    
    ElseIf ConnInfo.EmailClient = ATT Then
    
        Call SendAttFaxAndEmail("stckmaster1.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, Recipients, subject, Message, FieldName)

    ElseIf ConnInfo.EmailClient = Unknown Then
    
        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

    End If

   ' Call Recipients.Delete(adAffectAllChapters)

    Set Recipients = Nothing
    
Exit Sub

ErrHandler:

MsgBox "Errors occurred while trying to generate the Stock master report. Error Description : " & Err.Description, vbCritical, "Ims"

Err.Clear

End Sub

Private Sub LROleDBNavBar1_OnFirstClick()

If GFormmode <> mdVisualization Then

    If SaveToManufacturer = False Then Exit Sub
    
End If

If manufacturer.MoveFirst = True Then

    Call ClearManufacturer
    Call LoadFromManufacturer(manufacturer)
    Call ToogleManufacturerCombo
End If
    
End Sub

Private Sub LROleDBNavBar1_OnLastClick()


If GFormmode <> mdVisualization Then

    If SaveToManufacturer = False Then Exit Sub
    
End If

If manufacturer.MoveLast = True Then

    Call ClearManufacturer
    Call LoadFromManufacturer(manufacturer)
    Call ToogleManufacturerCombo
    
End If

End Sub

Private Sub LROleDBNavBar1_OnNewClick()
'Juan 2010-8-9
qtyPrimary = "1"
qtySecondary = "1"
'------------------


Select Case SSTab1.Tab

    Case 0
    
         If GFormmode = mdVisualization Then
            
            GFormmode = ChangeMode(mdCreation)
            
            Set StockHeader = InitializeStockHeader
             RsStockNameDesc.AddNew
            
            If StockHeader.AddNew = False Then
            
                MsgBox "Errors Occurred while trying to Add a new record(in the Dll).", vbCritical, "Imswin"
        
            End If
SSOleCharge = "N/A"
txt_Maximum = 0
txt_Minimum = 0
            SSDBHeader.Scroll 0, SSDBHeader.Rows
            
            SSDBHeader.MoveLast
            
SSOleCharge = "N/A"
txt_Maximum = 0
txt_Minimum = 0
            Call ClearStockMasterDetails
            
            Call SetInitialValuesStockmaster
                
            Call LoadCombos
            
            Call ToggleNavbar
            
            Call ToogleStockHeaderControls
            
            Call ToogleManufacturerControls
            
            If GFormmode = mdCreation Then
                TxtStockNumber.SetFocus
            End If
         End If
SSOleCharge = "N/A"
txt_Maximum = 0
txt_Minimum = 0
    Case 2
    
       
        If GFormmode <> mdVisualization Then
                
                If ValidatemanufacturerValues = False Then Exit Sub
                
                If SaveToManufacturer = False Then Exit Sub
                        
                Set manufacturer = InitializeManufacturer
                
                If manufacturer.AddNew = False Then
                
                        MsgBox "Errors Occurred while trying to Add a new record", vbCritical, "Imswin"
                        
                        Exit Sub
                    
                 End If
                
                Call ClearManufacturer
                
                Call SetInitialValuesManufacturer
                
                Call loadManufacturerCombos
        
        End If
        
    
End Select
End Sub


Private Sub LROleDBNavBar1_OnNextClick()


 If GFormmode <> mdVisualization Then
 
     If SaveToManufacturer = False Then Exit Sub
     
 End If

If manufacturer.MoveNext = True Then

    Call ClearManufacturer
    Call LoadFromManufacturer(manufacturer)
    Call ToogleManufacturerControls
End If

End Sub

Private Sub LROleDBNavBar1_OnPreviousClick()


 If GFormmode <> mdVisualization Then
     
     If ValidatemanufacturerValues = False Then Exit Sub
     
     If SaveToManufacturer = False Then Exit Sub
     
 End If

If manufacturer.MovePrevious = True Then

    Call ClearManufacturer
    Call LoadFromManufacturer(manufacturer)
    Call ToogleManufacturerCombo
    
End If

End Sub

Private Sub LROleDBNavBar1_OnPrintClick()
On Error GoTo ErrHandler
Screen.MousePointer = vbHourglass
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath & "Stckmaster1.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stcknumb;" & Trim$(TxtStockNumber) & ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00119") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock Master", msg1) 'J modified
        Call translator.Translate_Reports("Stckmaster1.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
    
 Screen.MousePointer = vbArrow
    
     Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    Screen.MousePointer = vbArrow
End Sub

Private Sub optPool_Click()
    If optPool.value Then
        SsoleSecUnit.Enabled = True
        unitsRatio.Enabled = True
    End If
End Sub

Private Sub optSpecific_Click()
    'Juan 2010/8/9
    If optSpecific.value Then
        SsoleSecUnit.Enabled = False
        unitsRatio.Enabled = False
        qtyPrimary.Text = "1"
        qtySecondary.Text = "1"
        lbl_CompFactor = ""
    End If
    '-----------
End Sub


Private Sub qtyPrimary_Change()
    'Juan 2010/8/9
    If IsNumeric(qtyPrimary.Text) And IsNumeric(qtySecondary.Text) Then
        qtyPrimary_Validate (False)
    End If
    '--------------
End Sub

Private Sub qtyPrimary_Validate(Cancel As Boolean)
    If Not Cancel Then
        lbl_CompFactor = CDbl(getRatio)
    End If
End Sub


Private Sub qtySecondary_Change()
    'Juan 2010/8/9
    If IsNumeric(qtyPrimary.Text) And IsNumeric(qtySecondary.Text) Then
        qtySecondary_Validate (False)
    End If
    '--------------
End Sub

Private Sub qtySecondary_Validate(Cancel As Boolean) ' Juan 2010-8-9
    If Not Cancel Then
        lbl_CompFactor = CDbl(getRatio)
    End If
End Sub


Private Sub SSDBHeader_BeforeRowColChange(Cancel As Integer)
'To precent the user from selecting any other Stock from the Gird while one is being modfied.
If GFormmode = mdModification Then Cancel = 1
End Sub

Private Sub SSDBHeader_Click()
If GFormmode = mdVisualization Then

    Call LoadStockMaster
    
 Else
 
    
 
 End If
    
    
End Sub

Private Sub SSDBHeader_KeyDown(KeyCode As Integer, Shift As Integer)

'KeyCode = 0

End Sub

Private Sub SSDBHeader_KeyPress(KeyAscii As Integer)
 
If KeyAscii = 13 Then

Call LoadStockMaster

Else

KeyAscii = 0
 
End If
End Sub

Private Sub SSDBHeader_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

If LastRow = RsStockNameDesc.AbsolutePosition Then Exit Sub

    Call LoadStockMaster
    

End Sub

Private Sub SSDBHeader_RowLoaded(ByVal Bookmark As Variant)
    Dim i As Integer
                    
 If IsArrayLoaded(GInitiliazeParams.StockAdded) Then
     
     For i = 0 To UBound(GInitiliazeParams.StockAdded)
         
            If UCase(Trim(SSDBHeader.Columns(0).Text)) = UCase(Trim(GInitiliazeParams.StockAdded(i))) Then
        
                
                    SSDBHeader.Columns(0).CellStyleSet "RowAdded"
                    SSDBHeader.Columns(1).CellStyleSet "RowAdded"
                    Exit Sub
            End If
        
      Next i
      
 End If
 
 
 If IsArrayLoaded(GInitiliazeParams.StocksModified) Then
     
     For i = 0 To UBound(GInitiliazeParams.StocksModified)
         
            If UCase(Trim(SSDBHeader.Columns(0).Text)) = UCase(Trim(GInitiliazeParams.StocksModified(i))) Then
        
                
                    SSDBHeader.Columns(0).CellStyleSet "RowModified"
                    SSDBHeader.Columns(1).CellStyleSet "RowModified"
                    Exit Sub
            End If
        
      Next i
      
 End If
 
End Sub

Private Sub SsOleCategory_Click()
SsOleCategory.Tag = UCase(Trim(SsOleCategory.Columns(0).Text))
End Sub

Private Sub SsOleCategory_Validate(Cancel As Boolean)


If SsOleCategory.IsItemInList = False And GFormmode <> mdVisualization Then
    
    MsgBox "Please select a valid Category.", vbInformation, "Imswin"
    
    Cancel = True
    
End If

End Sub

Private Sub SSOleCharge_GotFocus()
Call HighlightBackground(SSOleCharge)
End Sub

Private Sub SSOleCharge_LostFocus()
Call NormalBackground(SSOleCharge)
End Sub

Private Sub SsOleCategory_GotFocus()
Call HighlightBackground(SsOleCategory)
End Sub

Private Sub SsOleCategory_LostFocus()
Call NormalBackground(SsOleCategory)
End Sub

Private Sub SSOleCharge_Validate(Cancel As Boolean)

If SSOleCharge.IsItemInList = False And GFormmode <> mdVisualization Then
    
    MsgBox "Please select a valid Account.", vbInformation, "Imswin"
    
    Cancel = True
    
End If
End Sub

Private Sub SSoleEccnno_Validate(Cancel As Boolean)

If ConnInfo.Eccnactivate = Constno Then Exit Sub

If SSoleEccnno.IsItemInList = False And GFormmode <> mdVisualization Then

        MsgBox "Eccn # does not exist in the list, please select a valid one.", , "Imswin"
        SSoleEccnno.SetFocus
        Cancel = True
        
End If
End Sub

Private Sub SSoleManufacturer_Click()
SSoleManufacturer.Tag = Trim(UCase(SSoleManufacturer.Columns(0).Text))
End Sub

Private Sub SSoleManufacturer_GotFocus()
Call HighlightBackground(SSoleManufacturer)
End Sub

Private Sub SSoleManufacturer_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SSoleManufacturer.DroppedDown Then SSoleManufacturer.DroppedDown = True
 End If
End Sub

Private Sub SSoleManufacturer_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SSOleCharge.MoveNext
If KeyAscii = 13 Then
    TxtPartnumb.SetFocus
End If
End Sub

Private Sub SSoleManufacturer_LostFocus()
Call NormalBackground(SSoleManufacturer)
End Sub

Private Sub SSoleManufacturer_Validate(Cancel As Boolean)
If SSoleManufacturer.IsItemInList = False And GFormmode <> mdVisualization Then
    
    MsgBox "Please select a valid Manufacturer.", vbInformation, "Imswin"
    
    Cancel = True
    
End If
End Sub

Private Sub SSOlePrimUnit_Change()
If optPool Then
Else
    'juan 2010-11-9 to write down on sec unit
    If SSOlePrimUnit = "" Then
        If SsoleSecUnit = "" Then
        Else
            SsoleSecUnit = ""
        End If
    Else
        SsoleSecUnit = SSOlePrimUnit
    End If
End If
End Sub

Private Sub SSOlePrimUnit_Click()

SSOlePrimUnit.Tag = UCase(Trim(SSOlePrimUnit.Columns(0).Text))
If optPool Then
    SsoleSecUnit.Text = "EACH"
Else
    'juan 2010-11-9 to write down on sec unit
    SsoleSecUnit = SSOlePrimUnit
End If
SsoleSecUnit.Tag = UCase(Trim(SSOlePrimUnit.Columns(0).Text))

End Sub
Private Sub SSOleCharge_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SSOleCharge.DroppedDown Then SSOleCharge.DroppedDown = True
 End If
End Sub

Private Sub SSOleCharge_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SSOleCharge.MoveNext
End Sub

Private Sub SSOlePrimUnit_GotFocus()
Call HighlightBackground(SSOlePrimUnit)
End Sub

Private Sub SSOlePrimUnit_Validate(Cancel As Boolean)

If SSOlePrimUnit.IsItemInList = False And GFormmode <> mdVisualization Then
    
    MsgBox "Please select a valid Unit.", vbInformation, "Imswin"
    
    Cancel = True
    
End If

End Sub

Private Sub SsoleSecUnit_GotFocus()
Call HighlightBackground(SsoleSecUnit)
End Sub

Private Sub SsoleSecUnit_Validate(Cancel As Boolean)
If SsoleSecUnit.IsItemInList = False And GFormmode <> mdVisualization Then
    
    MsgBox "Please select a valid Unit.", vbInformation, "Imswin"
    
    Cancel = True
    
End If

End Sub

Private Sub SSoleSource_Click()
SSoleSource.Tag = Trim(UCase(SSoleSource.Columns(0).Text))
End Sub

Private Sub SSoleSource_GotFocus()
Call HighlightBackground(SSoleSource)
End Sub

Private Sub SSoleSource_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SSoleSource.DroppedDown Then SSoleSource.DroppedDown = True
 End If
End Sub

Private Sub SSoleSource_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SSoleSource.MoveNext
End Sub

Private Sub SSoleSource_LostFocus()
Call NormalBackground(SSoleSource)
End Sub
Private Sub SSoleEccnNo_Click()
SSoleEccnno.Tag = Trim(UCase(SSoleEccnno.Columns(0).Text))
End Sub

Private Sub SSoleEccnNo_GotFocus()
Call HighlightBackground(SSoleEccnno)
End Sub

Private Sub SSoleEccnNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SSoleEccnno.DroppedDown Then SSoleEccnno.DroppedDown = True
 End If
End Sub

Private Sub SSoleEccnNo_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SSoleEccnno.MoveNext
End Sub

Private Sub SSoleEccnNo_LostFocus()
Call NormalBackground(SSoleEccnno)
End Sub
Private Sub SsoleStockType_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SsoleStockType.DroppedDown Then SsoleStockType.DroppedDown = True
 End If
End Sub

Private Sub SsoleStockType_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SsoleStockType.MoveNext
End Sub
Private Sub SsoleSecUnit_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SsoleSecUnit.DroppedDown Then SsoleSecUnit.DroppedDown = True
 End If
End Sub

Private Sub SsoleSecUnit_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SsoleSecUnit.MoveNext
End Sub
Private Sub SSOlePrimUnit_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SSOlePrimUnit.DroppedDown Then SSOlePrimUnit.DroppedDown = True
 End If
End Sub

Private Sub SSOlePrimUnit_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SSOlePrimUnit.MoveNext
End Sub
Private Sub SsOleCategory_KeyDown(KeyCode As Integer, Shift As Integer)
 If GFormmode <> mdVisualization Then
    If Not SsOleCategory.DroppedDown Then SsOleCategory.DroppedDown = True
 End If
End Sub

Private Sub SsOleCategory_KeyPress(KeyAscii As Integer)
If GFormmode <> mdVisualization Then SsOleCategory.MoveNext
End Sub
Private Sub SSOlePrimUnit_LostFocus()

Dim str As String

'Juan 2010/8/9
'If Len(Trim(SSOlePrimUnit)) > 0 And Len(Trim(SsoleSecUnit)) > 0 And Trim(SSOlePrimUnit.Tag) <> Trim(StockHeader.PrimUOfMeasure) Then
'
'    str = InputBox("Please enter how many " & SsoleSecUnit & " it takes to make 1 " & SSOlePrimUnit)
'
'    If Len(str) > 0 And IsNumeric(str) Then
'
'        lbl_CompFactor = CDbl(1000 / str)
'
'    End If
'
'    End If
'----------------------

Call NormalBackground(SSOlePrimUnit)

End Sub

Private Sub SsoleSecUnit_Click()
SsoleSecUnit.Tag = UCase(Trim(SsoleSecUnit.Columns(0).Text))
End Sub

Private Sub SsoleSecUnit_LostFocus()

Dim str As String

Set StockHeader = InitializeStockHeader

'Juan 2010/8/9
'If Len(Trim(SSOlePrimUnit)) > 0 And Len(Trim(SsoleSecUnit)) > 0 And Trim(SsoleSecUnit.Tag) <> StockHeader.SecoUOfMeasure Then
'
'  If Trim(UCase(SSOlePrimUnit)) <> Trim(UCase(SsoleSecUnit)) Then
'
'        str = InputBox("Please enter how many " & SsoleSecUnit & " it takes to make 1 " & SSOlePrimUnit)
'
'        If Len(Trim(str)) > 0 And IsNumeric(str) Then lbl_CompFactor = CDbl(1000 / str)
'
'  Else
'
'        'lbl_CompFactor = 1 Modified on 06/26. The previous version would make that value null
'        lbl_CompFactor = ""
'
'  End If
'
'End If
'-----------------------------------

Call NormalBackground(SsoleSecUnit)

End Sub

Private Sub SsoleStockType_Click()
SsoleStockType.Tag = Trim(UCase(SsoleStockType.Columns(0).Text))
End Sub
Private Sub SsoleStockType_GotFocus()
Call HighlightBackground(SsoleStockType)
End Sub

Private Sub SsoleStockType_LostFocus()

Call NormalBackground(SsoleStockType)
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim RetCode As Integer

Select Case PreviousTab

    Case 0
    
        If mcheckManufac = False Then Exit Sub
    
        If GFormmode <> mdVisualization Then
        
            mcheckStockHeader = ValidateStockHeaderValues
            
            If mcheckStockHeader = True Then
            
                Call SaveToStockHeader
                
             Else
             
                SSTab1.Tab = 0
                
             End If
             
         ElseIf GFormmode = mdVisualization Then
         
            If Len(Trim(TxtStockNumber)) = 0 Then
            
            MsgBox "Please select a Stock Number before moving to any other tab.", vbInformation, "IMS"
            
            mcheckStockHeader = False
            
            SSTab1.Tab = 0
            
            End If
            
         End If
    
    Case 1
    
        If mcheckStockHeader = False Or mcheckManufac = False Then Exit Sub
        
        If GFormmode <> mdVisualization Then
        
        StockHeader.Techspec = Trim(txtTechSpec)
        
        End If
    
    Case 2

    
      If mcheckStockHeader = False Then Exit Sub
    
        If GFormmode <> mdVisualization And manufacturer.Count > 0 Then
        
            mcheckManufac = ValidatemanufacturerValues
           
           ' Call SaveToManufacturer
           
            If mcheckManufac = True Then
            
                Call SaveToManufacturer
                
             Else
             
                SSTab1.Tab = 2
                
             End If
             
         End If
         
    Case 3
         


         
End Select

Select Case SSTab1.Tab

 

    Case 0
    
        'Call LoadFromStockheader(StockHeader)
        
    
    Case 1
    
    
    Case 2

        
        
        Set manufacturer = InitializeManufacturer
        
        If UCase(Trim(manufacturer.StockNumber)) <> UCase(Trim(TxtStockNumber)) Or UCase(Trim(TxtManuStock)) <> UCase(Trim(TxtStockNumber)) Then
        
                RetCode = manufacturer.MoveToStocknumber(TxtStockNumber)
                
                'Call ClearManufacturer
                
                Select Case (RetCode)
                
                    Case 0
                
                        If manufacturer.Count > 0 Then
                
                            Call LoadFromManufacturer(manufacturer)
                            
                        ElseIf manufacturer.Count = 0 And GFormmode <> mdVisualization Then
                        
                            LROleDBNavBar1.AddNew
                            
                        ElseIf manufacturer.Count = 0 And GFormmode = mdVisualization Then
                            
                            Call ClearManufacturer
                            
                        End If
                    
                    Case 1
                    
                        MsgBox "Unidentified error occurred while trying to "
                    
                    Case 2
                    
                        'No Record Exists
                    
                 End Select
        
        End If
        
        Case 3
                
                Screen.MousePointer = vbHourglass
                opt_Email.value = True
                LROleDBNavBar1.Visible = False
                TxtRecpStockNumb = TxtStockNumber
                Screen.MousePointer = vbArrow
                
        
End Select


   Call ToggleNavbar
End Sub


Public Function loadManufacturerCombos() As Boolean
loadManufacturerCombos = False
On Error GoTo ErrHandler

    If GInitiliazeParams.ManufacturerCombosLoaded = False Then
    
       Call PopulateManufacturer
       
       GInitiliazeParams.ManufacturerCombosLoaded = True
       
    End If
    
loadManufacturerCombos = True

Exit Function

ErrHandler:
Err.Clear



End Function

Public Function PopulateManufacturer() As Boolean
Dim rsMANUFACTURER As ADODB.Recordset
PopulateManufacturer = False

On Error GoTo ErrHandler

Set lookups = InitializeLookup

If lookups.GetManufacturers(rsMANUFACTURER) = 1 Then

    MsgBox "Errors Occurrred while trying to Populate the Manufaturer Combo in the Dll. Please Try again.", vbCritical, "Imswin"

    Exit Function
    
End If

SSoleManufacturer.RemoveAll

Do While Not rsMANUFACTURER.EOF

    SSoleManufacturer.AddItem rsMANUFACTURER("man_code") & vbTab & rsMANUFACTURER("man_name")
    
    rsMANUFACTURER.MoveNext

Loop


PopulateManufacturer = True

Exit Function

ErrHandler:

MsgBox "Error Occurred while trying to Populate manufacturer in the EXE. Please try again.", vbCritical, "Imswin"

Err.Clear
End Function

Public Function SaveToStockHeader() As Boolean

SaveToStockHeader = False

On Error GoTo ErrHandler
            
        StockHeader.StockNumber = Trim(TxtStockNumber)
        
        StockHeader.CategoryCode = Trim(SsOleCategory.Tag)
        
        StockHeader.PrimUOfMeasure = Trim(SSOlePrimUnit.Tag)
        
        StockHeader.SecoUOfMeasure = Trim(SsoleSecUnit.Tag)
        
        StockHeader.Maximum = Trim(txt_Maximum)
        
        StockHeader.Minimum = Trim(txt_Minimum)
        
        StockHeader.Activeflag = CheckActive.value
        
        StockHeader.PoolOrSpecific = optPool.value
                
       If IsNumeric(lbl_CompFactor) Then StockHeader.ComputationFactor = lbl_CompFactor
        
        StockHeader.estmprice = Trim(txt_Estimate)
        
        StockHeader.characctcode = Trim(SSOleCharge.Tag)
        
        StockHeader.stdrcost = Trim(txt_Standard)
        
        StockHeader.hazmatclau = Trim(txt_ShortDescript)
        
        StockHeader.Description = Trim(txt_LongDescript)
        
        StockHeader.Techspec = Trim(txtTechSpec)
        
        StockHeader.stocktype = IIf(Len(Trim(SsoleStockType.Text)) = 0, "", Trim(SsoleStockType.Text))
        
        'Juan 2010-8-11
        StockHeader.ratio1 = qtyPrimary
        StockHeader.ratio2 = qtySecondary
        '----------------
        
        If GFormmode = mdCreation Then StockHeader.CreateUser = CurrentUser ' 07/02/02 Modified to make sure the creation user and modification user is stored.
        
        StockHeader.ModiUser = CurrentUser ' 07/02/02 Modified to make sure the creation user and modification user is stored.
        
        StockHeader.Eccnid = IIf(Len(SSoleEccnno.Tag) = 0, 0, SSoleEccnno.Tag)
        
        'StockHeader.Eccnno = Trim(SSoleEccnNo.Text)
        
        StockHeader.Eccnsource = IIf(Len(SSoleSource.Tag) = 0, 0, SSoleSource.Tag)
        
        'StockHeader.Eccnsourcename = SSoleSource.Text
        
        StockHeader.Eccnlicsreq = chkLicense.value
        
        SaveToStockHeader = True

Exit Function

ErrHandler:

MsgBox "Errors occurred while trying to Save the Stock Record. Please try again. Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear


End Function

Public Function SaveToManufacturer() As Boolean

SaveToManufacturer = False

On Error GoTo ErrHandler
            
         Set manufacturer = InitializeManufacturer
            
    If manufacturer.State = 1 And manufacturer.Count > 0 Then
            
         manufacturer.ManufactCode = SSoleManufacturer.Tag
        
         manufacturer.PartNumb = TxtPartnumb
        
         manufacturer.Estmpric = TxtEstPrice
        
         manufacturer.Techspec = txtManSpecs

         manufacturer.StockNumber = TxtManuStock
         
     End If
        
SaveToManufacturer = True

Exit Function

ErrHandler:

MsgBox "Errors occurred while trying to Save the manufacturer record. Please try again. Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function

Public Function ValidateStockHeaderValues() As Boolean

ValidateStockHeaderValues = False
On Error GoTo Handled
Dim i As Long

       
    If Len(Trim(TxtStockNumber)) = 0 Then
       MsgBox "StockNumber can not be left empty."
       TxtStockNumber.SetFocus
       Exit Function
    End If
    
''    If SSoleManufacturer.IsItemInList = False Then
''       MsgBox "Manufacturer is not valid."
''       SSoleManufacturer.SetFocus
''       Exit Function
''    End If
    
'    If Len(Trim$(SsOleCategory)) = 0 Then
'        MsgBox "Category can not be Left Empty."
'        TxtPartnumb.SetFocus
'        Exit Function
'    End If
    
    If Len(Trim$(SSOlePrimUnit)) = 0 Then
        MsgBox "Primary Unit can not be Left Empty."
        TxtPartnumb.SetFocus
        Exit Function
    End If
        
    If Len(Trim$(SsoleSecUnit)) = 0 Then
        MsgBox "Secondary Unit can not be Left Empty."
        TxtPartnumb.SetFocus
        Exit Function
    End If
    
'    If Len(Trim$(txt_Standard)) = 0 Then
'        MsgBox "Standard cost can not be Left Empty."
'        txt_Standard.SetFocus
'        Exit Function
'    End If
    
''        If Len(Trim$(txt_ShortDescript)) = 0 Then
''        MsgBox "Short Description can not be Left Empty."
''        txt_ShortDescript.SetFocus
''        Exit Function
''    End If
    
        If Len(Trim$(txt_LongDescript)) = 0 Then
        MsgBox "Long Description can not be Left Empty."
        txt_LongDescript.SetFocus
        Exit Function
    End If
    
''        If Len(Trim$(txtTechSpec)) = 0 Then
''        MsgBox "Technical Specifications can not be Left Empty."
''        txtTechSpec.SetFocus
''        Exit Function
''    End If
    
        
''    If Len(Trim$(TxtEstPrice.text)) = 0 Then
''        MsgBox "Estimated price can not be Left Empty."
''        TxtEstPrice.SetFocus
''        Exit Function
''    End If
''
''    If Len(Trim$(txtManSpecs.text)) = 0 Then
''         MsgBox "Technical Specs can not be Left Empty."
''         SSOleDBPriority.SetFocus: Exit Function
''    End If
    
    ValidateStockHeaderValues = True
    
    Exit Function
        
Handled:
    
MsgBox " Errors Occurred while trying to Validate the enteries made in the Stock record.", vbCritical, "Imswin"

Err.Clear

End Function

Public Function ValidatemanufacturerValues() As Boolean

ValidatemanufacturerValues = False
On Error GoTo Handled
Dim i As Long

    Set manufacturer = InitializeManufacturer

    If manufacturer.Count = 0 Then ValidatemanufacturerValues = True: Exit Function
       
    If Len(Trim(SSoleManufacturer)) = 0 Then
       MsgBox "Manufacturer can not be left empty."
       SSoleManufacturer.SetFocus
       Exit Function
    End If
    
    If SSoleManufacturer.IsItemInList = False Then
       MsgBox "Manufacturer is not valid."
       SSoleManufacturer.SetFocus
       Exit Function
    End If
    
    If Len(Trim$(TxtPartnumb.Text)) = 0 Then
        MsgBox "Partnumber can not be Left Empty."
        TxtPartnumb.SetFocus
        Exit Function
    End If
        
''    If Len(Trim$(TxtEstPrice.text)) = 0 Then
''        MsgBox "Estimated price can not be Left Empty."
''        TxtEstPrice.SetFocus
''        Exit Function
''    End If
''
''    If Len(Trim$(txtManSpecs.text)) = 0 Then
''         MsgBox "Technical Specs can not be Left Empty."
''         SSOleDBPriority.SetFocus: Exit Function
''    End If
    
    If IsNumeric(TxtEstPrice.Text) = False Then
        MsgBox "Please enter a valid value for the Estimated price."
        TxtEstPrice.SetFocus
        Exit Function
    End If
    
    ValidatemanufacturerValues = True
    
    Exit Function
        
Handled:
    
MsgBox " Errors Occurred while trying to Validate the enteries made in the manufacturer.", vbCritical, "Imswin"

Err.Clear

End Function

Public Function ToogleStockHeaderControls() As Boolean
On Error GoTo ErrHandler
ToogleStockHeaderControls = False

Select Case GFormmode

    Case mdCreation
    
        TxtStockNumber.Enabled = True
        FraStockHeader.Enabled = True
        txt_Standard.locked = False
        txt_ShortDescript.locked = False
        txt_LongDescript.locked = False
        txtTechSpec.locked = False
        TxtStockSearch.Enabled = False
        TxtStockSearch = ""
        lblManufacNo.Visible = False
        
        SSoleEccnno.Enabled = True
        SSoleSource.Enabled = True
        chkLicense.Enabled = True
                
                
    Case mdModification
    
        TxtStockNumber.Enabled = False
        FraStockHeader.Enabled = True
        txt_Standard.locked = False
        txt_ShortDescript.locked = False
        txt_LongDescript.locked = False
        txtTechSpec.locked = False
        TxtStockSearch.Enabled = False
        TxtStockSearch = ""
        lblManufacNo.Visible = False
        
        SSoleEccnno.Enabled = True
        SSoleSource.Enabled = True
        chkLicense.Enabled = True
        
        
    Case mdVisualization
    
        TxtStockNumber.Enabled = False
        FraStockHeader.Enabled = False
        txt_Standard.locked = True
        txt_ShortDescript.locked = True
        txt_LongDescript.locked = True
        txtTechSpec.locked = True
        TxtStockSearch.Enabled = True
        TxtStockSearch = ""
        lblManufacNo.Visible = True
        
        SSoleEccnno.Enabled = False
        SSoleSource.Enabled = False
        chkLicense.Enabled = False
        
End Select
LblManufno.Visible = lblManufacNo.Visible
ToogleStockHeaderControls = True

Exit Function
ErrHandler:

MsgBox " Errors occured while trying to Toggle the controls on stock header.Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function

Public Function ToogleManufacturerControls() As Boolean
On Error GoTo ErrHandler

ToogleManufacturerControls = False

Select Case GFormmode

    Case mdCreation

         SSoleManufacturer.Enabled = True
        
         TxtPartnumb.Enabled = True
        
         TxtEstPrice.Enabled = True
        
         txtManSpecs.locked = False

          Call ToggleNavbar

         'TxtManuStock.en
        
      Case mdModification
      
         
         Call ToogleManufacturerCombo
         
         'SSoleManufacturer.Enabled = True
        
         TxtPartnumb.Enabled = True
        
         TxtEstPrice.Enabled = True
        
         txtManSpecs.locked = False
         
         Call ToggleNavbar

      
      Case mdVisualization
      
          
         SSoleManufacturer.Enabled = False
        
         TxtPartnumb.Enabled = False
        
         TxtEstPrice.Enabled = False
        
         txtManSpecs.locked = True
         
         Call ToggleNavbar

 End Select

ToogleManufacturerControls = True
Exit Function
ErrHandler:

MsgBox " Errors occured while trying to Toggle the controls on manufacturer.Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function

Public Function ToggleNavbar() As Boolean

ToggleNavbar = False
On Error GoTo ErrHandler

Select Case GFormmode

    Case mdCreation
        
    If SSTab1.Tab = 0 Then
        LROleDBNavBar1.Visible = True
        LROleDBNavBar1.NewEnabled = False
        LROleDBNavBar1.SaveEnabled = True
        LROleDBNavBar1.CancelEnabled = True
        LROleDBNavBar1.NextVisible = False
        LROleDBNavBar1.PreviousVisible = False
        LROleDBNavBar1.FirstVisible = False
        LROleDBNavBar1.LastVisible = False
        LROleDBNavBar1.EditEnabled = False
        LROleDBNavBar1.EMailEnabled = False
    ElseIf SSTab1.Tab = 1 Then
    
        LROleDBNavBar1.Visible = False
        LROleDBNavBar1.EMailEnabled = False
    
    ElseIf SSTab1.Tab = 2 Then
        
        LROleDBNavBar1.Visible = True
        LROleDBNavBar1.EMailEnabled = False
        LROleDBNavBar1.NewEnabled = True
        LROleDBNavBar1.SaveEnabled = False
        LROleDBNavBar1.CancelEnabled = True
        LROleDBNavBar1.NextVisible = True
        LROleDBNavBar1.PreviousVisible = True
        LROleDBNavBar1.FirstVisible = True
        LROleDBNavBar1.LastVisible = True
        LROleDBNavBar1.EditEnabled = False
    End If
    
        
    Case mdModification
    
     If SSTab1.Tab = 0 Then
        
        LROleDBNavBar1.Visible = True
        
        LROleDBNavBar1.NewEnabled = False
        LROleDBNavBar1.SaveEnabled = True
        LROleDBNavBar1.CancelEnabled = True
        LROleDBNavBar1.NextVisible = False
        LROleDBNavBar1.PreviousVisible = False
        LROleDBNavBar1.FirstVisible = False
        LROleDBNavBar1.LastVisible = False
        LROleDBNavBar1.EditEnabled = False
        LROleDBNavBar1.EMailEnabled = False
    ElseIf SSTab1.Tab = 1 Then
    
        LROleDBNavBar1.Visible = False
        LROleDBNavBar1.EMailEnabled = False
    ElseIf SSTab1.Tab = 2 Then
    
        LROleDBNavBar1.Visible = True
    
        LROleDBNavBar1.NewEnabled = True
        LROleDBNavBar1.SaveEnabled = False
        LROleDBNavBar1.CancelEnabled = True
        LROleDBNavBar1.NextVisible = True
        LROleDBNavBar1.PreviousVisible = True
        LROleDBNavBar1.FirstVisible = True
        LROleDBNavBar1.LastVisible = True
        LROleDBNavBar1.EditEnabled = False
        LROleDBNavBar1.EMailEnabled = False
    End If
        
    Case mdVisualization

    
    If SSTab1.Tab = 0 Then
        
        LROleDBNavBar1.Visible = True
    
        LROleDBNavBar1.NextVisible = False
        LROleDBNavBar1.PreviousVisible = False
        LROleDBNavBar1.FirstVisible = False
        LROleDBNavBar1.LastVisible = False
        LROleDBNavBar1.EMailEnabled = True
        LROleDBNavBar1.EditEnabled = GInitiliazeParams.NavbarEditEnabled
        LROleDBNavBar1.NewEnabled = GInitiliazeParams.NavBarNewEnabled
        LROleDBNavBar1.SaveEnabled = GInitiliazeParams.NavbarSaveEnabled
            
    ElseIf SSTab1.Tab = 1 Then
    
        LROleDBNavBar1.Visible = False
        LROleDBNavBar1.EMailEnabled = False
    ElseIf SSTab1.Tab = 2 Then
    
        LROleDBNavBar1.Visible = True
    
        LROleDBNavBar1.NextVisible = True
        LROleDBNavBar1.PreviousVisible = True
        LROleDBNavBar1.FirstVisible = True
        LROleDBNavBar1.LastVisible = True
        LROleDBNavBar1.EditEnabled = False
        LROleDBNavBar1.NewEnabled = False
        LROleDBNavBar1.SaveEnabled = False
        LROleDBNavBar1.EMailEnabled = False
    End If

        'LROleDBNavBar1.NewEnabled = True
        'LROleDBNavBar1.SaveEnabled = False
        LROleDBNavBar1.CancelEnabled = False

        
End Select

ToggleNavbar = True

Exit Function
ErrHandler:


End Function
Public Function ChangeMode(FMode As FormMode) As FormMode
On Error Resume Next
Dim bl As Boolean
Dim msg1 As String
    
ChkManufActive.Enabled = True 'JCG 2007/01/10

    
    If FMode = mdCreation Then
        lblStatus.ForeColor = vbRed
        
        
        lblStatus.Caption = IIf(msg1 = "", "Creation", msg1)
        
        
    ElseIf FMode = mdModification Then
        lblStatus.ForeColor = vbBlue
                
        
       lblStatus.Caption = IIf(msg1 = "", "Modification", msg1)
        
  
     ElseIf FMode = mdVisualization Then
        lblStatus.ForeColor = vbGreen
        
        
        lblStatus.Caption = IIf(msg1 = "", "Visualization", msg1)
        
        ChkManufActive.Enabled = False 'JCG 2007/01/10
    
    End If
    
    

   ChangeMode = FMode
End Function

Public Function SetInitialValuesManufacturer() As Boolean

SetInitialValuesManufacturer = False

On Error GoTo ErrHandler

TxtLineNumber = manufacturer.Count

txtTotal = manufacturer.Count

 TxtManuStock = TxtStockNumber
 
Call ToogleManufacturerCombo

SetInitialValuesManufacturer = True

Exit Function

ErrHandler:

MsgBox "Errors Occurred while trying to Initialize Manufacturer record.Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function

Public Function SetInitialValuesStockmaster() As Boolean

SetInitialValuesStockmaster = False

On Error GoTo ErrHandler

SetInitialValuesStockmaster = True

Exit Function

ErrHandler:

MsgBox "Errors Occurred while trying to Initialize Stock header record. Error Description " & Err.Description, vbCritical, "Imswin"

Err.Clear

End Function

Private Sub txt_Estimate_Validate(Cancel As Boolean)
If Not IsNumeric(txt_Estimate) And Len(Trim(txt_Estimate)) > 0 Then
    MsgBox "Please enter a valid Estimate amount.", vbInformation, "Imswin"
    Cancel = True
    
 ElseIf IsNumeric(txt_Estimate) And Len(Trim(txt_Estimate)) > 0 Then
    
    txt_Estimate = Format(txt_Estimate, "00.00")
    
 End If
    
 
End Sub

Private Sub txt_LongDescript_KeyUp(KeyCode As Integer, Shift As Integer)
If GFormmode <> mdVisualization And Trim(SSDBHeader.Columns(0).Text) = Trim(TxtStockNumber) Then
    SSDBHeader.Columns(1).Text = Trim(txt_LongDescript) ' & Chr(KeyAscii)
End If
End Sub

Private Sub txt_LongDescript_LostFocus()
If TxtStockNumber.Enabled = True Then

    TxtStockNumber.SetFocus
    
 Else
 
    If SsOleCategory.Enabled Then SsOleCategory.SetFocus
 
 End If
 Call NormalBackground(txt_LongDescript)
End Sub
Private Sub txt_LongDescript_GotFocus()
Call HighlightBackground(txt_LongDescript)
End Sub

Private Sub txt_Maximum_GotFocus()
Call HighlightBackground(txt_Maximum)
End Sub

Private Sub txt_Maximum_LostFocus()

Call NormalBackground(txt_Maximum)
End Sub

Private Sub txt_Maximum_Validate(Cancel As Boolean)
If Not IsNumeric(txt_Maximum) And Len(Trim(txt_Maximum)) > 0 Then
    MsgBox "Please enter a valid Maximum amount.", vbInformation, "Imswin"
    Cancel = True
    
  ElseIf IsNumeric(txt_Maximum) And Len(Trim(txt_Maximum)) > 0 Then
    
    txt_Maximum = Format(txt_Maximum, "00.00")
    
 End If
    

End Sub

Private Sub txt_Minimum_GotFocus()
Call HighlightBackground(txt_Minimum)
End Sub

Private Sub txt_Minimum_LostFocus()

Call NormalBackground(txt_Minimum)
End Sub

Private Sub txt_Minimum_Validate(Cancel As Boolean)
If Not IsNumeric(txt_Minimum) And Len(Trim(txt_Minimum)) > 0 Then
    
    MsgBox "Please enter a valid Minimum amount.", vbInformation, "Imswin"
    Cancel = True
    
 ElseIf IsNumeric(txt_Minimum) And Len(Trim(txt_Minimum)) > 0 Then
    
    txt_Minimum = Format(txt_Minimum, "00.00")
    
 End If

End Sub

Private Sub txt_ShortDescript_GotFocus()
Call HighlightBackground(txt_ShortDescript)
End Sub

Private Sub txt_ShortDescript_LostFocus()
Call NormalBackground(txt_ShortDescript)
End Sub
Private Sub txt_Standard_Validate(Cancel As Boolean)
If Not IsNumeric(txt_Standard) And Len(Trim(txt_Standard)) > 0 Then
    MsgBox "Please enter a valid Standard Cost.", vbInformation, "Imswin"
    Cancel = True
    
   ElseIf IsNumeric(txt_Standard) And Len(Trim(txt_Standard)) > 0 Then
    
    txt_Standard = Format(txt_Standard, "00.00")
    
 End If
    
    

End Sub

Private Sub txt_Standard_GotFocus()
Call HighlightBackground(txt_Standard)
End Sub

Private Sub txt_Standard_LostFocus()
Call NormalBackground(txt_Standard)
End Sub
Private Sub TxtEstPrice_GotFocus()
Call HighlightBackground(TxtEstPrice)
End Sub

Private Sub TxtEstPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtEstPrice_Validate (True)
        txtManSpecs.SetFocus
    End If
End Sub

Private Sub TxtEstPrice_LostFocus()
Call NormalBackground(TxtEstPrice)
End Sub

Private Sub TxtEstPrice_Validate(Cancel As Boolean)

If Not IsNumeric(TxtEstPrice) And Len(Trim(TxtEstPrice)) > 0 Then
    MsgBox "Please enter a valid Estimated Cost.", vbInformation, "Imswin"
    Cancel = True
    
   ElseIf IsNumeric(TxtEstPrice) And Len(Trim(TxtEstPrice)) > 0 Then
    
    TxtEstPrice = Format(TxtEstPrice, "00.00")
    
 End If

End Sub

Private Sub txtManSpecs_GotFocus()
Call HighlightBackground(txtManSpecs)
End Sub

Private Sub txtManSpecs_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        TxtPartnumb.SetFocus
'    End If
End Sub

Private Sub txtManSpecs_LostFocus()
Call NormalBackground(txtManSpecs)
If SSoleManufacturer.Enabled = True Then
        SSoleManufacturer.SetFocus

Else

    TxtPartnumb.SetFocus

End If

    
End Sub

Private Sub TxtPartnumb_GotFocus()
Call HighlightBackground(TxtPartnumb)
End Sub

Private Sub TxtPartnumb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtEstPrice.SetFocus
    End If
End Sub

Private Sub TxtPartnumb_LostFocus()
Call NormalBackground(TxtPartnumb)
End Sub

Private Sub TxtStockNumber_GotFocus()
'Call HighlightBackground(TxtStockNumber)
End Sub

Private Sub TxtStockNumber_KeyUp(KeyCode As Integer, Shift As Integer)
If GFormmode <> mdVisualization Then 'And Trim(SSDBHeader.Columns(0).text) = Trim(TxtStockNumber) Then
    
    SSDBHeader.Columns(0).Text = Trim(TxtStockNumber) ' & Chr(KeyAscii)
    
End If
End Sub

Private Sub TxtStockNumber_LostFocus()
'Call NormalBackground(TxtStockNumber)
End Sub

Private Sub TxtStockNumber_Validate(Cancel As Boolean)

If GFormmode = mdCreation Then

    Set lookups = InitializeLookup
    
    If lookups.DoesStockExist(TxtStockNumber) > 0 Then
    
        MsgBox "Stock with that name already exist. Please use a different one.", vbInformation, "Imswin"
        
        Cancel = True
        
    End If
    
End If
    
End Sub

Private Sub TxtStockSearch_Change()

Dim Count As Integer

''If GFormmode = mdVisualization Then
''
''If Len(Trim(TxtStockSearch)) = 0 Then Exit Sub
''
''RsStockNameDesc.MoveFirst
''
''RsStockNameDesc.Find "Stk_stcknumb like '" & Trim(TxtStockSearch) & "%'"
''
''
''
''Else
''
''
''
''End If

Call MoveGridTo(TxtStockSearch)

End Sub


Public Function GenerateStyleSheets()


SSDBHeader.StyleSets.Add ("CellBeingModified")
SSDBHeader.StyleSets("CellBeingModified").BackColor = vbYellow

SSDBHeader.StyleSets.Add ("RowBeingModified")
SSDBHeader.StyleSets("RowBeingModified").BackColor = &H80C0FF

SSDBHeader.StyleSets.Add ("RowAdded")
SSDBHeader.StyleSets("RowAdded").BackColor = vbGreen

SSDBHeader.StyleSets.Add ("RowModified")
SSDBHeader.StyleSets("RowModified").BackColor = &HFFFFC0

SSDBHeader.ActiveRowStyleSet = "CellBeingModified"
SSDBHeader.activeCELL.StyleSet = "RowBeingModified"

'SSDBHeader.StyleSets.Add ("RowBeingModified")
'SSDBHeader.StyleSets("RowBeingModified").BackColor = &H80C0FF

SSDBHeader.ActiveRowStyleSet = "CellBeingModified"
SSDBHeader.activeCELL.StyleSet = "RowBeingModified"


End Function

Public Function StoreStocksPlayedwith(EditMode As Integer)
        If EditMode = 2 Then
        
           If IsArrayLoaded(GInitiliazeParams.StockAdded) = False Then
            
                ReDim Preserve GInitiliazeParams.StockAdded(0)
                
           Else
           
                 ReDim Preserve GInitiliazeParams.StockAdded(UBound(GInitiliazeParams.StockAdded) + 1)
                 
            End If
            
            GInitiliazeParams.StockAdded(UBound(GInitiliazeParams.StockAdded)) = StockHeader.StockNumber
            
            
        SSDBHeader.Columns(0).CellStyleSet "RowAdded"
        SSDBHeader.Columns(1).CellStyleSet "RowAdded"
            
        End If
        
        
        If EditMode = 1 Then
        
           If IsArrayLoaded(GInitiliazeParams.StocksModified) = False Then
            
                ReDim Preserve GInitiliazeParams.StocksModified(0)
                
           Else
           
                 ReDim Preserve GInitiliazeParams.StocksModified(UBound(GInitiliazeParams.StocksModified) + 1)
                 
            End If
            
            GInitiliazeParams.StocksModified(UBound(GInitiliazeParams.StocksModified)) = Trim(StockHeader.StockNumber)
            
            
        SSDBHeader.Columns(0).CellStyleSet "RowModified"
        SSDBHeader.Columns(1).CellStyleSet "RowModified"
            
        End If
End Function


Public Function AddRecepient(RecepientAddress As String)

Dim Count As Integer

'If dgRecipientList.Rows = 0 Then dgRecipientList.AddItem RecepientAddress: Exit Sub

'If dgRecipientList.Rows = 7 Then

 '  MsgBox "Can not Add more than 7 recepients.", vbInformation + vbOKOnly, "Imswin"

'ElseIf dgRecipientList.Rows < 7 Then

    'count = 1

    dgRecipientList.MoveFirst

    Do While Not dgRecipientList.Rows = Count

        If dgRecipientList.Columns(0).value = RecepientAddress Then

            MsgBox "Recepient already exists, Please choose a different one.", vbInformation + vbOKOnly, "Imswin"

            Exit Function

        End If

        dgRecipientList.MoveNext

        Count = Count + 1

    Loop

    dgRecipientList.AddItem RecepientAddress

'End If

End Function



Private Sub Text1_KeyPress(KeyAscii As Integer)


End Sub

Private Sub SSTabRequisitions_Click(PreviousTab As Integer)
opt_Email.value = True
End Sub

Private Sub Txt_search_Change()

Dim Grid As SSOleDBGrid

Dim x As Integer

Dim Count As Integer

Dim i As Integer

If SSOLEDBEmail.Visible = True Then Set Grid = SSOLEDBEmail

If SSOLEDBFax.Visible = True Then Set Grid = SSOLEDBFax

i = Len(Txt_search)

Count = 1

    Grid.MoveFirst

    Do While Not Grid.Rows = Count

        If UCase(Txt_search) = UCase(Mid(Grid.Columns(0).value, 1, i)) Then

           Grid.Scroll 0, Grid.row

           Exit Sub

        End If

        Grid.MoveNext

        Count = Count + 1

    Loop

End Sub

Private Sub SSOLEDBEmail_DblClick()
On Error Resume Next

 AddRecepient SSOLEDBEmail.Columns(1).value


End Sub

Private Sub opt_FaxNum_Click()

SSOLEDBFax.Visible = True
SSOLEDBEmail.Visible = False

If GGridFilledWithFax = False Then

    Call GetSupplierPhoneDirFAX

    GGridFilledWithFax = True

 End If


End Sub

Public Sub GetSupplierPhoneDirFAX()
Dim str As String
Dim cmd As Command
Dim rst As New Recordset
Dim Sql As String

Sql = "select sup_name Names,  upper( sup_contaFax) Fax  from supplier where sup_npecode='" & deIms.NameSpace & "' and sup_contaFax is not null and len(sup_contaFax) > 0   union"

Sql = Sql & " select phd_name Names, upper(phd_faxnumb) Fax from phonedir  where phd_npecode='" & deIms.NameSpace & "'and phd_faxnumb is not null and len(phd_faxnumb)>0 order by names"

rst.Source = Sql

rst.ActiveConnection = deIms.cnIms

rst.Open

    If rst.RecordCount = 0 Then GoTo clearup

    rst.MoveFirst

    Do While Not rst.EOF

        SSOLEDBFax.AddItem rst("Names") & Chr(9) & rst("FAX")

        rst.MoveNext

    Loop

    GGridFilledWithFax = True

clearup:

    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub


Public Sub GetSupplierPhoneDirEmails()

Dim str As String
Dim cmd As Command
Dim rst As New Recordset
Dim Sql As String

Sql = " select sup_name Names,  upper( sup_mail) Emails  from supplier where sup_npecode='" & deIms.NameSpace & "' and sup_mail is not null and len(sup_mail) > 0   union "

Sql = Sql & " select phd_name Names, upper(phd_mail) Emails from phonedir  where phd_npecode='" & deIms.NameSpace & "'and phd_mail is not null and len(phd_mail)>0 order by names "

rst.Source = Sql

rst.ActiveConnection = deIms.cnIms

rst.Open

    If rst.RecordCount = 0 Then GoTo clearup

    rst.MoveFirst

    Do While Not rst.EOF

        SSOLEDBEmail.AddItem rst("Names") & Chr(9) & rst("Emails")

        rst.MoveNext

    Loop

    GGridFilledWithEmails = True

clearup:

    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub opt_Email_Click()

SSOLEDBEmail.Visible = True
'SSOLEDBFax.Visible = False

If GGridFilledWithEmails = False Then

   Call GetSupplierPhoneDirEmails

   GGridFilledWithEmails = True

End If

End Sub


''''Private Sub NavBar1_OnEMailClick()
''''Dim IFile As IMSFile
''''Dim FileName(1) As String
''''Dim Recepients() As String
''''Dim rsr As ADODB.Recordset
''''Dim rptinfo As RPTIFileInfo
''''Dim Subject As String
''''Dim attention As String
''''Dim ParamsForCrystalReports() As String
''''Dim ParamsForRPTI() As String
''''Dim FieldName As String
''''Dim Message As String
''''
''''ReDim ParamsForCrystalReports(2)
''''ReDim ParamsForRPTI(2)
''''
''''    Set rsr = GetObsRecipients(deIms.NameSpace, ssOleDbPO, SScmbMessage.text)
''''
''''    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
''''
''''    ParamsForCrystalReports(1) = "mesgnumb;" + SScmbMessage + ";TRUE"
''''
''''    ParamsForCrystalReports(2) = "ponumb;" + ssOleDbPO + ";true"
''''
''''    ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
''''
''''    ParamsForRPTI(1) = "mesgnumb=" + SScmbMessage
''''
''''    ParamsForRPTI(2) = "ponumb=" + ssOleDbPO
''''
''''    FieldName = "Recipient"
''''
''''    Subject = "Tracking Message for PO -" & ssOleDbPO
''''
''''    If ConnInfo.EmailClient = Outlook Then
''''
''''        Call sendOutlookEmailandFax("obs.rpt", "Tracking Message", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, Subject, attention)
''''
''''    ElseIf ConnInfo.EmailClient = ATT Then
''''
''''        Call SendAttFaxAndEmail("obs.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, Subject, Message, FieldName)
''''
''''    ElseIf ConnInfo.EmailClient = Unknown Then
''''
''''        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"
''''
''''    End If
''''
''''
''''    If chkYesorNo.Value = 1 Then
''''
''''        ReDim ParamsForCrystalReports(1)
''''
''''        ReDim ParamsForRPTI(1)
''''
''''
''''            ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
''''
''''            ParamsForCrystalReports(1) = "ponumb;" + ssOleDbPO + ";true"
''''
''''            ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
''''
''''            ParamsForRPTI(1) = "ponumb=" + ssOleDbPO
''''
''''            FieldName = "Recipient"
''''
''''            If ConnInfo.EmailClient = Outlook Then
''''
''''                Call sendOutlookEmailandFax("PO.rpt", "Tracking Message", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, Subject, attention)
''''
''''            ElseIf ConnInfo.EmailClient = ATT Then
''''
''''                Call SendAttFaxAndEmail("PO.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsr, Subject, Message, FieldName)
''''
''''            ElseIf ConnInfo.EmailClient = Unknown Then
''''
''''                MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"
''''
''''            End If
''''
''''    End If
''''
''''End Sub

Private Sub cmd_Add_Click()
On Error Resume Next
If (OptEmail.value = True Or OptFax.value = True) Then

        If Len(Trim$(txt_Recipient)) > 0 Then
               txt_Recipient = UCase(txt_Recipient)

               If OptEmail.value = True Then txt_Recipient = (txt_Recipient)
               If OptFax.value = True Then txt_Recipient = (txt_Recipient)

              'dgRecipientList.AddItem txt_Recipient

              AddRecepient txt_Recipient

              txt_Recipient = ""

        End If
 Else
    MsgBox "Please check Email or Fax.", vbInformation, "Imswin"

 End If
End Sub

Private Sub cmdRemove_Click()

Dim x As Integer

If Len(dgRecipientList.SelBookmarks(0)) = 0 Then

    MsgBox "Please make a selection first.", vbInformation, "Imswin"

    Exit Sub

 End If



    dgRecipientList.DeleteSelected

' dgRecipientList.SelBookmarks.RemoveAll

End Sub
Private Sub SSOLEDBFax_DblClick()
On Error Resume Next

    'dgRecipientList.AddItem SSOLEDBFax.Columns(1).Value

    AddRecepient SSOLEDBFax.Columns(1).value

    If Err Then Err.Clear
End Sub


Public Function MoveGridTo(StockNumber As String)

Dim Count As Integer

If GFormmode = mdVisualization Then

If Len(Trim(StockNumber)) = 0 Then Exit Function

RsStockNameDesc.MoveFirst

RsStockNameDesc.Find "Stk_stcknumb like '" & Trim(StockNumber) & "%'"

Else



End If

End Function

Private Sub TxtStockSearch_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

    If RsStockNameDesc.AbsolutePosition <> adPosBOF And RsStockNameDesc.AbsolutePosition <> adPosEOF And RsStockNameDesc.AbsolutePosition <> adPosUnknown Then
    
        Call LoadStockMaster
    
    End If

End If

End Sub

Public Function LoadStockMaster()

Screen.MousePointer = vbHourglass

Set StockHeader = InitializeStockHeader

Call StockHeader.MoveToStocknumber(SSDBHeader.Columns(0).Text)

Call ClearStockMasterDetails

Call LoadFromStockheader(StockHeader)

Screen.MousePointer = vbArrow

End Function

Private Sub txtTechSpec_GotFocus()
Call HighlightBackground(txtTechSpec)
End Sub

Private Sub txtTechSpec_LostFocus()
Call NormalBackground(txtTechSpec)
End Sub

Private Sub txt_Estimate_GotFocus()
Call HighlightBackground(txt_Estimate)
End Sub

Private Sub txt_Estimate_LostFocus()

Call NormalBackground(txt_Estimate)
End Sub

Public Function ToogleManufacturerCombo()

Set manufacturer = InitializeManufacturer
        If manufacturer.Count > 0 Then
                
                 If manufacturer.EditMode = 0 Or manufacturer.EditMode = 1 Then
                
                    SSoleManufacturer.Enabled = False
                    
                 ElseIf manufacturer.EditMode = 2 Then
                 
                    SSoleManufacturer.Enabled = True
                    
                 End If
                 
         End If
         
End Function
