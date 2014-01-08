VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frm_Stock2 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Master"
   ClientHeight    =   6195
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8625
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   Tag             =   "02010100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   2760
      TabIndex        =   57
      Top             =   5760
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailVisible    =   -1  'True
      NewEnabled      =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DisableSaveOnSave=   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   609
      TabCaption(0)   =   "Stock"
      TabPicture(0)   =   "frm_Stock2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "frm_Stock2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1(1)"
      Tab(1).Control(1)=   "cmd_Remove"
      Tab(1).Control(2)=   "cmd_Add"
      Tab(1).Control(3)=   "ssdbRecepientList"
      Tab(1).Control(4)=   "lbl_Recipients(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Image"
      TabPicture(2)   =   "frm_Stock2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picHolder"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tech. Specs."
      TabPicture(3)   =   "frm_Stock2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtTechSpec"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Manufacturer"
      TabPicture(4)   =   "frm_Stock2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1"
      Tab(4).Control(1)=   "Label2"
      Tab(4).Control(2)=   "Label3"
      Tab(4).Control(3)=   "Label4"
      Tab(4).Control(4)=   "dcboManuFac"
      Tab(4).Control(5)=   "SSTab2"
      Tab(4).Control(6)=   "Text1"
      Tab(4).Control(7)=   "Text2"
      Tab(4).Control(8)=   "Check2"
      Tab(4).ControlCount=   9
      Begin VB.CheckBox Check2 
         DataField       =   "stm_flag"
         DataMember      =   "GetStockManufacturer"
         Height          =   195
         Left            =   -69240
         TabIndex        =   23
         Top             =   960
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   -73200
         MaxLength       =   9
         TabIndex        =   18
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         DataField       =   "stm_partnumb"
         DataMember      =   "GetStockManufacturer"
         Height          =   315
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   17
         Top             =   570
         Width           =   2055
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3855
         Left            =   -74640
         TabIndex        =   20
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Parts Specification"
         TabPicture(0)   =   "frm_Stock2.frx":008C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtManSpecs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Images"
         TabPicture(1)   =   "frm_Stock2.frx":00A8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "imgManImg"
         Tab(1).ControlCount=   1
         Begin VB.TextBox txtManSpecs 
            DataField       =   "stm_techspec"
            DataMember      =   "GetStockManufacturer"
            Height          =   3375
            Left            =   120
            MaxLength       =   3500
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   360
            Width           =   7215
         End
         Begin VB.Image imgManImg 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "stm_imge"
            DataMember      =   "GetStockManufacturer"
            Height          =   3255
            Left            =   -74880
            Stretch         =   -1  'True
            Top             =   480
            Width           =   7215
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3015
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   3015
         ScaleWidth      =   8055
         TabIndex        =   50
         Top             =   2400
         Width           =   8055
         Begin ImsMailVB.Imsmail Imsmail1 
            Height          =   3135
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
         End
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74595
         TabIndex        =   22
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74595
         TabIndex        =   21
         Top             =   1575
         Width           =   1215
      End
      Begin VB.TextBox txtTechSpec 
         DataField       =   "stk_techspec"
         DataMember      =   "STOCKMASTER"
         DataSource      =   "deIms"
         Height          =   4970
         Left            =   -74700
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   420
         Width           =   7600
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   4935
         Index           =   0
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   8055
         TabIndex        =   29
         Top             =   360
         Width           =   8055
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo dcboSecUnit 
            Bindings        =   "frm_Stock2.frx":00C4
            DataField       =   "stk_secouom"
            DataMember      =   "STOCKMASTER"
            DataSource      =   "deIms"
            Height          =   315
            Left            =   1680
            TabIndex        =   4
            Top             =   1110
            Width           =   2235
            DataFieldList   =   "uni_code"
            _Version        =   196617
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
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "uni_desc"
            Columns(0).FieldLen=   256
            Columns(1).Width=   1138
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "uni_code"
            Columns(1).FieldLen=   256
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "uni_desc"
         End
         Begin VB.TextBox txt_ShortDescript 
            DataField       =   "stk_hazmatclau"
            DataMember      =   "STOCKMASTER"
            Height          =   675
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   2340
            Width           =   6300
         End
         Begin VB.TextBox txt_LongDescript 
            DataField       =   "stk_desc"
            DataMember      =   "STOCKMASTER"
            Height          =   1755
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   3040
            Width           =   6300
         End
         Begin VB.CheckBox chkDescHist 
            Alignment       =   1  'Right Justify
            DataField       =   "stk_descflag"
            DataMember      =   "STOCKMASTER"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1200
            TabIndex        =   47
            Top             =   3600
            Width           =   180
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Active"
            DataField       =   "stk_flag"
            DataMember      =   "STOCKMASTER"
            Height          =   195
            Left            =   45
            TabIndex        =   12
            Top             =   2100
            Value           =   1  'Checked
            Width           =   1815
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
            Left            =   5820
            TabIndex        =   9
            Top             =   780
            Width           =   2160
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
            Left            =   5820
            TabIndex        =   11
            Top             =   1770
            Width           =   2160
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
            Left            =   5820
            TabIndex        =   10
            Top             =   1440
            Width           =   2160
         End
         Begin VB.TextBox txt_Minimum 
            DataField       =   "stk_mini"
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Top             =   1770
            Width           =   2235
         End
         Begin VB.OptionButton optPool 
            Alignment       =   1  'Right Justify
            Caption         =   "Pool"
            Height          =   255
            Left            =   5820
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optSpecific 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific"
            Height          =   255
            Left            =   6915
            TabIndex        =   8
            Top             =   480
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   5820
            TabIndex        =   30
            Top             =   120
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            BackColor       =   16777152
            ForeColor       =   16711680
            Text            =   ""
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSdcboCategory 
            Bindings        =   "frm_Stock2.frx":00DA
            DataField       =   "stk_catecode"
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   450
            Width           =   2235
            DataFieldList   =   "cate_catecode"
            _Version        =   196617
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
            stylesets(0).Picture=   "frm_Stock2.frx":00F0
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
            stylesets(1).Picture=   "frm_Stock2.frx":010C
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3942
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "cate_name"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "cate_catename"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1270
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "cate_catecode"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "cate_catecode"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483639
            DataFieldToDisplay=   "cate_catename"
         End
         Begin MSDataListLib.DataCombo dcboChargeAccount 
            Bindings        =   "frm_Stock2.frx":0128
            DataField       =   "stk_characctcode"
            DataMember      =   "STOCKMASTER"
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Top             =   1440
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "cha_acctname"
            BoundColumn     =   "cha_acctcode"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbo_StockNum 
            Bindings        =   "frm_Stock2.frx":015B
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            Top             =   120
            Width           =   2235
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
            ColumnHeaders   =   0   'False
            RowHeight       =   423
            Columns(0).Width=   3200
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo dcboPrimUnit 
            Bindings        =   "frm_Stock2.frx":0166
            DataField       =   "stk_primuon"
            DataMember      =   "STOCKMASTER"
            DataSource      =   "deIms"
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   780
            Width           =   2235
            DataFieldList   =   "uni_code"
            _Version        =   196617
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
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Description"
            Columns(0).DataField=   "uni_desc"
            Columns(0).FieldLen=   256
            Columns(1).Width=   1191
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "uni_code"
            Columns(1).FieldLen=   256
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "uni_desc"
         End
         Begin VB.Label lbl_ShortDescript 
            BackStyle       =   0  'Transparent
            Caption         =   "Haz. Mat."
            Height          =   225
            Left            =   45
            TabIndex        =   49
            Top             =   2385
            Width           =   1695
         End
         Begin VB.Label lbl_Long 
            BackStyle       =   0  'Transparent
            Caption         =   "Long Description"
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   48
            Top             =   3040
            Width           =   1700
         End
         Begin VB.Label lbl_CompFactor 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "stk_compfctr"
            DataMember      =   "STOCKMASTER"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5820
            TabIndex        =   43
            Top             =   1110
            Width           =   2160
         End
         Begin VB.Label lbl_Category 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   195
            Left            =   45
            TabIndex        =   42
            Top             =   470
            Width           =   1695
         End
         Begin VB.Label lbl_PrimaryUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Unit"
            Height          =   195
            Left            =   45
            TabIndex        =   41
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label lbl_SecondaryUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Unit"
            Height          =   195
            Left            =   45
            TabIndex        =   40
            Top             =   1100
            Width           =   1695
         End
         Begin VB.Label lbl_StockNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Number"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   39
            Top             =   135
            Width           =   1700
         End
         Begin VB.Label lbl_Charge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charge Account"
            Height          =   195
            Left            =   45
            TabIndex        =   38
            Top             =   1470
            Width           =   1700
         End
         Begin VB.Label lbl_Minimum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum"
            Height          =   195
            Left            =   45
            TabIndex        =   37
            Top             =   1785
            Width           =   1700
         End
         Begin VB.Label lbl_Maximum 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum"
            Height          =   225
            Left            =   4155
            TabIndex        =   36
            Top             =   1785
            Width           =   1695
         End
         Begin VB.Label lbl_Computed 
            BackStyle       =   0  'Transparent
            Caption         =   "Computed Factor"
            Height          =   225
            Left            =   4155
            TabIndex        =   35
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label lbl_Standard 
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Cost"
            Height          =   225
            Left            =   4155
            TabIndex        =   34
            Top             =   1455
            Width           =   1695
         End
         Begin VB.Label lbl_Estimate 
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Price"
            Height          =   225
            Left            =   4155
            TabIndex        =   33
            Top             =   830
            Width           =   1695
         End
         Begin VB.Label lbl_PoolSpecific 
            BackStyle       =   0  'Transparent
            Caption         =   "Pool/Specific"
            Height          =   225
            Left            =   4155
            TabIndex        =   32
            Top             =   515
            Width           =   1695
         End
         Begin VB.Label lbl_StockNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Type"
            Height          =   195
            Index           =   1
            Left            =   4155
            TabIndex        =   31
            Top             =   135
            Width           =   1695
         End
         Begin VB.Label lbl_Long 
            BackStyle       =   0  'Transparent
            Caption         =   "Description History"
            Height          =   555
            Index           =   1
            Left            =   45
            TabIndex        =   52
            Top             =   3480
            Width           =   1140
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picHolder 
         Height          =   4965
         Left            =   -74700
         ScaleHeight     =   4905
         ScaleWidth      =   7545
         TabIndex        =   26
         Top             =   420
         Width           =   7605
         Begin MSComCtl2.FlatScrollBar flsbVert 
            Height          =   4700
            Left            =   7330
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   8281
            _Version        =   393216
            Orientation     =   1245184
            SmallChange     =   10
         End
         Begin MSComCtl2.FlatScrollBar flsbHoriz 
            Height          =   210
            Left            =   0
            TabIndex        =   27
            Top             =   4700
            Visible         =   0   'False
            Width           =   7540
            _ExtentX        =   13309
            _ExtentY        =   370
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1245185
         End
         Begin VB.Image imgImage 
            DataField       =   "stk_imge"
            DataMember      =   "STOCKMASTER"
            DataSource      =   "deIms"
            Height          =   4905
            Left            =   0
            Top             =   0
            Width           =   7545
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   1605
         Left            =   -73080
         TabIndex        =   51
         Top             =   600
         Width           =   5850
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
         _ExtentX        =   10319
         _ExtentY        =   2831
         _StockProps     =   79
         Caption         =   "Recipient List"
         BackColor       =   -2147483633
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
      Begin MSDataListLib.DataCombo dcboManuFac 
         Bindings        =   "frm_Stock2.frx":017C
         DataField       =   "stm_manucode"
         DataMember      =   "GetStockManufacturer"
         Height          =   315
         Left            =   -73200
         TabIndex        =   16
         Top             =   570
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "man_name"
         BoundColumn     =   "man_code"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "Active"
         Height          =   255
         Left            =   -70320
         TabIndex        =   56
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Estimated Price"
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Part Number"
         Height          =   255
         Left            =   -70320
         TabIndex        =   54
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ManuFacturer"
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label lbl_Recipients 
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient(s)"
         Height          =   300
         Index           =   0
         Left            =   -74640
         TabIndex        =   24
         Top             =   720
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdSave 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2160
      Picture         =   "frm_Stock2.frx":019B
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5730
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdOpen 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   1800
      Picture         =   "frm_Stock2.frx":029D
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5730
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox VisM1 
      Height          =   480
      Left            =   48
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   7776
      Width           =   1200
   End
End
Attribute VB_Name = "frm_Stock2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReceptList As ADODB.Recordset
Dim WithEvents stock As ADODB.Recordset
Attribute stock.VB_VarHelpID = -1
Dim mFromDesc As Boolean
Dim mIsItInsert As Boolean
Dim mFromClick As Boolean
Dim PriUnit As String
Dim SecUnit As String
Dim rights As UserRights
Dim rowguid, locked As Boolean, dbtablename As String       'jawdat



Private Sub cbo_StockNum_Change()
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
'Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)


'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock
ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02

If locked = True Then 'sets locked = true because another user has this record open in edit mode

optSpecific.Enabled = False
optPool.Enabled = False
SSdcboCategory.Enabled = False
dcboPrimUnit.Enabled = False
dcboSecUnit.Enabled = False
dcboChargeAccount.Enabled = False

NavBar1.SaveEnabled = False
  
  
      Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
       Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = False
        End If

    Next checkboxes
  
'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

optSpecific.Enabled = True
optPool.Enabled = True
SSdcboCategory.Enabled = True
dcboPrimUnit.Enabled = True
dcboSecUnit.Enabled = True
dcboChargeAccount.Enabled = True

NavBar1.SaveEnabled = True
  
  
    '  Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = True
        End If

    Next textboxes
  '    Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = True
        End If

    Next checkboxes



End If
'
End Sub

'call function to find stock numbers

'''Private Sub cbo_StockNum_Click(Area As Integer)
'''On Error Resume Next
'''Dim BK As Variant
'''
'''    If Area = 2 Then
'''        NavBar1.CancelUpdate
'''        BK = NavBar1.Recordset.Bookmark
'''        Call NavBar1.Recordset.Find("stk_stcknumb = '" & cbo_StockNum & "'", 0, adSearchForward, adBookmarkFirst)
'''        If NavBar1.Recordset.EOF Then NavBar1.Recordset.Bookmark = BK
'''    End If
'''
'''    If Err Then Err.Clear
'''End Sub

Private Sub cbo_StockNum_Click()


            


On Error Resume Next
Dim BK As Variant

        Call SetDatasourceForUnits
        
        mFromClick = True
        NavBar1.CancelUpdate
        mFromClick = False
        BK = NavBar1.Recordset.Bookmark
        
        NavBar1.Recordset.MoveFirst
        Call NavBar1.Recordset.Find("stk_stcknumb = '" & cbo_StockNum & "'", 0, adSearchForward)
        If NavBar1.Recordset.EOF Then NavBar1.Recordset.Bookmark = BK
        'Added by Muzammil 03/17/01
        'Reason - This sets the datasource of the  combo boxes,It sets it to a recordset with inactive Units
        'while just browsing thrught the records
    
    If Err Then Err.Clear
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
'Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)


'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock
ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02

If locked = True Then 'sets locked = true because another user has this record open in edit mode

optSpecific.Enabled = False
optPool.Enabled = False
SSdcboCategory.Enabled = False
dcboPrimUnit.Enabled = False
dcboSecUnit.Enabled = False
dcboChargeAccount.Enabled = False

NavBar1.SaveEnabled = False
  
  
      Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
       Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = False
        End If

    Next checkboxes
  
'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

optSpecific.Enabled = True
optPool.Enabled = True
SSdcboCategory.Enabled = True
dcboPrimUnit.Enabled = True
dcboSecUnit.Enabled = True
dcboChargeAccount.Enabled = True

NavBar1.SaveEnabled = True
  
  
    '  Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = True
        End If

    Next textboxes
  '    Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = True
        End If

    Next checkboxes



End If
End Sub

Private Sub cbo_StockNum_DropDown()
If NavBar1.Recordset.EditMode = 2 Then cbo_StockNum.DroppedDown = False
End Sub

Private Sub cbo_StockNum_GotFocus()
Call HighlightBackground(cbo_StockNum)
End Sub

Private Sub cbo_StockNum_KeyDown(KeyCode As Integer, Shift As Integer)
cbo_StockNum = Trim$(cbo_StockNum)
If deIms.rsSTOCKMASTER.EditMode = 2 Then Exit Sub
If Not cbo_StockNum.DroppedDown = True Then cbo_StockNum.DroppedDown = True
End Sub

Private Sub cbo_StockNum_LostFocus()
Call NormalBackground(cbo_StockNum)
End Sub

Private Sub cbo_StockNum_Scroll(Cancel As Integer)
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
'Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)


'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock
ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02

If locked = True Then 'sets locked = true because another user has this record open in edit mode

optSpecific.Enabled = False
optPool.Enabled = False
SSdcboCategory.Enabled = False
dcboPrimUnit.Enabled = False
dcboSecUnit.Enabled = False
dcboChargeAccount.Enabled = False

NavBar1.SaveEnabled = False
  
  
      Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
       Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = False
        End If

    Next checkboxes
  
'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

optSpecific.Enabled = True
optPool.Enabled = True
SSdcboCategory.Enabled = True
dcboPrimUnit.Enabled = True
dcboSecUnit.Enabled = True
dcboChargeAccount.Enabled = True

NavBar1.SaveEnabled = True
  
  
    '  Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = True
        End If

    Next textboxes
  '    Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = True
        End If

    Next checkboxes



End If
End Sub

Private Sub cbo_StockNum_Validate(Cancel As Boolean)
cbo_StockNum = Trim$(cbo_StockNum)
If Len(Trim$(cbo_StockNum)) = 0 Then Exit Sub

If Len(Trim$(cbo_StockNum)) > 20 Then
    MsgBox "Stock number can not be greater than 20 characters."
    Cancel = True
    cbo_StockNum.SetFocus
    Exit Sub
End If

If cbo_StockNum.IsItemInList And deIms.rsSTOCKMASTER.EditMode = 2 Then
     MsgBox "Stock number Already exists.Please use a different one."
     Cancel = True
     cbo_StockNum.SetFocus
    Exit Sub
End If



End Sub

Private Sub Check1_GotFocus()
Call HighlightBackground(Check1)
End Sub

Private Sub Check1_LostFocus()
Call NormalBackground(Check1)
End Sub

'call function to add current receptient to receptient list

Private Sub cmd_Add_Click()
Imsmail1.AddCurrentRecipient
End Sub

'call function delete current receptient from receptient list

Private Sub cmd_Remove_Click()

 If IsNothing(rsReceptList) = False Then
   
      rsReceptList.Find ("Recipients ='" & ssdbRecepientList.Columns(0).Text & "'")
      
      If Not rsReceptList.AbsolutePosition = adPosEOF Then
      
            rsReceptList.Delete
            If rsReceptList.RecordCount > 0 Then rsReceptList.MoveFirst
      
      End If
      
   End If
   
End Sub

Private Sub dcboChargeAccount_GotFocus()
Call HighlightBackground(dcboChargeAccount)
End Sub

Private Sub dcboChargeAccount_LostFocus()
Call NormalBackground(dcboChargeAccount)
End Sub

Private Sub dcboChargeAccount_Validate(Cancel As Boolean)
If Len(dcboChargeAccount) > 20 Then
MsgBox "Code number can not be greater than 20 characters."
Cancel = True
dcboChargeAccount.SetFocus
End If
End Sub

Private Sub dcboManuFac_GotFocus()
Call HighlightBackground(dcboManuFac)
End Sub

Private Sub dcboManuFac_LostFocus()
Call NormalBackground(dcboManuFac)
End Sub

'call function to get primary unit
''Private Sub dcboPrimUnit_Click(Area As Integer)
''Dim rs As ADODB.Recordset
''
''    If Area = 0 Then
''        Set rs = OpenUnit
''
''        Set dcboPrimUnit.RowSource = rs
''
''    End If
''End Sub
Private Sub dcboPrimUnit_dropdown()
Dim Rs As ADODB.Recordset

    'If Area = 0 Then
        Set Rs = OpenUnit
        
        Set dcboPrimUnit.DataSourceList = Rs
   
    ' End If
End Sub
'call function to get primary unit
Private Sub dcboPrimUnit_GotFocus()
'cOMMENTED OUT BY MUZAMMIL 03/17/01

''Dim rs As ADODB.Recordset
''
''    Set rs = OpenUnit
''
''    Set dcboPrimUnit.RowSource = rs
    Call HighlightBackground(dcboPrimUnit)
End Sub

'function get combo data

Function GetNearestDataComboItem(cbo As DataCombo, Optional KeyAscii As Integer, Optional sItem As String) As Boolean
On Error Resume Next
Dim Y As Integer, i As Integer

    #If DBUG = 0 Then
        On Error Resume Next
    #End If
    

    If sItem = "" Then
    
            'If KeyAscii = 0 Then _
                Err.Raise 9999, , "Specify a char code or a string": Exit Function
            
            cbo.SelText = ""
            sItem = cbo.Text
            Y = Len(cbo)
            i = cbo.SelLength
            
            If KeyAscii = 0 Then
            
            ElseIf KeyAscii > 31 Then
                cbo.SelText = Chr$(KeyAscii)
                
           
            Else
                Y = Y - 1
                i = i + 1
                
            
                cbo.SelStart = Y
                cbo.SelLength = i
                cbo.SelText = ""
                cbo.SelStart = Y
            End If
        
        sItem = cbo.Text: KeyAscii = 0:
        Y = cbo.SelStart: i = cbo.SelLength
        cbo.SetFocus: cbo.SelStart = Y: cbo.SelLength = i
    End If
    
    i = SendMessageStr(cbo.HWND, CB_FINDSTRING, CLng(-1), sItem)
                
    
        ' If i = CB_ERR Then i = cbo.ListIndex
        
        GetNearestDataComboItem = i <> CB_ERR
        Call SendMessage(cbo.HWND, &H14E, i, 0)
        
        
        If TypeName(cbo) = "ComboBox" And i <> CB_ERR Then
            cbo.SelStart = Len(sItem)
            cbo.SelLength = Len(cbo.Text) - cbo.SelStart
        End If
End Function

Private Sub dcboPrimUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If dcboPrimUnit.DroppedDown = False Then dcboPrimUnit.DroppedDown = True
End Sub

'call function to primary unit

Private Sub dcboPrimUnit_KeyPress(KeyAscii As Integer)
dcboPrimUnit.MoveNext
End Sub


Private Sub dcboPrimUnit_LostFocus()
Call NormalBackground(dcboPrimUnit)
End Sub

'validate secoundary unit and set data to recordset

Private Sub dcboPrimUnit_Validate(Cancel As Boolean)
On Error Resume Next
Dim str As String
Dim msg1 As String
Dim msg2 As String
Dim msg As String

    If Len(Trim$(dcboPrimUnit)) = 0 Then Exit Sub
   '  If ISAValidUnit(dcboPrimUnit.BoundText) = False Then 'And stock.editmode = 2 Then
     If ISAValidUnit(dcboPrimUnit.value) = False Then  'And stock.editmode = 2 Then
           Cancel = True
           MsgBox "Please select a unit from the list.", vbInformation, "Imswin"
           dcboPrimUnit.SetFocus
           Call HighlightBackground(dcboPrimUnit)
           Exit Sub
     End If
   
   If validateFromTable("UNIT", dcboPrimUnit.value, dcboPrimUnit.Text) = False Then

    MsgBox "Please enter a valid unit."
    Cancel = True
    dcboPrimUnit.SetFocus
    Exit Sub
    
   End If
   
   
   
   If Len(dcboPrimUnit) > 10 Then
        MsgBox "Primary unit can not be greater than 10 characters."
        Cancel = True
        dcboPrimUnit.SetFocus
        Exit Sub
        
    End If
    
    
    
    str = Trim$(stock!stk_secouom & "")
'    stock!stk_primuon = dcboPrimUnit.BoundText
     stock!stk_primuon = dcboPrimUnit.value
    
    If Len(str) = 0 Then Exit Sub
    
    'Added by Muzammil 04/02/01
    'Reason - To pop up the message box to take in the Computaion Factor Value.
    '-----------------------------------------------------------------------------
    
    
      If IsStringEqual(dcboSecUnit, dcboPrimUnit) = False Then
    
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00317") 'J added
            msg2 = translator.Trans("M00318") 'J modified
            msg = IIf(msg1 = "", "Please enter how many ", msg1 + " ") & dcboSecUnit 'J modified
            msg = msg & IIf(msg2 = "", " it takes to make 1 ", " " + msg2 + " ") & dcboPrimUnit 'J modified
           
            
            If ((stock.EditMode <> adEditNone) And _
                (stock!stk_primuon <> stock("stk_primuon").originalVALUE & "")) Then
                
                str = InputBox(msg, , 1)
                
                If Len(Trim$(str)) = 0 Then str = 1
                
                    If Not IsNumeric(str) Then
                    
                        Do Until IsNumeric(str)
                        
                            str = InputBox(msg, , 1)
                            If Len(Trim$(str)) = 0 Then str = 1
                        Loop
                        
                    End If
    
                    If str = 1 Then
                        stock!stk_compfctr = 1
                    Else
                        stock!stk_compfctr = CDbl(10000 / str)
                    End If
                    
                    
        End If
    
    End If
    
    
    '--------------------------------------------------------------------------------
    
    
    
''    If Len(str) = 0 Then
''        stock!stk_secouom = dcboPrimUnit.BoundText
''    Else
''        Call dcboSecUnit_Validate(False)
''    End If
    
    
End Sub

Private Sub dcboSecUnit_Click()
   
     ' Secunit = dcboSecUnit.Columns(0).text
      
End Sub

'call function toget secoundary unit recordset

''Private Sub dcboSecUnit_Click(Area As Integer)
''Dim rs As ADODB.Recordset
''
''    If Area = 0 Then
''        Set rs = OpenUnit
''
''        Set dcboSecUnit.RowSource = rs
''
''    End If
''End Sub

''Private Sub dcboSecUnit_Click()
''Dim rs As ADODB.Recordset
''
''    'If Area = 0 Then
''        Set rs = OpenUnit
''
''        Set dcboSecUnit.DataSourceList = rs
''
''     'End If
''End Sub

Private Sub dcboSecUnit_DropDown()

Dim Rs As ADODB.Recordset

    'If Area = 0 Then
        Set Rs = OpenUnit

        Set dcboSecUnit.DataSourceList = Rs

     'End If
End Sub


Private Sub dcboSecUnit_GotFocus()
   Call HighlightBackground(dcboSecUnit)
End Sub

Private Sub dcboSecUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If dcboSecUnit.DroppedDown = False Then dcboSecUnit.DroppedDown = True
End Sub

Private Sub dcboSecUnit_KeyPress(KeyAscii As Integer)
dcboSecUnit.MoveNext
End Sub

Private Sub dcboSecUnit_LostFocus()
Call NormalBackground(dcboSecUnit)
'''
'''If deIms.rsUNIT.State And adStateOpen Then
'''        Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
'''        'Set dcboSecUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
'''    Else
'''        Call deIms.Unit(deIms.NameSpace)
'''        Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
''''        Set dcboPrimUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
'''
'''        deIms.rsUNIT.Close
'''    End If
End Sub

'calculate secondary unit and assign data to recordset

Private Sub dcboSecUnit_Validate(Cancel As Boolean)
On Error Resume Next

    Dim msg As String
    Dim str As String
    Dim stockCODE As String
    Dim StockDesc As String
'    stockCODE = dcboSecUnit.BoundText
    
   ' stock!stk_secouom = stockCODE
     If Len(Trim$(dcboSecUnit)) = 0 Then Exit Sub
     'If Len(Trim$(dcboSecUnit.BoundText)) = 0 Then Exit Sub
    ' If Len(Trim$(dcboSecUnit.Columns(0).text)) = 0 Then Exit Sub
    If Len(Trim$(dcboSecUnit.value)) = 0 Then Exit Sub
     
     
     'If ISAValidUnit(dcboSecUnit.BoundText) = False Then
     ' If ISAValidUnit(dcboSecUnit.Columns(0).text) = False Then
      If ISAValidUnit(dcboSecUnit.value) = False Then
           Cancel = True
           MsgBox "Please Enter a valid unit.", vbInformation, "Imswin"
           dcboSecUnit.SetFocus
           Call HighlightBackground(dcboSecUnit)
           Exit Sub
     End If
         
     If validateFromTable("UNIT", dcboSecUnit.value, dcboSecUnit.Text) = False Then

        MsgBox "Please enter a valid unit."
        Cancel = True
        dcboSecUnit.SetFocus
        Exit Sub
        
     End If
         
         
         
     If Len(dcboSecUnit) > 20 Then
        MsgBox "Unit can not be greater than 20 characters."
        Cancel = True
        dcboSecUnit.SetFocus
     End If
         
    'stock!stk_secouom = dcboSecUnit.BoundText
    'stock!stk_secouom = dcboSecUnit.Columns(0).text
     stock!stk_secouom = dcboSecUnit.value
    'If stock!stk_secouom = dcboPrimUnit.BoundText Then
     If stock!stk_secouom = dcboPrimUnit.value Then _
        stock!stk_compfctr = Null: Exit Sub
    
    
    'Msg = Trim$(dcboSecUnit.BoundText)
    'Msg = Trim$(dcboSecUnit.Columns(0).text)
    msg = Trim$(dcboSecUnit.value)
    'str = Trim$(dcboPrimUnit.BoundText)
    str = Trim$(dcboPrimUnit.value)
    
    If Len(str) = 0 Then Exit Sub
    If Len(msg) = 0 Then Exit Sub
    
    str = ""
    If IsStringEqual(dcboSecUnit, dcboPrimUnit) = False Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00317") 'J added
        msg2 = translator.Trans("M00318") 'J modified
        msg = IIf(msg1 = "", "Please enter how many ", msg1 + " ") & dcboSecUnit 'J modified
        msg = msg & IIf(msg2 = "", " it takes to make 1 ", " " + msg2 + " ") & dcboPrimUnit 'J modified
        '---------------------------------------------
        
        If ((stock.EditMode <> adEditNone) And _
            (stock!stk_secouom <> stock("stk_secouom").originalVALUE & "")) Then
            
            str = InputBox(msg, , 1)
            
            If Len(Trim$(str)) = 0 Then str = 1
            
            If Not IsNumeric(str) Then
            
                Do Until IsNumeric(str)
                
                    str = InputBox(msg, , 1)
                    If Len(Trim$(str)) = 0 Then str = 1
                Loop
                
            End If

                If str = 1 Then
                    stock!stk_compfctr = 1
                Else
                    stock!stk_compfctr = CDbl(10000 / str)
                End If
                
                
        End If
    Else
        stock!stk_compfctr = 1
    End If
 
 Cancel = False
 
End Sub

'set image values

Private Sub flsbHoriz_Change()
    imgImage.Left = -flsbHoriz.value
End Sub

'set image values

Private Sub flsbHoriz_Scroll()
    imgImage.Left = -flsbHoriz.value
End Sub

'set image values

Private Sub flsbVert_Change()
    imgImage.Top = -flsbVert.value
End Sub

'set image values

Private Sub flsbVert_Scroll()
    imgImage.Top = -flsbVert.value
End Sub

'close recordset and free memory

Private Sub Form_Unload(Cancel As Integer)
 
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  

On Error Resume Next

    Hide
    stock.Close
    Set stock = Nothing
    deIms.rsGetStockManufacturer.Close
    deIms.rsStockNumbers.Close
    'Set dcboSecUnit.RowSource = Nothing
    Set dcboSecUnit.DataSourceList = Nothing
      Set dcboPrimUnit.DataSourceList = Nothing
    'Set dcboPrimUnit.RowSource = Nothing
    Set dcboChargeAccount.RowSource = Nothing
    Set SSdcboCategory.DataSourceLis = Nothing
    'Set ssdbddManufacturer.DataSource = Nothing
    
    If IsNothing(rsReceptList) = False Then Set rsReceptList = Nothing
    
    If Err Then Err.Clear
    'Imsmail1.Connected = False 'M
    If open_forms <= 5 Then ShowNavigator
End Sub

'call function to load image files

Private Sub cmdOpen_Click()
Dim st As String, i As Long
    
    st = sGetFileName
    If Len(st) > 3 Then
        If SSTab1.Tab <> 4 Then
            stock!stk_imge = FileToField(st, i)
        Else
            deIms.rsGetStockManufacturer!stm_imge = FileToField(st, i)
        End If
    End If

End Sub

'call function to image files

Private Sub cmdSave_Click()
Dim i As IPictureDisp
On Error GoTo Handled
    
    MDI_IMS.cmdDialog.ShowSave
    
    Set i = imgImage.Picture
    Call SavePicture(i, MDI_IMS.cmdDialog.FileName)
    
    Exit Sub
Handled:
    If Err = 32755 Then
        Err.Clear
    Else
        MsgBox Err.Description
    End If
End Sub



'add receptient to receptient list

Private Sub IMSMail1_OnAddClick(ByVal address As String)
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

'set recordset position to add new position

Private Sub NavBar1_BeforeCancelClick()
    NavBar1.AllowAddNew = True
    NavBar1.AllowUpdate = True
  
End Sub

Private Sub NavBar1_BeforeFirstClick()
  

  
  Call SetDatasourceForUnits
End Sub

Private Sub NavBar1_BeforeLastClick()
  

  Call SetDatasourceForUnits
End Sub

'get stock manufacturer recordset

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
On Error Resume Next
    deIms.rsGetStockManufacturer.UpdateBatch
    If Err Then Err.Clear
End Sub



Private Sub NavBar1_BeforeNextClick()
  

  Call SetDatasourceForUnits
End Sub

Private Sub NavBar1_BeforePreviousClick()


  
  Call SetDatasourceForUnits
End Sub

Private Sub NavBar1_OnCancelClick()
If mFromClick = True Then Exit Sub
  If Not NavBar1.Recordset.RecordCount = 0 Then cbo_StockNum = NavBar1.Recordset!stk_stcknumb
  
''  If NavBar1.Recordset.editmode = 1 Then
''
''      Priunit = NavBar1.Recordset("stk_primuon").originalVALUE
''      Secunit = NavBar1.Recordset("stk_secouom").originalVALUE
''  ElseIf NavBar1.Recordset.editmode = 2 Then
''          Priunit = ""
''          Secunit = ""
''  End If
  
End Sub

'unload form

Private Sub NavBar1_OnCloseClick()
    Unload Me
    
End Sub

Private Sub NavBar1_OnFirstClick()

'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
''Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)
'
'cbo_StockNum.Refresh
''jawdat, start copy
'Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
''Dim imsLock As imsLock.lock
'Dim ListOfPrimaryControls() As String
'Set imsLock = New imsLock.Lock
'ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
'
'If locked = True Then 'sets locked = true because another user has this record open in edit mode
'
'optSpecific.Enabled = False
'optPool.Enabled = False
'SSdcboCategory.Enabled = False
'dcboPrimUnit.Enabled = False
'dcboSecUnit.Enabled = False
'dcboChargeAccount.Enabled = False
'
'NavBar1.SaveEnabled = False
'
'
'      Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'       Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = False
'        End If
'
'    Next checkboxes
'
''Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
'Else
'
'optSpecific.Enabled = True
'optPool.Enabled = True
'SSdcboCategory.Enabled = True
'dcboPrimUnit.Enabled = True
'dcboSecUnit.Enabled = True
'dcboChargeAccount.Enabled = True
'
'NavBar1.SaveEnabled = True
'
'
'    '  Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = True
'        End If
'
'    Next textboxes
'  '    Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = True
'        End If
'
'    Next checkboxes
'
'
'
'End If
''


If SSTab1.Tab = 0 Then
cbo_StockNum = NavBar1.Recordset!stk_stcknumb
End If

'Priunit = NavBar1.Recordset!stk_primuon
'Secunit = NavBar1.Recordset!stk_secouom

End Sub

Private Sub NavBar1_OnLastClick()
'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
''Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)
'
'
''jawdat, start copy
'Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
''Dim imsLock As imsLock.lock
'Dim ListOfPrimaryControls() As String
'Set imsLock = New imsLock.Lock
'ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
'
'If locked = True Then 'sets locked = true because another user has this record open in edit mode
'
'optSpecific.Enabled = False
'optPool.Enabled = False
'SSdcboCategory.Enabled = False
'dcboPrimUnit.Enabled = False
'dcboSecUnit.Enabled = False
'dcboChargeAccount.Enabled = False
'
'NavBar1.SaveEnabled = False
'
'
'      Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'       Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = False
'        End If
'
'    Next checkboxes
'
''Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
'Else
'
'optSpecific.Enabled = True
'optPool.Enabled = True
'SSdcboCategory.Enabled = True
'dcboPrimUnit.Enabled = True
'dcboSecUnit.Enabled = True
'dcboChargeAccount.Enabled = True
'
'NavBar1.SaveEnabled = True
'
'
'    '  Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = True
'        End If
'
'    Next textboxes
'  '    Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = True
'        End If
'
'    Next checkboxes
'
'
'
'End If


'End If
If SSTab1.Tab = 0 Then
  cbo_StockNum = NavBar1.Recordset!stk_stcknumb
End If
'Priunit = NavBar1.Recordset!stk_primuon
'Secunit = NavBar1.Recordset!stk_secouom
End Sub

'assige datas to recordset, create user and modify user to current
'user and name space to current name space

Private Sub NavBar1_OnNewClick()

If locked = True Then
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
End If
  



On Error Resume Next
Dim i As Integer

    
    If SSTab1.Tab <> 4 Then
        i = SSTab1.Tab
        
        PriUnit = ""
        SecUnit = ""
        cbo_StockNum = ""
        
        
        deIms.rsSTOCKMASTER!stk_modiuser = CurrentUser
        deIms.rsSTOCKMASTER!stk_creauser = CurrentUser
        deIms.rsSTOCKMASTER!stk_npecode = deIms.NameSpace
        Check1.value = 1
        optPool.value = True
        'cbo_StockNum = ""
        SSTab1.Tab = 4
        deIms.rsGetStockManufacturer.Close
        Call deIms.rsGetStockManufacturer.AddNew
        AssignManufacturesDefault
        
        SSTab1.Tab = i
    Else
        AssignManufacturesDefault
    End If
    
    If Err Then Err.Clear
    cbo_StockNum.SetFocus
    

End Sub

Private Sub NavBar1_OnNextClick()
'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
''Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)
'
'
''jawdat, start copy
'Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
''Dim imsLock As imsLock.lock
'Dim ListOfPrimaryControls() As String
'Set imsLock = New imsLock.Lock
'ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
'
'If locked = True Then 'sets locked = true because another user has this record open in edit mode
'
'optSpecific.Enabled = False
'optPool.Enabled = False
'SSdcboCategory.Enabled = False
'dcboPrimUnit.Enabled = False
'dcboSecUnit.Enabled = False
'dcboChargeAccount.Enabled = False
'
'NavBar1.SaveEnabled = False
'
'
'      Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'       Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = False
'        End If
'
'    Next checkboxes
'
''Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
'Else
'
'optSpecific.Enabled = True
'optPool.Enabled = True
'SSdcboCategory.Enabled = True
'dcboPrimUnit.Enabled = True
'dcboSecUnit.Enabled = True
'dcboChargeAccount.Enabled = True
'
'NavBar1.SaveEnabled = True
'
'
'    '  Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = True
'        End If
'
'    Next textboxes
'  '    Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = True
'        End If
'
'    Next checkboxes
'
'
'
'End If
''

If SSTab1.Tab = 0 Then
   cbo_StockNum = NavBar1.Recordset!stk_stcknumb
End If
'Priunit = NavBar1.Recordset!stk_primuon
'Secunit = NavBar1.Recordset!stk_secouom
End Sub

Private Sub NavBar1_OnPreviousClick()
'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
''Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)
'
'
''jawdat, start copy
'Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
''Dim imsLock As imsLock.lock
'Dim ListOfPrimaryControls() As String
'Set imsLock = New imsLock.Lock
'ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
'
'If locked = True Then 'sets locked = true because another user has this record open in edit mode
'
'optSpecific.Enabled = False
'optPool.Enabled = False
'SSdcboCategory.Enabled = False
'dcboPrimUnit.Enabled = False
'dcboSecUnit.Enabled = False
'dcboChargeAccount.Enabled = False
'
'NavBar1.SaveEnabled = False
'
'
'      Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'       Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = False
'        End If
'
'    Next checkboxes
'
''Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
'Else
'
'optSpecific.Enabled = True
'optPool.Enabled = True
'SSdcboCategory.Enabled = True
'dcboPrimUnit.Enabled = True
'dcboSecUnit.Enabled = True
'dcboChargeAccount.Enabled = True
'
'NavBar1.SaveEnabled = True
'
'
'    '  Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = True
'        End If
'
'    Next textboxes
'  '    Dim checkboxes As Control
'
'    For Each checkboxes In Controls
'        If (TypeOf checkboxes Is CheckBox) Then
'            checkboxes.Enabled = True
'        End If
'
'    Next checkboxes
'
'
'
'End If

If SSTab1.Tab = 0 Then
cbo_StockNum = NavBar1.Recordset!stk_stcknumb
End If
'Priunit = NavBar1.Recordset!stk_primuon
'Secunit = NavBar1.Recordset!stk_secouom
End Sub

'get crystal report parameters and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Errhandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath & "Stckmaster1.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stcknumb;" & Trim$(cbo_StockNum) & ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00119") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock Master", msg1) 'J modified
        Call translator.Translate_Reports("Stckmaster1.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
     Exit Sub
    
Errhandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If

End Sub

'get email function parameters

Private Sub NavBar1_OnEMailClick()

Dim ParamsForRPTI(1) As String

Dim rptinf As RPTIFileInfo

Dim ParamsForCrystalReports(1) As String

Dim Subject As String

Dim FieldName As String

Dim Message As String

Dim attention As String

On Error Resume Next

If rsReceptList Is Nothing Then Exit Sub
                
    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    
    ParamsForCrystalReports(1) = "stcknumb;" + Trim$(cbo_StockNum) + ";TRUE"
    
    ParamsForRPTI(0) = "namespace=" & deIms.NameSpace
    
    ParamsForRPTI(1) = "stcknumb=" & Trim$(cbo_StockNum)
    
    FieldName = "Recipients"
    
    Subject = "Stock Master Record " & Trim$(cbo_StockNum)
    
    If ConnInfo.EmailClient = Outlook Then
    
        'Call sendOutlookEmailandFax("stckmaster1.rpt", "StockMaster", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, Subject, attention) MM 030209
        Call sendOutlookEmailandFax(Report_EmailFax_Stockmaster_name, "StockMaster", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, Subject, attention)
    
    ElseIf ConnInfo.EmailClient = ATT Then
    
        Call SendAttFaxAndEmail("stckmaster1.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, Subject, Message, FieldName)

    ElseIf ConnInfo.EmailClient = Unknown Then
    
        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

    End If

    Call rsReceptList.Delete(adAffectAllChapters)

    Set rsReceptList = Nothing
    
    Set ssdbRecepientList.DataSource = Nothing

'''' This is the old piece of code.
''''
''''On Error Resume Next
''''
''''    If rsReceptList Is Nothing Then Exit Sub
''''
''''
''''
''''    BeforePrint
''''
''''    With rptinf
''''        .ReportFileName = ReportPath & "stckmaster1.rpt"
''''
''''        'Modified by Juan (8/28/2000) for Multilingual
''''        Call translator.Translate_Reports("Stckmaster1.rpt") 'J added
''''        '---------------------------------------------
''''
''''        Params(0) = "namespace=" & deIms.NameSpace
''''        Params(1) = "stcknumb=" & Trim$(cbo_StockNum)
''''        .Parameters = Params
''''    End With
''''
''''    Params(0) = ""
''''    Call WriteRPTIFile(rptinf, Params(0))
''''    Call SendEmailAndFax(rsReceptList, "Recipients", "Stock Master Record " & Trim$(cbo_StockNum), "Stock Master Record", Params(0))
''''
''''
''''
''''    Call rsReceptList.Delete(adAffectAllChapters)
''''
''''    Set rsReceptList = Nothing
''''    Set ssdbRecepientList.DataSource = Nothing
''''
''''    If Err Then Call LogErr(Name & "::NavBar1_OnEMailClick", Err.Description, Err)
''''    Err.Clear
End Sub

'get crystal report parameters

Public Sub BeforePrint()
On Error GoTo Errhandler

      With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\stckmaster1.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stcknumb;" + Trim$(cbo_StockNum) + ";TRUE"
        
    End With
    Exit Sub
    
Errhandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'before save validate data format
'check stock number exist or not

Private Sub NavBar1_BeforeSaveClick()
On Error Resume Next
Dim msg As String
Dim bl As Boolean
Dim str As String, OldNum As String
Dim Cancel As Boolean


    Dim ls_full_record As String
    Dim ls_stocknum As String
    Dim li_StoreYesNo As Integer
    
    NavBar1.AllowUpdate = False
    'Call cbo_StockNum_Validate(bl)
    
    If bl Then Exit Sub
    
    ValidateControls
    NavBar1.AllowUpdate = False
    
    If SSTab1.Tab = 0 Then
        
             deIms.rsSTOCKMASTER!stk_stcknumb = Trim$(cbo_StockNum)
             deIms.rsSTOCKMASTER!stk_modiuser = CurrentUser
             deIms.rsSTOCKMASTER!stk_creauser = CurrentUser
             deIms.rsSTOCKMASTER!stk_npecode = deIms.NameSpace
             
             If Len(Trim$(cbo_StockNum.Text)) = 0 Then
             
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00119") 'J added
                 msg = IIf(msg1 = "", "Stock Number", msg1) 'J modified
                 '---------------------------------------------
                 
             Else
                 deIms.rsSTOCKMASTER!stk_stcknumb = cbo_StockNum
             End If
             
             If Len(Trim$(txt_LongDescript)) = 0 Then
             
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00278") 'J added
                 msg = IIf(msg1 = "", "Long Description", msg1) 'J modified
                 '---------------------------------------------
                 
             End If
                 
            If Len(Trim$(dcboPrimUnit)) = 0 Then
            
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00112") 'J added
                 msg = IIf(msg1 = "", "Primary Unit", msg1) 'J modified
                 '---------------------------------------------
                 
             Else
                'deIms.rsSTOCKMASTER!stk_primuon = dcboPrimUnit.BoundText
                deIms.rsSTOCKMASTER!stk_primuon = dcboPrimUnit.value
             End If
                 
              If Len(Trim$(SSdcboCategory)) = 0 Then
              
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00027") 'J added
                 msg = IIf(msg1 = "", "Category", msg1) 'J modified
                 '---------------------------------------------
                 
              End If
              
              If Len(Trim$(dcboSecUnit)) = 0 Then
              
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00115") 'J added
                 msg = IIf(msg1 = "", "Secondary Unit", msg1) 'J modified
                 '---------------------------------------------
                 
              Else
                 'deIms.rsSTOCKMASTER!stk_secouom = dcboSecUnit.BoundText
                 'deIms.rsSTOCKMASTER!stk_secouom = dcboSecUnit.Columns(0).text
                 deIms.rsSTOCKMASTER!stk_secouom = dcboSecUnit.value
              End If
              
              If Len(Trim$(txt_Estimate)) = 0 Then
              
                 'Modified by Juan (9/14/2000) for Multilingual
                 msg1 = translator.Trans("L00283") 'J added
                 msg = IIf(msg1 = "", "Estimated Price", msg1) 'J modified
                 '---------------------------------------------
                 
             Else
                 'deIms.rsSTOCKMASTER!stk_estmprice = CInt(txt_Estimate)
              End If
         
             AddValuesToStockMan
             
    ElseIf SSTab1.Tab = 4 Then
            
            NavBar1.AllowUpdate = True
            If Len(Trim$(dcboManuFac)) = 0 Then
                NavBar1.AllowUpdate = False
                MsgBox "Manufacturer field can not be left empty.", vbInformation, "Imswin"
                dcboManuFac.SetFocus
                Call HighlightBackground(dcboManuFac)
                Exit Sub
            End If
            
            If Len(Trim$(Text1)) = 0 Then
                NavBar1.AllowUpdate = False
                MsgBox "Part number field can not be left empty.", vbInformation, "Imswin"
                Text1.SetFocus
                Call HighlightBackground(Text1)
                Exit Sub
            End If
            
            
               AssignManufacturesDefault
               deIms.rsGetStockManufacturer.Update
            
            
   End If
    
    
    If Len(Trim$(msg)) <> 0 Then
        
            NavBar1.AllowAddNew = False
            NavBar1.AllowUpdate = False
            
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00016") 'J added
            MsgBox msg & IIf(msg1 = "", " cannot be left empty", " " + msg1) 'J modified
            '---------------------------------------------
            
    Else
        
            NavBar1.AllowAddNew = True
            NavBar1.AllowUpdate = True
        
    End If
    
    If Err Then Err.Clear
    
    DoEvents
    DoEvents
    
    
    
    If SSTab1.Tab = 0 Then
    

            str = LCase(Trim$(cbo_StockNum.Text))
            OldNum = LCase(Trim$(stock!stk_stcknumb.originalVALUE & ""))
              
            If Len(OldNum) Then
        
                    If OldNum <> str Then
                        cbo_StockNum.Text = OldNum
                        
                        'Modified by Juan (9/14/2000) for Multilingual
                        msg1 = translator.Trans("M00262") 'J added
                        MsgBox IIf(msg1 = "", "Stock Number cannot be changed once saved", msg1) 'J modified
                        '---------------------------------------------
                        
                        NavBar1.AllowUpdate = False
                        Exit Sub
                    End If
                
            Else
        
                    If Len(str) Then
                        If deIms.StockNumberExist(str, False) Then
                            Cancel = True
                            
                            'Modified by Juan (9/14/2000) for Multilingual
                            msg1 = translator.Trans("L00119") 'J added
                            msg1 = translator.Trans("L00541") 'J modified
                            MsgBox IIf(msg1 = "", "Stock number ", msg1 + " ") & str & IIf(msg2 = "", " already exist", " " + msg2) 'Modified
                            '---------------------------------------------
                            
                            NavBar1.AllowUpdate = False
                        End If
            
                    End If
        
            End If
            
            
End If

     mIsItInsert = True
     
     
End Sub

'call function to get datas for data grids and combo

Private Sub Form_Load()
On Error Resume Next
Dim ctl As Control
Dim Rs As ADODB.Recordset

    'Added by Juan (9/14/2000) for Multilngual
    Call translator.Translate_Forms("frm_Stock2")
    '-----------------------------------------

    For Each ctl In Controls
         Call gsb_fade_to_black(ctl)
    Next
    
    lbl_CompFactor.BackColor = &HFFFFC0
    
    If deIms.rsMANUFACTURER.State And adStateOpen Then
         
         Set dcboManuFac.RowSource = deIms.rsMANUFACTURER.Clone()
        
    Else
        
        deIms.manufacturer (deIms.NameSpace)
        
         Set dcboManuFac.RowSource = deIms.rsMANUFACTURER.Clone(adLockReadOnly)
        
        deIms.rsMANUFACTURER.Close
    
    End If
    
    Set Rs = OpenCategory
    Set SSdcboCategory.DataSourceList = Rs
    
    Set Rs = OpenUnit
   'Set dcboPrimUnit.RowSource = rs
'    Set dcboSecUnit.RowSource = rs
    Set dcboPrimUnit.DataSourceList = Rs
  '  Set dcboPrimUnit.DataSourceList = rs
    Set dcboSecUnit.DataSourceList = Rs
   
    
    
    If deIms.rsUNIT.State And adStateOpen Then

        
        'Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
        
        Set dcboSecUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
        'Set dcboPrimUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)

    Else
        Call deIms.Unit(deIms.NameSpace)
        
        'Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
        
        Set dcboSecUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
      '  Set dcboPrimUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)

        deIms.rsUNIT.Close
    End If

    If deIms.rsCHARGE.State And adStateOpen Then
        Set dcboChargeAccount.RowSource = deIms.rsCHARGE.Clone(adLockReadOnly)

    Else
        Call deIms.CHARGE(deIms.NameSpace)
        Set dcboChargeAccount.RowSource = deIms.rsCHARGE.Clone(adLockReadOnly)

        deIms.rsCHARGE.Close
    End If
'
        
    deIms.rsSTOCKMASTER.Close
    Call deIms.STOCKMASTER(deIms.NameSpace)
    If Err Then Err.Clear
    
    Set stock = deIms.rsSTOCKMASTER
    Set NavBar1.Recordset = deIms.rsSTOCKMASTER
        
    deIms.rsSTOCKMASTER.Move 0
        
        
    Call deIms.StockNumbers(deIms.NameSpace)
    If deIms.rsStockNumbers.RecordCount > 0 Then
    
            deIms.rsStockNumbers.MoveFirst
            cbo_StockNum.Text = deIms.rsStockNumbers!stk_stcknumb
            Do While Not deIms.rsStockNumbers.EOF
                cbo_StockNum.AddItem deIms.rsStockNumbers!stk_stcknumb
                deIms.rsStockNumbers.MoveNext
            Loop
            
    End If
    
    deIms.rsStockNumbers.Close
    
    Call DisableButtons(Me, NavBar1)
    
    'The user only has read only properties.
    
    If NavBar1.NewEnabled = False Then
    
       rights = mdReadonly
       SSdcboCategory.Enabled = False
       dcboPrimUnit.Enabled = False
       dcboSecUnit.Enabled = False
       cbo_StockNum.AllowInput = True
       
    Else
       rights = mdReadWriteOnly
       SSdcboCategory.Enabled = True
       dcboPrimUnit.Enabled = True
       dcboSecUnit.Enabled = True
       
    End If
    
    
    deIms.rsGetStockManufacturer.Close
    Call deIms.GetStockManufacturer(deIms.NameSpace, stock!stk_stcknumb)

    Call BindAll(Me, deIms)
    Imsmail1.NameSpace = deIms.NameSpace
    Call BindControlsToDataMenber("GetStockManufacturer", Me)
    
    Imsmail1.SetActiveConnection deIms.cnIms 'M
    Imsmail1.Language = Language 'M
    
    cbo_StockNum.locked = False
    Caption = Caption + " - " + Tag
    
    Check1.value = vbChecked
    Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
'Call imsLock.UnLock_table(dbtablename, Me.Name, deIms.cnIms, CurrentUser)


'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock
ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02

If locked = True Then 'sets locked = true because another user has this record open in edit mode

optSpecific.Enabled = False
optPool.Enabled = False
SSdcboCategory.Enabled = False
dcboPrimUnit.Enabled = False
dcboSecUnit.Enabled = False
dcboChargeAccount.Enabled = False

NavBar1.SaveEnabled = False
  
  
      Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
       Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = False
        End If

    Next checkboxes
  
'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else

optSpecific.Enabled = True
optPool.Enabled = True
SSdcboCategory.Enabled = True
dcboPrimUnit.Enabled = True
dcboSecUnit.Enabled = True
dcboChargeAccount.Enabled = True

NavBar1.SaveEnabled = True
  
  
    '  Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = True
        End If

    Next textboxes
  '    Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = True
        End If

    Next checkboxes



End If

End Sub

'set recordset update position

Private Sub NavBar1_OnSaveClick()
On Error Resume Next

    If SSTab1.Tab = 0 Then
        deIms.rsGetStockManufacturer.Update
        Call deIms.rsSTOCKMASTER.UpdateBatch(adAffectCurrent)
        
        
        deIms.rsGetStockManufacturer.Update
        deIms.rsGetStockManufacturer.UpdateBatch
    ElseIf deIms.rsSTOCKMASTER.EditMode <> adEditAdd Then
        deIms.rsGetStockManufacturer.UpdateBatch
    End If
    If Err Then Err.Clear
    If mIsItInsert = True Then
    
    If cbo_StockNum.IsItemInList = False Then cbo_StockNum.AddItem Trim$(cbo_StockNum)

    End If
    
    MsgBox "Record saved Successfully.", vbInformation, "Imswin"
End Sub

'assign data to recordset

Private Sub optPool_Click()
    If deIms.rsSTOCKMASTER!stk_poolspec = 0 Then _
        deIms.rsSTOCKMASTER!stk_poolspec = 1
End Sub

Private Sub optPool_GotFocus()
Call HighlightBackground(optPool)
End Sub

Private Sub optPool_LostFocus()
Call NormalBackground(optPool)
End Sub

'assign data to recordset

Private Sub optSpecific_Click()
    If deIms.rsSTOCKMASTER!stk_poolspec <> 0 Then _
        deIms.rsSTOCKMASTER!stk_poolspec = 0
End Sub

Private Sub optSpecific_GotFocus()
Call HighlightBackground(optSpecific)
End Sub

Private Sub optSpecific_LostFocus()
Call NormalBackground(optSpecific)
End Sub

'set data to be delete

Private Sub ssdbRecepientList_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    Cancel = False
    DispPromptMsg = False
End Sub

'delete receptient fromreceptient list

Private Sub ssdbRecepientList_DblClick()
    
  '  If Len(ssdbRecepientList.SelBookmarks(0)) > 0 Then
    
   If IsNothing(rsReceptList) = False Then
   
      rsReceptList.Find ("Recipients ='" & ssdbRecepientList.Columns(0).Text & "'")
      
      If Not rsReceptList.AbsolutePosition = adPosEOF Then
      
            rsReceptList.Delete
            If rsReceptList.RecordCount > 0 Then rsReceptList.MoveFirst
      
      End If
      
   End If
  '    ssdbRecepientList.DeleteSelected
  '  End If
End Sub

'get recordset data and assign values to data grid

Private Sub SSdcboCategory_Click()

End Sub

'select active recordset to fill data grid

Private Sub SSdcboCategory_DropDown()
Dim Rs As ADODB.Recordset

'If Rights = mdReadonly Then SSdcboCategory.DroppedDown = False: Exit Sub

    Set Rs = OpenCategory
    Rs.Filter = "cate_actvflag <> 0"
    Set SSdcboCategory.DataSourceList = Rs
End Sub



Private Sub SSdcboCategory_GotFocus()
Call HighlightBackground(SSdcboCategory)
End Sub

Private Sub SSdcboCategory_KeyDown(KeyCode As Integer, Shift As Integer)

If Not SSdcboCategory.DroppedDown Then SSdcboCategory.DroppedDown = True
End Sub

Private Sub SSdcboCategory_KeyPress(KeyAscii As Integer)
SSdcboCategory.MoveNext
End Sub

Private Sub SSdcboCategory_LostFocus()
Call NormalBackground(SSdcboCategory)
End Sub

Private Sub SSdcboCategory_Validate(Cancel As Boolean)

'If SSdcboCategory.IsItemInList = False Then




If Len(SSdcboCategory) > 25 Then
    MsgBox "Category can not be greater than 25 characters."
    Cancel = True
    SSdcboCategory.SetFocus
    Exit Sub
End If

If validateFromTable("category", SSdcboCategory.value, SSdcboCategory.Text) = False Then

    MsgBox "Please enter a valid category."
    Cancel = True
    SSdcboCategory.SetFocus
    Exit Sub
End If



If Not SSdcboCategory.IsItemInList Then
   MsgBox "Please selected a valid category from the list.", vbInformation, "Imswin"
   Cancel = True
   SSdcboCategory.SetFocus
   Exit Sub
End If
 
End Sub

'depend on tab to set buttom and get stock data information

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim cmd As ADODB.Command
On Error Resume Next


    stock.Update
    If SSTab1.Tab = 4 Then
         
         dcboManuFac.SetFocus
        If Len(cbo_StockNum) Then stock!stk_stcknumb = cbo_StockNum
        If deIms.rsGetStockManufacturer!stm_stcknumb <> stock!stk_stcknumb Then
            deIms.rsGetStockManufacturer.Close
            
            Set cmd = deIms.Commands("GetStockManufacturer")
            
            cmd.parameters("@NameSpace") = deIms.NameSpace
            cmd.parameters("@STOCKNUMBER") = stock!stk_stcknumb
            
            cmd.Execute
            Call deIms.rsGetStockManufacturer.Close
            Call deIms.rsGetStockManufacturer.Open(, , adOpenStatic, adLockBatchOptimistic)
            
            Call BindControlsToDataMenber("GetStockManufacturer", Me)
            
        End If
    
    If Err Then Err.Clear
        Set NavBar1.Recordset = deIms.rsGetStockManufacturer
    Else
    
        If SSTab1.Tab = 2 Then
            cmdOpen.Visible = True: cmdSave.Visible = True
        Else
            cmdOpen.Visible = False: cmdSave.Visible = False
        End If
        
        Set NavBar1.Recordset = deIms.rsSTOCKMASTER
    End If
    
    
    
End Sub

'select active data to recordset

Private Sub SSTab1_DragDrop(Source As Control, x As Single, Y As Single)
Dim Rs As ADODB.Recordset

    Set Rs = OpenCategory
    Rs.Filter = "cate_actvflag <> 0"
    Set SSdcboCategory.DataSourceList = Rs
End Sub

'call function to get category recordset

Private Function OpenCategory() As ADODB.Recordset

    If deIms.rsCATEGORY.State And adStateOpen Then
        Set OpenCategory = deIms.rsCATEGORY.Clone
    Else
        deIms.Category (deIms.NameSpace)
        Set OpenCategory = deIms.rsCATEGORY.Clone
        deIms.rsCATEGORY.Close
    End If
    
End Function

'enable open and save command buttom

Private Sub SSTab2_Click(PreviousTab As Integer)
    If SSTab2.Tab = 1 Then
        cmdOpen.Visible = True: cmdSave.Visible = True
    Else
        cmdOpen.Visible = False: cmdSave.Visible = False
    End If
End Sub

Private Sub SSTab2_LostFocus()
dcboManuFac.SetFocus
End Sub

'set category values pool or specify

Private Sub Stock_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 On Error Resume Next
 
    If SSTab1.Tab = 0 Then
        If Not ((stock.EOF) Or (stock.Bof)) Then
        
            If stock!stk_poolspec <> 0 Then
                optPool.value = True
                
            ElseIf stock!stk_poolspec = 0 Then
                optSpecific.value = True
                
            Else
                optPool = False
                optSpecific = False
            End If
            
        End If
    End If
    
    imgImage.Move 0, 0, 0, 0
    imgImage.Stretch = True
    imgImage.Stretch = False
    
    If imgImage.Width > picHolder.Width Then
        flsbHoriz.value = 0
        flsbHoriz.Visible = True
        flsbHoriz.Max = imgImage.Width - picHolder.Width + 270
        flsbHoriz.LargeChange = flsbHoriz.Max / 5
    Else
        flsbHoriz.Visible = False
    End If
    
    If imgImage.Height > picHolder.Height Then
        flsbVert.value = 0
        flsbVert.Visible = True
        flsbVert.Max = imgImage.Height - picHolder.Height + 270
        flsbVert.LargeChange = flsbVert.Max / 5
    Else
        flsbVert.Visible = False
    End If
    
End Sub



Private Sub Text1_GotFocus()
Call HighlightBackground(Text1)
End Sub

Private Sub Text1_LostFocus()
Call NormalBackground(Text1)
End Sub

Private Sub Text2_GotFocus()
Call HighlightBackground(Text2)
End Sub

Private Sub Text2_LostFocus()
Call NormalBackground(Text2)
End Sub

Private Sub txt_Estimate_GotFocus()
Call HighlightBackground(txt_Estimate)
End Sub

'
Private Sub txt_Estimate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    txt_Estimate.MaxLength = 9
End Sub

Private Sub txt_Estimated1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    End Sub

Private Sub txt_Estimated2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
End Sub

Private Sub txt_Estimated3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    End Sub

Private Sub txt_Estimate_LostFocus()
Call NormalBackground(txt_Estimate)
End Sub

'assign stock master values

Private Sub txt_LongDescript_Change()
On Error Resume Next

    If txt_LongDescript <> deIms.rsSTOCKMASTER("stk_desc").originalVALUE Then
        chkDescHist.value = 1
    Else
       chkDescHist.value = 0
    End If
    
    If Err Then Err.Clear
    txt_LongDescript.MaxLength = 1500
    mFromDesc = True
End Sub

Private Sub txt_LongDescript_GotFocus()
Call HighlightBackground(txt_LongDescript)
End Sub

Private Sub txt_LongDescript_LostFocus()
cbo_StockNum.SetFocus
Call NormalBackground(txt_LongDescript)
End Sub

Private Sub txt_Maximum_GotFocus()
Call HighlightBackground(txt_Maximum)
End Sub

Private Sub txt_Maximum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    txt_Maximum.MaxLength = 6
End Sub

Private Sub txt_Maximum_LostFocus()
Call NormalBackground(txt_Maximum)
End Sub

Private Sub txt_Minimum_GotFocus()
Call HighlightBackground(txt_Minimum)
End Sub

Private Sub txt_Minimum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    txt_Minimum.MaxLength = 6
End Sub




Private Sub txt_Minimum_LostFocus()
Call NormalBackground(txt_Minimum)
End Sub

Private Sub txt_ShortDescript_Change()
txt_ShortDescript.MaxLength = 210
End Sub

Private Sub txt_ShortDescript_GotFocus()
Call HighlightBackground(txt_ShortDescript)
End Sub

Private Sub txt_ShortDescript_LostFocus()
Call NormalBackground(txt_ShortDescript)
End Sub

Private Sub txt_Standard_GotFocus()
Call HighlightBackground(txt_Standard)
End Sub

Private Sub txt_Standard_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0 To 7    ' KeyAscii is 0 - 7.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 8"
            KeyAscii = 0
        Case 8
            'backspace key
        Case 9 To 45    ' KeyAscii is 9 - 45.
            'Debug.Print KeyAscii; "Greater than or equal to 0 Less than 46"
            KeyAscii = 0
        Case 46
            'decimal
        Case 47
            KeyAscii = 0
        Case 48 To 57  ' KeyAscii is 48- 57.
            'Debug.Print "48-57"
        Case Is > 57 ' KeyAscii is > 57.
            'Debug.Print "> 57"
            KeyAscii = 0
        Case Else
            'Do Nothing FOOBAR
    End Select
    txt_Standard.MaxLength = 9
End Sub

'get file name

Public Function sGetFileName() As String
On Error Resume Next
    
    MDI_IMS.cmdDialog.ShowOpen

    If Err = 32755 Then
        Exit Function
    Else: sGetFileName = MDI_IMS.cmdDialog.FileName
    End If
End Function

'assign data to recordset

Private Sub AssignManufacturesDefault()
    deIms.rsGetStockManufacturer!stm_flag = True
    deIms.rsGetStockManufacturer!stm_npecode = deIms.NameSpace
    deIms.rsGetStockManufacturer!stm_stcknumb = stock!stk_stcknumb
End Sub

'assign data to recordset

Private Sub AddValuesToStockMan()
On Error Resume Next

    If stock.EditMode = adEditAdd Then
    
        If deIms.rsGetStockManufacturer.RecordCount = 0 Then Exit Sub
        
        deIms.rsGetStockManufacturer.MoveFirst
        Do Until deIms.rsGetStockManufacturer.EOF
            deIms.rsGetStockManufacturer!stm_npecode = deIms.NameSpace
            deIms.rsGetStockManufacturer!stm_stcknumb = stock!stk_stcknumb
            deIms.rsGetStockManufacturer.MoveNext
        Loop
    End If
    
    If Err Then Err.Clear
End Sub

'call function to get unit recordset

Public Function OpenUnit() As ADODB.Recordset
    If deIms.rsGetUnit.State And adStateOpen Then
        Set OpenUnit = deIms.rsGetUnit.Clone
    Else
        deIms.Getunit (deIms.NameSpace)
        Set OpenUnit = deIms.rsGetUnit.Clone
        deIms.rsGetUnit.Close
    End If
    
End Function

'SQL statement to get unit recordset

Private Sub Getunit()
Dim Rs As ADODB.Recordset
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        
        .CommandText = " select uni_desc from unit"
        .CommandText = .CommandText & "where uni_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " and uni_actvflag = 1"
        Set Rs = .Execute
    End With
    
    If Rs.RecordCount = 0 Then GoTo clearup
       
     Rs.MoveFirst
        Do While ((Not Rs.EOF))
'            dcboPrimUnit. rs!uni_desc
            Rs.MoveNext
        Loop


clearup:
    Rs.Close
    Set Rs = Nothing
    Set cmd = Nothing
End Sub

Private Sub txt_Standard_LostFocus()
Call NormalBackground(txt_Standard)
End Sub

Private Sub txtManSpecs_Change()
'txtManSpecs.MaxLength = 16
End Sub

Private Sub txtManSpecs_GotFocus()
Call HighlightBackground(txtManSpecs)
End Sub

Private Sub txtManSpecs_LostFocus()
Call NormalBackground(txtManSpecs)
End Sub

Public Function ISAValidUnit(Unit As String) As Boolean
ISAValidUnit = False
On Error GoTo Handler
Unit = Trim$(Unit)



Dim RsActUnit As ADODB.Recordset
Set RsActUnit = OpenUnit

RsActUnit.MoveFirst
RsActUnit.Find "UNI_CODE='" & Unit & "'", , adSearchForward

If RsActUnit.AbsolutePosition <> adPosEOF Then ISAValidUnit = True
 RsActUnit.Close
 Set RsActUnit = Nothing
Exit Function
Handler:
   MsgBox "Errors Occurred in processing the unit." & vbCrLf & "Error description -- " & Err.Description
   Err.Clear
   
End Function

Public Function SetDatasourceForUnits() As Boolean
On Error GoTo Handler

    If deIms.rsUNIT.State And adStateOpen Then
'        Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
        'Set dcboPrimUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
     
     
        Set dcboSecUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
'        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)

    Else
        Call deIms.Unit(deIms.NameSpace)
'        Set dcboSecUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
'        Set dcboPrimUnit.RowSource = deIms.rsUNIT.Clone(adLockReadOnly)
        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)

        Set dcboSecUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
'        Set dcboPrimUnit.DataSourceList = deIms.rsUNIT.Clone(adLockReadOnly)
        

        deIms.rsUNIT.Close
    End If
    
Exit Function

Handler:
  MsgBox "Errors Occurred while trying to set the datasource of the Two Units combos boxes." & vbCrLf & "Error Description -- " & Err.Description
  Err.Clear
  
End Function

Public Function validateFromTable(table As String, Code As String, desc As String) As Boolean

Dim Rs As ADODB.Recordset
Dim query As String

validateFromTable = False

Set Rs = New ADODB.Recordset

table = UCase(Trim$(table))
desc = Trim$(desc)
Code = Trim$(Code)

    Select Case table

       Case "CATEGORY"
         
          query = "select count(*) countIt from category where cate_catecode = '" & Code & "' and cate_catename='" & desc & "' and cate_npecode='" & deIms.NameSpace & "' and cate_actvflag=1"
       
       Case "UNIT"
    
          query = "select count(*) countIt from unit where uni_code = '" & Code & "' and uni_desc='" & desc & "' and uni_npecode='" & deIms.NameSpace & "' and uni_actvflag=1"
  
    End Select
    

    Rs.Source = query
    Rs.ActiveConnection = deIms.cnIms
    Rs.Open
    
    If Rs!countIt > 0 Then
       validateFromTable = True
    End If
    
    

End Function

