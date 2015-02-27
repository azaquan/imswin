VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_NewPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Order"
   ClientHeight    =   8145
   ClientLeft      =   2520
   ClientTop       =   2190
   ClientWidth     =   9015
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   9015
   Tag             =   "02020100"
   Begin VB.CommandButton CmdConvert 
      Caption         =   "Convert a transaction"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6705
      TabIndex        =   148
      Top             =   120
      Width           =   2175
   End
   Begin TabDlg.SSTab sst_PO 
      Height          =   7185
      Left            =   120
      TabIndex        =   135
      Top             =   120
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   758
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Transaction Order"
      TabPicture(0)   =   "NewPurchaseOrder.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_PO"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Purchase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frm_FromFQA"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "NewPurchaseOrder.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl_Recipients"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lbl_search"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dgRecipientList"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fra_FaxSelect"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd_Add"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt_Recipient"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "dgRecepients"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmdAddSupEmail"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "OptFax"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "OptEmail"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Line Items"
      TabPicture(2)   =   "NewPurchaseOrder.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_LineItem"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Fra_ToFqa"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fra_LI"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Remarks"
      TabPicture(3)   =   "NewPurchaseOrder.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtRemarks"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "CmdcopyLI(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Txt_RemNo"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Notes/Instructions"
      TabPicture(4)   =   "NewPurchaseOrder.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Addterms"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtClause"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "CmdcopyLI(2)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Txt_ClsNo"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.Frame fra_LI 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   -74880
         TabIndex        =   87
         Top             =   480
         Width           =   8520
         Begin VB.Label lbl_DocType 
            Alignment       =   1  'Right Justify
            Caption         =   "Document Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   173
            Top             =   135
            Width           =   1695
         End
         Begin VB.Label LblPOI_Doctype 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            TabIndex        =   172
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label lbl_PO2 
            Alignment       =   1  'Right Justify
            Caption         =   "Transaction #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   171
            Top             =   135
            Width           =   1785
         End
         Begin VB.Label LblPOi_PONUMB 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "poi_ponumb"
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1935
            TabIndex        =   170
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Frame Fra_ToFqa 
         Caption         =   "To FQA"
         Height          =   615
         Left            =   -74900
         TabIndex        =   163
         Top             =   6480
         Width           =   8480
         Begin VB.TextBox TxtToStocktypeFQA 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5685
            MaxLength       =   4
            TabIndex        =   52
            Top             =   240
            Width           =   450
         End
         Begin VB.TextBox TxtToCompanyFQA 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   885
            MaxLength       =   1
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   230
            Width           =   330
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBToLocationFQA 
            Bindings        =   "NewPurchaseOrder.frx":008C
            Height          =   315
            Left            =   2080
            TabIndex        =   50
            Top             =   240
            Width           =   975
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(0).Picture=   "NewPurchaseOrder.frx":00B8
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":00D4
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBtoUSChartFQA 
            Bindings        =   "NewPurchaseOrder.frx":00F0
            Height          =   315
            Left            =   3960
            TabIndex        =   51
            Top             =   225
            Width           =   1215
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(0).Picture=   "NewPurchaseOrder.frx":011C
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0138
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBToCamChartFQA 
            Bindings        =   "NewPurchaseOrder.frx":0154
            Height          =   315
            Left            =   7320
            TabIndex        =   53
            Top             =   225
            Width           =   1095
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0180
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":019C
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "US Chart#"
            Height          =   255
            Left            =   3120
            TabIndex        =   169
            Top             =   225
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Location"
            Height          =   255
            Left            =   1320
            TabIndex        =   168
            Top             =   225
            Width           =   735
         End
         Begin VB.Label LblType 
            Alignment       =   1  'Right Justify
            Caption         =   "Type"
            Height          =   255
            Index           =   2
            Left            =   5220
            TabIndex        =   167
            Top             =   225
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Cam. Chart #"
            Height          =   255
            Left            =   6240
            TabIndex        =   166
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Company"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   230
            Width           =   735
         End
      End
      Begin VB.Frame Frm_FromFQA 
         Caption         =   "From FQA"
         Height          =   615
         Left            =   240
         TabIndex        =   152
         Top             =   6360
         Width           =   8295
         Begin VB.TextBox TxtFromCompany 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   885
            MaxLength       =   1
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   240
            Width           =   330
         End
         Begin VB.TextBox TxtFromCamChart 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            MaxLength       =   8
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox TxtFromType 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            MaxLength       =   4
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   240
            Width           =   450
         End
         Begin VB.TextBox TxtFromUsChart 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3960
            MaxLength       =   9
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox TxtFromLocation 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Company"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Cam. Chart #"
            Height          =   255
            Left            =   6045
            TabIndex        =   160
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LblType 
            Alignment       =   1  'Right Justify
            Caption         =   "Type"
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   159
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Location"
            Height          =   255
            Left            =   1200
            TabIndex        =   158
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "US Chart#"
            Height          =   255
            Left            =   3045
            TabIndex        =   157
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton OptEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -72000
         TabIndex        =   27
         Top             =   3000
         Width           =   735
      End
      Begin VB.OptionButton OptFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   -72840
         TabIndex        =   26
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   -72840
         MaxLength       =   60
         TabIndex        =   30
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CommandButton CmdAddSupEmail 
         Caption         =   "Add Supplier Email"
         Height          =   645
         Left            =   -74760
         TabIndex        =   33
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Txt_ClsNo 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Left            =   -67440
         TabIndex        =   145
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Txt_RemNo 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Left            =   -67440
         TabIndex        =   144
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame fra_Purchase 
         ClipControls    =   0   'False
         Height          =   5460
         Left            =   240
         TabIndex        =   66
         Top             =   890
         Width           =   8295
         Begin VB.CheckBox chk_USExportH 
            Caption         =   "US Export"
            Height          =   255
            Left            =   4515
            TabIndex        =   19
            Top             =   1040
            Width           =   2175
         End
         Begin VB.CheckBox chk_FrmStkMst 
            Caption         =   "From Stock Master"
            Height          =   285
            Left            =   4560
            TabIndex        =   23
            Top             =   4800
            Width           =   3585
         End
         Begin VB.CheckBox chk_FreightFard 
            Caption         =   "Freight Forwarder Receipt Mandatory"
            Height          =   195
            Left            =   4515
            TabIndex        =   15
            Top             =   160
            Width           =   3705
         End
         Begin VB.TextBox Txt_supContaName 
            Height          =   315
            Left            =   1920
            MaxLength       =   35
            TabIndex        =   8
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Txt_supContaPh 
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   9
            Top             =   3450
            Width           =   2295
         End
         Begin VB.CheckBox chk_ConfirmingOrder 
            Caption         =   "Confirming Order"
            DataField       =   "po_confordr"
            DataMember      =   "PO"
            Height          =   288
            Left            =   4515
            TabIndex        =   17
            Top             =   720
            Width           =   1815
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboDelivery 
            Bindings        =   "NewPurchaseOrder.frx":01B8
            Height          =   315
            Left            =   6480
            TabIndex        =   24
            Top             =   5085
            Width           =   1665
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":01E4
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0200
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   1984
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3889
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSoledbSupplier 
            Bindings        =   "NewPurchaseOrder.frx":021C
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Top             =   2790
            Width           =   2295
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
            FieldSeparator  =   ";"
            ForeColorEven   =   8388608
            BackColorOdd    =   16771818
            RowHeight       =   423
            Columns.Count   =   5
            Columns(0).Width=   3200
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "sup_code"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "sup_name"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Caption=   "City"
            Columns(2).Name =   "sup_city"
            Columns(2).CaptionAlignment=   0
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Caption=   "Phone Number"
            Columns(3).Name =   "sup_phonnumb"
            Columns(3).CaptionAlignment=   0
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Caption=   "Fax NUmber"
            Columns(4).Name =   "sup_faxnumb"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.TextBox txtSite 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6480
            TabIndex        =   132
            Top             =   4470
            Width           =   1665
         End
         Begin VB.CheckBox chk_Forwarder 
            Caption         =   "Forwarder"
            Height          =   288
            Left            =   6360
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox txt_ChargeTo 
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   3
            Top             =   810
            Width           =   2295
         End
         Begin VB.TextBox txt_Buyer 
            BackColor       =   &H00FFFFC0&
            CausesValidation=   0   'False
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   1470
            Width           =   2295
         End
         Begin VB.CheckBox chk_Requ 
            Caption         =   "Print Required date for each LI ? Y/N"
            Height          =   288
            Left            =   4515
            TabIndex        =   16
            Top             =   400
            Width           =   3705
         End
         Begin VB.Frame fra_Stat 
            BackColor       =   &H8000000A&
            Caption         =   "Status"
            Enabled         =   0   'False
            Height          =   1620
            Left            =   4560
            TabIndex        =   67
            Top             =   1440
            Width           =   3600
            Begin VB.Label LblStatus7 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1260
               TabIndex        =   139
               Top             =   1200
               Width           =   2250
            End
            Begin VB.Label LblStatus6 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1260
               TabIndex        =   138
               Top             =   880
               Width           =   2250
            End
            Begin VB.Label LblStatus5 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1260
               TabIndex        =   137
               Top             =   560
               Width           =   2250
            End
            Begin VB.Label LblStatus4 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1260
               TabIndex        =   136
               Top             =   240
               Width           =   2250
            End
            Begin VB.Label lbl_Shipping 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               Caption         =   "Shipping"
               Height          =   225
               Left            =   105
               TabIndex        =   71
               Top             =   885
               Width           =   1080
            End
            Begin VB.Label lbl_Delivery 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               Caption         =   "Delivery"
               Height          =   225
               Left            =   105
               TabIndex        =   70
               Top             =   585
               Width           =   1080
            End
            Begin VB.Label lbl_Status 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               Caption         =   "PO"
               Height          =   225
               Left            =   120
               TabIndex        =   69
               Top             =   300
               Width           =   1080
            End
            Begin VB.Label lbl_Inventory 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               Caption         =   "Inventory"
               Height          =   225
               Left            =   105
               TabIndex        =   68
               Top             =   1215
               Width           =   1080
            End
         End
         Begin MSComCtl2.DTPicker dtpRequestedDate 
            Bindings        =   "NewPurchaseOrder.frx":0227
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   6480
            TabIndex        =   21
            Top             =   3795
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55050243
            CurrentDate     =   36402
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboShipper 
            Bindings        =   "NewPurchaseOrder.frx":024D
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   480
            Width           =   3015
            DataFieldList   =   "Column 0"
            AutoRestore     =   0   'False
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0283
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":029F
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2434
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3995
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   5318
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
            Bindings        =   "NewPurchaseOrder.frx":02BB
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Top             =   4770
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":02E7
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0303
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   5292
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 1"
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 0"
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBPriority 
            Bindings        =   "NewPurchaseOrder.frx":031F
            Height          =   315
            Left            =   1920
            TabIndex        =   4
            Top             =   1140
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":034B
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0367
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2275
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4101
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBOriginator 
            Bindings        =   "NewPurchaseOrder.frx":0383
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   1800
            Width           =   2295
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":03AF
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":03CB
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBToBeUsedFor 
            Bindings        =   "NewPurchaseOrder.frx":03E7
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   2460
            Width           =   2295
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0413
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":042F
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
            Bindings        =   "NewPurchaseOrder.frx":044B
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            Top             =   3780
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0477
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0493
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2090
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCompany 
            Bindings        =   "NewPurchaseOrder.frx":04AF
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Top             =   4110
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":04DB
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":04F7
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3360
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4763
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBInvLocation 
            Bindings        =   "NewPurchaseOrder.frx":0513
            Height          =   315
            Left            =   1920
            TabIndex        =   12
            Top             =   4440
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":053F
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":055B
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2381
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4524
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBShipTo 
            Bindings        =   "NewPurchaseOrder.frx":0577
            Height          =   315
            Left            =   5760
            TabIndex        =   22
            Top             =   4125
            Width           =   2385
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":05A3
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":05BF
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2064
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4128
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4207
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin MSComCtl2.DTPicker DTPicker_poDate 
            Bindings        =   "NewPurchaseOrder.frx":05DB
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   6480
            TabIndex        =   20
            Top             =   3120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55050243
            CurrentDate     =   36402
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOledbSrvCode 
            Bindings        =   "NewPurchaseOrder.frx":0601
            Height          =   315
            Left            =   1920
            TabIndex        =   14
            Top             =   5100
            Width           =   2295
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":062D
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0649
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2275
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4101
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Service Utility"
            Height          =   225
            Left            =   90
            TabIndex        =   147
            Top             =   5100
            Width           =   1725
         End
         Begin VB.Label DTPicker_poDate1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6840
            TabIndex        =   143
            Top             =   1080
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Ph"
            Height          =   255
            Left            =   90
            TabIndex        =   141
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Name"
            Height          =   255
            Left            =   90
            TabIndex        =   140
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Company"
            Height          =   225
            Left            =   90
            TabIndex        =   134
            Top             =   4125
            Width           =   1710
         End
         Begin VB.Label lbl_InvLoc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Location"
            Height          =   225
            Left            =   90
            TabIndex        =   133
            Top             =   4455
            Width           =   1710
         End
         Begin VB.Label lbl_Revision 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Number"
            Height          =   225
            Left            =   90
            TabIndex        =   131
            Top             =   130
            Width           =   1245
         End
         Begin VB.Label LblRevNumb 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "po_revinumb"
            DataMember      =   "PO"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1440
            TabIndex        =   130
            Top             =   135
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Term"
            Height          =   225
            Left            =   4320
            TabIndex        =   125
            Top             =   5085
            Width           =   2085
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T && C"
            Height          =   225
            Left            =   90
            TabIndex        =   124
            Top             =   4800
            Width           =   1725
         End
         Begin VB.Label LblRevDate 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   3240
            TabIndex        =   122
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label LblDateSent 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6480
            TabIndex        =   121
            Top             =   3450
            Width           =   1665
         End
         Begin VB.Label LblAppBy 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   120
            Top             =   2130
            Width           =   2295
         End
         Begin VB.Label lbl_Supplier 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            Height          =   225
            Left            =   90
            TabIndex        =   86
            Top             =   2745
            Width           =   1725
         End
         Begin VB.Label lbl_ToBe 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To Be Used For"
            Height          =   225
            Left            =   90
            TabIndex        =   85
            Top             =   2415
            Width           =   1725
         End
         Begin VB.Label lbl_Shipper 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Shipper"
            Height          =   225
            Left            =   90
            TabIndex        =   84
            Top             =   465
            Width           =   1005
         End
         Begin VB.Label lbl_Currency 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            Height          =   225
            Left            =   90
            TabIndex        =   83
            Top             =   3795
            Width           =   1665
         End
         Begin VB.Label lbl_DelivDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Required"
            Height          =   225
            Left            =   4200
            TabIndex        =   82
            Top             =   3795
            Width           =   2205
         End
         Begin VB.Label lbl_RevisionDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Date"
            Height          =   225
            Left            =   2160
            TabIndex        =   81
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label lbl_RequDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PO Creation Date"
            Height          =   225
            Left            =   4320
            TabIndex        =   80
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label lbl_ShipTo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ship To"
            Height          =   225
            Left            =   4320
            TabIndex        =   79
            Top             =   4125
            Width           =   1335
         End
         Begin VB.Label lbl_ChargeTo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Charge To/A.F.E"
            Height          =   285
            Left            =   90
            TabIndex        =   78
            Top             =   840
            Width           =   1725
         End
         Begin VB.Label lbl_Priority 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Shipping Mode"
            Height          =   225
            Left            =   90
            TabIndex        =   77
            Top             =   1125
            Width           =   1725
         End
         Begin VB.Label lbl_Buyer 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Buyer"
            Height          =   225
            Left            =   90
            TabIndex        =   76
            Top             =   1455
            Width           =   1725
         End
         Begin VB.Label lbl_Originator 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Originator"
            Height          =   225
            Left            =   90
            TabIndex        =   75
            Top             =   1785
            Width           =   1725
         End
         Begin VB.Label lbl_DateSent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Approved"
            Height          =   225
            Left            =   4320
            TabIndex        =   74
            Top             =   3450
            Width           =   2055
         End
         Begin VB.Label lbl_Site 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
            Height          =   225
            Left            =   4320
            TabIndex        =   73
            Top             =   4470
            Width           =   2010
         End
         Begin VB.Label lbl_ApprovedBy 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
            Height          =   225
            Left            =   90
            TabIndex        =   72
            Top             =   2115
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   2
         Left            =   -74760
         TabIndex        =   56
         Top             =   550
         Width           =   2175
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   1
         Left            =   -74760
         TabIndex        =   59
         Top             =   550
         Width           =   2415
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74760
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fra_LineItem 
         BorderStyle     =   0  'None
         Height          =   6090
         Left            =   -74880
         TabIndex        =   88
         Top             =   1080
         Width           =   8520
         Begin VB.CheckBox Chk_license 
            Caption         =   "License required"
            Height          =   255
            Left            =   1080
            TabIndex        =   41
            Top             =   2400
            Width           =   2655
         End
         Begin VB.CheckBox chk_usexportLI 
            Caption         =   "US Export"
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   330
            Width           =   2025
         End
         Begin VB.CommandButton CmdAssignFQA 
            Caption         =   "Assign FQA to all"
            Height          =   305
            Left            =   0
            TabIndex        =   49
            Top             =   5040
            Width           =   1935
         End
         Begin VB.CommandButton CmdcopyLI 
            Caption         =   "Copy From ...."
            Height          =   305
            Index           =   0
            Left            =   0
            TabIndex        =   46
            Top             =   4080
            Width           =   1935
         End
         Begin VB.TextBox txt_SerialNum 
            DataField       =   "poi_serlnumb"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   5760
            MaxLength       =   25
            TabIndex        =   42
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txt_remk 
            DataField       =   "poi_remk"
            DataMember      =   "POITEM"
            DataSource      =   "deIms"
            Height          =   675
            Left            =   2040
            MaxLength       =   12000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   4755
            Width           =   6420
         End
         Begin VB.TextBox txt_Descript 
            DataField       =   "poi_desc"
            DataMember      =   "POITEM"
            Height          =   675
            Left            =   2040
            MaxLength       =   1500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   4035
            Width           =   6420
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboRequisition 
            Bindings        =   "NewPurchaseOrder.frx":0665
            Height          =   315
            Left            =   5760
            TabIndex        =   36
            Top             =   120
            Width           =   1575
            DataFieldList   =   "Column 0"
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
            BorderStyle     =   0
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0670
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":068C
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   5
            Columns(0).Width=   2672
            Columns(0).Caption=   "Number"
            Columns(0).Name =   "Number"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   3836
            Columns(1).Caption=   "Type"
            Columns(1).Name =   "Type"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            Columns(2).Width=   1005
            Columns(2).Caption=   "Item"
            Columns(2).Name =   "Item"
            Columns(2).DataField=   "Column 2"
            Columns(2).FieldLen=   256
            Columns(3).Width=   5292
            Columns(3).Caption=   "Description"
            Columns(3).Name =   "Description"
            Columns(3).DataField=   "Column 3"
            Columns(3).FieldLen=   256
            Columns(4).Width=   1693
            Columns(4).Caption=   "Qty"
            Columns(4).Name =   "Qty"
            Columns(4).DataField=   "Column 4"
            Columns(4).FieldLen=   256
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.TextBox txt_LI 
            BackColor       =   &H00FFFFC0&
            DataField       =   "poi_liitnumb"
            DataMember      =   "POITEM"
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   119
            Top             =   0
            Width           =   435
         End
         Begin MSComCtl2.DTPicker DTP_Required 
            Bindings        =   "NewPurchaseOrder.frx":06A8
            DataField       =   "poi_liitreqddate"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   5760
            TabIndex        =   54
            Top             =   450
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55050243
            CurrentDate     =   36405
         End
         Begin VB.TextBox txt_TotalLIs 
            BackColor       =   &H00FFFFC0&
            DataField       =   "LCount"
            DataMember      =   "LineItemCount"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3000
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   0
            Width           =   420
         End
         Begin VB.TextBox txt_AFE 
            DataField       =   "poi_afe"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   38
            Top             =   1320
            Width           =   2190
         End
         Begin VB.Frame fra_Status 
            Caption         =   "Status"
            Enabled         =   0   'False
            Height          =   1770
            Left            =   3960
            TabIndex        =   95
            Top             =   1125
            Width           =   4530
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "NewPurchaseOrder.frx":06D9
               Height          =   315
               Index           =   0
               Left            =   1320
               TabIndex        =   113
               Top             =   270
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "NewPurchaseOrder.frx":0700
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   114
               Top             =   600
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "NewPurchaseOrder.frx":0727
               Height          =   315
               Index           =   2
               Left            =   1320
               TabIndex        =   115
               Top             =   930
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "NewPurchaseOrder.frx":074E
               Height          =   315
               Index           =   3
               Left            =   1320
               TabIndex        =   116
               Top             =   1260
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.Label lbl_StatInventory 
               Alignment       =   1  'Right Justify
               Caption         =   "Inventory"
               Height          =   225
               Left            =   105
               TabIndex        =   99
               Top             =   1245
               Width           =   1170
            End
            Begin VB.Label lbl_StatShipping 
               Alignment       =   1  'Right Justify
               Caption         =   "Shipping"
               Height          =   225
               Left            =   105
               TabIndex        =   98
               Top             =   930
               Width           =   1170
            End
            Begin VB.Label lbl_StatDelivery 
               Alignment       =   1  'Right Justify
               Caption         =   "Delivery"
               Height          =   225
               Left            =   105
               TabIndex        =   97
               Top             =   615
               Width           =   1170
            End
            Begin VB.Label lbl_StatItem 
               Alignment       =   1  'Right Justify
               Caption         =   "Item"
               Height          =   225
               Left            =   105
               TabIndex        =   96
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Frame fra_Quantity 
            Height          =   1020
            Left            =   0
            TabIndex        =   63
            Top             =   2955
            Width           =   8480
            Begin VB.TextBox txt_Price 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   315
               Left            =   6840
               TabIndex        =   45
               Top             =   240
               Width           =   1275
            End
            Begin VB.TextBox txt_Total 
               BackColor       =   &H00FFFFC0&
               DataField       =   "poi_totaprice"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               DataMember      =   "POITEM"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   149
               Top             =   600
               Width           =   1275
            End
            Begin VB.TextBox txt_Requested 
               DataField       =   "poi_primreqdqty"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               DataMember      =   "POITEM"
               Height          =   315
               Left            =   1800
               TabIndex        =   43
               Top             =   240
               Width           =   945
            End
            Begin VB.TextBox txt_Inventory2 
               BackColor       =   &H00FFFFC0&
               DataField       =   "poi_qtyinvt"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               DataMember      =   "POITEM"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   4800
               TabIndex        =   118
               Top             =   600
               Width           =   720
            End
            Begin VB.TextBox txt_Shipped 
               BackColor       =   &H00FFFFC0&
               DataField       =   "poi_qtyship"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               DataMember      =   "POITEM"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3000
               TabIndex        =   117
               Top             =   600
               Width           =   720
            End
            Begin VB.TextBox txt_Delivered 
               BackColor       =   &H00FFFFC0&
               DataField       =   "poi_qtydlvd"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               DataMember      =   "POITEM"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   1080
               TabIndex        =   123
               Top             =   600
               Width           =   720
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBUnit 
               Bindings        =   "NewPurchaseOrder.frx":0775
               Height          =   315
               Left            =   4320
               TabIndex        =   44
               Top             =   240
               Width           =   1215
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
               stylesets(0).Picture=   "NewPurchaseOrder.frx":07A1
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
               stylesets(1).Picture=   "NewPurchaseOrder.frx":07BD
               stylesets(1).AlignmentText=   1
               HeadFont3D      =   4
               DefColWidth     =   5292
               BeveColorScheme =   1
               ForeColorEven   =   8388608
               BackColorEven   =   16771818
               BackColorOdd    =   16777215
               RowHeight       =   423
               Columns.Count   =   2
               Columns(0).Width=   1799
               Columns(0).Caption=   "Code"
               Columns(0).Name =   "Code"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               Columns(1).Width=   3493
               Columns(1).Caption=   "Name"
               Columns(1).Name =   "Name"
               Columns(1).DataField=   "Column 1"
               Columns(1).DataType=   8
               Columns(1).FieldLen=   256
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   -2147483643
               DataFieldToDisplay=   "Column 0"
            End
            Begin VB.Label lbl_Cost 
               Alignment       =   1  'Right Justify
               Caption         =   "Unit Price"
               Height          =   225
               Left            =   5640
               TabIndex        =   151
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label lbl_Total 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Price"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   5520
               TabIndex        =   150
               Top             =   600
               Width           =   1275
            End
            Begin VB.Label lbl_Delivered 
               Alignment       =   1  'Right Justify
               Caption         =   "Delivered"
               Height          =   225
               Left            =   120
               TabIndex        =   93
               Top             =   600
               Width           =   930
            End
            Begin VB.Label lbl_Shipped 
               Alignment       =   1  'Right Justify
               Caption         =   "Shipped"
               Height          =   225
               Left            =   1920
               TabIndex        =   92
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl_Requested 
               Alignment       =   1  'Right Justify
               Caption         =   "Quantity Required"
               Height          =   225
               Left            =   120
               TabIndex        =   91
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lbl_Unit 
               Alignment       =   1  'Right Justify
               Caption         =   "Purchase Unit"
               Height          =   195
               Left            =   2880
               TabIndex        =   90
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label lbl_Inventory2 
               Alignment       =   1  'Right Justify
               Caption         =   "Inventory"
               Height          =   225
               Left            =   3840
               TabIndex        =   89
               Top             =   600
               Width           =   855
            End
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboManNumber 
            Bindings        =   "NewPurchaseOrder.frx":07D9
            Height          =   315
            Left            =   2040
            TabIndex        =   37
            Top             =   960
            Width           =   1830
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":07E4
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0800
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3440
            Columns(0).Caption=   "Manufacturer"
            Columns(0).Name =   "Part Number"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   4683
            Columns(1).Caption=   "Part Number"
            Columns(1).Name =   "Manufacturer"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCustCategory 
            Bindings        =   "NewPurchaseOrder.frx":081C
            Height          =   315
            Left            =   2040
            TabIndex        =   39
            Top             =   1680
            Width           =   1830
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0848
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0864
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   6879
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCommoditty 
            Bindings        =   "NewPurchaseOrder.frx":0880
            Height          =   315
            Left            =   2040
            TabIndex        =   35
            Top             =   600
            Width           =   1830
            ListAutoValidate=   0   'False
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
            FieldSeparator  =   "(Space)"
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":08AC
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":08C8
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            HeadStyleSet    =   "RowFont"
            StyleSet        =   "RowFont"
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   291
            Columns.Count   =   0
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSoleEccnno 
            Bindings        =   "NewPurchaseOrder.frx":08E4
            Height          =   315
            Left            =   1080
            TabIndex        =   40
            Top             =   2040
            Width           =   2775
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
            FieldSeparator  =   "(Space)"
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0910
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":092C
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            HeadStyleSet    =   "RowFont"
            StyleSet        =   "RowFont"
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   291
            Columns.Count   =   0
            _ExtentX        =   4895
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleSourceofinfo 
            Bindings        =   "NewPurchaseOrder.frx":0948
            Height          =   315
            Left            =   2040
            TabIndex        =   176
            Top             =   2640
            Width           =   1815
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
            FieldSeparator  =   "(Space)"
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0974
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0990
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            HeadStyleSet    =   "RowFont"
            StyleSet        =   "RowFont"
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   291
            Columns.Count   =   0
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Label lblsourceofinfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Source Of Info"
            Height          =   225
            Left            =   0
            TabIndex        =   175
            Top             =   2685
            Width           =   1920
         End
         Begin VB.Label lblEccn 
            Alignment       =   1  'Right Justify
            Caption         =   "Eccn#"
            Height          =   225
            Left            =   0
            TabIndex        =   174
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label lbl_PartNum 
            Alignment       =   1  'Right Justify
            Caption         =   "Manufacturer P/N"
            Height          =   225
            Left            =   0
            TabIndex        =   129
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lbl_SerialNum 
            Alignment       =   1  'Right Justify
            Caption         =   "Serial Number"
            Height          =   225
            Left            =   3960
            TabIndex        =   128
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Remarks"
            Height          =   255
            Left            =   0
            TabIndex        =   127
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label lblReqLineitem 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "poi_requliitnumb"
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7440
            TabIndex        =   126
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lbl_Of 
            Alignment       =   1  'Right Justify
            Caption         =   "of"
            Height          =   225
            Left            =   2520
            TabIndex        =   107
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lbl_Description 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            Height          =   225
            Left            =   120
            TabIndex        =   106
            Top             =   4440
            Width           =   1815
         End
         Begin VB.Label lbl_Item 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   225
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lbl_Commodity 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity"
            Height          =   225
            Left            =   0
            TabIndex        =   104
            Top             =   645
            Width           =   1980
         End
         Begin VB.Label lbl_AFE 
            Alignment       =   1  'Right Justify
            Caption         =   "Charge To/A.F.E"
            Height          =   225
            Left            =   0
            TabIndex        =   103
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lbl_Custom 
            Alignment       =   1  'Right Justify
            Caption         =   "Customs Category"
            Height          =   225
            Left            =   0
            TabIndex        =   102
            Top             =   1680
            Width           =   1965
         End
         Begin VB.Label lbl_Requisition 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Req/Quot/BO#"
            Height          =   195
            Left            =   3960
            TabIndex        =   101
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label lbl_RequDate2 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Required"
            Height          =   225
            Left            =   3960
            TabIndex        =   100
            Top             =   450
            Width           =   1680
         End
      End
      Begin VB.TextBox txtClause 
         DataField       =   "poc_clau"
         DataMember      =   "POCLAUSE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   -74760
         MaxLength       =   12000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   1020
         Width           =   8300
      End
      Begin VB.TextBox txtRemarks 
         DataField       =   "por_remk"
         DataMember      =   "POREM"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   -74760
         MaxLength       =   12000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   1020
         Width           =   8295
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   2055
         Left            =   -72840
         TabIndex        =   61
         Top             =   4080
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   3625
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
      Begin VB.TextBox txt_Recipient 
         Height          =   288
         Left            =   -72840
         MaxLength       =   60
         TabIndex        =   29
         Top             =   3360
         Width           =   6150
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74760
         TabIndex        =   28
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmd_Addterms 
         Caption         =   "Add Clause"
         Height          =   288
         Left            =   -72480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   550
         Width           =   2415
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   1170
         Left            =   -74760
         TabIndex        =   62
         Top             =   4000
         Width           =   1755
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1515
         End
      End
      Begin VB.Frame fra_PO 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   225
         TabIndex        =   108
         Top             =   450
         Width           =   8430
         Begin VB.OptionButton showAll 
            Caption         =   "2 yrs only"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   178
            Top             =   220
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton showAll 
            Caption         =   "Show all"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   177
            Top             =   20
            Width           =   1335
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssOleDbPO 
            Bindings        =   "NewPurchaseOrder.frx":09AC
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   75
            Width           =   1815
            DataFieldList   =   "Column 0"
            ListAutoValidate=   0   'False
            MinDropDownItems=   8
            _Version        =   196617
            Cols            =   1
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":09D8
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":09F4
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            RowSelectionStyle=   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBDocType 
            Bindings        =   "NewPurchaseOrder.frx":0A10
            Height          =   315
            Left            =   5880
            TabIndex        =   1
            Top             =   75
            Width           =   2415
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0A3C
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0A58
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   2275
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3043
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label lbl_DocumentType 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4680
            TabIndex        =   110
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label lbl_Purchase 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   109
            Top             =   120
            Width           =   1245
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
         Height          =   2085
         Left            =   -72840
         TabIndex        =   60
         Top             =   660
         Width           =   6135
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         FieldSeparator  =   ";"
         Col.Count       =   3
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
         stylesets(0).Picture=   "NewPurchaseOrder.frx":0A74
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
         stylesets(1).Picture=   "NewPurchaseOrder.frx":0A90
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
         Columns.Count   =   3
         Columns(0).Width=   5292
         Columns(0).Caption=   "Column 0"
         Columns(0).Name =   "Column 0"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   5292
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "supplierCode"
         Columns(1).Name =   "supplierCode"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   5292
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Supplier"
         Columns(2).Name =   "Supplier"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10821
         _ExtentY        =   3678
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
      Begin VB.Label Lbl_search 
         Alignment       =   1  'Right Justify
         Caption         =   "Search by name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   146
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lbl_Recipients 
         Alignment       =   1  'Right Justify
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74715
         TabIndex        =   111
         Top             =   570
         Width           =   1740
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -66720
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   7560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailVisible    =   -1  'True
      FirstEnabled    =   0   'False
      LastEnabled     =   0   'False
      NewEnabled      =   -1  'True
      NextEnabled     =   0   'False
      PreviousEnabled =   0   'False
      AllowDelete     =   0   'False
      DeleteVisible   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      SaveToolTipText =   "Save changes made to the current record"
      CancelToolTipText=   "Undo the changes made to the current Screen"
      EditToolTipText =   "Allows you to make modification"
   End
   Begin VB.Label LblCompanyCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Code"
      Height          =   375
      Left            =   4680
      TabIndex        =   142
      Top             =   7320
      Visible         =   0   'False
      Width           =   4215
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
      Left            =   4440
      TabIndex        =   65
      Top             =   7560
      Width           =   4500
   End
End
Attribute VB_Name = "frm_NewPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mainpo As New IMSPODLL.Mainpo
Dim Poheader As IMSPODLL.Poheader
Dim WithEvents PoItem As IMSPODLL.POITEMS
Attribute PoItem.VB_VarHelpID = -1
Dim PORemark As IMSPODLL.POREMARKS
Dim POClause As IMSPODLL.POClauses
Dim PoReceipients As IMSPODLL.PoReceipients
Dim POFqa As IMSPODLL.FQA
Dim CheckLoad As Boolean
Dim CheckErrors As Boolean
Dim CheckIfCombosLoaded As Boolean
Dim IntiClass As InitialValuesPOheader
Dim FNameSpace  As String
Dim FormMode As FormMode
Dim mIsPoheaderCombosLoaded As Boolean
Dim mIsPoNumbComboLoaded As Boolean
Dim mIsPoItemsComboLoaded As Boolean
Dim mSaveToPoRevision As Boolean
Dim mIsPoItemCombosLoaded As Boolean
Dim mIsDocTypeLoaded As Boolean
Dim mIsInvLocationLoaded As Boolean
Dim lookups As IMSPODLL.lookups
Dim GRsDoctype As ADODB.Recordset
Dim RsUNits As ADODB.Recordset
Dim objUnits As IMSPODLL.PoUnits
Dim GPOnumb As String
Dim mLoadMode As LoadMode
Dim rsDOCTYPE As ADODB.Recordset
Dim IsThisADifferentPO As Boolean
Dim mCheckPoFields As Boolean
Dim mCheckLIFields As Boolean
Dim MCheckClause As Boolean
Dim mCheckRemarks As Boolean
Dim mIsPoHeaderRsetsInit As Boolean
Dim msg1 As String
Dim msg2 As String
Dim RsEmailFax As ADODB.Recordset
Dim mSelection As Boolean
Dim WithEvents st As frm_ShipTerms
Attribute st.VB_VarHelpID = -1
Dim WithEvents comsearch As frm_StockSearch
Attribute comsearch.VB_VarHelpID = -1
Dim rowguid, locked As Boolean       'jawdat
Dim GToFQAComboLoaded As Boolean
Dim GRsEccnNo As ADODB.Recordset
Dim GRSSourceOfInfo As ADODB.Recordset

Dim newSupplier As Boolean  'JCG 2008/01/14
Dim suppContacts() As String
Private Sub getSupplierContacts()
On Error GoTo getError
            '-----JCG 2008/1/20 to remove old recipients in case changing of supplier
            If newSupplier = False Then
                Dim i As Integer
                Dim cont
                dgRecipientList.MoveFirst
                For i = 0 To dgRecipientList.Rows - 1
                    For Each cont In suppContacts
                        If cont = dgRecipientList.Columns(0).Text Then
                            PoReceipients.MoveFirst
                            Do While Not PoReceipients.EOF
                                If PoReceipients.Receipient = cont Then
                                    PoReceipients.DeleteCurrentLI (dgRecipientList.Columns(0).Text)
                                    Exit Do
                                End If
                                PoReceipients.MoveNext
                            Loop
                            dgRecipientList.RemoveItem (i)
                        End If
                    Next cont
                    dgRecipientList.MoveNext
                Next
            End If
            
            Dim rs As ADODB.Recordset
            Dim sql As String
            sql = "SELECT * FROM SUPPLIERCONTACT WHERE sct_supcode='" + Trim(SSoledbSupplier.Tag) + "'"
            Set rs = New ADODB.Recordset
            Call rs.Open(sql, deIms.cnIms, adOpenForwardOnly, adLockReadOnly)
            rs.MoveFirst
            If rs.RecordCount > 0 Then ReDim suppContacts(rs.RecordCount)
            Dim RecipientName As String
            PoReceipients.MoveLast
            Dim n As Integer

            Do While Not rs.EOF
                If LTrim(RecipientName) <> "" Then 'JCG 2008/8/16
                    If rs!sct_email = "" Then
                        RecipientName = rs!sct_fax
                    Else
                        RecipientName = rs!sct_email
                    End If
                    dgRecipientList.AddItem RecipientName
                    dgRecipientList.MoveLast
                    dgRecipientList.Columns(1).value = "supplierContact"
                    suppContacts(n) = RecipientName
                    With PoReceipients
                        .AddNew
                        .Linenumb = PoReceipients.Count
                        .Ponumb = Poheader.Ponumb
                        .Receipient = RecipientName
                        .NameSpace = deIms.NameSpace
                    End With
                End If ' JCG 2008/8/16
                rs.MoveNext
                n = n + 1
            Loop
            rs.Close
            Set rs = Nothing
            newSupplier = False
            
getError:
    Resume Next
    '---------------------
End Sub

Private Sub chk_ConfirmingOrder_GotFocus()
Call HighlightBackground(chk_ConfirmingOrder)
End Sub

Private Sub chk_ConfirmingOrder_LostFocus()
Call NormalBackground(chk_ConfirmingOrder)
End Sub

Private Sub chk_Forwarder_GotFocus()
Call HighlightBackground(chk_Forwarder)
End Sub

Private Sub chk_Forwarder_LostFocus()
Call NormalBackground(chk_Forwarder)
End Sub

Private Sub chk_FreightFard_GotFocus()
Call HighlightBackground(chk_FreightFard)
End Sub

Private Sub chk_FreightFard_LostFocus()
Call NormalBackground(chk_FreightFard)
End Sub

Private Sub chk_FrmStkMst_GotFocus()
Call HighlightBackground(chk_FrmStkMst)
End Sub

Private Sub chk_FrmStkMst_LostFocus()
Call NormalBackground(chk_FrmStkMst)
End Sub

Private Sub chk_Requ_GotFocus()
Call HighlightBackground(chk_Requ)
End Sub

Private Sub chk_Requ_LostFocus()
Call NormalBackground(chk_Requ)
End Sub

Private Sub chk_USExportH_Click()

If mLoadMode = NoLoadInProgress Then

    If FormMode <> mdvisualization Then
        
        If ConnInfo.Eccnactivate = Constyes And chk_USExportH.value = 0 Then
        
                chk_USExportH.value = 1
                MsgBox "System is configured for using Eccn, can not uncheck it.", vbInformation
                Exit Sub
                
        ElseIf ConnInfo.Eccnactivate = ConstOptional And chk_USExportH.value = 0 Then
            
                If MsgBox("This will uncheck the Eccn Flag on all the line items, Are you sure you want to do this?", vbInformation + vbYesNo) = vbYes Then
                
                   chk_USExportH.value = 0
                   UpdateAllLineItemsWithUSExport (chk_USExportH.value)
                   
                Else
                    mLoadMode = LoadingPOheader
                    chk_USExportH.value = 1
                    mLoadMode = NoLoadInProgress
                End If
                
                Exit Sub
                
        ElseIf ConnInfo.Eccnactivate = ConstOptional And chk_USExportH.value = 1 Then
            
                If MsgBox("This will check the Eccn Flag on all the line items, Are you sure you want to do this?", vbInformation + vbYesNo) = vbYes Then
                  
                   chk_USExportH.value = 1
                   UpdateAllLineItemsWithUSExport (chk_USExportH.value)
                   
                Else
                    mLoadMode = LoadingPOheader
                    chk_USExportH.value = 0
                    mLoadMode = NoLoadInProgress
                End If
                
                Exit Sub
                
        End If
        
        
        
     End If
        
End If

End Sub

Public Function UpdateAllLineItemsWithUSExport(UsExportValue As Boolean) As Boolean
Dim i As Integer
On Error GoTo ErrHand

            If PoItem Is Nothing Then Set PoItem = Mainpo.POITEMS
               
            If Trim$(PoItem.Ponumb) <> Poheader.Ponumb Then
                    GPOnumb = Poheader.Ponumb
                       
                    Call PoItem.Move(GPOnumb)
                    
            End If
                     
                     
             If PoItem.Count > 0 Then
                     
                    PoItem.MoveFirst
                    For i = 0 To PoItem.Count - 1
                    'Do While PoItem.EOF = False
                    
                       PoItem.usexport = UsExportValue
                       Call PoItem.MoveNext
                       
                       'Debug.Print PoItem.Linenumb
                       
                    Next
                    
                    PoItem.MoveFirst
                     
                    LoadFromPOITEM
                    POFqa.MoveLineTo (PoItem.Linenumb)
                    LoadFromTOFQA
                
                Else
                   'This means that there are no Line Items Corresponding to This PO
                     ClearAllPoLineItems
                     
                End If
            
Exit Function
ErrHand:

MsgBox "Errors Occurred while trying to update the line items with the new US Export value." & Err.Description, vbCritical
End Function
Private Sub chk_usexportLI_Click()

If mLoadMode = NoLoadInProgress Then

    If FormMode <> mdvisualization Then
    
        If chk_usexportLI.value = 1 And chk_USExportH.value = 0 Then
        
            If ConnInfo.Eccnactivate = Constyes Or ConnInfo.Eccnactivate = ConstOptional Then
            
                MsgBox "Please make sure that the US export on the PO Header is checked before checking the Line item's US Export.", vbInformation
                If chk_USExportH.value = False Then chk_usexportLI.value = False
            
            End If
        
        ElseIf chk_usexportLI.value = 0 And chk_USExportH.value = 1 Then
        
            chk_usexportLI.value = 1
            MsgBox "Can not Un-check the US Export flag at the line item level when the US Export flag at the header level is checked.", vbInformation
        
        End If
        
        If ConnInfo.Eccnactivate = Constyes And chk_usexportLI.value = 0 Then
                 chk_usexportLI.value = 1
                MsgBox "System is configured for using Eccn, can not uncheck it.", vbInformation
        
        
        End If
        
    End If

End If
End Sub

Private Sub cmd_Addterms_LostFocus()

If txtClause.Enabled = True Then
   txtClause.SetFocus
   Call HighlightBackground(txtClause)
End If
End Sub

Private Sub CmdAddSupEmail_Click()
On Error GoTo Handler
 
 If Len(Trim$(CmdAddSupEmail.Tag)) Then
                
    CmdAddSupEmail.Tag = UCase(CmdAddSupEmail.Tag)
    If InStr(1, CmdAddSupEmail.Tag, "") = 0 Then CmdAddSupEmail.Tag = ("" & CmdAddSupEmail.Tag)
    If Len(Trim$(CmdAddSupEmail.Tag)) > 60 Then
       MsgBox "Recepient Email can not be more than 60 characters.", vbInformation, "Imswin"
       CmdAddSupEmail.Tag = ""
       Exit Sub
    End If
    Call AddRecepient(CmdAddSupEmail.Tag, , True)
    'CmdAddSupEmail.Tag = ""
    
 End If
 Exit Sub
Handler:
  MsgBox "Error occurred while adding email to the recepient's list.Error Description  " & Err.Description, vbInformation, "Imswin"
  Err.Clear
End Sub

Private Sub CmdAddSupEmail_LostFocus()
cmdRemove.SetFocus
End Sub

Private Sub CmdAssignFQA_Click()
 Dim Company As String
 Dim Location As String
 Dim us As String
 Dim cc As String
 Dim stocktype As String
 
 On Error GoTo ErrHand
 
 If MsgBox("This action will assign the FQA code listed in the FQA fields below  '" & Trim(TxtToCompanyFQA) & "-" & Trim(SSOleDBToLocationFQA) & "-" & Trim(SSOleDBtoUSChartFQA) & "-" & Trim(TxtToStocktypeFQA) & "-" & Trim(SSOleDBToCamChartFQA) & "' to all the remaining line items. Are you sure you want to go ahead?", vbInformation + vbYesNo, "Ims") = vbYes Then
 
 If POFqa.Count > 2 Then
 
    Company = TxtToCompanyFQA
    Location = SSOleDBToLocationFQA
    us = SSOleDBtoUSChartFQA
    cc = SSOleDBToCamChartFQA
    stocktype = TxtToStocktypeFQA
    
    POFqa.MoveFirst
    
           Do While Not POFqa.EOF
                      
               POFqa.ToCompany = Company
               POFqa.Tolocation = Location
               POFqa.ToUSChart = us
               POFqa.ToCamChart = cc
               POFqa.ToStockType = stocktype
               
               If POFqa.MoveNext = False Then Exit Do
               
           Loop
       
    End If
    
    POFqa.MoveLineTo PoItem.Linenumb
    
 End If
 
 Exit Sub
ErrHand:

MsgBox "Errors occurred while trying to copy the FQA to all the line items. Please try again." & Err.Description, vbCritical, "Ims"
 
Err.Clear
End Sub

Private Sub CmdConvert_Click()


    Screen.MousePointer = vbHourglass
    Load frm_ConvertToPO
    
    frm_ConvertToPO.Show
    
    Screen.MousePointer = vbArrow
    
End Sub

Private Sub CmdcopyLI_LostFocus(Index As Integer)


Select Case Index
   Case 1
  If txtRemarks.Enabled = True Then
       txtRemarks.SetFocus
       Call txtRemarks_GotFocus
  End If
Case 2
        cmd_Addterms.SetFocus
End Select

End Sub

Private Sub dgRecipientList_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
If PoReceipients.Count > 0 Then PoReceipients.MoveFirst
 If Trim$(Len(dgRecipientList.Columns(0).Text)) > 0 Then
   Do While Not PoReceipients.EOF
       If PoReceipients.Receipient = Trim$(oldVALUE) Then
         PoReceipients.Receipient = Trim$(dgRecipientList.Columns(0).Text)
         Exit Sub
       End If
       PoReceipients.MoveNext
   Loop
 Else
   dgRecipientList.Columns(0).Text = oldVALUE
 End If
End Sub

Private Sub dgRecipientList_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
DispPromptMsg = 0
If MsgBox("Are you sure you want to Delete the Recepient?", vbYesNo) = vbYes Then
     PoReceipients.DeleteCurrentLI (dgRecipientList.Columns(0).Text)
Else
     Cancel = -1
End If
End Sub

Private Sub dgRecipientList_Click()
Dim x As Integer
mSelection = True
End Sub

Private Sub dgRecipientList_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub dgRecipientList_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
'dgRecipientList.
End Sub

Private Sub DTP_Required_GotFocus()
Call HighlightBackground(DTP_Required)
End Sub

Private Sub DTP_Required_LostFocus()
Call NormalBackground(DTP_Required)
If ssdcboCommoditty.Enabled = True Then
  Call HighlightBackground(ssdcboCommoditty)
  ssdcboCommoditty.SetFocus
Else
  Call HighlightBackground(ssdcboManNumber)
  ssdcboManNumber.SetFocus
End If
End Sub

Private Sub dtpRequestedDate_GotFocus()
Call HighlightBackground(dtpRequestedDate)
End Sub

Private Sub dtpRequestedDate_LostFocus()
Call NormalBackground(dtpRequestedDate)
End Sub

Private Sub dtpRequestedDate_Validate(Cancel As Boolean)
If FormMode <> mdvisualization And dtpRequestedDate.value < Now() Then
   Cancel = True
   MsgBox "Date Required cannot be less than Today's Date."
   dtpRequestedDate.SetFocus
End If
End Sub

Private Sub NavBar1_OnCloseClick()
'frm_NewPurchase.Visible = False
Unload frm_NewPurchase

 Unload frm_ConvertToPO
    

End Sub

Private Sub NavBar1_OnDeleteClick()
Dim y As Integer
Dim mpo As String
Dim PoHeaderErrors As Boolean
    y = sst_PO.Tab
    
  Select Case y
        
   Case 0
   CmdConvert.Enabled = False
   
   mpo = Poheader.Ponumb
   If CInt(LblRevNumb) = 0 And Len(Trim(LblAppBy)) = 0 Then
   
        If MsgBox("Are you sure you want to Delete  Transaction number " & mpo & "?", vbCritical + vbYesNo, "Imswin") = vbNo Then Exit Sub
                   Me.MousePointer = vbHourglass
             
                             
                             deIms.cnIms.Errors.Clear
                             deIms.cnIms.BeginTrans
                             PoHeaderErrors = True
                             PoHeaderErrors = Poheader.Delete(mpo)
                             
                    If deIms.cnIms.Errors.Count = 0 Or PoHeaderErrors = False Then
                                 deIms.cnIms.CommitTrans
                                 Poheader.Requery
                                 
                                 If Poheader.MoveFirst = True Then Call LoadFromPOHEADER
                                   mIsPoNumbComboLoaded = False
                                   MsgBox "Transaction Order # " & mpo & " was deleted successfully."
                                 
                                 
                                  
                                  Set PoItem = Nothing
                                  Set PoReceipients = Nothing
                                  Set PORemark = Nothing
                                  Set POClause = Nothing
                                
                                  FormMode = ChangeMode(mdvisualization)
                            
                                  If FormMode = mdvisualization Then
                               
                                    NavBar1.NewEnabled = True
                                    NavBar1.NextEnabled = True
                                    NavBar1.PreviousEnabled = True
                                    NavBar1.LastEnabled = True
                                    NavBar1.FirstEnabled = True
                                    NavBar1.DeleteEnabled = True
                                    NavBar1.EditEnabled = True
                                    NavBar1.CancelEnabled = False
                                    ssOleDbPO.Enabled = True
                                    NavBar1.SaveEnabled = False
                                    NavBar1.EMailEnabled = True 'JCG 2008/11/14
                                 End If
                               
                                
                                 
                      Else
                                  deIms.cnIms.RollbackTrans
                                   MsgBox "Errors Occured.Could Not Delete The Transaction Order # " & mpo & ".Please Close the Form and start it once more.", vbCritical, "Imswin"
                                   Poheader.CancelUpdate: LoadFromPOHEADER
                                 
                                 
                                 
                      End If
           Me.MousePointer = vbArrow
  Else
       MsgBox "This Transaction order # " & mpo & " can not be deleted.", vbInformation, "Imswin"
  End If
   
 Dim imsLock As imsLock.Lock
 Set imsLock = New imsLock.Lock
 Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

   
  Case 2
  
        Dim x As Integer
        If Not IsNothing(PoItem) Then
             
           If PoItem.Count > 0 Then
                  If lookups Is Nothing Then Set lookups = Mainpo.lookups
                  x = lookups.CanUserDeleteLineItem(PoItem.Ponumb, IIf(Poheader.revinumb = 0, 0, Poheader.Originalrevinumb - 1), PoItem.Linenumb)
                  
                  If x = 0 Then
                    
                         If UCase(Trim$(PoItem.Stasliit)) = "OP" Then
                              
                                  MsgBox "Can not delete this Line Item.It is being carried from Previous Revision.", vbInformation, "Imswin"
                         
                         Else
                         
                                 If MsgBox("Are you sure you want to Delete this Record?", vbCritical + vbYesNo, "Imswin") = vbYes Then
                                         If PoItem.DeleteCurrentLI Then
                                            If PoItem.Count > 0 Then
                                              LoadFromPOITEM
                                              POFqa.DeleteLine
                                            Else
                                              ClearAllPoLineItems
                                            End If
                                         End If
                                 End If
                              
                         End If
                    
                    
                  ElseIf x = 1 Then
                  
                      If Poheader.revinumb = 1 And PoItem.EditMode = 2 Then
                           If PoItem.DeleteCurrentLI Then LoadFromPOITEM
                      ElseIf (Poheader.revinumb = 1 Or Poheader.revinumb > 1) And PoItem.EditMode <> 2 Then
                            MsgBox "Can not delete the Line Item.It is being carried over from the previous Revisions.", vbInformation, "Imswin"
                      End If
                      
                  End If
           End If
        End If
        
   End Select
End Sub

Private Sub NavBar1_OnEMailClick()
On Error Resume Next

Dim i As RPTIFileInfo
Dim Params(1) As String

'Call sendOutlookEmailandFax
If Poheader.stas <> "OH" Then 'JCG 2008/11/15
    If Poheader.stas <> "CL" Then  'JCG 2008/11/15
        If Poheader.stas <> "CA" Then  'JCG 2008/11/15
            Call SelectGatewayAndSendOutMails
        End If
    End If
End If

End Sub

Private Sub NavBar1_OnPrintClick()

On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath + "po.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Poheader.Ponumb + ";true"
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("M00392") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction", msg1) 'J modified
        Call translator.Translate_Reports("po.rpt") 'J added
        msg1 = translator.Trans("M00091") 'J added
        If msg1 = "" Then msg1 = "Total Price of"
        msg2 = translator.Trans("M00093") 'J added
        If msg2 = "" Then msg2 = "in"
        Dim curr
        curr = " : "
        .Formulas(99) = "gttext = ' " + msg1 + " ' + {DOCTYPE.doc_desc} + ' " + msg2 + " ' + {CURRENCY.curr_desc} + ' " + curr + "' + totext(Sum ({@total}, {PO.po_ponumb}))" 'J modified
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)
    End If
End Sub



Private Sub PoItem_AfterCancelUpdate()
Dim i As Integer

 If Poheader.fromstckmast = False Then
    
    PoItem.MoveFirst
    
    For i = 1 To PoItem.Count

        PoItem.Linenumb = i
        
        PoItem.Comm = PoItem.Ponumb & "/" & i
        
        PoItem.MoveNext
    
    Next
    
    PoItem.MoveFirst

  End If
  
  
End Sub

Private Sub PoItem_AfterDeleteCurrentLI()
Dim i As Integer

 If Poheader.fromstckmast = False Then
    
    PoItem.MoveFirst
    
    For i = 1 To PoItem.Count

        PoItem.Linenumb = i
        
        PoItem.Comm = Trim(PoItem.Ponumb) & "/" & i
        
        PoItem.MoveNext
    
    Next
    
    PoItem.MoveFirst

  End If
  

End Sub

Private Sub showAll_Click(Index As Integer)
    mIsPoNumbComboLoaded = False
End Sub

Private Sub ssdcboCommoditty_DblClick()
On Error Resume Next
    Set comsearch = New frm_StockSearch
    
    comsearch.Execute
End Sub

Private Sub ssdcboCommoditty_DropDown()

    deIms.rsActiveStockmasterLookup.Close
   
    Call deIms.ActiveStockMasterLooKUP(deIms.NameSpace)
    Set ssdcboCommoditty.DataSourceList = Nothing
    Set ssdcboCommoditty.DataSourceList = deIms.rsActiveStockmasterLookup

    ssdcboCommoditty.DataFieldToDisplay = "stk_stcknumb"
    ssdcboCommoditty.DataFieldList = "stk_desc"

    mDidUserOpenStkMasterForm = False


End Sub

Private Sub ssdcboCommoditty_GotFocus()
ssdcboCommoditty.SelStart = 0
ssdcboCommoditty.SelLength = 0
Call HighlightBackground(ssdcboCommoditty)
End Sub

Private Sub ssdcboCommoditty_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboCommoditty.DroppedDown Then ssdcboCommoditty.DroppedDown = True
End Sub

Private Sub ssdcboCommoditty_KeyPress(KeyAscii As Integer)
'ssdcboCommoditty.MoveNext
End Sub

Private Sub ssdcboCommoditty_LostFocus()
Call NormalBackground(ssdcboCommoditty)
End Sub

Private Sub ssdcboCommoditty_Validate(Cancel As Boolean)

If FormMode = mdvisualization Then Exit Sub
If Len(Trim(ssdcboCommoditty)) = 0 Then Exit Sub
If chk_FrmStkMst.value = 0 Then Exit Sub

deIms.rsActiveStockmasterLookup.MoveFirst
deIms.rsActiveStockmasterLookup.Find " stk_stcknumb ='" & Trim(ssdcboCommoditty) & "'"

If deIms.rsActiveStockmasterLookup.AbsolutePosition = adPosBOF Or deIms.rsActiveStockmasterLookup.AbsolutePosition = adPosEOF Or deIms.rsActiveStockmasterLookup.AbsolutePosition = adPosUnknown Then

        Cancel = True
        MsgBox " Stock Number does not exist. Please enter one from the list.", vbInformation, "Imswin"
        
ElseIf Trim(ssdcboCommoditty.Text) <> Trim(ssdcboCommoditty.Tag) Then
    
        'This case arises when the user uses the TAB instead of ENTER to move to the next tab.
    
        Call ssdcboCommoditty_Click

End If

End Sub

Private Sub ssdcboCondition_Click()
ssdcboCondition.SelStart = 0
ssdcboCondition.SelLength = 0
ssdcboCondition.Tag = ssdcboCondition.Columns(0).Text
End Sub

Private Sub ssdcboCondition_GotFocus()
ssdcboCondition.SelStart = 0
ssdcboCondition.SelLength = 0
Call HighlightBackground(ssdcboCondition)
End Sub

Private Sub ssdcboCondition_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboCondition.DroppedDown Then ssdcboCondition.DroppedDown = True

End Sub

Private Sub ssdcboCondition_LostFocus()
Call NormalBackground(ssdcboCondition)
End Sub

Private Sub ssdcboCondition_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboCondition.Text)) > 0 And Not ssdcboCondition.IsItemInList Then
  Cancel = True
   ssdcboCondition.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub ssdcboDelivery_Click()
On Error GoTo Handler
ssdcboDelivery.SelStart = 0
ssdcboDelivery.SelLength = 0
ssdcboDelivery.Tag = ssdcboDelivery.Columns(0).Text
Exit Sub
Handler:
 MsgBox "Error occurred during ssdcboDelivery_Click.Please try again.Error description  " & Err.Description, vbInformation, "Imswin"
 Err.Clear
End Sub

Private Sub ssdcboDelivery_DropDown()
ssdcboDelivery.RemoveAll
If deIms.rsTermDelivery.State = 1 Then deIms.rsTermDelivery.Close
Call deIms.TermDelivery(FNameSpace)
 deIms.rsTermDelivery.Filter = "tod_actvflag<>0"
Do While Not deIms.rsTermDelivery.EOF
       ssdcboDelivery.AddItem deIms.rsTermDelivery!tod_termcode & ";" & deIms.rsTermDelivery!tod_desc
       deIms.rsTermDelivery.MoveNext
Loop

deIms.rsTermDelivery.Filter = ""

End Sub

Private Sub ssdcboDelivery_GotFocus()
ssdcboDelivery.SelStart = 0
ssdcboDelivery.SelLength = 0
Call HighlightBackground(ssdcboDelivery)
End Sub

Private Sub ssdcboDelivery_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboDelivery.DroppedDown Then ssdcboDelivery.DroppedDown = True
End Sub

Private Sub ssdcboDelivery_LostFocus()
Call NormalBackground(ssdcboDelivery)
If ssdcboShipper.Enabled = True Then ssdcboShipper.SetFocus
End Sub

Private Sub ssdcboDelivery_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboDelivery.Text)) > 0 And Not ssdcboDelivery.IsItemInList Then
  Cancel = True
   ssdcboDelivery.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub ssdcboManNumber_Click()
ssdcboManNumber.SelStart = 0
ssdcboManNumber.SelLength = 0
End Sub

Private Sub ssdcboManNumber_GotFocus()
ssdcboManNumber.SelStart = 0
ssdcboManNumber.SelLength = 0
Call HighlightBackground(ssdcboManNumber)
End Sub

Private Sub ssdcboManNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboManNumber.DroppedDown Then ssdcboManNumber.DroppedDown = True
End Sub
Private Sub ssdcboManNumber_KeyPress(KeyAscii As Integer)

 If Len(ssdcboManNumber & Chr(KeyAscii)) > 30 And KeyAscii <> 8 Then
   MsgBox "Input is restricted to 30 cahracters"
   KeyAscii = 0
End If

End Sub

Private Sub ssdcboManNumber_LostFocus()
Call NormalBackground(ssdcboManNumber)
End Sub

Private Sub ssdcboManNumber_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboManNumber.Text)) > 0 Then
'''   If ssdcboManNumber.IsItemInList = False Then
'''        Cancel = True
'''        ssdcboManNumber.SetFocus
'''        MsgBox "Invalid Value For Manufacturer.", , "Imswin"
'''    End If
    If Len(ssdcboManNumber) > 30 Then
       Cancel = True
       ssdcboManNumber.SetFocus
       MsgBox "Manufacturer text can not be more than 30 characters."
    End If
End If
End Sub

Private Sub ssdcboRequisition_GotFocus()
ssdcboRequisition.SelLength = 0
ssdcboRequisition.SelStart = 0
Call HighlightBackground(ssdcboRequisition)
End Sub

Private Sub ssdcboRequisition_LostFocus()
Call NormalBackground(ssdcboRequisition)
End Sub

Private Sub ssdcboRequisition_Validate(Cancel As Boolean)

If Len(Trim(ssdcboRequisition.Text)) = 0 Then

    lblReqLineitem.Caption = ""
    Exit Sub
    
End If
    
If ssdcboRequisition.IsItemInList = False Then

    MsgBox "Requisition does not exist. Please select one from the list.", vbInformation, "Ims"
    Cancel = True
    ssdcboRequisition.SetFocus
Else

    If Trim(ssdcboRequisition.value) <> Trim(ssdcboRequisition.Tag) Then
        Call ssdcboRequisition_Click
    End If
    
End If

End Sub

Private Sub ssdcboShipper_DropDown()
ssdcboShipper.RemoveAll
If deIms.rsSHIPPER.State = 1 Then deIms.rsSHIPPER.Close
Call deIms.Shipper(FNameSpace)
Do While Not deIms.rsSHIPPER.EOF
       ssdcboShipper.AddItem deIms.rsSHIPPER!shi_code & ";" & deIms.rsSHIPPER!shi_name
       deIms.rsSHIPPER.MoveNext
   Loop
End Sub

Private Sub ssdcboShipper_GotFocus()
ssdcboShipper.SelLength = 0
ssdcboShipper.SelStart = 0
 Call HighlightBackground(ssdcboShipper)
End Sub

Private Sub ssdcboShipper_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboShipper.DroppedDown Then ssdcboShipper.DroppedDown = True
End Sub

Private Sub ssdcboShipper_LostFocus()
Call NormalBackground(ssdcboShipper)
End Sub

Private Sub ssdcboShipper_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboShipper.Text)) > 0 And Not ssdcboShipper.IsItemInList Then
  Cancel = True
   ssdcboShipper.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBCompany_DropDown()
SSOleDBcompany.RemoveAll
If deIms.rsActiveCompany.State = 1 Then deIms.rsActiveCompany.Close
Call deIms.ActiveCompany(FNameSpace)

Do While Not deIms.rsActiveCompany.EOF
       SSOleDBcompany.AddItem deIms.rsActiveCompany!com_compcode & ";" & deIms.rsActiveCompany!com_name
       deIms.rsActiveCompany.MoveNext
       
   Loop
End Sub

Private Sub SSOleDBCompany_GotFocus()
SSOleDBcompany.SelLength = 0
SSOleDBcompany.SelStart = 0
Call HighlightBackground(SSOleDBcompany)
End Sub

Private Sub SSOleDBCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBcompany.DroppedDown Then SSOleDBcompany.DroppedDown = True
End Sub

Private Sub SSOleDBCompany_LostFocus()
Call NormalBackground(SSOleDBcompany)
End Sub

Private Sub SSOleDBCompany_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBcompany.Text)) > 0 Then
   If SSOleDBcompany.IsItemInList = False Then
        Cancel = True
        SSOleDBcompany.SetFocus
        MsgBox "Invalid Value For Company Code.", , "Imswin"
   
     
    End If
End If
End Sub

Private Sub SSOleDBCurrency_DropDown()
SSOleDBCurrency.RemoveAll
If deIms.rsCURRENCY.State = 1 Then deIms.rsCURRENCY.Close
Call deIms.Currency(FNameSpace)

 Do While Not deIms.rsCURRENCY.EOF
       SSOleDBCurrency.AddItem deIms.rsCURRENCY!curr_code & ";" & deIms.rsCURRENCY!curr_desc
       deIms.rsCURRENCY.MoveNext
   Loop
End Sub

Private Sub SSOleDBCurrency_GotFocus()
SSOleDBCurrency.SelStart = 0
SSOleDBCurrency.SelLength = 0
 Call HighlightBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCurrency.DroppedDown Then SSOleDBCurrency.DroppedDown = True
End Sub

Private Sub SSOleDBCurrency_LostFocus()
Call NormalBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCurrency.Text)) > 0 Then
    If Not SSOleDBCurrency.IsItemInList Then
         Cancel = True
          SSOleDBCurrency.SetFocus
        MsgBox "Invalid Value", , "Imswin"
    Else
      If lookups Is Nothing Then Set lookups = Mainpo.lookups
      If lookups.CurrencyDetlExist(SSOleDBCurrency.Columns(0).Text) = False Then
         MsgBox "No Currency Detail for today.Please Update Currency Table"
         SSOleDBcompany.Text = ""
         SSOleDBcompany.SetFocus
         Cancel = True
      End If
          
    End If
End If
End Sub

Private Sub SSOleDBCustCategory_Click()
SSOleDBCustCategory.SelLength = 0
SSOleDBCustCategory.SelStart = 0
End Sub

Private Sub SSOleDBCustCategory_DropDown()
Dim rsCUSTOM As ADODB.Recordset
 SSOleDBCustCategory.RemoveAll
If lookups Is Nothing Then Set lookups = Mainpo.lookups
   
   Set rsCUSTOM = lookups.GetCustom
   Do While Not rsCUSTOM.EOF
     SSOleDBCustCategory.AddItem rsCUSTOM!cust_cate
     rsCUSTOM.MoveNext
   Loop
   rsCUSTOM.Close
 Set rsCUSTOM = Nothing
 
End Sub

Private Sub SSOleDBCustCategory_GotFocus()
SSOleDBCustCategory.SelLength = 0
SSOleDBCustCategory.SelStart = 0
Call HighlightBackground(SSOleDBCustCategory)
End Sub

Private Sub SSOleDBCustCategory_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCustCategory.DroppedDown Then SSOleDBCustCategory.DroppedDown = True
End Sub

Private Sub SSOleDBCustCategory_LostFocus()
Call NormalBackground(SSOleDBCustCategory)
End Sub

Private Sub SSOleDBCustCategory_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCustCategory.Text)) > 0 Then
   If SSOleDBCustCategory.IsItemInList = False Then
        Cancel = True
        SSOleDBCustCategory.SetFocus
        MsgBox "Invalid Value For Customs category.", , "Imswin"
    End If
End If
End Sub

Private Sub SSOleDBDocType_Click()
 
Dim Recepients() As String
Dim PreviousDoctype As String
Dim i As Integer
Dim True_False As Boolean

SSOleDBDocType.SelLength = 0

SSOleDBDocType.SelStart = 0

PreviousDoctype = SSOleDBDocType.Tag

SSOleDBDocType.Tag = SSOleDBDocType.Columns(0).Text
 
'JCG 2008/10/26
    ClearPoReceipients
    Set PoReceipients = Nothing
'----------------------
 
 If lookups Is Nothing Then Set lookups = Mainpo.lookups
 
 'If FormMode <> mdCreation Then Exit Sub
  
 If DeleteDefaultRecepientsForDoctype(PreviousDoctype) = 1 Then Exit Sub
 
  If lookups.CanDocTypeAutoDist(SSOleDBDocType.Tag, True_False) = 1 Then
 
    MsgBox "Either you do not have any document rights or error occurred while trying to check if the document type can be auto-distributed.", vbCritical, "Imswin"
    
    Exit Sub
    
 End If
 
 If True_False = False Then
 
    MsgBox "Document Type is set on NO AUTO-DISTRIBUTION mode. Stored Electronic distribution recepients will not be added.", vbInformation, "Imswin"
 
 Exit Sub

 End If

 If lookups.GetDefaultRecForDoctype(SSOleDBDocType.Tag, Recepients) = 1 Then
 
    MsgBox "Errors Occurred while Trying to Access the Distribution List for the Document type. Please Try again.", vbCritical, "Imswin"
    
    Exit Sub
    
 End If
  
 If IsArrayLoaded(Recepients) Then
 
    For i = 0 To UBound(Recepients)
        If LTrim(RTrim(SSOleDBDocType.Columns(0).Text)) = "R" Then 'JCG 2008-10-26
            If InStr(Recepients(i), "@") > 0 Then 'JCG 2008-10-26
                Call AddRecepient(Recepients(i)) 'JCG 2008-10-26
            End If 'JCG 2008-10-26
        Else 'JCG 2008-10-26
            Call AddRecepient(Recepients(i))
        End If 'JCG 2008-10-26
    Next i
    
 End If
 
End Sub

Private Sub SSOleDBDocType_DropDown()
'If mIsDocTypeLoaded = True Then
      SSOleDBDocType.RemoveAll
      If lookups Is Nothing Then Set lookups = Mainpo.lookups
        Dim GRsDoctype As ADODB.Recordset
        Set GRsDoctype = lookups.GetDoctypeForUser(CurrentUser)

        Do While Not GRsDoctype.EOF
           rsDOCTYPE.MoveFirst
           rsDOCTYPE.Find ("DOC_CODE='" & Trim$(GRsDoctype!buyr_docutype) & "'")

           SSOleDBDocType.AddItem GRsDoctype!buyr_docutype & ";" & rsDOCTYPE!doc_desc
           GRsDoctype.MoveNext
        Loop
  '  mIsDocTypeLoaded = True
       Set GRsDoctype = Nothing
 ' End If

    
End Sub

Private Sub SSOleDBDocType_GotFocus()
SSOleDBDocType.SelLength = 0
SSOleDBDocType.SelStart = 0
 Call HighlightBackground(SSOleDBDocType)
End Sub

Private Sub SSOleDBDocType_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBDocType.DroppedDown Then SSOleDBDocType.DroppedDown = True
End Sub

Private Sub SSOleDBDocType_KeyPress(KeyAscii As Integer)
SSOleDBDocType.MoveNext
End Sub

Private Sub SSOleDBDocType_LostFocus()
Call NormalBackground(SSOleDBDocType)
End Sub

Private Sub SSOleDBDocType_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBDocType.Text)) > 0 And Not SSOleDBDocType.IsItemInList Then
  Cancel = True
   SSOleDBDocType.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBInvLocation_Click()
SSOleDBInvLocation.Tag = SSOleDBInvLocation.Columns(0).Text
SSOleDBInvLocation.SelLength = 0
SSOleDBInvLocation.SelStart = 0
'Call ModifyToFQAWithNewLocation
End Sub

Private Sub SSOleDBInvLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBInvLocation.DroppedDown Then SSOleDBInvLocation.DroppedDown = True
End Sub

Private Sub SSOleDBInvLocation_KeyPress(KeyAscii As Integer)
SSOleDBInvLocation.MoveNext
End Sub

Private Sub SSOleDBInvLocation_LostFocus()
Call NormalBackground(SSOleDBInvLocation)
End Sub

Private Sub SSOleDBInvLocation_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBInvLocation.Text)) > 0 And Not SSOleDBInvLocation.IsItemInList Then
  Cancel = True
   SSOleDBInvLocation.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBOriginator_Click()
SSOleDBOriginator.SelLength = 0
SSOleDBOriginator.SelStart = 0

'SSOleDBOriginator.Tag = SSOleDBOriginator.Columns(0).text
End Sub

Private Sub SSOleDBOriginator_DropDown()
SSOleDBOriginator.RemoveAll
If deIms.rsActiveOriginator.State = 1 Then deIms.rsActiveOriginator.Close
Call deIms.ActiveOriginator(FNameSpace)
Do While Not deIms.rsActiveOriginator.EOF
       SSOleDBOriginator.AddItem deIms.rsActiveOriginator!ori_code '& ";" & deIms.rsActiveOriginator!ori_code
       deIms.rsActiveOriginator.MoveNext
   Loop
End Sub

Private Sub SSOleDBOriginator_GotFocus()
SSOleDBOriginator.SelLength = 0
SSOleDBOriginator.SelStart = 0
 Call HighlightBackground(SSOleDBOriginator)
End Sub

Private Sub SSOleDBOriginator_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBOriginator.DroppedDown Then SSOleDBOriginator.DroppedDown = True
End Sub

Private Sub SSOleDBOriginator_KeyPress(KeyAscii As Integer)
SSOleDBOriginator.MoveNext
End Sub

Private Sub SSOleDBOriginator_LostFocus()
Call NormalBackground(SSOleDBOriginator)
End Sub

Private Sub SSOleDBOriginator_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SSOleDBOriginator.SetFocus
End Sub

Private Sub SSOleDBOriginator_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBOriginator.Text)) > 0 And Not SSOleDBOriginator.IsItemInList Then
  Cancel = True
   SSOleDBOriginator.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBPO_Change()
'''If FormMode = mdCreation Then
'''    SSoledbSupplier.text = ""
'''    SSoledbSupplier.Tag = ""
'''End If
End Sub

Private Sub SSOleDBPO_DropDown()
If Not FormMode = mdCreation Then
    If mIsPoNumbComboLoaded = False Then
      If deIms.rsPonumb.State = 1 Then
         deIms.rsPonumb.Close
      End If
       Call deIms.Ponumb(deIms.NameSpace)
       If showAll(1).value = True Then
        Dim sql As String
            sql = "Select po_ponumb from po where  po_npecode = '" + FNameSpace + "' and " _
                & "po_date >=  dateadd(year, -1, getdate())" _
                & "ORDER BY po_ponumb, po_date, po_reqddelvdate"
            deIms.rsPonumb.Close
            deIms.rsPonumb.Source = sql
            deIms.rsPonumb.Open
        End If

        '2012-9-30 juan commented and added
        'ssOleDbPO.RemoveAll

        Set ssOleDbPO.DataSourceList = deIms.rsPonumb
        ssOleDbPO.DataFieldList = deIms.rsPonumb.Fields(0).Name
       
        'Do While Not deIms.rsPonumb.EOF
        '   ssOleDbPO.AddItem deIms.rsPonumb!po_ponumb
        '   deIms.rsPonumb.MoveNext
        'Loop
        mIsPoNumbComboLoaded = True
    
    End If
  Else
    ssOleDbPO.DroppedDown = False
 End If
End Sub

Private Sub SSOleDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
 If Not FormMode = mdCreation Then
    If Not ssOleDbPO.DroppedDown Then ssOleDbPO.DroppedDown = True
 End If
End Sub

Private Sub ssOleDbPO_KeyPress(KeyAscii As Integer)
If Not FormMode = mdCreation Then ssOleDbPO.MoveNext
End Sub

Private Sub SSOleDBPriority_Click()
SSOleDBPriority.SelLength = 0
SSOleDBPriority.SelStart = 0
SSOleDBPriority.Tag = SSOleDBPriority.Columns(0).Text
End Sub

Private Sub SSOleDBPriority_DropDown()
SSOleDBPriority.RemoveAll
If deIms.rsPRIORITY.State = 1 Then deIms.rsPRIORITY.Close
Call deIms.Priority(FNameSpace)
Do While Not deIms.rsPRIORITY.EOF
       SSOleDBPriority.AddItem deIms.rsPRIORITY!pri_code & ";" & deIms.rsPRIORITY!pri_desc
       deIms.rsPRIORITY.MoveNext
   Loop
End Sub

Private Sub SSOleDBPriority_GotFocus()
SSOleDBPriority.SelLength = 0
SSOleDBPriority.SelStart = 0
 Call HighlightBackground(SSOleDBPriority)
End Sub

Private Sub SSOleDBPriority_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBPriority.DroppedDown Then SSOleDBPriority.DroppedDown = True
End Sub

Private Sub SSOleDBPriority_KeyPress(KeyAscii As Integer)
SSOleDBPriority.MoveNext
End Sub

Private Sub SSOleDBPriority_LostFocus()
Call NormalBackground(SSOleDBPriority)
End Sub

Private Sub SSOleDBPriority_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBPriority.Text)) > 0 And Not SSOleDBPriority.IsItemInList Then
  Cancel = True
   SSOleDBPriority.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBShipTo_Click()
SSOleDBShipTo.SelLength = 0
SSOleDBShipTo.SelStart = 0
SSOleDBShipTo.Tag = SSOleDBShipTo.Columns(0).Text
End Sub

Private Sub SSOleDBShipTo_DropDown()
SSOleDBShipTo.RemoveAll

If deIms.rsActiveShipTo.State = 1 Then deIms.rsActiveShipTo.Close
Call deIms.ActiveShipTo(FNameSpace)
Do While Not deIms.rsActiveShipTo.EOF
       SSOleDBShipTo.AddItem deIms.rsActiveShipTo!sht_code & ";" & deIms.rsActiveShipTo!sht_name
       deIms.rsActiveShipTo.MoveNext
Loop
   
End Sub

Private Sub SSOleDBShipTo_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBShipTo.DroppedDown Then SSOleDBShipTo.DroppedDown = True
End Sub

Private Sub SSOleDBShipTo_KeyPress(KeyAscii As Integer)
SSOleDBShipTo.MoveNext
End Sub

Private Sub SSOleDBShipTo_LostFocus()
Call NormalBackground(SSOleDBShipTo)
End Sub

Private Sub SSOleDBShipTo_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBShipTo.Text)) > 0 And Not SSOleDBShipTo.IsItemInList Then
  Cancel = True
   SSOleDBShipTo.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOledbSrvCode_Click()
SSOledbSrvCode.SelLength = 0
SSOledbSrvCode.SelStart = 0
SSOledbSrvCode.Tag = SSOledbSrvCode.Columns(0).Text
End Sub

Private Sub SSOledbSrvCode_GotFocus()
SSOledbSrvCode.SelStart = 0
SSOledbSrvCode.SelLength = 0
 Call HighlightBackground(SSOledbSrvCode)
End Sub

Private Sub SSOledbSrvCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOledbSrvCode.DroppedDown Then SSOledbSrvCode.DroppedDown = True
End Sub

Private Sub SSOledbSrvCode_KeyUp(KeyCode As Integer, Shift As Integer)
'SSOledbSrvCode.MoveNext
End Sub

Private Sub SSOledbSrvCode_LostFocus()
Call NormalBackground(SSOledbSrvCode)
End Sub

Private Sub SSOleDBsupplier_DropDown()
'If mIsSupplierComboLoaded = False Then
On Error GoTo Handler
Dim rsSUPPLIER As ADODB.Recordset

If DoesDocTypeExist = False Then
   Exit Sub
End If

 
 

SSoledbSupplier.RemoveAll
If lookups Is Nothing Then Set lookups = Mainpo.lookups

If lookups.GetUserMenuLevel(CurrentUser) = 5 Then
  Set rsSUPPLIER = lookups.GetLocalSuppliers
  If deIms.rsActiveSupplier.State = 1 Then
     deIms.rsActiveSupplier.Close
     Call deIms.ActiveSupplier(deIms.NameSpace)
  End If
Else
  If deIms.rsActiveSupplier.State = 1 Then
     deIms.rsActiveSupplier.Close
     Call deIms.ActiveSupplier(deIms.NameSpace)
  End If
  
  Set rsSUPPLIER = deIms.rsActiveSupplier
End If

If rsSUPPLIER.RecordCount > 0 Then
   rsSUPPLIER.Filter = "sup_actvflag=1"
    rsSUPPLIER.MoveFirst
    Do While Not rsSUPPLIER.EOF
       SSoledbSupplier.AddItem rsSUPPLIER!sup_code & ";" & rsSUPPLIER!sup_name & ";" & rsSUPPLIER!sup_city & ";" & rsSUPPLIER!sup_phonnumb & ";" & rsSUPPLIER!sup_faxnumb
       rsSUPPLIER.MoveNext
    Loop
   rsSUPPLIER.Filter = ""
End If

Exit Sub
Handler:
MsgBox "Errors occured while querying the Database for Suppliers.Error Description is '" & Err.Description & "'"
Err.Clear
End Sub

Private Sub SSOleDBsupplier_GotFocus()
SSoledbSupplier.SelStart = 0
SSoledbSupplier.SelLength = 0
 Call HighlightBackground(SSoledbSupplier)
End Sub

Private Sub SSoledbSupplier_InitColumnProps()
SSoledbSupplier.Columns(0).Visible = False
End Sub

Private Sub SSOleDBsupplier_KeyDown(KeyCode As Integer, Shift As Integer)

If DoesDocTypeExist = False Then
   KeyCode = 0
   Exit Sub
End If

If Not SSoledbSupplier.DroppedDown Then SSoledbSupplier.DroppedDown = True
End Sub

Private Sub SSOleDBsupplier_LostFocus()
Call NormalBackground(SSoledbSupplier)
End Sub

Private Sub SSOleDBsupplier_Validate(Cancel As Boolean)
If Len(Trim$(SSoledbSupplier.Text)) > 0 And Not SSoledbSupplier.IsItemInList Then
  Cancel = True
   SSoledbSupplier.SetFocus
 MsgBox "Invalid Value", , "Imswin"
Else 'JCG 2008/01/14
    newSupplier = True 'JCG 2008/01/14
End If
End Sub

Private Sub SSOleDBToBeUsedFor_Click()
SSOleDBToBeUsedFor.SelLength = 0
SSOleDBToBeUsedFor.SelStart = 0
'SSOleDBToBeUsedFor.Tag = SSOleDBToBeUsedFor.Columns(0).text
End Sub

Private Sub SSOleDBToBeUsedFor_DropDown()
SSOleDBToBeUsedFor.RemoveAll
If deIms.rsActiveTbu.State = 1 Then deIms.rsActiveTbu.Close
Call deIms.ActiveTbu(FNameSpace)
Do While Not deIms.rsActiveTbu.EOF
       SSOleDBToBeUsedFor.AddItem deIms.rsActiveTbu!tbu_name '& ";" & deIms.rsActiveOriginator!tbu_name
       deIms.rsActiveTbu.MoveNext
   Loop
End Sub

Private Sub SSOleDBToBeUsedFor_GotFocus()
SSOleDBToBeUsedFor.SelStart = 0
SSOleDBToBeUsedFor.SelLength = 0
 Call HighlightBackground(SSOleDBToBeUsedFor)
End Sub

Private Sub SSOleDBToBeUsedFor_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBToBeUsedFor.DroppedDown Then SSOleDBToBeUsedFor.DroppedDown = True
End Sub

Private Sub SSOleDBToBeUsedFor_LostFocus()
Call NormalBackground(SSOleDBToBeUsedFor)
End Sub

Private Sub SSOleDBToBeUsedFor_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBToBeUsedFor.Text)) > 0 And Not SSOleDBToBeUsedFor.IsItemInList Then
  Cancel = True
   SSOleDBToBeUsedFor.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBToCamChartFQA_InitColumnProps()
SSOleDBToCamChartFQA.DroppedDown = True
End Sub

Private Sub SSOleDBToCamChartFQA_KeyPress(KeyAscii As Integer)
If Not SSOleDBToCamChartFQA.DroppedDown Then SSOleDBToCamChartFQA.DroppedDown = True
SSOleDBToCamChartFQA.DroppedDown = True

End Sub

Private Sub SSOleDBToLocationFQA_InitColumnProps()
SSOleDBToLocationFQA.DroppedDown = True
End Sub

Private Sub SSOleDBToLocationFQA_KeyPress(KeyAscii As Integer)
If Not SSOleDBToLocationFQA.DroppedDown Then SSOleDBToLocationFQA.DroppedDown = True
SSOleDBToLocationFQA.DroppedDown = True

End Sub

Private Sub SSOleDBtoUSChartFQA_KeyUp(KeyCode As Integer, Shift As Integer)
If Not SSOleDBtoUSChartFQA.DroppedDown Then SSOleDBtoUSChartFQA.DroppedDown = True
SSOleDBtoUSChartFQA.DroppedDown = True
End Sub

Private Sub SSOleDBUnit_DropDown()
Dim str As String
On Error GoTo Handler
If Len(ssdcboCommoditty.Text) > 0 And chk_FrmStkMst.value = 1 Then
        If objUnits Is Nothing Then Set objUnits = Mainpo.PoUnits
       
        objUnits.StockNumber = Trim$(ssdcboCommoditty.Text)
       
        SSOleDBUnit.RemoveAll
        
        If RsUNits Is Nothing Then
          If lookups Is Nothing Then Set lookups = Mainpo.lookups
          Set RsUNits = lookups.GetAllUnits
        End If
        RsUNits.MoveFirst
        RsUNits.Find ("uni_code='" & Trim$(objUnits.PrimaryUnit) & "'")
        SSOleDBUnit.AddItem objUnits.PrimaryUnit & ";" & RsUNits("uni_desc")
       
       If Not Trim$(objUnits.PrimaryUnit) = Trim$(objUnits.SecondaryUnit) Then
            RsUNits.MoveFirst
            RsUNits.Find ("uni_code='" & Trim$(objUnits.SecondaryUnit) & "'")
            SSOleDBUnit.AddItem objUnits.SecondaryUnit & ";" & RsUNits("uni_desc")
                      
        End If


        

        
End If
Exit Sub
Handler:
 MsgBox "Please check the Units.Error in processing units."
 Err.Clear
End Sub

Private Sub SSOleDBUnit_LostFocus()
Call NormalBackground(SSOleDBUnit)
End Sub

Private Sub SSOleDBUnit_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBUnit.Text)) > 0 Then
   
   'If SSOleDBUnit.IsItemInList Then
    If Len(Trim$(ssdcboCommoditty.Text)) > 0 And Len(txt_Requested) > 0 Then
''''                  If Not objUnits Is Nothing Then
''''                      Set objUnits = MainPO.PoUnits
''''                      objUnits.StockNumber = Trim$(ssdcboCommoditty.text)
''''                   End If
                
                If IsPrimQuantLessThanONE = False Then
                   txt_Requested.SetFocus
                   'Cancel = True
                End If
                       
   End If
   Else
      
      If chk_FrmStkMst.value = 1 And Len(ssdcboCommoditty) > 0 Then
        Cancel = True
        
        MsgBox "Unit missing", , "Imswin"
        SSOleDBUnit.SetFocus
       End If
   End If
   
'End If
End Sub

Private Sub SSoleEccnNo_Click()
SSoleEccnNo.Tag = Trim(UCase(SSoleEccnNo.Columns(0).Text))
End Sub

Private Sub SSoleEccnNo_DropDown()
Call FillEccnCombos(lookups)
End Sub

Private Sub SSoleEccnNo_GotFocus()
Call HighlightBackground(SSoleEccnNo)
End Sub

Private Sub SSoleEccnNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If FormMode <> mdvisualization Then
    If Not SSoleEccnNo.DroppedDown Then SSoleEccnNo.DroppedDown = True
 End If
End Sub

Private Sub SSoleEccnNo_KeyPress(KeyAscii As Integer)
If FormMode <> mdvisualization Then SSoleEccnNo.MoveNext
End Sub

Private Sub SSoleEccnNo_LostFocus()
Call NormalBackground(SSoleEccnNo)
End Sub

Private Sub SSoleEccnno_Validate(Cancel As Boolean)

If chk_usexportLI.value = 0 Then Exit Sub

If SSoleEccnNo.IsItemInList = False Then

        MsgBox "Eccn # does not exist in the list, please select a valid one.", , "Imswin"
        SSoleEccnNo.SetFocus
        Cancel = True
        
End If
   
End Sub
Private Sub SSoleSourceofInfo_Click()
SSOleSourceofinfo.Tag = Trim(UCase(SSOleSourceofinfo.Columns(0).Text))
End Sub

Private Sub SSoleSourceofInfo_DropDown()
Call FillSourceOfinfoCombos(lookups)
End Sub

Private Sub SSoleSourceofInfo_GotFocus()
Call HighlightBackground(SSOleSourceofinfo)
End Sub

Private Sub SSoleSourceofInfo_KeyDown(KeyCode As Integer, Shift As Integer)
 If FormMode <> mdvisualization Then
    If Not SSOleSourceofinfo.DroppedDown Then SSOleSourceofinfo.DroppedDown = True
 End If
End Sub

Private Sub SSoleSourceofInfo_KeyPress(KeyAscii As Integer)
If FormMode <> mdvisualization Then SSOleSourceofinfo.MoveNext
End Sub

Private Sub SSoleSourceofInfo_LostFocus()
Call NormalBackground(SSOleSourceofinfo)
End Sub


Private Sub SSoleSourceofInfo_Validate(Cancel As Boolean)

If chk_usexportLI.value = 0 Then Exit Sub

If SSOleSourceofinfo.IsItemInList = False Then

        MsgBox "Source Of Info does not exist in the list, please select a valid one.", , "Imswin"
        SSOleSourceofinfo.SetFocus
        Cancel = True
        
End If
   
End Sub
Private Sub st_Completed(Cancelled As Boolean, Terms As String)
On Error Resume Next

    If Not Cancelled Then
        'txtClause.SelText = Terms
        txtClause.Text = txtClause.Text & Terms
        txtClause.SelStart = Len(txtClause)
        Terms = txtClause.Text
        POClause.Clause = Terms
        
    End If
    
    Set st = Nothing
    
End Sub

Private Sub chk_FrmStkMst_Click()
mIsPoItemsComboLoaded = False
End Sub

Private Sub cmd_Add_Click()
On Error GoTo errorHandler
If (OptEmail.value = True Or OptFax.value = True) Then
    
        If Len(Trim$(txt_Recipient)) > 0 Then
               txt_Recipient = UCase(txt_Recipient)
               
               If OptEmail.value = True Then txt_Recipient = (txt_Recipient)
               If OptFax.value = True Then txt_Recipient = (txt_Recipient)
            Call AddRecepient(txt_Recipient, , False)

            dgRecipientList.Refresh ' JCG 2008/07/30


            txt_Recipient = ""
        'Else
        '    dgRecepients_DblClick
        End If
 Else
    MsgBox "Please check Email or Fax.", vbInformation, "Imswin"
    
 End If
    Exit Sub
errorHandler:
    MsgBox "Error in cmd_Add_Click: " + Err.Description
    
End Sub

Private Sub cmd_Addterms_Click()
On Error Resume Next
 Me.MousePointer = vbHourglass
    ' Load frm_ShipTerms
    If st Is Nothing Then Set st = New frm_ShipTerms
    st.Show
    st.txt_Description.SetFocus
    'frm_ShipTerms.Show
    If Err Then Err.Clear
   Me.MousePointer = vbArrow
End Sub

Private Sub CmdcopyLI_Click(Index As Integer)
  Me.MousePointer = vbHourglass
 Select Case (Index)
 
   Case 0
    
    'If rsPOITEM.State <> adStateClosed Then
    
    If PoItem.Count > 0 Then
     
      If FormMode = mdCreation Or FormMode = mdModification Then
    
     
       Load FrmCopyPOItems
       FrmCopyPOItems.Show
     
      End If
    
     End If
     
    
   
   Case 1
       
       
   'If PORemark.State <> adStateClosed Then
    
    If PORemark.Count > 0 Then
     
      If FormMode = mdCreation Or FormMode = mdModification Then
     
        Load FrmCopyPORemarks
        FrmCopyPORemarks.Show
     
      End If
      
     End If
     

    
   Case 2
   
    '  If POClause.State <> adStateClosed Then
    
    If POClause.Count > 0 Then
     
      If FormMode = mdCreation Or FormMode = mdModification Then
     
       Load FrmCopyPOClause
       FrmCopyPOClause.Show
     
      End If
    
     End If
     
  
   
 End Select
 Me.MousePointer = vbArrow
End Sub

Private Sub cmdRemove_Click()
Dim x As Integer
On Error GoTo errorHandler

' JCG 2008/7/30
'If Len(dgRecipientList.SelBookmarks(0)) = 0 Then
'    MsgBox "Please make a selection first.", vbInformation, "Imswin"
'    Exit Sub
' End If
'-------------------

If FormMode = mdCreation Then

    dgRecipientList.DeleteSelected
    
ElseIf FormMode = mdModification Then
    If IsNothing(lookups) Then Set lookups = Mainpo.lookups
      x = lookups.CanUserDeleteRecepient(Poheader.Ponumb, IIf(Poheader.revinumb = 0, 0, Poheader.Originalrevinumb - 1), dgRecipientList.Columns(0).Text)
         
         If x = 0 Then
                If Poheader.revinumb = 1 And Poheader.Originalrevinumb = 0 Then
                        If PoReceipients.EditMode = 2 Then
                               dgRecipientList.DeleteSelected
                           Else
                               MsgBox "Can not Delete the Recepient ,it is being carried over from the Previous Revisions.", vbInformation, "Imswin"
                           End If
                 Else
                               dgRecipientList.DeleteSelected
                 End If
                 
         ElseIf x = 1 Then
                                 MsgBox "Can not Delete the Recepient ,it is being carried over from the Previous Revisions.", vbInformation, "Imswin"
                
         End If
                   
End If

Exit Sub
errorHandler:
    MsgBox "There is an error: " + Err.Description
End Sub

Private Sub dgRecepients_DblClick()
On Error Resume Next
    Call AddRecepient(dgRecepients.Columns(1).value, , True)
    
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim x As Integer

If FormMode = mdModification Or FormMode = mdCreation Then
   x = MsgBox("You will lose any unsaved changes you might have made. Are you sure you want to exit?", vbCritical + vbYesNo, "Imswin")
Else
  x = vbYes
End If

If x = vbYes Then

            frm_NewPurchase.TxtToStocktypeFQA.Enabled = False
            If Not rsDOCTYPE Is Nothing Then Set rsDOCTYPE = Nothing
            If Not st Is Nothing Then Set st = Nothing
            If Not comsearch Is Nothing Then Set comsearch = Nothing
            If Not RsEmailFax Is Nothing Then Set RsEmailFax = Nothing
            If Not rsDOCTYPE Is Nothing Then Set rsDOCTYPE = Nothing
            If Not objUnits Is Nothing Then Set objUnits = Nothing
            If Not RsUNits Is Nothing Then Set RsUNits = Nothing
            If Not GRsDoctype Is Nothing Then Set GRsDoctype = Nothing
            
            If deIms.rsPonumb.State = 1 Then Call deIms.rsPonumb.Close
            If deIms.rsSHIPPER.State = 1 Then Call deIms.rsSHIPPER.Close
            If deIms.rsCURRENCY.State = 1 Then Call deIms.rsCURRENCY.Close
            If deIms.rsPRIORITY.State = 1 Then Call deIms.rsPRIORITY.Close
            If deIms.rsTermDelivery.State = 1 Then Call deIms.rsTermDelivery.Close
             
            If deIms.rsActiveSupplier.State = 1 Then Call deIms.rsActiveSupplier.Close
            If deIms.rsTermCondition.State = 1 Then Call deIms.rsTermCondition.Close
            
            If deIms.rsCOMPANY.State = 1 Then Call deIms.rsCOMPANY.Close
            
            If deIms.rsActiveOriginator.State = 1 Then Call deIms.rsActiveOriginator.Close
            If deIms.rsActiveTbu.State = 1 Then Call deIms.rsActiveTbu.Close
            If deIms.rsSERVCODECAT.State = 1 Then Call deIms.rsSERVCODECAT.Close
            If deIms.rsActiveShipTo.State = 1 Then Call deIms.rsActiveShipTo.Close
            If deIms.rsActiveCompany.State = 1 Then Call deIms.rsActiveCompany.Close
            If deIms.rsCompanyLocations.State = 1 Then Call deIms.rsCompanyLocations.Close
            If deIms.rsActiveStockmasterLookup.State = 1 Then Call deIms.rsActiveStockmasterLookup.Close
            
            If Not Poheader Is Nothing Then Set Poheader = Nothing
            If Not PoReceipients Is Nothing Then Set PoReceipients = Nothing
            If Not PoItem Is Nothing Then Set PoItem = Nothing
            If Not PORemark Is Nothing Then Set PORemark = Nothing
            If Not POClause Is Nothing Then Set POClause = Nothing
            If Not Mainpo Is Nothing Then Set Mainpo = Nothing
            If Not lookups Is Nothing Then Set lookups = Nothing
            If Not POFqa Is Nothing Then Set POFqa = Nothing
            
            mIsPoheaderCombosLoaded = False
            mIsDocTypeLoaded = False
            mIsPoNumbComboLoaded = False
            mIsPoItemsComboLoaded = False
            GToFQAComboLoaded = False
            
Else

   Cancel = True
End If


 Dim imsLock As imsLock.Lock
 Set imsLock = New imsLock.Lock
 Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


End Sub

Private Sub NavBar1_BeforeSaveClick()
Dim mpo As String
Dim IsErrorinTransaction As Integer
Dim ErrorSource As String
Dim str As String
Dim madeARevision As Boolean
Dim NoSetPOnumbErr As Integer 'AN
Dim poFqaErrors As Boolean
 Dim imsLock As imsLock.Lock

On Error GoTo Handler

Select Case sst_PO.Tab
        Case 0
        
        Dim Response_ As Integer  'AM
            
            If Len(Trim(ssOleDbPO)) = 0 And FormMode = mdCreation Then 'AM
            
                    Response_ = MsgBox("You did not enter any Transaction number. Would you like" & _
                      " the system to generate one for you ?", vbInformation + vbYesNo, "Imswin") 'AM
                         
                    If Response_ = vbNo Then 'AM
                        NavBar1.SaveEnabled = True
                        Exit Sub 'AM
                    
                    End If 'AM
                    
            End If 'AM

            Me.MousePointer = vbHourglass
            If CheckPoFields Then
                    
                    Load FrmShowApproving
                    Screen.MousePointer = 11
                    FrmShowApproving.Top = 4620
                    FrmShowApproving.Left = 3330
                    FrmShowApproving.Width = 3000
                    FrmShowApproving.Height = 1140
                    
                    FrmShowApproving.Show
                    FrmShowApproving.Label2.Caption = " Saving PO ......"
                    
                    Screen.MousePointer = 11
                    FrmShowApproving.Refresh
                    Screen.MousePointer = 11
                                         
                     
                     Dim PoHeaderErrors As Boolean
                     Dim POITEMErrors As Boolean
                     Dim PoremarksErrors As Boolean
                     Dim poClauseErrors As Boolean
                     Dim poRecepientsErrors As Boolean
                     
                      PoHeaderErrors = True
                      POITEMErrors = True
                      poRecepientsErrors = True
                      PoremarksErrors = True
                      poClauseErrors = True
                    
                     SaveToPOHEADER
                     
                     deIms.cnIms.Errors.Clear
                     IsErrorinTransaction = 0
                     deIms.cnIms.BeginTrans
                     IsErrorinTransaction = 1
                     
                     If mSaveToPoRevision Then InsertPoRevision (Poheader.Ponumb): madeARevision = True
                     
                     mSaveToPoRevision = False
                     
                     'If the user has entered a Po number then use that, else make use of the AutoGenerated _
                     one which is stored in the PONUMB property of all the PO objects
                  
                     If FormMode = mdCreation Then 'AM
                        
                                          
                                   NoSetPOnumbErr = Mainpo.SetPONUMBforAllPoObjects(Trim$(ssOleDbPO)) 'AM
                                          
                                      If NoSetPOnumbErr = 0 Then ssOleDbPO = Poheader.Ponumb 'AM
                                    
                                                     
                     Else 'AM
                     
                            NoSetPOnumbErr = 0 'AM
                            
                     End If 'AM
                     
                     
                    
                
                      
                        Select Case (NoSetPOnumbErr)
                        
                        'This is the case when the error is unidentified
                        
                        Case 1
                        
                            MsgBox "Could not save the Transaction. Error occurred while trying to implement autonumbering. Please try again.", vbOKOnly, "Imswin" 'AM
                            GoTo Handler
                            Exit Sub 'AM
                            
                        'This is the case when there is no AutoNumbering is associated with that perticular Document Type
                            
                        Case 2
                            
                            MsgBox "Could not save the Transaction. There is no Auto-numbering associated with this document type. Please fill in a Transaction Number and try saving again.", vbCritical + vbOKOnly, "Imswin" 'AM
                                                deIms.cnIms.RollbackTrans
                                                NavBar1.SaveEnabled = True
                                                Screen.MousePointer = vbArrow
                                                 Unload FrmShowApproving
                             
                            Exit Sub 'AM
                            
                       End Select
                            
                     'End If 'AM
                          
                     mpo = Poheader.Ponumb  'AM
                          
                     ErrorSource = "PoHeader"
                      
                     PoHeaderErrors = Poheader.SAVE
                      
                      If PoHeaderErrors = True Then
                         WriteStatus ("Poheader saved successfully.")
                      Else
                         WriteStatus ("Errors Occurred While Trying to Save Poitems")
                      End If
                      
                       If Not PoItem Is Nothing Then
                           
                           ErrorSource = "Poitem"
                           
                           POITEMErrors = PoItem.Update
                           
                           If madeARevision = True And POITEMErrors = True Then
                                
                                Dim rs As New ADODB.Recordset
                                
                                rs.Source = "UPDATE POitem SET poi_stasliit = 'OH' WHERE poi_npecode ='" & deIms.NameSpace & "' and poi_ponumb ='" & Poheader.Ponumb & "' and  poi_stasliit = 'OP'"
                                
                                rs.ActiveConnection = deIms.cnIms
                                
                                rs.Open
                                
                                If Err.number <> 0 Then Err.Clear: POITEMErrors = False
                                
                                madeARevision = False
                                
                           End If
                           
                          If POITEMErrors = True Then
                            WriteStatus ("Poitems Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Poitems")
                          End If
                       End If
                       
                       madeARevision = False
                       
                      If Not PoReceipients Is Nothing Then
                            ErrorSource = "PoRecepients"
                            
                            poRecepientsErrors = PoReceipients.Update
                         If poRecepientsErrors = True Then
                           WriteStatus ("Recipients Saved Successfully")
                            CmdAddSupEmail.Tag = ""
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Recipients")
                          End If
                      End If
                      
                      If Not PORemark Is Nothing Then
                             ErrorSource = "Poremark"
                             
                             PoremarksErrors = PORemark.Update
                             If PoremarksErrors = True Then
                           WriteStatus ("Remarks Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Remarks")
                          End If
                      End If
                             
                      If Not POClause Is Nothing Then
                            ErrorSource = "Poclause"
                            poClauseErrors = POClause.Update
                          If poClauseErrors = True Then
                           WriteStatus ("Clause Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Clause")
                          End If
                      End If
                      
                     If Not POFqa Is Nothing Then
                            ErrorSource = "PoFqa"
                            poFqaErrors = POFqa.Update
                          If poFqaErrors = True Then
                           WriteStatus ("FQA Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save FQA")
                          End If
                      End If
                      
                     If PoHeaderErrors = True And POITEMErrors = True And poRecepientsErrors = True And PoremarksErrors = True And poClauseErrors = True And poFqaErrors = True Then
                         
                         deIms.cnIms.CommitTrans
                         
                         
                          If Response_ = vbYes Then
                                MsgBox "The Transaction number you just saved is " & mpo & " .", vbOKOnly, "Imswin"
                          End If
                        
                           Poheader.Requery
                           'So that we are on the Po which the user has just saved
                           'LoadFromPOHEADER
                           If Poheader.Move(mpo) = True Then Call LoadFromPOHEADER
                           mIsPoNumbComboLoaded = False
                           FrmShowApproving.Label2.Caption = "Transaction Order # " & mpo & " saved successfully"
                          'So that we are on the Po which the user has just saved
                        
                           WriteStatus ("")
                       Else
                         
                         deIms.cnIms.RollbackTrans
                         FrmShowApproving.Label2.Caption = "Errors Occured. Could not save the transaction order."
                         MsgBox "Errors Occured.Could Not Save The Transaction Order.", vbCritical, "Ims" ' Added on 06/25
                         WriteStatus ("Rolling Back the Transaction")
                         Poheader.CancelUpdate: LoadFromPOHEADER
                         
                         WriteStatus ("")
                       
                       End If
                       
                      
                          Set PoItem = Nothing
                          Set PoReceipients = Nothing
                          Set PORemark = Nothing
                          Set POClause = Nothing
                          
                          CmdConvert.Enabled = False

                    CheckErrors = True
                    
                    FormMode = ChangeMode(mdvisualization)
                    
                    If FormMode = mdvisualization Then
                       
                       NavBar1.NewEnabled = True
                       NavBar1.NextEnabled = True
                       NavBar1.PreviousEnabled = True
                       NavBar1.LastEnabled = True
                       NavBar1.FirstEnabled = True
                       
                       NavBar1.EditEnabled = True
                       NavBar1.CancelEnabled = False
                       ssOleDbPO.Enabled = True
                       NavBar1.SaveEnabled = False
                       NavBar1.DeleteEnabled = True
                       NavBar1.EMailEnabled = True 'JCG 2008/11/14
                    End If
                    
                    Screen.MousePointer = vbArrow
                    Unload FrmShowApproving
                    
                     Set imsLock = New imsLock.Lock
                     Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

                       
                    
          Else
                  NavBar1.SaveEnabled = True
                    
          End If
 End Select
 
   Me.MousePointer = vbArrow
   CmdAddSupEmail.Tag = ""
   

   
Exit Sub
Handler:
  str = Err.Description
 
 Err.Clear
  Unload FrmShowApproving
  Err.Clear
 MsgBox "Unknown errors occurred while saving " & ErrorSource & " .Could not save the po." & vbCrLf & "Error description   " & str, vbCritical, "Imswin"

If IsErrorinTransaction = 1 Then deIms.cnIms.RollbackTrans
IsErrorinTransaction = 0
Err.Clear

FormMode = ChangeMode(mdvisualization)
                    
                    If FormMode = mdvisualization Then
                       
                       NavBar1.NewEnabled = True
                       NavBar1.NextEnabled = True
                       NavBar1.PreviousEnabled = True
                       NavBar1.LastEnabled = True
                       NavBar1.FirstEnabled = True
                       
                       NavBar1.EditEnabled = True
                       NavBar1.CancelEnabled = False
                       ssOleDbPO.Enabled = True
                       NavBar1.SaveEnabled = False
                       NavBar1.DeleteEnabled = True
                       NavBar1.EMailEnabled = True 'JCG 2008/11/14
                    End If
                    
                    CmdAddSupEmail.Tag = ""
                    Screen.MousePointer = vbArrow



 Set imsLock = New imsLock.Lock
 Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode



End Sub

Private Sub NavBar1_OnCancelClick()

If FormMode = mdModification Or FormMode = mdCreation Then

  
   'Cancelling all the Changes made by the user
        Select Case sst_PO.Tab
             
             Case 0
             
             CmdConvert.Enabled = False
             ssOleDbPO.Enabled = True
              
              If FormMode = mdModification Then
                  
                  Call LoadFromPOHEADER
                  
                   If Not PoReceipients Is Nothing Then
                      
                      Set PoReceipients = Nothing
                      
                   End If
                   
                    Dim imsLock As imsLock.Lock
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
                    
              ElseIf FormMode = mdCreation Then
                  
                  Poheader.CancelUpdate
                  POFqa.CancelUpdate
                  Call CleanFROMFQA
                  If Not PoReceipients Is Nothing Then Set PoReceipients = Nothing
                  If Not PoItem Is Nothing Then Set PoItem = Nothing
                  If Not PORemark Is Nothing Then Set PORemark = Nothing
                  If Not POClause Is Nothing Then Set POClause = Nothing
                  If Not POFqa Is Nothing Then Set POFqa = Nothing
                  Poheader.MoveFirst
                  Call LoadFromPOHEADER
              
              End If
              
              CmdAddSupEmail.Tag = ""
              FormMode = ChangeMode(mdvisualization)
             mSaveToPoRevision = False
             
             If FormMode = mdvisualization Then
                       
                       NavBar1.NewEnabled = True
                       NavBar1.NextEnabled = True
                       NavBar1.PreviousEnabled = True
                       NavBar1.LastEnabled = True
                       NavBar1.FirstEnabled = True
                       NavBar1.SaveEnabled = False
                       NavBar1.EditEnabled = True
                       NavBar1.CancelEnabled = False
                       NavBar1.DeleteEnabled = True
                       ssOleDbPO.Enabled = True
                       NavBar1.EMailEnabled = True 'JCG 2008/11/14
                    End If
             
             
             Case 1
             Case 2
              
              If FormMode = mdModification And PoItem.EditMode = 0 Then
                  Call LoadFromPOITEM
              ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And PoItem.EditMode = 2) Then
              'CAncels only the Latest REcord
                  PoItem.CancelUpdate
                  POFqa.CancelUpdateline
                  
                  If PoItem.Count > 0 Then
                      POFqa.MoveLineTo (PoItem.Linenumb)
                      Call LoadFromPOITEM
                      Call LoadFromTOFQA
                      FirstTimeAssignmentsPOITEM
                  Else
                    ClearAllPoLineItems
                    CleanToFQAControls
                  End If
              End If
              If PoItem.Count = 0 And FormMode <> mdvisualization Then
                    fra_LineItem.Enabled = False
              Else
                    fra_LineItem.Enabled = True
              End If
             Case 3
               If PORemark.Count > 0 Then
                        If FormMode = mdModification And PORemark.EditMode = 0 Then
                             Call LoadFromPORemarks
                        ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And PORemark.EditMode = 2) Then
                              PORemark.CancelUpdate
                           
                           If PORemark.Count > 0 Then
                               Call LoadFromPORemarks
                               HandleEdittingOfRemarks
                           Else
                               Call ClearPoRemarks
                           End If
                           mCheckRemarks = True
                        End If
                End If
             Case 4
                If POClause.Count > 0 Then
                
                    If FormMode = mdModification And POClause.EditMode = 0 Then
                       Call LoadFromPOClause
                       
                    ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And POClause.EditMode = 2) Then
                        POClause.CancelUpdate
                        
                        If POClause.Count > 0 Then
                           Call LoadFromPOClause
                           HandleEdittingOfClause
                        Else
                           Call ClearPoclause
                        End If
                        MCheckClause = True
                    End If
                    
               End If
         End Select
 End If



End Sub

Private Sub NavBar1_OnEditClick()

                                                    'jawdat, start copy
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
                        
                        'jawdat, end copy

    If Trim$(Poheader.stas) = "CA" Or Trim$(Poheader.stas) = "CL" Then
         MsgBox " Can not Edit this Document ,It is Closed"
         GoTo CANNOTEDIT
    End If
     
     
     
    If deIms.CanUserEditDocType(CurrentUser, Trim$(Poheader.Docutype)) Then
         
         If Len(Trim$(LblAppBy)) > 0 And CanDocTypeBeRevised(Poheader.Docutype) Then
                
                If MsgBox(" You will Create a new Revision. Do You want to Continue ?", vbYesNo) = vbYes Then
                    
                    LblRevNumb.Caption = IIf(Len(LblRevNumb.Caption) = 0, 0, CInt(LblRevNumb.Caption) + 1)
                    LblRevDate = Format(Now(), "MM/DD/YY")
                    LblAppBy = ""
                    LblDateSent = ""
                    mSaveToPoRevision = True
                Else
                    GoTo CANNOTEDIT
                End If
                 
        ElseIf Len(Trim$(LblAppBy)) > 0 And CanDocTypeBeRevised(Poheader.Docutype) = False Then
                 
            Call MsgBox("The Transaction Order can not be modified. The Document type does not allow any Revisions.", vbInformation, "Imswin")
             
            GoTo CANNOTEDIT
        
        End If
           
           If locked = False Then locked = True
        
           Me.MousePointer = vbHourglass
           
           
           FormMode = ChangeMode(mdModification)
        
    
             FirstTimeAssignmentsHeader
    
           
           Select Case sst_PO.Tab
             Case 0
                  If mIsDocTypeLoaded = False Then LoadDocType
                  If mIsPoheaderCombosLoaded = False Then
                    CheckErrors = LoadPoHeaderCombos
    
                    'Disabling the PO so that the user can not Navigate in Create Mode
                    ssOleDbPO.Enabled = False
                  End If
                  
                Me.MousePointer = vbArrow
          End Select
              
         
                       
                       NavBar1.NewEnabled = False
                       NavBar1.NextEnabled = False
                       NavBar1.PreviousEnabled = False
                       NavBar1.LastEnabled = False
                       NavBar1.FirstEnabled = False
                       NavBar1.SaveEnabled = True
                       NavBar1.EditEnabled = False
                       NavBar1.CancelEnabled = True
                       NavBar1.DeleteEnabled = False
                       NavBar1.EMailEnabled = False 'JCG 2008/10/14
                       
                      
     Else
        MsgBox " User can not Edit this Document "
             
     End If
    
CANNOTEDIT:




     '
End Sub

Private Sub NavBar1_OnFirstClick()
Select Case sst_PO.Tab
        Case 0
            If Poheader.MoveFirst Then LoadFromPOHEADER
            If POFqa.GetFQAInfo(ssOleDbPO) Then
    
                    LoadFromFROMFQA
            Else
    
                    CleanFROMFQA
        
            End If
    
        Case 1
        Case 2
          If Not FormMode = mdvisualization And PoItem.BOF = False Then
            If CheckLIFields = True Then
                       SaveToPOITEM
                       SaveToTOFQA
                    Else
                       Exit Sub
                    End If
          End If
          
               If PoItem.MoveFirst Then
                  FirstTimeAssignmentsPOITEM
                  LoadFromPOITEM
                  
               End If
               
                If POFqa.MoveLineTo(PoItem.Linenumb) Then
                  LoadFromTOFQA
                End If
        Case 3
        
              If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtRemarks)) > 0 Then
                             savetoPORemarks
                          Else
                             MsgBox "Remarks can not be empty."
                             Exit Sub
                          End If
               End If
               
               
            If PORemark.MoveFirst Then
               LoadFromPORemarks
               HandleEdittingOfRemarks
            End If
        Case 4
        
              If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtClause)) > 0 Then
                             savetoPOclause
                          Else
                             MsgBox "Clause can not be empty."
                             Exit Sub
                          End If
               End If
           
            If POClause.MoveFirst Then
               LoadFromPOClause
               HandleEdittingOfClause
            End If
 End Select
End Sub

Private Sub NavBar1_OnLastClick()
 
 Select Case sst_PO.Tab
        Case 0
            If Poheader.MoveLast Then LoadFromPOHEADER
            If POFqa.GetFQAInfo(ssOleDbPO) Then
    
                    LoadFromFROMFQA
            Else
    
                    CleanFROMFQA
        
            End If
    
        Case 1
        Case 2
            
               If FormMode <> mdvisualization And PoItem.EOF = False Then
                    If CheckLIFields = True Then
                       SaveToPOITEM
                    Else
                       Exit Sub
                    End If
                    
                    SaveToTOFQA
                    
                End If
                
                If PoItem.MoveLast Then
                  FirstTimeAssignmentsPOITEM
                  LoadFromPOITEM
                End If
                
                
                'f POFqa.MoveLast Then
                If POFqa.MoveLineTo(PoItem.Linenumb) Then
                  LoadFromTOFQA
                End If
                
        Case 3
        
           If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtRemarks)) > 0 Then
                             savetoPORemarks
                          Else
                             MsgBox "Remarks can not be empty."
                             Exit Sub
                          End If
            End If
        
        
            If PORemark.MoveLast Then
                LoadFromPORemarks
                HandleEdittingOfRemarks
            End If
        Case 4
          
             If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtClause)) > 0 Then
                             savetoPOclause
                          Else
                             MsgBox "Clause can not be empty."
                             Exit Sub
                          End If
               End If
        
        
            If POClause.MoveLast Then
               LoadFromPOClause
               HandleEdittingOfClause
            End If
 End Select

 
End Sub

Private Sub NavBar1_OnNewClick()
  
  Select Case (sst_PO.Tab)
   
   Case 0
   
             'JCG 2008/9/19
          ClearPoReceipients
            Set PoReceipients = Nothing
        '----------------------
        
        ssOleDbPO.SetFocus
        newSupplier = True 'JDCG 2008/1/20
        CmdConvert.Enabled = True
        chk_FrmStkMst.value = 0
        FormMode = ChangeMode(mdCreation)
          
          If mIsDocTypeLoaded = False Then
              
              SSOleDBDocType.Text = ""
              
              LoadDocType
           
           End If
         
         If mIsPoheaderCombosLoaded = False Then
           
           CheckErrors = LoadPoHeaderCombos
         
         End If
        
        If mIsPoheaderCombosLoaded = True Then
           
           CheckErrors = Poheader.AddNew
           
           FirstTimeAssignmentsHeader
           
           If CheckErrors = False Then
           
               MsgBox "Some error has occurred, Could not create a New Transaction. Please close the form and try once more, also check if AUTONUMBERING is also configured." 'AM
           
           Else
           
                Set POFqa = Nothing
                      If POFqa Is Nothing Then Set POFqa = Mainpo.FQA
                      CheckErrors = POFqa.AddNew
                                
                                If CheckErrors = False Then
                                
                                    MsgBox "Error occurred while trying to create a new record for FQA. Please close the form try again." & Err.Description, vbCritical, "Ims"
                                
                                Else
                                
                                     Call FillFromFQAControls(SSOleDBcompany.Tag, "Purch")
                                     Call SavetoFROMFQA
                                     POFqa.Ponumb = Poheader.Ponumb
                                     SetInitialVAluesPoHeader
                                     ToggleNavButtons (mdCreation)
                                End If
                       End If
               
               mLoadMode = LoadingPOheader
               
                If ConnInfo.Eccnactivate = Constyes Or ConnInfo.Eccnactivate = ConstOptional Then
                    chk_USExportH.value = IIf(ConnInfo.usexport = True, 1, 0)
                    
                ElseIf ConnInfo.Eccnactivate = Constno Then
                    chk_USExportH.value = 0
                    
                End If
                    
               mLoadMode = NoLoadInProgress
               
            End If
  
  Case 1

        
  Case 2

        If mIsPoItemsComboLoaded = False Then
            CheckErrors = LoadPoItemCombos
        End If
      
         If mIsPoItemsComboLoaded = True Then
           
            If PoItem.Count > 0 Then
                    
                If CheckLIFields = True Then
                    SaveToPOITEM
                    SaveToTOFQA
                Else
                    Exit Sub
                End If
                
            End If
            
                CheckErrors = PoItem.AddNew
                
                If CheckErrors = False Then
                   
                   MsgBox "Error In Adding A POITEM"
                
                Else
                   
'                   CheckErrors = POFqa.AddNewLine
                   
                   'It means it is one of those OLD POs without any FQA
                   
                  If POFqa.Count > 0 Then
                            
                            CheckErrors = POFqa.AddNew
                           
                           Call InitializeNewTOFQARecord(SSOleDBcompany.Tag, SSOleDBInvLocation.Tag)
                           
                  Else
                  
                  CheckErrors = True
                           
                  End If
                           
                           If CheckErrors = False Then
                           
                                MsgBox "Error In Adding An FQA"
                        
                           Else
                           
                 
                   
                        ClearAllPoLineItems
                        SetInitialVAluesPOITEM
                        FirstTimeAssignmentsPOITEM
                        
                        If ssdcboCommoditty.Enabled = True Then
                             
                             ssdcboCommoditty.SetFocus
                             Call HighlightBackground(ssdcboCommoditty)
                        
                        End If
                        
                   End If
                   
                End If
           
        End If
        
        
      fra_LineItem.Enabled = True
        
        
         
  Case 3
        If PORemark.Count > 0 Then
            
                If Len(Trim$(txtRemarks)) > 0 Then
                   savetoPORemarks
                Else
                   MsgBox "Remarks can not be empty."
                   Exit Sub
                End If
            
        End If
        
        
        
        
        CheckErrors = PORemark.AddNew
        If CheckErrors = False Then
             MsgBox "Error In Adding A Remarks"
        Else
             ClearPoRemarks
             txtRemarks.Text = "REVISION " & Poheader.revinumb & "********************************************************" & vbCrLf
             txtRemarks.SelStart = Len(txtRemarks)
             txtRemarks.locked = False
             txtRemarks.SetFocus
             
        End If
  Case 4
  
         If POClause.Count > 0 Then
            
                         If Len(Trim$(txtClause)) > 0 Then
                             savetoPOclause
                          Else
                             MsgBox "Clause can not be empty."
                             Exit Sub
                          End If
          End If
               
          
        CheckErrors = POClause.AddNew
        If CheckErrors = False Then
             MsgBox "Error In Adding A Notes/Clause"
        Else
             ClearPoclause
             txtClause.Text = "REVISION " & Poheader.revinumb & "********************************************************" & vbCrLf
             txtClause.SelStart = Len(txtClause)
             txtClause.locked = False
             txtClause.SetFocus
        End If
        
        
 End Select
End Sub

Private Sub NavBar1_OnNextClick()

 Select Case sst_PO.Tab
        Case 0
            
            If Poheader.MoveNext Then LoadFromPOHEADER
            If POFqa Is Nothing Then Set POFqa = Mainpo.FQA
            If POFqa.GetFQAInfo(Poheader.Ponumb) Then
    
                    LoadFromFROMFQA
            Else
    
                    CleanFROMFQA
        
            End If
    
        Case 1
        Case 2
             If Not FormMode = mdvisualization And PoItem.EOF = False Then
                    If CheckLIFields = True Then
                       SaveToPOITEM
                       SaveToTOFQA
                    Else
                       Exit Sub
                    End If
              End If
            
               If PoItem.MoveNext Then
                  FirstTimeAssignmentsPOITEM
                  LoadFromPOITEM
               End If
               
               If POFqa Is Nothing Then Set POFqa = Mainpo.FQA
               'If POFqa.MoveNext Then
               If POFqa.MoveLineTo(PoItem.Linenumb) Then
                  
                  LoadFromTOFQA
                End If
                
                If POFqa.EOF = True Then POFqa.MoveLast
        Case 3
        
             If Not FormMode = mdvisualization Then
              If Len(Trim$(txtRemarks)) > 0 Then
                       savetoPORemarks
                    Else
                       MsgBox "Remarks can not be empty."
                       Exit Sub
                    End If
              End If
        
        
            If PORemark.MoveNext Then
              LoadFromPORemarks
              HandleEdittingOfRemarks
            End If
        Case 4
        
              If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtClause)) > 0 Then
                             savetoPOclause
                          Else
                             MsgBox "Clause can not be empty."
                             Exit Sub
                          End If
               End If
        
             If POClause.MoveNext Then
               LoadFromPOClause
               HandleEdittingOfClause
             End If
 End Select


End Sub

Private Sub NavBar1_OnPreviousClick()


 Select Case sst_PO.Tab
        Case 0
            If Poheader.MovePrevious Then LoadFromPOHEADER
            If POFqa.GetFQAInfo(ssOleDbPO) Then
    
                    LoadFromFROMFQA
            Else
    
                    CleanFROMFQA
        
            End If
    
        Case 1
        Case 2
            If Not FormMode = mdvisualization And PoItem.BOF = False Then
              If CheckLIFields = True Then
                       SaveToPOITEM
                       SaveToTOFQA
                    Else
                       Exit Sub
                    End If
             End If
            If PoItem.MovePrevious Then
               FirstTimeAssignmentsPOITEM
               LoadFromPOITEM
            End If
            
               'If POFqa.MovePrevious Then
               If POFqa.MoveLineTo(PoItem.Linenumb) Then
                  LoadFromTOFQA
                End If
        Case 3
        
           If Not FormMode = mdvisualization Then
              If Len(Trim$(txtRemarks)) > 0 Then
                       savetoPORemarks
                    Else
                       MsgBox " can not be empty"
                       Exit Sub
                    End If
              End If
        
           
            If PORemark.MovePrevious Then
              LoadFromPORemarks
              HandleEdittingOfRemarks
            End If
        Case 4
        
              If Not FormMode = mdvisualization Then
                          If Len(Trim$(txtClause)) > 0 Then
                             savetoPOclause
                          Else
                             MsgBox "Clause can not be empty."
                             Exit Sub
                          End If
               End If
        
        
            If POClause.MovePrevious Then
               LoadFromPOClause
               HandleEdittingOfClause
            End If
 End Select
   

End Sub

Private Sub opt_Email_Click()
On Error GoTo Handler
Dim co As MSDataGridLib.column

If lookups Is Nothing Then Set lookups = Mainpo.lookups

    Set co = dgRecepients.Columns(1)


    co.Caption = "Email Address"

    co.DataField = "phd_mail"


    dgRecepients.Columns(0).DataField = "phd_name"

     Set RsEmailFax = Nothing
     
     Set RsEmailFax = lookups.GetAddresses("ATEMAIL")
     Set dgRecepients.DataSource = RsEmailFax
      
         
    
     
     

   Exit Sub
        
     
Handler:
     Err.Raise Err.number, , Err.Description
     Err.Clear
     
End Sub

Private Sub opt_FaxNum_Click()
Dim co As MSDataGridLib.column

On Error GoTo Handler
If lookups Is Nothing Then Set lookups = Mainpo.lookups
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
    
     co.Caption = "Fax Numbers"
     co.DataField = "phd_faxnumb"
    
     dgRecepients.Columns(0).DataField = "phd_name"
     Set RsEmailFax = Nothing
     Set RsEmailFax = lookups.GetAddresses("ATFAX")
     
     Set dgRecepients.DataSource = RsEmailFax
     

Exit Sub
Handler:

  Err.Raise Err.number, , Err.Description
  Err.Clear
End Sub

Private Sub opt_SupFax_Click()
On Error Resume Next
Dim rs As ADODB.Recordset
Dim co As MSDataGridLib.column
    
    Set co = dgRecepients.Columns(1)
    co.Caption = "Supplier Email"
    co.DataField = "sup_mail"
    
    dgRecepients.Columns(0).DataField = "sup_name"
    
    Set rs = New ADODB.Recordset
    
    With rs
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        Set .ActiveConnection = deIms.cnIms
        .Open ("select sup_name, sup_mail from SUPPLIER where sup_npecode = '" & deIms.NameSpace & "' and sup_mail IS NOT NULL and len(sup_mail) > 3 order by 1")
        Set dgRecepients.DataSource = .DataSource
    End With
End Sub

Private Sub ssdcboCommoditty_Click()
Dim Eccnid As Integer
Dim Eccnno As String
Dim Sourceid As Integer
Dim Sourceno As String
Dim EcnLicense As Boolean
On Error GoTo Handler
  If Len(Trim$(ssdcboCommoditty.Text)) > 0 Then
    If lookups Is Nothing Then Set lookups = Mainpo.lookups
     Dim rsMANUFACTURER As ADODB.Recordset
    
        
        Set rsMANUFACTURER = lookups.GetManuFActurer(Trim$(ssdcboCommoditty))
        
        ssdcboManNumber.RemoveAll
        
        If Not rsMANUFACTURER.EOF Then
        
        '2012-9-30 juan
        Set ssdcboManNumber.DataSourceList = rsMANUFACTURER
        ssdcboManNumber.DataFieldList = rsMANUFACTURER.Fields(0).Name
'            Do While Not rsMANUFACTURER.EOF
'                  ssdcboManNumber.AddItem rsMANUFACTURER!stm_manucode & ";" & rsMANUFACTURER!stm_partnumb & ";" & rsMANUFACTURER!stm_estmpric
'                  rsMANUFACTURER.MoveNext
'            Loop
            
        End If
           
        Set rsMANUFACTURER = Nothing
        'This is a Global Variable which Stores the info about the Stock number.
        'Can use the Public Type "StockDesc" instead.
        
        If objUnits Is Nothing Then Set objUnits = Mainpo.PoUnits

        objUnits.StockNumber = Trim$(ssdcboCommoditty.Text)

                  SSOleDBUnit.RemoveAll
                  'NON-STOCK.Append a "N".In SaveToPOitem ,we checkif it is in-stock or non-stock.
                       
                       If RsUNits Is Nothing Then
                          If lookups Is Nothing Then lookups = Mainpo.lookups
                          Set RsUNits = lookups.GetAllUnits
                        End If
                        
                        RsUNits.MoveFirst
                        RsUNits.Find ("uni_code='" & Trim$(objUnits.PrimaryUnit) & "'")
                        SSOleDBUnit.AddItem objUnits.PrimaryUnit & ";" & RsUNits("uni_desc")
                        
                        RsUNits.MoveFirst
                        RsUNits.Find ("uni_code='" & Trim$(objUnits.SecondaryUnit) & "'")
                        SSOleDBUnit.AddItem objUnits.SecondaryUnit & ";" & RsUNits("uni_desc")
                        SSOleDBUnit = ""
                        SSOleDBUnit = objUnits.SecondaryUnit  'Juan 2010-9-26 to put a default vale on it
                                         
                       txt_Descript = objUnits.Description
                      
        Set lookups = Nothing
        
        ssdcboCommoditty.Tag = Trim(ssdcboCommoditty.Text)   'M 12/16/2002
        'If the user uses the TAB instead of the ENTER, the desc and units were not being populated. By using this option we know if the event was
        'fired, in a case when it is not fired, the tag would have a differnt value from the value selected. We would check this on the Validation event
        'and take an action appropriatly
        
        If chk_usexportLI.value = 1 Then
            
            If GetEccnForSelectedStock(Trim(ssdcboCommoditty.Text), Eccnid, Eccnno, EcnLicense, Sourceid, Sourceno) = False Then
            
                Exit Sub
            
            Else
                
                SSoleEccnNo.Tag = Eccnid
                SSoleEccnNo = Eccnno
                Chk_license.value = IIf(EcnLicense = True, 1, 0)
                SSOleSourceofinfo.Tag = Sourceid
                SSOleSourceofinfo = Sourceno
                
            End If
            
        End If
        
        If Len(Trim(SSoleEccnNo.Tag)) = 0 Then SSoleEccnNo.Tag = 0
        If Len(Trim(SSOleSourceofinfo.Tag)) = 0 Then SSOleSourceofinfo.Tag = 0
        
        If SSoleEccnNo.Tag > 0 Then
                   
                   SSoleEccnNo.Enabled = False
        ElseIf SSoleEccnNo.Tag = 0 And ConnInfo.Eccnactivate <> Constno And chk_FrmStkMst.value = 1 Then
                   
                   SSoleEccnNo.Enabled = True
        End If
        
        If SSOleSourceofinfo.Tag > 0 Then
                   
                   SSOleSourceofinfo.Enabled = False
        ElseIf SSOleSourceofinfo.Tag = 0 And ConnInfo.Eccnactivate <> Constno And chk_FrmStkMst.value = 1 Then
                   
                   SSOleSourceofinfo.Enabled = True
        End If
        
  End If
  
  
 Exit Sub
Handler:
  MsgBox "Error occurred while processing the units of the selected commoditty number." & vbCrLf & "Error Description   " & Err.Description
  Err.Clear
  ssdcboCommoditty = ""
  
End Sub

Private Sub ssdcboRequisition_Click()
Dim RsReqPO As ADODB.Recordset
Set lookups = Mainpo.lookups
ssdcboRequisition.Tag = ssdcboRequisition.value
lblReqLineitem = Trim$(ssdcboRequisition.Columns(2).Text)
Set RsReqPO = lookups.GetReqisitionLineItem(lblReqLineitem, Trim$(ssdcboRequisition.Text))
If RsReqPO.RecordCount > 0 Then Call LoadPOLINEFromRequsition(RsReqPO)

ssdcboRequisition.SelLength = 0
ssdcboRequisition.SelStart = 0


End Sub
Private Sub ssdcboRequisition_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not ssdcboRequisition.DroppedDown Then ssdcboRequisition.DroppedDown = True
End Sub

Private Sub ssdcboRequisition_KeyPress(KeyAscii As Integer)
If Not FormMode = mdCreation Then ssdcboRequisition.MoveNext
End Sub
Private Sub ssdcboShipper_Click()
If CheckIfCombosLoaded = False Then FillUPCOMBOS
ssdcboShipper.Tag = ssdcboShipper.Columns(0).Text
End Sub

Private Sub SSOleDBCompany_Click()
 SSOleDBInvLocation = ""
 SSOleDBInvLocation.RemoveAll
 If CheckIfCombosLoaded = False Then FillUPCOMBOS

    Dim value As String
    
    deIms.rsCompanyLocations.Filter = ""
    deIms.rsCompanyLocations.Filter = "loc_compcode='" & Trim$(SSOleDBcompany.Columns(0).Text) & "'"

    If deIms.rsCompanyLocations.EOF Then
         SSOleDBInvLocation.RemoveAll
         SSOleDBInvLocation.Enabled = False

    Else

       SSOleDBInvLocation.Enabled = True
        SSOleDBInvLocation.RemoveAll
       Do While Not deIms.rsCompanyLocations.EOF

           SSOleDBInvLocation.AddItem deIms.rsCompanyLocations!loc_locacode & ";" & deIms.rsCompanyLocations!loc_name
           deIms.rsCompanyLocations.MoveNext

       Loop

    End If
SSOleDBcompany.SelStart = 0
SSOleDBcompany.SelLength = 0
SSOleDBcompany.Tag = SSOleDBcompany.Columns(0).Text
End Sub

Private Sub SSOleDBCurrency_Click()
If CheckIfCombosLoaded = False Then FillUPCOMBOS
SSOleDBCurrency.SelLength = 0
SSOleDBCurrency.SelStart = 0
SSOleDBCurrency.Tag = SSOleDBCurrency.Columns(0).Text
End Sub

Private Sub Form_Load()

NavBar1.CancelLastSepVisible = False
NavBar1.LastPrintSepVisible = False
NavBar1.PrintSaveSepVisible = False
NavBar1.DeleteVisible = True

FNameSpace = deIms.NameSpace

Dim mLoadForm As Boolean
Dim x As String
Dim Count As Integer
    newSupplier = False

    NavBar1.EditEnabled = True
    mDidUserOpenStkMasterForm = False
    Mainpo.Configure deIms.NameSpace, deIms.cnIms

    Set Poheader = Mainpo.Poheader
    

    InitializePOheaderRecordset

    'Added by Juan (2015-02-13) for Multilingual
    Call translator.Translate_Forms("frm_NewPurchase")
    '------------------------------------------

    If Poheader.EOF = False Then LoadFromPOHEADER
    Set POFqa = Mainpo.FQA
       
   If POFqa.GetFQAInfo(ssOleDbPO) Then
    
        LoadFromFROMFQA
    Else
    
        CleanFROMFQA
        
    End If
       
    PoReceipeintsInit
      
    sst_PO.Tab = 0
 
    mCheckLIFields = True
    mCheckPoFields = True
    MCheckClause = True
    mCheckRemarks = True
    
    Call DisableButtons(frm_NewPurchase, NavBar1)
    
    ssOleDbPO.AllowInput = True
    
    NavBar1.PreviousEnabled = True
    NavBar1.LastEnabled = True
    NavBar1.FirstEnabled = True
    NavBar1.NextEnabled = True
    NavBar1.SaveEnabled = False
    NavBar1.DeleteEnabled = NavBar1.EditEnabled
    
    
    FormMode = ChangeMode(mdvisualization)
    
    If NavBar1.EditEnabled = False Then
         
         NavBar1.EditVisible = False
         NavBar1.DeleteVisible = False
         NavBar1.DeleteEnabled = False
         
    Else
    
         NavBar1.EditVisible = True
         NavBar1.DeleteVisible = True
         NavBar1.DeleteEnabled = True
         
    End If
    
    If NavBar1.NewEnabled = False Then
    
           NavBar1.NewVisible = False
           
    Else
    
            NavBar1.NewVisible = True
            
    End If
    
    Caption = Caption + " - " + Tag
    
    With frm_NewPurchase
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub


Public Function LoadFromPOHEADER() As Boolean
Dim RsStatus As New ADODB.Recordset
Dim rsDOCTYPE As New ADODB.Recordset
Dim RsSrvCode As ADODB.Recordset
mLoadMode = LoadingPOheader

RsStatus.Source = "select sts_code,sts_name from status where sts_npecode='" & deIms.NameSpace & "'"
RsStatus.ActiveConnection = deIms.cnIms
RsStatus.CursorType = adOpenKeyset
RsStatus.Open

rsDOCTYPE.Source = "select doc_code,doc_desc from doctype where doc_npecode='" & deIms.NameSpace & "'"
rsDOCTYPE.ActiveConnection = deIms.cnIms
rsDOCTYPE.CursorType = adOpenKeyset
rsDOCTYPE.Open


LoadFromPOHEADER = True
On Error GoTo Handler

rsDOCTYPE.MoveFirst
rsDOCTYPE.Find "doc_code='" & Poheader.Docutype & "'"

If Not rsDOCTYPE.AbsolutePosition = adPosEOF Then
    
    SSOleDBDocType.Tag = Poheader.Docutype

    SSOleDBDocType.Text = rsDOCTYPE!doc_desc

Else
    
    SSOleDBDocType.Tag = Poheader.Docutype

    SSOleDBDocType.Text = Poheader.Docutype
    MsgBox "Document code does not exist.Code will be used instead of the name."

End If

    ssOleDbPO = Poheader.Ponumb
    LblRevNumb = Poheader.revinumb
    LblRevDate = Format(Poheader.daterevi, "MM/DD/YY")

ssdcboShipper.Tag = Poheader.shipcode

If Len(Poheader.shipcode) > 0 Then
    deIms.rsSHIPPER.MoveFirst
    deIms.rsSHIPPER.Find ("shi_code='" & Poheader.shipcode & "'")
    
   If Not deIms.rsSHIPPER.AbsolutePosition = adPosEOF Then
      ssdcboShipper.Text = deIms.rsSHIPPER!shi_name
   Else
         ssdcboShipper.Text = Poheader.shipcode
         MsgBox "Invalid Shippier Code.The code will be used instead of the name."
   End If
   
Else
    ssdcboShipper.Text = ""
    ssdcboShipper.Tag = ""
End If

txt_ChargeTo = Poheader.chrgto

SSOleDBPriority.Tag = Poheader.priocode

deIms.rsPRIORITY.MoveFirst
deIms.rsPRIORITY.Find ("pri_code='" & Poheader.priocode & "'")

  If Not deIms.rsPRIORITY.AbsolutePosition = adPosEOF Then
         SSOleDBPriority.Text = deIms.rsPRIORITY!pri_desc
   Else
        'Added/Modified by Juan Gonzalez 12/27/2006
        Dim rsPrio As New ADODB.Recordset
        rsPrio.Source = "SELECT pri_desc FROM PRIORITY WHERE pri_npecode='" & deIms.NameSpace & "' AND pri_code =  '" & Poheader.priocode & "'"
        rsPrio.ActiveConnection = deIms.cnIms
        rsPrio.CursorType = adOpenForwardOnly
        rsPrio.Open
        
        If rsPrio.AbsolutePosition = adPosEOF Then
            SSOleDBPriority.Text = Poheader.priocode 'Previous
            MsgBox "Invalid Priority Code. The code will be used instead of the name." 'Previous
        Else
            SSOleDBPriority.AddItem Poheader.priocode & ";" & rsPrio!pri_desc, 0
            SSOleDBPriority.Text = rsPrio!pri_desc
        End If
        rsPrio.Close
        '-----------------
   End If



txt_Buyer = Poheader.buyr

SSOleDBOriginator = Poheader.orig

LblAppBy = Poheader.apprby
SSOleDBToBeUsedFor = Poheader.tbuf

SSoledbSupplier.Tag = Poheader.suppcode

deIms.rsActiveSupplier.MoveFirst
deIms.rsActiveSupplier.Find ("sup_code='" & Poheader.suppcode & "'")
If Not deIms.rsActiveSupplier.AbsolutePosition = adPosEOF Then
    SSoledbSupplier.Text = deIms.rsActiveSupplier!sup_name
Else
    SSoledbSupplier.Text = Poheader.suppcode
    MsgBox "Invalid Supplier Code. The Supplier Does not exist."
End If

Txt_supContaName = Poheader.SuppContactName
Txt_supContaPh = Poheader.SuppContaPH

SSOleDBCurrency.Tag = Poheader.currCODE

deIms.rsCURRENCY.MoveFirst
deIms.rsCURRENCY.Find ("curr_code='" & Poheader.currCODE & "'")
If Not deIms.rsCURRENCY.AbsolutePosition = adPosEOF Then
   SSOleDBCurrency.Text = deIms.rsCURRENCY!curr_desc
Else
   SSOleDBCurrency.Text = Poheader.currCODE
   MsgBox "Invalid Currency Code. The currency does not Exist."
End If

LblCompanyCode.Caption = Poheader.CompCode
deIms.rsActiveCompany.MoveFirst
deIms.rsActiveCompany.Find ("com_compcode='" & Poheader.CompCode & "'")

SSOleDBcompany.Tag = Poheader.CompCode

If Not deIms.rsActiveCompany.AbsolutePosition = adPosEOF Then
       SSOleDBcompany.Text = deIms.rsActiveCompany!com_name
Else
       SSOleDBcompany.Text = Poheader.CompCode
       MsgBox "Invalid company code. Company Does not exist"
End If



If Len(Trim$(Poheader.invloca)) > 0 Then
    
    deIms.rsCompanyLocations.MoveFirst
    deIms.rsCompanyLocations.Find ("loc_locacode='" & Poheader.invloca & "'")
    
    If Not deIms.rsCompanyLocations.AbsolutePosition = adPosEOF Then
        SSOleDBInvLocation.Tag = Poheader.invloca
        SSOleDBInvLocation.Text = deIms.rsCompanyLocations!loc_name
    Else
            SSOleDBInvLocation.Text = Poheader.invloca
            SSOleDBInvLocation.Tag = Poheader.invloca
    End If
    
Else
 
 SSOleDBInvLocation = ""
 SSOleDBInvLocation.Tag = ""
 
End If

Set RsSrvCode = GetServiceCode '03/08

If Len(Trim$(Poheader.srvccode)) > 0 And RsSrvCode.RecordCount > 0 Then '03/08
    
        RsSrvCode.MoveFirst '03/08
        RsSrvCode.Find ("srvc_code='" & Poheader.srvccode & "'") '03/08
        
        If Not RsSrvCode.AbsolutePosition = adPosEOF Then '03/08
        
            SSOledbSrvCode.Tag = Poheader.srvccode '03/08
            SSOledbSrvCode.Text = RsSrvCode("srvc_desc") '03/08
            
        Else '03/08
        
        'Added/Modified by Juan Gonzalez 12/27/2006
        Dim rsSrvCde As New ADODB.Recordset
        rsSrvCde.Source = "SELECT srvc_desc FROM SERVCODE WHERE srvc_npecode='" + deIms.NameSpace + "' AND srvc_code =  '" & Poheader.srvccode & "'"
        rsSrvCde.ActiveConnection = deIms.cnIms
        rsSrvCde.CursorType = adOpenForwardOnly
        rsSrvCde.Open
        
        If rsSrvCde.AbsolutePosition = adPosEOF Then
            SSOledbSrvCode.Text = Poheader.srvccode '03/08
            SSOledbSrvCode.Tag = Poheader.srvccode '03/08
        Else
            SSOledbSrvCode.AddItem Poheader.srvccode & ";" & rsSrvCde!srvc_desc, 0
            SSOledbSrvCode.Text = rsSrvCde!srvc_desc
        End If
        rsSrvCde.Close
        '-----------------
            
        End If '03/08
        
Else '03/08
 
 SSOledbSrvCode = "" '03/08
 SSOledbSrvCode.Tag = "" '03/08
 
End If '03/08

chk_ConfirmingOrder = IIf(Poheader.confordr = True, 1, 0)


If Len(Poheader.taccode) > 0 Then
    
    deIms.rsTermCondition.MoveFirst
    
    deIms.rsTermCondition.Find ("tac_taccode='" & Poheader.taccode & "'")
    
    If Not deIms.rsTermCondition.AbsolutePosition = adPosEOF Then
       ssdcboCondition.Tag = Poheader.taccode
       ssdcboCondition.Text = deIms.rsTermCondition!tac_desc
    Else
    
        'Added/Modified by Juan Gonzalez 12/28/2006
        Dim rsCond As New ADODB.Recordset
        rsCond.Source = "SELECT tac_desc FROM TERMSANDCONDITION WHERE tac_npecode='" + deIms.NameSpace + "' AND tac_code =  '" & Poheader.taccode & "'"
        rsCond.ActiveConnection = deIms.cnIms
        rsCond.CursorType = adOpenForwardOnly
        rsCond.Open
        
        If rsSrvCde.AbsolutePosition = adPosEOF Then
            ssdcboCondition.Tag = Poheader.taccode 'previous
            ssdcboCondition.Text = Poheader.taccode 'previous
            MsgBox "Invalid condition Code. Condition code does not Exist" 'previous
        Else
            ssdcboCondition.AddItem Poheader.taccode & ";" & rsCond!tac_desc, 0
            ssdcboCondition.Text = rsCond!tac_desc
        End If
        rsSrvCde.Close
        '-----------------
    End If

Else

    ssdcboCondition.Text = ""
 
End If

deIms.rsTermDelivery.MoveFirst
deIms.rsTermDelivery.Find ("tod_termcode='" & Poheader.termcode & "'")

If Not deIms.rsTermDelivery.AbsolutePosition = adPosEOF Then
   
    ssdcboDelivery.Text = deIms.rsTermDelivery!tod_desc
    ssdcboDelivery.Tag = Poheader.termcode
Else
    ssdcboDelivery.Text = Poheader.termcode
    ssdcboDelivery.Tag = Poheader.termcode
    MsgBox "Invalid Term of delivery code. The code will be used instead of the name."
End If

If Len(Poheader.shipto) > 0 Then
  If deIms.rsActiveShipTo.State = 0 Then Call deIms.ActiveShipTo(deIms.NameSpace)
  deIms.rsActiveShipTo.MoveFirst
    deIms.rsActiveShipTo.Find ("sht_code='" & Poheader.shipto & "'")
   If Not deIms.rsActiveShipTo.AbsolutePosition = adPosEOF Then
      
        SSOleDBShipTo.Text = deIms.rsActiveShipTo!sht_name
        SSOleDBShipTo.Tag = Poheader.shipto
   Else
        SSOleDBShipTo.Tag = Poheader.shipto
         SSOleDBShipTo.Text = Poheader.shipto
       MsgBox "Invalid Shipto Code. The code will be used instead of the name"
   End If
Else
    SSOleDBShipTo.Text = ""
    SSOleDBShipTo.Tag = ""
End If

chk_FrmStkMst = IIf(Poheader.fromstckmast = True, 1, 0)
txtSite = Poheader.Site

dtpRequestedDate = Poheader.reqddelvdate
LblDateSent = Format(Poheader.datesent, "mm/dd/yy")
DTPicker_poDate = Format(Poheader.Createdate, "mm/dd/yy")
chk_USExportH = IIf(Poheader.usexport = True, 1, 0)
Dim StasINvt As String
Dim Stasship As String
Dim stasdelv As String
Dim stas As String

If Len(Poheader.StasINvt) > 0 Then
    RsStatus.MoveFirst
    RsStatus.Find "sts_code='" & Poheader.StasINvt & "'"
    StasINvt = RsStatus!sts_name
Else
    StasINvt = ""
End If

If Len(Poheader.Stasship) > 0 Then
    RsStatus.MoveFirst
    RsStatus.Find "sts_code='" & Poheader.Stasship & "'"
    Stasship = RsStatus!sts_name
Else
    Stasship = ""
End If

If Len(Poheader.stasdelv) > 0 Then
    RsStatus.MoveFirst
    RsStatus.Find "sts_code='" & Poheader.stasdelv & "'"
    stasdelv = RsStatus!sts_name
Else
    stasdelv = ""
End If


If Len(Poheader.stas) > 0 Then
    RsStatus.MoveFirst
    RsStatus.Find "sts_code='" & Poheader.stas & "'"
    stas = RsStatus!sts_name
Else
    stas = ""
End If

LblStatus7.Caption = StasINvt
LblStatus6.Caption = Stasship
LblStatus5.Caption = stasdelv
'dcbostatus(4) = stas
LblStatus4.Caption = stas

  chk_Forwarder = IIf(Poheader.forwr = True, 1, 0)
  chk_Requ = IIf(Poheader.reqddelvflag = True, 1, 0)
  chk_FreightFard = IIf(Poheader.freigforwr = True, 1, 0)
  LoadFromPOHEADER = True
  mLoadMode = NoLoadInProgress
  Exit Function
Handler:
   MsgBox Err.Description
   Err.Clear
   mLoadMode = NoLoadInProgress
End Function
Public Function LoadFromPOITEM() As Boolean
On Error GoTo Handler
LoadFromPOITEM = False

Dim RsStatus As New ADODB.Recordset
Dim frameenabled As Boolean

mLoadMode = loadingPoItem
RsStatus.Source = "select sts_code,sts_name from status where sts_npecode='" & deIms.NameSpace & "'"
RsStatus.ActiveConnection = deIms.cnIms
RsStatus.CursorType = adOpenKeyset
RsStatus.Open



Dim Stasliit As String
Dim Stasdlvy As String
Dim Stasship As String
Dim StasINvt As String

If Len(PoItem.Stasliit) > 0 Then
RsStatus.MoveFirst
RsStatus.Find "sts_code='" & PoItem.Stasliit & "'"
Stasliit = RsStatus!sts_name
End If

If Len(PoItem.Stasdlvy) > 0 Then
RsStatus.MoveFirst
RsStatus.Find "sts_code='" & PoItem.Stasdlvy & "'"
Stasdlvy = RsStatus!sts_name
End If

If Len(PoItem.Stasship) > 0 Then
RsStatus.MoveFirst
RsStatus.Find "sts_code='" & PoItem.Stasship & "'"
Stasship = RsStatus!sts_name
End If

If Len(PoItem.StasINvt) > 0 Then
RsStatus.MoveFirst
RsStatus.Find "sts_code='" & PoItem.StasINvt & "'"
StasINvt = RsStatus!sts_name
End If


dcbostatus(0).Text = Stasliit
dcbostatus(1) = Stasdlvy
dcbostatus(2) = Stasship
dcbostatus(3) = StasINvt



LblPOi_PONUMB = IIf(InStr(1, PoItem.Ponumb, "_"), "", PoItem.Ponumb) ' "_" is the first chanrater of a Ponumb which is generated on the fly, before An Auonumber is Generated and the PO number is saved to the database.
txt_LI = PoItem.Linenumb
txt_TotalLIs = PoItem.Count
ssdcboCommoditty.Text = IIf(InStr(1, PoItem.Comm, "_"), "", PoItem.Comm)    ' "_" is the first chanrater of a Ponumb which is generated on the fly, before An Auonumber is Generated and the PO number is saved to the database.
ssdcboManNumber = PoItem.Manupartnumb
txt_AFE = PoItem.Afe
SSOleDBCustCategory = PoItem.Custcate
txt_SerialNum = PoItem.Serlnumb
txt_Requested = Replace(FormatNumber$(PoItem.Primreqdqty, 4), ",", "")
'txtSecRequested = PoItem.Secoreqdqty

'SSOleDBSecUnit = PoItem.Secouom  'Row Member - SECONDARYUNIT,ListField - uni_desc , BoundColumns - uni_code
txt_Delivered = Replace(FormatNumber$(PoItem.PriQtydlvd, 4), ",", "")
txt_Shipped = Replace(FormatNumber$(PoItem.PriQtyship, 4), ",", "")
txt_Inventory2 = Replace(FormatNumber$(PoItem.PriQtyinvt, 4), ",", "")

txt_Descript = PoItem.Description
txt_remk = PoItem.Remk

 
'If the PO was created in Primary mode
If Len(PoItem.UnitOfPurch) = 0 Or Trim$(PoItem.UnitOfPurch) = "P" Then
   
    txt_Requested = Replace(FormatNumber$(PoItem.Primreqdqty, 4), ",", "")
    SSOleDBUnit = PoItem.Primuom     'w MemBer -GET_UNIT_OF_MEASURE ,ListField-uni_desc,BoundColumns-Uni_code
    txt_Price = Replace(FormatNumber$(PoItem.PrimUnitprice, 2), ",", "")
    txt_Total = FormatNumber$(PoItem.PriTotaprice, 2)
    
'else If the PO was created in Secondary mode
ElseIf Trim$(PoItem.UnitOfPurch) = "S" Then

    txt_Requested = Replace(FormatNumber$(PoItem.Secoreqdqty, 4), ",", "")
    SSOleDBUnit = PoItem.Secouom      'w MemBer -GET_UNIT_OF_MEASURE ,ListField-uni_desc,BoundColumns-Uni_code
    txt_Price = Replace(FormatNumber$(PoItem.SecUnitPrice, 2), ",", "")
    txt_Total = FormatNumber$(PoItem.SecTotaprice, 2)
    
End If

ssdcboRequisition = PoItem.Requnumb
lblReqLineitem = PoItem.Requliitnumb



If Len(PoItem.Liitreqddate) = 0 Then
   DTP_Required.value = ""
Else
   DTP_Required.checkBox = True
   DTP_Required.value = PoItem.Liitreqddate
End If

'eccn fields

'frameenabled = fra_LineItem.Enabled

'fra_LineItem.Enabled = True
 
   chk_usexportLI = IIf(PoItem.usexport = True, 1, 0)
   Chk_license = IIf(PoItem.Eccnlicsreq = True, 1, 0)
 
 SSoleEccnNo.Tag = PoItem.Eccnid
 'SSoleEccnno.Text = PoItem.Eccnno
 
If GRsEccnNo Is Nothing Then

    Call FillEccnCombos(lookups)

End If


If Len(PoItem.Eccnid) > 0 And GRsEccnNo.RecordCount > 0 Then

 GRsEccnNo.MoveFirst
 GRsEccnNo.Find "eccnid='" & PoItem.Eccnid & "'"
         If GRsEccnNo.EOF = False Then
            SSoleEccnNo.Text = GRsEccnNo!eccn_no
        Else
            SSoleEccnNo.Text = ""
        End If
         
  
End If

SSOleSourceofinfo.Tag = PoItem.Sourceofinfoid

If GRSSourceOfInfo Is Nothing Then

    Call FillSourceOfinfoCombos(lookups)

End If
 
If Len(PoItem.Sourceofinfoid) > 0 And GRSSourceOfInfo.RecordCount > 0 Then

        GRSSourceOfInfo.MoveFirst
        GRSSourceOfInfo.Find "sourceid='" & PoItem.Sourceofinfoid & "'"
        If GRSSourceOfInfo.EOF = False Then
            SSOleSourceofinfo.Text = GRSSourceOfInfo!Source
        Else
            SSOleSourceofinfo.Text = ""
        End If
         
End If
 
 LoadFromPOITEM = True
 mLoadMode = NoLoadInProgress
 
 Exit Function
 
Handler:
MsgBox Err.Description
  Err.Clear
  mLoadMode = NoLoadInProgress
  
End Function

Public Function LoadPoHeaderCombos() As Boolean

Dim FNameSpace As String * 5

'Dim RsBRQ As ADODB.Recordset

'Dim defsite As String

On Error GoTo Handler
  
   LoadPoHeaderCombos = False
   
   FNameSpace = deIms.NameSpace

    If Not mIsPoHeaderRsetsInit = True Then InitializePOheaderRecordset
    LoadPoHeaderCombos = FillUPCOMBOS
    mIsPoheaderCombosLoaded = LoadPoHeaderCombos
    
    sst_PO.Tab = 0

    NavBar1.NextEnabled = sst_PO.Tab <> 0
    NavBar1.LastEnabled = sst_PO.Tab <> 0
    NavBar1.FirstEnabled = sst_PO.Tab <> 0
    NavBar1.PreviousEnabled = sst_PO.Tab <> 0
    
    Exit Function
    
Handler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FillUPCOMBOS() As Boolean
 
 Dim rsSUPPLIER As ADODB.Recordset
 Dim Count As Integer
 Dim rsSERVCODE As ADODB.Recordset
 
If lookups Is Nothing Then Set lookups = Mainpo.lookups

If lookups.GetUserMenuLevel(CurrentUser) = 5 Then
 Set rsSUPPLIER = lookups.GetLocalSuppliers
Else
 Set rsSUPPLIER = deIms.rsActiveSupplier
End If
 
 FillUPCOMBOS = False
 
 On Error GoTo Handler
 
     Set IntiClass = New InitialValuesPOheader
 
   If Not deIms.rsSHIPPER.EOF Then
        deIms.rsSHIPPER.MoveFirst
        IntiClass.InitShipperCode = Trim$(deIms.rsSHIPPER!shi_code)
        IntiClass.InitShipperName = Trim$(deIms.rsSHIPPER!shi_name)
   End If
ssdcboShipper.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsSHIPPER.EOF
       ssdcboShipper.AddItem deIms.rsSHIPPER!shi_code & ";" & deIms.rsSHIPPER!shi_name
       deIms.rsSHIPPER.MoveNext
   Loop
       
   If Not deIms.rsTermDelivery.EOF Then deIms.rsTermDelivery.Filter = "tod_actvflag<>0"
   If Not deIms.rsTermDelivery.EOF Then
        deIms.rsTermDelivery.MoveFirst
        IntiClass.InitDelivery = Trim$(deIms.rsTermDelivery!tod_desc)
   End If
ssdcboDelivery.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsTermDelivery.EOF
       ssdcboDelivery.AddItem deIms.rsTermDelivery!tod_termcode & ";" & deIms.rsTermDelivery!tod_desc
       deIms.rsTermDelivery.MoveNext
   Loop
   deIms.rsTermDelivery.Filter = ""
   
   If Not deIms.rsTermCondition.EOF Then deIms.rsTermCondition.Filter = "tac_actvflag<>0"
   If Not deIms.rsTermCondition.EOF Then
    deIms.rsTermCondition.MoveFirst
    IntiClass.InitCondition = Trim$(deIms.rsTermCondition!tac_desc)
   End If
ssdcboCondition.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsTermCondition.EOF
       ssdcboCondition.AddItem deIms.rsTermCondition!tac_taccode & ";" & deIms.rsTermCondition!tac_desc
       deIms.rsTermCondition.MoveNext
   Loop
   deIms.rsTermCondition.Filter = ""
   
   If Not deIms.rsCURRENCY.EOF Then
     deIms.rsCURRENCY.MoveFirst
     IntiClass.InitCurrency = Trim$(deIms.rsCURRENCY!curr_desc)
   End If
SSOleDBCurrency.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsCURRENCY.EOF
       SSOleDBCurrency.AddItem deIms.rsCURRENCY!curr_code & ";" & deIms.rsCURRENCY!curr_desc
       deIms.rsCURRENCY.MoveNext
   Loop
   
   If Not deIms.rsPRIORITY.EOF Then deIms.rsPRIORITY.Filter = "pri_actvflag <>0"
   If Not deIms.rsPRIORITY.EOF Then
     
     deIms.rsPRIORITY.MoveFirst
     IntiClass.InitPriority = Trim$(deIms.rsPRIORITY!pri_desc)
   End If
SSOleDBPriority.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsPRIORITY.EOF
       SSOleDBPriority.AddItem deIms.rsPRIORITY!pri_code & ";" & deIms.rsPRIORITY!pri_desc
       deIms.rsPRIORITY.MoveNext
   Loop
   deIms.rsPRIORITY.Filter = ""
   
   If Not rsSUPPLIER.EOF Then
      rsSUPPLIER.MoveFirst
      IntiClass.InitSupplier = Trim$(rsSUPPLIER!sup_name)
   End If
SSoledbSupplier.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not rsSUPPLIER.EOF
       SSoledbSupplier.AddItem rsSUPPLIER!sup_code & ";" & rsSUPPLIER!sup_name & ";" & rsSUPPLIER!sup_city & ";" & rsSUPPLIER!sup_phonnumb
       rsSUPPLIER.MoveNext
   Loop
   
   If Not deIms.rsActiveOriginator.EOF Then
     deIms.rsActiveOriginator.MoveFirst
     IntiClass.InitOriginator = Trim$(deIms.rsActiveOriginator!ori_code)
   End If
SSOleDBOriginator.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsActiveOriginator.EOF
       SSOleDBOriginator.AddItem deIms.rsActiveOriginator!ori_code '& ";" & deIms.rsActiveOriginator!ori_code
       deIms.rsActiveOriginator.MoveNext
   Loop
   
   If Not deIms.rsActiveTbu.EOF Then
      deIms.rsActiveTbu.MoveFirst
      IntiClass.InitToBeUsedFor = Trim$(deIms.rsActiveTbu!tbu_name)
   End If
SSOleDBToBeUsedFor.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsActiveTbu.EOF
       SSOleDBToBeUsedFor.AddItem deIms.rsActiveTbu!tbu_name '& ";" & deIms.rsActiveOriginator!tbu_name
       deIms.rsActiveTbu.MoveNext
   Loop
   
   If Not deIms.rsActiveShipTo.EOF Then
       deIms.rsActiveShipTo.MoveFirst
       IntiClass.InitShipTo = Trim$(deIms.rsActiveShipTo!sht_name)
   End If
SSOleDBShipTo.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsActiveShipTo.EOF
       SSOleDBShipTo.AddItem deIms.rsActiveShipTo!sht_code & ";" & deIms.rsActiveShipTo!sht_name
       deIms.rsActiveShipTo.MoveNext
   Loop
   
   If Not deIms.rsActiveCompany.EOF Then
       deIms.rsActiveCompany.MoveFirst
       IntiClass.InitCompanyCode = Trim$(deIms.rsActiveCompany!com_compcode)
       IntiClass.InitCompanyName = Trim$(deIms.rsActiveCompany!com_name)
   End If
SSOleDBcompany.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not deIms.rsActiveCompany.EOF
       SSOleDBcompany.AddItem deIms.rsActiveCompany!com_compcode & ";" & deIms.rsActiveCompany!com_name
       deIms.rsActiveCompany.MoveNext
       
   Loop
   '-----------------------------
   'code of serivce code, Added on 03/09 by Muz
   
   Set rsSERVCODE = GetServiceCode
   
   If Not rsSERVCODE.EOF Then
       rsSERVCODE.MoveFirst
       IntiClass.initServiceCode = Trim$(rsSERVCODE!srvc_code)
       IntiClass.initServicedesc = Trim$(rsSERVCODE!srvc_desc)
   End If
   SSOledbSrvCode.RemoveAll 'JCGFIXES 2007/24/1
   Do While Not rsSERVCODE.EOF
       SSOledbSrvCode.AddItem rsSERVCODE!srvc_code & ";" & rsSERVCODE!srvc_desc
       rsSERVCODE.MoveNext
   Loop
   
   rsSERVCODE.Close
   
   Set rsSERVCODE = Nothing
   
   '------------------------------
   
   IntiClass.InitpoDate = Format(Now(), "mm/dd/yy")
   IntiClass.InitBuyer = CurrentUser
    
    Set rsSUPPLIER = Nothing
    CheckIfCombosLoaded = True
   FillUPCOMBOS = True
   Exit Function
Handler:
   
       
    Err.Clear
End Function

Private Sub SSOleDBInvLocation_DropDown()

'If Not mIsInvLocationLoaded = False Then
        SSOleDBInvLocation = ""
         SSOleDBInvLocation.RemoveAll
         If CheckIfCombosLoaded = False Then FillUPCOMBOS
        
            Dim value As String
        '
        LblCompanyCode.Caption = Trim$(SSOleDBcompany.Columns(0).Text)
            deIms.rsCompanyLocations.Filter = ""
            deIms.rsCompanyLocations.Filter = "LOC_COMPCODE= '" & Trim$(LblCompanyCode.Caption) & "'"
        
            If deIms.rsCompanyLocations.EOF Then
        
                 SSOleDBInvLocation.Enabled = False
        
            Else
        
               SSOleDBInvLocation.Enabled = True
               SSOleDBInvLocation.RemoveAll
               Do While Not deIms.rsCompanyLocations.EOF
        
                   SSOleDBInvLocation.AddItem deIms.rsCompanyLocations!loc_locacode & ";" & deIms.rsCompanyLocations!loc_name
                   deIms.rsCompanyLocations.MoveNext
        
               Loop
        
            End If
            
        mIsDocTypeLoaded = True
 'End If
End Sub

Private Sub SSOleDBInvLocation_GotFocus()

SSOleDBInvLocation = ""
 SSOleDBInvLocation.RemoveAll
 If CheckIfCombosLoaded = False Then FillUPCOMBOS

    Dim value As String
'
    
    deIms.rsCompanyLocations.Filter = ""
    deIms.rsCompanyLocations.Filter = "LOC_COMPCODE= '" & Trim$(LblCompanyCode.Caption) & "'"

    If deIms.rsCompanyLocations.EOF Then

         SSOleDBInvLocation.Enabled = False

    Else

       SSOleDBInvLocation.Enabled = True
       SSOleDBInvLocation.RemoveAll
       Do While Not deIms.rsCompanyLocations.EOF

           SSOleDBInvLocation.AddItem deIms.rsCompanyLocations!loc_locacode & ";" & deIms.rsCompanyLocations!loc_name
           deIms.rsCompanyLocations.MoveNext

       Loop

    End If
    SSOleDBInvLocation.SelLength = 0
    SSOleDBInvLocation.SelStart = 0
    Call HighlightBackground(SSOleDBInvLocation)
    
End Sub

Private Sub SSOleDBPO_Click()

If Not FormMode = mdCreation Then
    
    Poheader.Move Trim$(ssOleDbPO)
    
    Call LoadFromPOHEADER

    If POFqa Is Nothing Then Set POFqa = Mainpo.FQA

    If POFqa.GetFQAInfo(ssOleDbPO) Then
    
        LoadFromFROMFQA
    Else
    
        CleanFROMFQA
        
    End If
    
End If
End Sub

Public Function SaveToPOHEADER() As Boolean
On Error GoTo Handler
SaveToPOHEADER = False

 Poheader.Npecode = deIms.NameSpace
 
 Poheader.Docutype = SSOleDBDocType.Tag
  'Since the number would be autogenerated in the POHEADER.ADDNEW method
  
 'Poheader.Ponumb = SSOleDBPO 'AM
 
 If FormMode <> mdCreation Then 'AM
 
    Poheader.Ponumb = ssOleDbPO 'AM
 
 End If 'AM
  
 Poheader.revinumb = LblRevNumb
 
 If Len(LblRevDate.Caption) > 0 Then Poheader.daterevi = LblRevDate
 
 Poheader.shipcode = Trim$(ssdcboShipper.Tag)
 
 Poheader.chrgto = txt_ChargeTo
 
Poheader.priocode = Trim$(SSOleDBPriority.Tag)
 
 Poheader.buyr = txt_Buyer
 
 Poheader.orig = SSOleDBOriginator
 
  Poheader.tbuf = SSOleDBToBeUsedFor
  
  Poheader.suppcode = Trim$(SSoledbSupplier.Tag)
  
  Poheader.SuppContactName = Txt_supContaName
  
  Poheader.SuppContaPH = Txt_supContaPh
  
  Poheader.currCODE = Trim$(SSOleDBCurrency.Tag)
  
  Poheader.CompCode = Trim$(SSOleDBcompany.Tag)
   
   Poheader.invloca = Trim$(SSOleDBInvLocation.Tag)
  
  Poheader.confordr = IIf(chk_ConfirmingOrder = 1, True, False)
  
  Poheader.taccode = Trim$(ssdcboCondition.Tag)
  
  Poheader.termcode = Trim$(ssdcboDelivery.Tag)
  
  Poheader.fromstckmast = chk_FrmStkMst
  
  Poheader.Site = txtSite
  
 Poheader.shipto = Trim$(SSOleDBShipTo.Tag)
  
  Poheader.reqddelvdate = dtpRequestedDate
  
  Poheader.datesent = LblDateSent
  
  Poheader.Createdate = DTPicker_poDate
  
  Poheader.forwr = chk_Forwarder
  
  Poheader.freigforwr = chk_FreightFard
  
  Poheader.reqddelvflag = chk_Requ
  
  Poheader.srvccode = Trim(Trim(SSOledbSrvCode.Tag))
  
  'If this is a POREVISION
  If mSaveToPoRevision = True Then
     
     Poheader.apprby = ""
     Poheader.stas = "OH"
  
  Else
     
     Poheader.apprby = LblAppBy
  
  End If
    Poheader.usexport = chk_USExportH
  Call SavetoFROMFQA
  

  
  SaveToPOHEADER = True
  
  Exit Function
Handler:
   MsgBox Err.number
   Err.Clear
   
End Function
Private Function GetDocumentType(All As Boolean) As ADODB.Recordset
On Error Resume Next
Dim rs As ADODB.Recordset
Dim Count As Long
Dim CurrentUser As String

    With deIms
    
        All = True
        If All Then
            Set GetDocumentType = .UserDocumentType("", Count)
        Else
            Set GetDocumentType = .UserDocumentType(CurrentUser, Count)
        End If
        
    End With
    
    
End Function

Private Sub SSOleDBPO_GotFocus()
ssOleDbPO.SelLength = 0
ssOleDbPO.SelStart = 0
 Call HighlightBackground(ssOleDbPO)
End Sub

Public Function SetInitialVAluesPoHeader()
Dim RsStatus As New ADODB.Recordset



If lookups Is Nothing Then Set lookups = Mainpo.lookups

 SetInitialVAluesPoHeader = False
  
 SSOleDBDocType = ""
   
 SSOleDBDocType.Tag = ""
   
  ssOleDbPO = ""
  
  LblRevNumb = 0
  LblRevDate = ""
 
  LblAppBy = ""
  
  ssdcboShipper.Tag = ""
  ssdcboShipper.Text = ""
  
  txt_ChargeTo = ""
  
  SSOleDBPriority.Tag = ""
  SSOleDBPriority = ""
  
  txt_Buyer = IntiClass.InitBuyer
  
  SSOleDBOriginator.Tag = ""
  SSOleDBOriginator = ""
  
  LblDateSent.Caption = ""
 
  
  SSOleDBToBeUsedFor.Tag = ""
  SSOleDBToBeUsedFor = ""
  
  SSoledbSupplier.Tag = ""
  SSoledbSupplier = ""
  
  Txt_supContaName = ""
  Txt_supContaPh = ""
  
  SSOleDBCurrency.Tag = ""
  SSOleDBCurrency = ""
  
  
  SSOleDBcompany.Tag = ""
  SSOleDBcompany.Text = ""
  
  SSOleDBInvLocation.Tag = ""
  SSOleDBInvLocation = ""

  ssdcboCondition.Tag = ""
  ssdcboCondition = ""
  
  ssdcboDelivery.Tag = ""
  ssdcboDelivery = ""
  
  txtSite = lookups.GetMYSite
  
  SSOleDBShipTo.Tag = ""
  SSOleDBShipTo = ""
  
  SSOledbSrvCode.Tag = "" ' 03/09
  SSOledbSrvCode.Text = "" ' 03/09
  
  DTPicker_poDate = IntiClass.InitpoDate
  
  dtpRequestedDate = DTPicker_poDate
  
  SetInitialVAluesPoHeader = True
  
  '-----------------------------------
  RsStatus.Source = "select sts_code,sts_name from status where sts_npecode='" & deIms.NameSpace & "'"
RsStatus.ActiveConnection = deIms.cnIms
RsStatus.CursorType = adOpenKeyset
RsStatus.Open
  RsStatus.MoveFirst
  RsStatus.Find ("sts_code='OH'")
  LblStatus4.Caption = RsStatus!sts_name
  RsStatus.MoveFirst
  RsStatus.Find ("sts_code='NR'")
  LblStatus5.Caption = RsStatus!sts_name
  RsStatus.MoveFirst
  RsStatus.Find ("sts_code='NS'")
  LblStatus6.Caption = RsStatus!sts_name
  RsStatus.MoveFirst
  RsStatus.Find ("sts_code='NI'")
  LblStatus7.Caption = RsStatus!sts_name

  RsStatus.Close
  Set RsStatus = Nothing
End Function


Private Sub SSOleDBPO_LostFocus()
Call NormalBackground(ssOleDbPO)
SSOleDBPO_Validate (False)

End Sub

Private Sub SSOleDBPO_Validate(Cancel As Boolean)

'If FormMode = mdCreation And Len(Trim$(SSOleDBPO.text)) = 0 Then 'AM
 '   MsgBox "This field can not be left empty. Please fill it to perform any other operation." 'AM
  '  SSOleDBPO.SetFocus 'AM
   ' Cancel = True 'AM
'End If 'AM


If FormMode = mdCreation And Len(ssOleDbPO.Text) > 0 Then
     If Len(Trim$(ssOleDbPO)) > 15 Then
        MsgBox "Transaction order number can not be greater than 15 characters."
        Cancel = True
        Exit Sub
     End If
     
     If deIms.rsPonumb.State = 0 Then Call deIms.Ponumb(deIms.NameSpace)
        deIms.rsPonumb.MoveFirst
        deIms.rsPonumb.Find "PO_PONUMB='" & Trim$(ssOleDbPO.Text) & "'"
        If Not deIms.rsPonumb.EOF Then
        
          MsgBox "PO Number Already Exists"
          ssOleDbPO.Text = ""
          Cancel = True
            ssOleDbPO.SetFocus
        End If
 End If
End Sub
Private Sub SSOleDBsupplier_Click()
Dim suppAddress As String
Dim DocCode As String
Dim RSDocautodist As ADODB.Recordset
On Error GoTo Handler
If Len(SSoledbSupplier.Text) > 0 Then

    deIms.rsActiveSupplier.MoveFirst
    deIms.rsActiveSupplier.Find ("sup_code='" & SSoledbSupplier.Columns(0).Text & "'")
    Txt_supContaName.Text = IIf(IsNull(deIms.rsActiveSupplier!sup_contaname), "", deIms.rsActiveSupplier!sup_contaname)
    Txt_supContaPh.Text = IIf(IsNull(deIms.rsActiveSupplier!sup_contaph), "", deIms.rsActiveSupplier!sup_contaph)
    
    Set RSDocautodist = New ADODB.Recordset
    
    If FormMode = mdCreation Then
       DocCode = SSOleDBDocType.Tag
       
    Else
       DocCode = Poheader.Docutype
       
    End If
    
    RSDocautodist.Source = "select doc_autodist from doctype where doc_code='" & DocCode & "' and doc_npecode='" & deIms.NameSpace & "'"
    
    RSDocautodist.ActiveConnection = deIms.cnIms
    
    RSDocautodist.Open
    
    If RSDocautodist!doc_autodist = False Then
         
         RSDocautodist.Close
         Set RSDocautodist = Nothing
         SSoledbSupplier.Tag = SSoledbSupplier.Columns(0).Text
         Exit Sub
         
     End If
    
    SSoledbSupplier.Tag = SSoledbSupplier.Columns(0).Text
    
    If Not IsNull(deIms.rsActiveSupplier!sup_faxnumb) And Not Len(deIms.rsActiveSupplier!sup_faxnumb) = 0 Then
        If LTrim(RTrim(DocCode)) <> "R" Then 'JCG 2009/10/26
             suppAddress = "" & Trim$(deIms.rsActiveSupplier!sup_faxnumb)  'D
        End If 'JCG 2009/10/26
    End If
    
        If Not IsNull(deIms.rsActiveSupplier!sup_mail) And Not Len(deIms.rsActiveSupplier!sup_mail) = 0 Then
       If InStr(UCase(deIms.rsActiveSupplier!sup_mail), "") = 0 Then 'D
        CmdAddSupEmail.Tag = "" & IIf(Len(Trim$(deIms.rsActiveSupplier!sup_mail & "")) > 0, Trim$(deIms.rsActiveSupplier!sup_mail), "") 'D
       End If
       
    Else
        CmdAddSupEmail.Tag = ""
    End If
    
      
     
     
    If Len(Trim$(suppAddress)) > 0 Then

             If PoReceipients Is Nothing Then
                     Set PoReceipients = Mainpo.PoReceipients
                     'PoReceipients.Move SSOleDBPO 'AM
                     PoReceipients.Move Poheader.Ponumb  'AM

             End If


                suppAddress = PrefixFaxorEmail(suppAddress)

               If IsRecipientInList(suppAddress) Then Exit Sub


    End If

               If PoReceipients Is Nothing Then
                     Set PoReceipients = Mainpo.PoReceipients
                     PoReceipients.Move ssOleDbPO
               End If

               'Call PoReceipients.SubmitSupplier(suppAddress, SSOleDBPO.text) 'AM
                Call PoReceipients.SubmitSupplier(suppAddress, Poheader.Ponumb)  'AM

    SSoledbSupplier.SelLength = 0
    SSoledbSupplier.SelStart = 0
    SSoledbSupplier.Tag = SSoledbSupplier.Columns(0).Text
End If
Exit Sub
Handler:
 MsgBox "Errors occurred while working with the supplier.Error description is '" & Err.Description & "'"
 Err.Clear
End Sub

Private Sub SSOleDBUnit_Click()
     Dim lookups As IMSPODLL.lookups
    
End Sub

Private Sub SSOleDBUnit_GotFocus()
SSOleDBUnit.SelStart = 0
SSOleDBUnit.SelLength = 0
Call HighlightBackground(SSOleDBUnit)
SSOleDBUnit.DroppedDown = True
End Sub

Private Sub sst_PO_Click(PreviousTab As Integer)
    On Error Resume Next

Dim EditMode(1) As Long, str As String

        
Screen.MousePointer = vbHourglass
        
        Select Case PreviousTab
            
            Case 0
                 If FormMode <> mdvisualization Then
                    If mCheckLIFields = True And MCheckClause = True And mCheckRemarks = True Then
                        If FormMode <> mdvisualization Then mCheckPoFields = CheckPoFields
                        
                        
                        If FormMode <> mdvisualization And mCheckPoFields = True Then
                              SaveToPOHEADER
                              SavetoFROMFQA
                           Else
                             sst_PO.Tab = 0
                        End If
                        
                    End If
                    
                    NavBar1.SaveEnabled = False
                   
                  End If
                  
                  
                     NavBar1.DeleteEnabled = False
                  
                
            
            Case 1

                
            Case 2
                
            
               'Save the POItem back to  POITEMs Object
               If FormMode <> mdvisualization Then
                If mCheckPoFields = True And MCheckClause = True And mCheckRemarks = True Then
                   
                   If PoItem.Count > 0 And FormMode <> mdvisualization Then
                      
                      If CheckLIFields Then
                        
                        mCheckLIFields = True
                        SaveToPOITEM
                        SaveToTOFQA
                        NavBar1.DeleteEnabled = False
                        
                        Else
                         
                        mCheckLIFields = False
                        sst_PO.Tab = 2
                       
                       End If
                       
                    End If
                    
                End If
               End If
               
              If FormMode <> mdvisualization Then
                    fra_LineItem.Enabled = True
              End If
               
               
            Case 3
                If FormMode <> mdvisualization Then
                   txtRemarks.SetFocus
                    If mCheckPoFields = True And mCheckLIFields = True And MCheckClause = True Then
                     If PORemark.Count > 0 And FormMode <> mdvisualization Then
                    
                   
                            If Len(Trim$(txtRemarks)) > 0 Then
                            
                                txtRemarks = FixTheFirstCarriageReturn(txtRemarks)
                            
                                 savetoPORemarks
                                 mCheckRemarks = True
                            Else
                                MsgBox "Remarks can not be left empty."
                                mCheckRemarks = False
                                sst_PO.Tab = 3
                            End If
                            
                      End If
                   
                   End If
                   
                   
                   
                End If
            
            Case 4
              '  Call txtClause_Validate(False)
              
              If FormMode <> mdvisualization Then
                  txtClause.SetFocus
                    If mCheckPoFields = True And mCheckLIFields = True And mCheckRemarks = True Then
                         If POClause.Count > 0 And FormMode <> mdvisualization Then
                                
                                If Len(Trim$(txtClause)) > 0 Then
                                
                                        txtClause = FixTheFirstCarriageReturn(txtClause)
                                        savetoPOclause
                                        MCheckClause = True
                                 Else
                                    MsgBox "Clause can not be left empty."
                                    MCheckClause = False
                                    sst_PO.Tab = 4
                                 End If
                                 
                          End If
                    End If
              End If
              
        End Select
        
    
    If Err Then Err.Clear
    NavBar1.NextEnabled = sst_PO.Tab <> 0
    NavBar1.LastEnabled = sst_PO.Tab <> 0
    NavBar1.FirstEnabled = sst_PO.Tab <> 0
    NavBar1.PreviousEnabled = sst_PO.Tab <> 0
    
    Select Case sst_PO.Tab
    
        Case 0
            
           FirstTimeAssignmentsHeader

            
            NavBar1.NextEnabled = True
            NavBar1.LastEnabled = True
            NavBar1.FirstEnabled = True
            NavBar1.PreviousEnabled = True
           
            
      If FormMode = mdvisualization Then
                    
                    NavBar1.EditEnabled = True
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = False
                    NavBar1.SaveEnabled = False
                 
                    If Not Poheader.EOF Then
                        NavBar1.NextEnabled = True
                        NavBar1.LastEnabled = True
                    End If
                    
                    If Not Poheader.BOF Then
                        NavBar1.PreviousEnabled = True
                        NavBar1.FirstEnabled = True
                    End If
                    'If NavBar1.EditVisible = False Then
                                  
         NavBar1.DeleteEnabled = NavBar1.EditEnabled
                 ElseIf FormMode = mdCreation Then
                 
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    
                    NavBar1.NextEnabled = False
                    NavBar1.LastEnabled = False
                    NavBar1.PreviousEnabled = False
                    NavBar1.FirstEnabled = False
                    NavBar1.DeleteEnabled = False
                    NavBar1.EMailEnabled = False 'JCG 2008/11/14
                 ElseIf FormMode = mdModification Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    NavBar1.NextEnabled = False
                    NavBar1.LastEnabled = False
                    NavBar1.PreviousEnabled = False
                    NavBar1.FirstEnabled = False
                    NavBar1.DeleteEnabled = False
                    NavBar1.EMailEnabled = False 'JCG 2008/11/14
                 End If
                 
        Case 1
        
             If PoReceipients Is Nothing Then Set PoReceipients = Mainpo.PoReceipients
             FirstTimeAssignmentsRec
                    
                         If PoReceipients.Move(Poheader.Ponumb) Then
                    
                           LoadFromPOReceipients
                        Else
                          'This means that there are no Line Items Corresponding to This PO
                           ClearPoReceipients
                        End If
             
             If FormMode = mdModification Then
                  NavBar1.NewEnabled = True
                  NavBar1.EditEnabled = False
                  NavBar1.CancelEnabled = True
                  NavBar1.EMailEnabled = False 'JCG 2008/11/14
                  cmdRemove.SetFocus
                  
             ElseIf FormMode = mdvisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
                  NavBar1.EMailEnabled = True 'JCG 2008/11/14
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  NavBar1.EMailEnabled = False 'JCG 2008/11/14
                  cmdRemove.SetFocus
    
                '------- JCG 2008/01/20
                'If newSupplier Then
                    getSupplierContacts
                'End If
                '-------
            End If
            
            
        Case 2
        
        
             If PoItem Is Nothing Then Set PoItem = Mainpo.POITEMS
               

               
            If Trim$(PoItem.Ponumb) <> Poheader.Ponumb Then
                'This is When the User Selects a New PO on POHEADER and Click POITEMS
                'the First Time.
        
                   GPOnumb = Poheader.Ponumb
                   
                   If PoItem.Move(GPOnumb) Then
                        LoadFromPOITEM
                        POFqa.MoveLineTo (PoItem.Linenumb)
                        LoadFromTOFQA
                   Else
                      'This means that there are no Line Items Corresponding to This PO
                        ClearAllPoLineItems
                   End If
                   
            
                    
                 If Not mIsPoItemsComboLoaded Then LoadPoItemCombos
              Else
                 'This CAse Arises when the User Clicks on POitem and Discovers
                 'that he forgot to click on CHK_FRMSTkmaster ,Gos back ,checks it on
                 'and comes back.
                 If mIsPoItemsComboLoaded = False Then LoadPoItemCombos
              
              End If
              
                  Call LoadToFQACombos
                  
                  FirstTimeAssignmentsPOITEM
                If FormMode = mdvisualization Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = False
                    NavBar1.DeleteEnabled = NavBar1.EditEnabled
                    NavBar1.EMailEnabled = True 'JCG 2008/11/14
                 ElseIf FormMode = mdCreation Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = True
                    NavBar1.DeleteEnabled = NavBar1.EditEnabled
                    NavBar1.EMailEnabled = False 'JCG 2008/11/14
                    If FormMode = mdCreation And PoItem.Count = 0 Then
                       NavBar1_OnNewClick
                    End If
                    txt_AFE = txt_ChargeTo
                    
                    If chk_FrmStkMst.value = 1 Then
                        ssdcboCommoditty.SetFocus
                    Else
                       txt_Requested.SetFocus
                    End If
                    
                 ElseIf FormMode = mdModification Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = False
                    NavBar1.DeleteEnabled = True
                    NavBar1.EMailEnabled = False 'JCG 2008/11/14
                    If FormMode = mdModification And PoItem.Count = 0 Then
                       NavBar1_OnNewClick
                    End If
                    
                    'If chk_FrmStkMst.Value = 1 Then
                     If ssdcboCommoditty.Enabled = True Then
                        Call HighlightBackground(ssdcboCommoditty)
                        ssdcboCommoditty.SetFocus
                        
                    Else
                    
                       Call HighlightBackground(txt_Requested)
                       txt_Requested.SetFocus
                       
                    End If
                    
                    If CInt(LblRevNumb) = 0 Then
                       NavBar1.DeleteEnabled = True
                    End If
                    
                    txt_AFE = txt_ChargeTo
                    
                    
                 End If
                 
                     
        
            
        Case 3

             If PORemark Is Nothing Then Set PORemark = Mainpo.POREMARKS
             
             If PORemark.Ponumb <> Poheader.Ponumb Then
                'This is When the User Selects a New PO on POHEADER and Click POClause
                'the First Time.
         
                        GPOnumb = Poheader.Ponumb
                    
                         If PORemark.Move(GPOnumb) Then
                         
                           LoadFromPORemarks
                        Else
                          'This means that there are no Line Items Corresponding to This PO
                           ClearPoRemarks
                        End If
              
             End If
             
             
             If FormMode = mdModification Then
                   NavBar1.NewEnabled = True
                   NavBar1.EditEnabled = False
                   NavBar1.CancelEnabled = True
                   NavBar1.EMailEnabled = False 'JCG 2008/11/14
                   txtRemarks.SetFocus
                   CmdcopyLI(1).Enabled = True
                   If PORemark.Count = 0 Then
                       NavBar1_OnNewClick
                   Else
                      HandleEdittingOfRemarks
                   End If
             ElseIf FormMode = mdvisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
                  NavBar1.EMailEnabled = True 'JCG 2008/11/14
                  CmdcopyLI(1).Enabled = False
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  NavBar1.EMailEnabled = False 'JCG 2008/11/14
                  txtRemarks.SetFocus
                  CmdcopyLI(1).Enabled = True
                 If PORemark.Count = 0 Then
                       NavBar1_OnNewClick
                 End If
            End If
            
        Case 4
        
             
             If POClause Is Nothing Then Set POClause = Mainpo.POClauses
             
             If POClause.Ponumb <> Poheader.Ponumb Then
                'This is When the User Selects a New PO on POHEADER and Click POClause
                'the First Time.
         
                    GPOnumb = Poheader.Ponumb
                    'POClause.Ponumb = GPOnumb
                    
                      If POClause.Move(GPOnumb) Then
                          'PORemark.Ponumb = GPOnumb
                    
                        'If Not (PORemark.EOF = True And PORemark.BOF = True) Then
                           LoadFromPOClause
                    
                        Else
                          'This means that there are no Line Items Corresponding to This PO
                           ClearPoclause
                          
                        End If
             End If
             
                        
            If FormMode = mdModification Then
                   
                   NavBar1.NewEnabled = True
                   NavBar1.EditEnabled = False
                   NavBar1.CancelEnabled = True
                   NavBar1.EMailEnabled = False 'JCG 2008/11/14
                   txtClause.SetFocus
                   
                   If POClause.Count = 0 Then
                       NavBar1_OnNewClick
                   ElseIf POClause.Count > 0 Then
                       HandleEdittingOfClause
                   End If
             ElseIf FormMode = mdvisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
                  NavBar1.EMailEnabled = True 'JCG 2008/11/14
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  NavBar1.EMailEnabled = False 'JCG 2008/11/14
                  txtClause.SetFocus
            
                  If POClause.Count = 0 Then
                       NavBar1_OnNewClick
                 End If
                  
            End If
                   
    End Select

Screen.MousePointer = vbArrow

End Sub

Public Function SetInitialVAluesPOITEM() As String
txt_LI = PoItem.Count
txt_TotalLIs = PoItem.Count
' DTP_Required = Format(Now(), "MM/DD/YY")
DTP_Required = dtpRequestedDate 'juan 2011-1-22
txt_AFE = txt_ChargeTo
chk_usexportLI = chk_USExportH
End Function

Public Function LoadPoItemCombos() As Boolean
Dim lookups As IMSPODLL.lookups
Dim RSStockNos As ADODB.Recordset
Dim RsRequsition As ADODB.Recordset
'Dim RsUNits As ADODB.Recordset

On Error GoTo Handler

mIsPoItemsComboLoaded = False
LoadPoItemCombos = False
Set lookups = Mainpo.lookups



     'Stockmaster
     If chk_FrmStkMst.value = 1 And mIsPoItemsComboLoaded = False Then
             
             ssdcboCommoditty.Enabled = True
             SSOleDBUnit.RemoveAll
             
              Set ssdcboCommoditty.DataSourceList = deIms.rsActiveStockmasterLookup  ' RSStockNos
              ssdcboCommoditty.DataFieldToDisplay = "stk_stcknumb"
              ssdcboCommoditty.DataFieldList = "stk_desc"
              
             
          mIsPoItemsComboLoaded = True
        
     
     ElseIf chk_FrmStkMst.value = 0 And mIsPoItemsComboLoaded = False Then
        
              ssdcboCommoditty.Enabled = False
        
             Set RsUNits = lookups.GetAllUnits
             
             Do While Not RsUNits.EOF
                SSOleDBUnit.AddItem RsUNits!uni_code & ";" & RsUNits!uni_desc
                RsUNits.MoveNext
             Loop
            
              mIsPoItemsComboLoaded = True
        ''    Load ALL Units of Stockmaster
    End If
         
    ' Requisition Number
    
         Set RsRequsition = lookups.GetRequisitions(deIms.NameSpace)
               
              If Not RsRequsition.EOF = True Then
                       ssdcboRequisition.Enabled = True
                  
                '2012-9-30 juan
                Set ssdcboRequisition.DataSourceList = RsRequsition
                ssdcboRequisition.DataFieldList = RsRequsition.Fields(0).Name
                   
'                   Do While Not RsRequsition.EOF
'                      ssdcboRequisition.AddItem RsRequsition!po_ponumb & ";" & RsRequsition!doc_desc & ";" & RsRequsition!poi_liitnumb & ";" & RsRequsition!poi_desc & ";" & RsRequsition!poi_primreqdqty
'                      RsRequsition.MoveNext
'                   Loop
                   
              Else
                    ssdcboRequisition.Text = ""
                    ssdcboRequisition.Enabled = False
              End If
               
    ' Eccn
    
    Call FillEccnCombos(lookups)
               
          Set lookups = Nothing
          
    LoadPoItemCombos = True
    mIsPoItemsComboLoaded = True
    Exit Function
          
Handler:
   MsgBox "Coud Not load all the Poitem Combos.  " & Err.Description
  Err.Clear
End Function


Public Function ClearAllPoLineItems() As Boolean
On Error GoTo Handler
ClearAllPoLineItems = False

mLoadMode = loadingPoItem

LblPOi_PONUMB = Trim$(ssOleDbPO) 'AM
txt_LI = 0
txt_TotalLIs = 0
ssdcboCommoditty = ""
ssdcboManNumber = ""
txt_AFE = ""
SSOleDBCustCategory = ""
txt_SerialNum = ""
txt_Requested = ""

SSOleDBUnit = ""

txt_Delivered = ""
txt_Shipped = ""
txt_Inventory2 = ""
txt_Price = ""
txt_Total = ""
txt_Descript = ""
txt_remk = ""
dcbostatus(1) = ""
dcbostatus(2) = ""
dcbostatus(3) = ""
dcbostatus(0) = ""
ssdcboRequisition = ""
lblReqLineitem = ""


SSoleEccnNo.Tag = 0
SSoleEccnNo = ""

SSOleSourceofinfo.Tag = 0
SSOleSourceofinfo = ""

chk_usexportLI.value = False
Chk_license.value = 0


mLoadMode = NoLoadInProgress

ClearAllPoLineItems = True
 
Exit Function
Handler:

    mLoadMode = NoLoadInProgress

  Err.Clear
End Function


Public Function LoadFromPORemarks() As Boolean
mLoadMode = loadingPoRemark
Txt_RemNo = PORemark.Linenumb
txtRemarks.Text = PORemark.remarks
mLoadMode = NoLoadInProgress
End Function

Public Function ClearPoRemarks() As Boolean
txtRemarks.Text = ""
Txt_RemNo = PORemark.Count
End Function

Public Function ClearPoclause() As Boolean
txtClause.Text = ""
Txt_ClsNo = POClause.Count
End Function
Public Function SaveToPOITEM() As Boolean

PoItem.NameSpace = deIms.NameSpace

'PoItem.Ponumb = LblPOi_PONUMB AM

'In case of autoNumbering for a new PO , the POheader object would always have a POnumber in it.
PoItem.Ponumb = Poheader.Ponumb 'AM

PoItem.Linenumb = txt_LI

'Incase this is a NON-STOCK case.

    If PoItem.EditMode = 2 Then
        If chk_FrmStkMst.value = 0 Then
              'PoItem.Comm = Trim$(ssOleDbPO.text) & "/" & Trim$(txt_LI) AM
               PoItem.Comm = Trim$(Poheader.Ponumb) & "/" & Trim$(txt_LI) 'AM
        Else
              PoItem.Comm = Trim$(ssdcboCommoditty)
        End If
    End If
    
PoItem.Manupartnumb = ssdcboManNumber

PoItem.Afe = txt_AFE

PoItem.Custcate = SSOleDBCustCategory

PoItem.Serlnumb = txt_SerialNum

Call SaveUnitsToPoItem

PoItem.Description = txt_Descript

PoItem.Remk = txt_remk

PoItem.Liitreqddate = IIf(IsNull(DTP_Required.value), "", DTP_Required.value)

PoItem.Requnumb = Trim$(ssdcboRequisition)

PoItem.Requliitnumb = (lblReqLineitem)

PoItem.usexport = chk_usexportLI
PoItem.Eccnlicsreq = Chk_license
PoItem.Eccnid = SSoleEccnNo.Tag
PoItem.Sourceofinfoid = IIf(Len(SSOleSourceofinfo.Tag) = 0, 0, SSOleSourceofinfo.Tag)
'PoItem.Eccnno = Trim(SSoleEccnno)

If Not objUnits Is Nothing Then Set objUnits = Nothing

End Function

Public Sub FirstTimeAssignmentsPOITEM()

On Error GoTo ErrHand

mLoadMode = loadingPoItem

LblPOI_Doctype.Caption = SSOleDBDocType.Text

If chk_FrmStkMst.value = 1 Then
    
    txt_Descript.locked = True

    ssdcboCommoditty.Enabled = True
    
        If FormMode = mdModification And PoItem.EditMode <> 2 Then
           
           ssdcboCommoditty.Enabled = False
        
        Else
               
           ssdcboCommoditty.Enabled = True
        
        End If
        
 Else
       txt_Descript.locked = False

       ssdcboCommoditty.Enabled = False
       
 End If
 

If Trim(UCase(SSOleDBDocType.Tag)) = "R" Then

    ssdcboRequisition.Enabled = False
    
Else

    ssdcboRequisition.Enabled = True
    
End If

chk_usexportLI = chk_USExportH
'SSoleEccnno.Tag = 0
'SSoleEccnno = ""
'Chk_license.Value = 0
mLoadMode = NoLoadInProgress

Exit Sub
ErrHand:

    MsgBox Err.Description
    Err.Clear

End Sub

Private Sub Text1_Change()

If opt_Email.value = False And opt_FaxNum.value = False Then Exit Sub
If Len(Trim$(Text1)) = 0 Then RsEmailFax.MoveFirst: Exit Sub
   RsEmailFax.MoveFirst
   RsEmailFax.Find "phd_name like '" & Text1.Text & "%'"

End Sub

Private Sub txt_AFE_GotFocus()

Call HighlightBackground(txt_AFE)
End Sub

Private Sub txt_AFE_LostFocus()
Call NormalBackground(txt_AFE)
End Sub

Private Sub txt_ChargeTo_GotFocus()

 Call HighlightBackground(txt_ChargeTo)
End Sub

Private Sub txt_ChargeTo_LostFocus()
Call NormalBackground(txt_ChargeTo)
End Sub

Private Sub txt_Descript_GotFocus()
Call HighlightBackground(txt_Descript)
End Sub

Private Sub txt_Descript_LostFocus()
Call NormalBackground(txt_Descript)
End Sub

Private Sub txt_Price_Change()
'If PoItem.editmode <> 2 And PoItem.editmode <> -1 And Len(Trim$(txt_Price)) > 0 And mLoadMode = NoLoadInProgress Then Call SaveUnitsToPoItem
End Sub

Private Sub txt_Price_GotFocus()
Call HighlightBackground(txt_Price)
End Sub

Private Sub txt_Price_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
If KeyAscii = 13 Then
    If Not IsNumeric(txt_Price) Then
       MsgBox "Unit Price has to be numeric.", vbInformation, "Imswin"
       KeyAscii = 0
    End If
End If
End Sub

Private Sub txt_Price_LostFocus()
Call NormalBackground(txt_Price)
End Sub

Private Sub txt_Price_Validate(Cancel As Boolean)
Dim x As Integer
Dim y As Integer
If Len(txt_Price) > 0 Then
       
     If Not IsNumeric(txt_Price) Then
       Cancel = True
       MsgBox "Unit Price should be numeric."
        txt_Price.SetFocus
       Exit Sub
    End If
            
    If CDbl(txt_Price) < 0 Then
       Cancel = True
       MsgBox "Unit Price should be greater than 0."
       txt_Price.SetFocus
       Exit Sub
    End If
     
     txt_Price = Replace(FormatNumber(txt_Price, 2), ",", "")
     x = InStr(1, txt_Price, ".")
     
     If Len(Mid(txt_Price, 1, x - 1)) > 7 Then
        Cancel = True
       MsgBox "Unit Price can not be more than 7 digits before the decimal point."
       txt_Price.SetFocus
       Exit Sub
    End If
        
    If Len(Mid(txt_Price, x + 1, Len(txt_Price))) > 2 Then
        Cancel = True
       MsgBox "Unit Price can not be more than 2 digits after the decimal point."
       txt_Price.SetFocus
       Exit Sub
    End If
        
   If Len(txt_Requested) > 0 Then
        txt_Total = FormatNumber(CDbl(txt_Price) * txt_Requested, 2)
   End If
   
End If
End Sub

Private Sub txt_remk_GotFocus()
Call HighlightBackground(txt_remk)
End Sub

Private Sub txt_remk_LostFocus()
Call NormalBackground(txt_remk)
End Sub

Private Sub txt_Requested_Change()

'If PoItem.editmode <> 2 And PoItem.editmode <> -1 And Len(Trim$(txt_Requested)) > 0 And mLoadMode = NoLoadInProgress Then Call SaveUnitsToPoItem

End Sub

Private Sub txt_Requested_GotFocus()
Call HighlightBackground(txt_Requested)
End Sub

Private Sub txt_Requested_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
If KeyAscii = 13 Then
    If Not IsNumeric(txt_Requested) Then
       MsgBox "Quantity requested has to be numeric.", vbInformation, "Imswin"
       KeyAscii = 0
    End If
End If
End Sub

Private Sub txt_Requested_LostFocus()
Call NormalBackground(txt_Requested)
End Sub

Private Sub txt_Requested_Validate(Cancel As Boolean)
Dim x As Integer
Dim y As Integer

If Len(Trim$(txt_Requested)) > 0 Then
    If Not IsNumeric(txt_Requested) Then
       Cancel = True
       MsgBox "Quantity Requested should be numeric."
        txt_Requested.SetFocus
       Exit Sub
    End If
       
     If CDbl(txt_Requested) < 0 Then
       Cancel = True
       MsgBox "Quantity Requested should be greater than 0."
       txt_Requested.SetFocus
       Exit Sub
    End If
       
    If IsPrimQuantLessThanONE = False Then
        txt_Requested.SetFocus
        Cancel = True
        'txt_Requested = 0
        Exit Sub
    End If
    
    If CDbl(txt_Requested) = 0 Then
       
        If Not txt_remk = "THIS ITEM HAS BEEN CANCELLED." Then
                If MsgBox("Are you sure you want to cancel this Line Item?", vbCritical + vbYesNo, "Imswin") = vbYes Then
                 txt_remk = "THIS ITEM HAS BEEN CANCELLED."
                 txt_Price = 0
                 Else
                     Cancel = True
                      Exit Sub
                 End If
        End If
    End If
    
       txt_Requested = Replace(FormatNumber(txt_Requested, 4), ",", "")


     x = InStr(1, txt_Requested, ".")
     
     If Len(Mid(txt_Requested, 1, x - 1)) > 7 Then
        Cancel = True
       MsgBox "Quantity Requested can not be more than 7 digits before the decimal point."
       txt_Requested.SetFocus
       Exit Sub
    End If
        
    If Len(Mid(txt_Requested, x + 1, Len(txt_Requested))) > 4 Then
        Cancel = True
       MsgBox "Quantity Requested can not be more than 4 digits after the decimal point."
       txt_Requested.SetFocus
       Exit Sub
    End If


       If Len(txt_Price) > 0 Then
          txt_Total = FormatNumber(txt_Price * txt_Requested, 2)
       End If
       
 End If
   
End Sub


Public Function LoadPOLINEFromRequsition(rs As ADODB.Recordset) As Boolean

ssdcboCommoditty.Text = rs("poi_comm") & ""
ssdcboManNumber = rs("poi_manupartnumb") & ""
txt_AFE = rs("poi_afe") & ""
SSOleDBCustCategory = rs("poi_custcate") & ""
txt_SerialNum = rs("poi_serlnumb") & ""

txt_Delivered = ""
txt_Shipped = ""
txt_Inventory2 = ""

txt_Price = rs("poi_unitprice") & ""

SSOleDBUnit.RemoveAll
SSOleDBUnit = rs("poi_primuom")
txt_Total = FormatNumber(rs("poi_unitprice") * rs("poi_primreqdqty"), 2)
txt_Requested = rs("poi_primreqdqty") & ""

If RsUNits Is Nothing Then
  If lookups Is Nothing Then lookups = Mainpo.lookups
  Set RsUNits = lookups.GetAllUnits
End If

RsUNits.MoveFirst
RsUNits.Find ("uni_code='" & Trim$(rs("poi_primuom")) & "'")
SSOleDBUnit.AddItem rs("poi_primuom") & ";" & RsUNits("uni_desc")

RsUNits.MoveFirst
RsUNits.Find ("uni_code='" & Trim$(rs("poi_primuom")) & "'")
SSOleDBUnit.AddItem rs("poi_secouom") & ";" & RsUNits("uni_desc")

txt_Descript = rs("poi_desc") & ""
txt_remk = ""
dcbostatus(1) = ""
dcbostatus(2) = ""
dcbostatus(3) = ""
dcbostatus(0) = ""

SSoleEccnNo.Tag = IIf(IsNull(rs("poi_eccnid")), 0, rs("poi_eccnid"))
SSoleEccnNo = rs("eccn_no") & ""

SSOleSourceofinfo.Tag = IIf(IsNull(rs("poi_sourceid")), 0, rs("poi_sourceid"))
SSOleSourceofinfo = rs("source") & ""

Dim EccnLicsReq1 As Boolean
EccnLicsReq1 = IIf(IsNull(rs("poi_eccnlicsreq")), False, rs("poi_eccnlicsreq"))
Chk_license = IIf(EccnLicsReq1 = True, 1, 0)
'ssdcboRequisition =
'lblReqLineitem = ""
End Function

Private Sub txt_SerialNum_Change()
'If PoItem.editmode <> 2 Or PoItem.editmode <> -1 And mLoadMode = NoLoadInProgress Then PoItem.Serlnumb = txt_SerialNum
End Sub



Public Function SaveUnitsToPoItem() As Boolean
Dim Pqty As Double
SaveUnitsToPoItem = False
On Error GoTo Handler

If chk_FrmStkMst.value = 1 Then
 
   'This means the PO is In-Stock
        If objUnits Is Nothing Then
           Set objUnits = Mainpo.PoUnits
           objUnits.StockNumber = ssdcboCommoditty.Text
        End If
        
        If objUnits.SecondaryUnit = Trim$(SSOleDBUnit.Text) And objUnits.SecondaryUnit <> objUnits.PrimaryUnit Then
           'It means it is in Seconday mode
                 PoItem.Secoreqdqty = txt_Requested
                 PoItem.Secouom = SSOleDBUnit.Text
                 PoItem.SecUnitPrice = (CDbl(txt_Price))
                 PoItem.SecTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
                 
                 PoItem.UnitOfPurch = "S"
                 'Juan 2010-9-7
                 'Pqty = CDbl(txt_Requested) * objUnits.ComputationFactor / 10000
                 Pqty = CDbl(txt_Requested) / objUnits.ratioValue
                 PoItem.Primreqdqty = Pqty
                 PoItem.Primuom = objUnits.PrimaryUnit
                 'PoItem.PrimUnitprice = CDbl(txt_Price) / objUnits.ComputationFactor * 10000
                 PoItem.PrimUnitprice = CDbl(txt_Price) * objUnits.ratioValue
                 '----------
                 PoItem.PriTotaprice = CDbl(PoItem.Primreqdqty) * CDbl(PoItem.PrimUnitprice)
                 PoItem.PriQtytobedlvd = Pqty - PoItem.PriQtydlvd
                 
        ElseIf objUnits.PrimaryUnit = Trim$(SSOleDBUnit.Text) And objUnits.SecondaryUnit <> objUnits.PrimaryUnit Then
                 'It is in Primary Mode
          
                 PoItem.Primreqdqty = CDbl(txt_Requested)
                 PoItem.Primuom = SSOleDBUnit.Text
                 PoItem.PrimUnitprice = CDbl(txt_Price)
                 PoItem.PriTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
                 PoItem.PriQtytobedlvd = CDbl(txt_Requested) - PoItem.PriQtydlvd
                 PoItem.UnitOfPurch = "P"
                 'Juan 2010-9-7
                 'PoItem.Secoreqdqty = CDbl(txt_Requested) / objUnits.ComputationFactor * 10000
                 PoItem.Secoreqdqty = CDbl(txt_Requested) * objUnits.ratioValue
                 PoItem.Secouom = objUnits.SecondaryUnit ' RsUNits!stk_secouom
                 'PoItem.SecUnitPrice = CDbl(txt_Price) * objUnits.ComputationFactor / 10000
                 PoItem.SecUnitPrice = CDbl(txt_Price) / objUnits.ratioValue
                 PoItem.SecTotaprice = CDbl(PoItem.Secoreqdqty) * CDbl(PoItem.SecUnitPrice)
        
        ElseIf objUnits.SecondaryUnit = objUnits.PrimaryUnit And objUnits.ComputationFactor = 0 Then
                
                If Len(txt_Requested) > 0 Then
                     PoItem.Secoreqdqty = txt_Requested
                End If
                PoItem.Secouom = SSOleDBUnit.Text
                If Len(txt_Price) > 0 Then
                PoItem.SecUnitPrice = IIf(Len(txt_Price) > 0, CDbl(txt_Price), "")
                End If
                If Len(txt_Requested) > 0 Then
                PoItem.SecTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
                End If
                PoItem.UnitOfPurch = "P"
                
                If Len(txt_Requested) > 0 Then
                   PoItem.Primreqdqty = CDbl(txt_Requested)
                End If
                PoItem.Primuom = SSOleDBUnit.Text
                If Len(txt_Price) > 0 Then
                   PoItem.PrimUnitprice = CDbl(txt_Price)
                End If
                
                If Len(txt_Price) And Len(txt_Requested) > 0 Then
                PoItem.PriTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
                End If
                
                PoItem.PriQtytobedlvd = CDbl(txt_Requested) - PoItem.PriQtydlvd
                
        End If
   
   Else
   
    'This means the po is Non-Stock
        If Len(txt_Requested) > 0 Then
         PoItem.Secoreqdqty = txt_Requested
        End If
        
        
        PoItem.Secouom = SSOleDBUnit.Text
        If Len(txt_Price) > 0 Then
           PoItem.SecUnitPrice = IIf(Len(txt_Price) > 0, txt_Price, 0)
        End If
        
        If Len(txt_Price) And Len(txt_Requested) > 0 Then
          PoItem.SecTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
        End If
        
        PoItem.UnitOfPurch = "P"
        If Len(txt_Requested) > 0 Then
          PoItem.Primreqdqty = CDbl(txt_Requested)
        End If
        PoItem.Primuom = SSOleDBUnit.Text
        If Len(txt_Price) > 0 Then
        PoItem.PrimUnitprice = CDbl(txt_Price)
        End If
        
        PoItem.PriQtytobedlvd = CDbl(txt_Requested) - PoItem.PriQtydlvd
        
        If Len(txt_Price) And Len(txt_Requested) > 0 Then
            PoItem.PriTotaprice = CDbl(txt_Requested) * CDbl(txt_Price)
        End If
   End If
    SaveUnitsToPoItem = True
  Exit Function
Handler:
   MsgBox Err.Description
   Err.Clear


End Function

Public Function ToggleNavButtons(FMode As FormMode) As Boolean


 
        If FormMode = mdvisualization Then
                    
                    NavBar1.EditEnabled = True
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = False
                    NavBar1.SaveEnabled = False
                 
'''                    If Not Poheader.EOF Then
'''                        NavBar1.NextEnabled = True
'''                        NavBar1.LastEnabled = True
'''                    End If
'''
'''                    If Not Poheader.BOF Then
'''                        NavBar1.PreviousEnabled = True
'''                        NavBar1.FirstEnabled = True
'''                    End If
'''
                 ElseIf FormMode = mdCreation Then
                 
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    
                    NavBar1.NextEnabled = False
                    NavBar1.LastEnabled = False
                    NavBar1.PreviousEnabled = False
                    NavBar1.FirstEnabled = False
                    
                 ElseIf FormMode = mdModification Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = True
                    NavBar1.SaveEnabled = True
                    
                    NavBar1.NextEnabled = False
                    NavBar1.LastEnabled = False
                    NavBar1.PreviousEnabled = False
                    NavBar1.FirstEnabled = False
                    
                 End If
End Function

Public Function LoadFromPOClause() As Boolean

mLoadMode = loadingPoClause
Txt_ClsNo = POClause.Linenumb
txtClause.Text = POClause.Clause
mLoadMode = NoLoadInProgress

End Function

Public Function savetoPOclause() As Boolean

 If POClause.EditMode = 2 Then
    POClause.NameSpace = deIms.NameSpace
    'POClause.Linenumb = POClause.Count
    POClause.Linenumb = CInt(Txt_ClsNo)
    POClause.Ponumb = Poheader.Ponumb
 End If
 
 POClause.Clause = Trim(txtClause.Text)

End Function

Public Function savetoPORemarks() As Boolean

If PORemark.EditMode = 2 Then
    PORemark.NameSpace = deIms.NameSpace
    'PORemark.Linenumb = PORemark.Count
    PORemark.Linenumb = CInt(Txt_RemNo)
    PORemark.Ponumb = Poheader.Ponumb
End If

PORemark.remarks = Trim(txtRemarks.Text)

End Function

Public Function LoadFromPOReceipients() As Boolean
mLoadMode = loadingPoRemark
dgRecipientList.RemoveAll
If PoReceipients.Count > 0 Then PoReceipients.MoveFirst
 Do While Not PoReceipients.EOF
    If LTrim(PoReceipients.Receipient) <> "" Then ' JCG 2008/8/16
        dgRecipientList.AddItem PoReceipients.Receipient
    End If ' JCG 2008/8/16
    PoReceipients.MoveNext
 Loop
mLoadMode = NoLoadInProgress
 
End Function

Public Function ClearPoReceipients()
dgRecipientList.RemoveAll
End Function

Private Sub AddRecepient(RecipientName As String, Optional ShowMessage As Boolean = True, Optional DoPrefix As Boolean, Optional RecepientType As String)
On Error GoTo errorHandler
Dim retval As Long

    If PoReceipients Is Nothing Then
            Set PoReceipients = Mainpo.PoReceipients
            PoReceipients.Move Poheader.Ponumb
            LoadFromPOReceipients
    End If

    If Len(Trim$(RecipientName)) = 0 Then Exit Sub
      RecipientName = UCase(RecipientName)
    If DoPrefix = True Then
        RecipientName = PrefixFaxorEmail(RecipientName)
    End If

        If IsRecipientInList(RecipientName, ShowMessage) Then Exit Sub
    
    If DoPrefix = True Then
        If ((opt_FaxNum) And (InStr(1, RecipientName, "", vbTextCompare) = 0)) And (InStr(1, RecipientName, "", vbTextCompare) = 0) Then _
            RecipientName = FixFaxNumber(RecipientName)
    End If

    If DoPrefix = False Then
        If ((OptFax) And (InStr(1, RecipientName, "", vbTextCompare) = 0)) And (InStr(1, RecipientName, "", vbTextCompare) = 0) Then _
            RecipientName = FixFaxNumber(RecipientName)
    End If

    ' JCG 2008/7/30
    If LTrim(RecipientName) <> "" Then ' JCG 2008/8/16
        dgRecipientList.AddItem RecipientName
    End If 'JCG 2008/8/16
    dgRecipientList.MoveLast
    'dgRecipientList.Columns(1).value = "supplierContact"
    '----------------------
    With PoReceipients
        
        
        .AddNew
        .Linenumb = PoReceipients.Count
'''        If Len(Poheader.Ponumb) = 0 Then AM
'''             .Ponumb = Trim$(ssOleDbPO.text) AM
'''         Else AM
'''             .Ponumb = Poheader.Ponumb AM
'''         End If AM
        
        .Ponumb = Poheader.Ponumb 'AM
        .Receipient = RecipientName
        .NameSpace = deIms.NameSpace
        
       
    End With
    
    Exit Sub
errorHandler:
    MsgBox "Error in cmd_Add_Click: " + Err.Description
End Sub

Public Sub SortGrid(rs As ADODB.Recordset, Grid As DataGrid, Col As Integer)
On Error Resume Next
    Dim SortOrder As String
    Dim BK As Variant

    BK = rs.Bookmark
    SortOrder = Grid.Tag
    SortOrder = IIf(UCase(SortOrder) = "ASC", "ASC", "DESC")
    Grid.Tag = IIf(UCase(SortOrder) = "ASC", "DESC", "ASC")

    rs.Sort = ""
    rs.Sort = ((Grid.Columns(Col).DataField) + " " + SortOrder)
    rs.Bookmark = BK
    If Err Then Err.Clear
End Sub
Private Function IsRecipientInList(RecepientName As String, Optional ShowMessage As Boolean = True) As Boolean
IsRecipientInList = False
On Error GoTo Handler

If PoReceipients.Count > 0 Then PoReceipients.MoveFirst

Do While Not PoReceipients.EOF
  
   If PoReceipients.Receipient = RecepientName Then
              IsRecipientInList = True
              Exit Do
   End If
   
   PoReceipients.MoveNext
Loop

   Exit Function

Handler:
   
   Err.Raise Err.number, , Err.Description
   
   Err.Clear
   
End Function

Private Function FixFaxNumber(Faxnumber As String) As String
On Error Resume Next

    If Len(Faxnumber) < 7 Then Exit Function

    If Left$(Faxnumber, 1) = "+" Then
        Faxnumber = Right$(Faxnumber, Len(Faxnumber) - 1)
    End If
    
    If Mid$(Faxnumber, 1, 4) <> "" Then _
        FixFaxNumber = "" & Faxnumber

    'Modified by Juan (9/14/2000) for Multilingual
  '  msg1 = translator.Trans("M00078") 'J added
  '  If Err Then Err.Clear: MsgBox IIf(msg1 = "", "err occured", msg1) 'J modified
    '---------------------------------------------

End Function


Public Sub PoReceipeintsInit()
dgRecipientList.Columns(0).Width = dgRecipientList.Width
End Sub

Public Sub FirstTimeAssignmentsRec()

If FormMode = mdCreation Then
   cmdRemove.Visible = True
ElseIf FormMode = mdModification Then
   cmdRemove.Visible = True
End If
dgRecipientList.Caption = "Recipients"
End Sub




Public Function LoadDocType()
If mIsDocTypeLoaded = False Then
      If lookups Is Nothing Then Set lookups = Mainpo.lookups
     Dim GRsDoctype As ADODB.Recordset
     Set GRsDoctype = lookups.GetDoctypeForUser(CurrentUser)
        
        Do While Not GRsDoctype.EOF
           rsDOCTYPE.MoveFirst
           rsDOCTYPE.Find ("DOC_CODE='" & Trim$(GRsDoctype!buyr_docutype) & "'")
          
           SSOleDBDocType.AddItem GRsDoctype!buyr_docutype & ";" & rsDOCTYPE!doc_desc
           GRsDoctype.MoveNext
        Loop
    mIsDocTypeLoaded = True
       Set GRsDoctype = Nothing
  End If
End Function
Private Function CheckPoFields() As Boolean

CheckPoFields = False
On Error GoTo Handled
Dim i As Long

   ' i = rsPO.editmode
    'If i = adEditNone Then CheckPoFields = True: Exit Function
    
    
    'If Len(Trim$(SSOleDBPO)) = 0 Then 'AM
     '  MsgBox "Transaction number can not be left empty." 'AM
      ' SSOleDBPO.SetFocus 'AM
      ' Exit Function 'AM
   ' End If 'AM
        
    If Len(Trim$(ssOleDbPO)) > 15 Then
       MsgBox "Transaction number can not be more than 15 characters."
       ssOleDbPO.SetFocus
       Exit Function
    End If
    
    If Len(Trim$(SSOleDBDocType.Text)) = 0 Then
        'Call MsgBox(LoadResString(101)): dcboDocumentType.SetFocus: Exit Function
        MsgBox "Document Type can not be Left Empty."
        SSOleDBDocType.SetFocus
        Exit Function
    End If
        
    If Len(Trim$(ssdcboShipper.Text)) = 0 Then
        'Call MsgBox(LoadResString(102)): ssdcboShipper.SetFocus: Exit Function
        MsgBox "Shipper can not be Left Empty."
        ssdcboShipper.SetFocus
        Exit Function
    End If
    'Else
    
   '     rsPO!po_shipcode = ssdcboShipper.Value
   ' End If
        
    
    If Len(Trim$(SSOleDBPriority.Text)) = 0 Then
        'Call MsgBox(LoadResString(103)): dcboPriority.SetFocus: Exit Function
         MsgBox "Priority can not be Left Empty."
         SSOleDBPriority.SetFocus: Exit Function
    'Else
      '  rsPO!po_priocode = dcboPriority.BoundText
    End If
    
    If Len(Trim$(SSOleDBCurrency.Text)) = 0 Then
        'Call MsgBox(LoadResString(104)):
        MsgBox "Currency can not Be left Empty"
        SSOleDBCurrency.SetFocus: Exit Function
        
   ' Else
    '    rsPO!po_currcode = dcboCurrency.BoundText
    End If
    
    
    If Len(Trim$(SSOleDBOriginator.Text)) = 0 Then
       ' Call MsgBox(LoadResString(105)): dcboOriginator.SetFocus: Exit Function
       MsgBox "Originator can not be Left Empty"
       SSOleDBCurrency.SetFocus
       Exit Function
    'Else
     '   rsPO!po_orig = dcboOriginator.BoundText
    End If
    
        
    If Len(Trim$(SSOleDBShipTo.Text)) = 0 Then
        'Call MsgBox(LoadResString(106)):
        MsgBox "Ship to can not be left empty."
        SSOleDBShipTo.SetFocus
        
        Exit Function
    End If
        
''''
''''    'Else
''''     '   rsPO!po_shipto = dcboShipto.BoundText
''''
''''    End If
    
    
    If Len(Trim$(SSOleDBcompany.Text)) = 0 Then  'M
    
        'Modified by Juan (9/13/2000) for Multilingual
       ' msg1 = translator.Trans("M00023") 'J added
        'MsgBox IIf(msg1 = "", "Company Can not be left empty", msg1), , "Imswin" 'J modified
        '---------------------------------------------
        MsgBox "Company can not be left Empty."
      SSOleDBcompany.SetFocus
      Exit Function 'M
        
    'Else  'M
     '  rsPO!po_compcode = dcboCompany.BoundText 'M
    End If 'M
    
    

    
    If Len(Trim$(SSOleDBInvLocation.Text)) = 0 Then
      '  Call MsgBox(LoadResString(107)): SSOleDBInvLocation.SetFocus: Exit Function
        MsgBox "Inventory Location Can not be left Empty."
        SSOleDBInvLocation.SetFocus: Exit Function
    'Else
     '   rsPO!po_invloca = dcboInvLocation.BoundText
    End If
    
    If Len(Trim$(SSoledbSupplier.Text)) = 0 Then
       ' Call MsgBox(LoadResString(108)): dcboSupplier.SetFocus: Exit Function
        MsgBox "Supplier Can not be Left Empty"
        SSoledbSupplier.SetFocus: Exit Function
    'Else
     '   rsPO!po_suppcode = dcboSupplier.Value
    End If
    
    'Modified by Muzammil 08/14/00
    'Reason - Should scream at the user when left empty and the user tries clicking some
    'other tab.
    
    
    If Len(Trim$(ssdcboCondition.Text)) = 0 Then          'M
    
        'Modified by Juan (9/13/2000) for Multilingual
        'msg1 = translator.Trans("M00034") 'J added
'        MsgBox IIf(msg1 = "", "T & C can not be left empty ", msg1) 'J modified
        '---------------------------------------------
        MsgBox "T & C can not be left empty "
        ssdcboCondition.SetFocus
        Exit Function 'M
    'Else   'M
     '   rsPO!po_taccode = ssdcboCondition.Value  'M
    End If  'M
    
    If Len(Trim$(ssdcboDelivery.Text)) = 0 Then  'M
    
        'Modified by Juan (9/14/2000) for Multilingual
       ' msg1 = translator.Trans("M00035") 'J added
'        MsgBox IIf(msg1 = "", "Payment Term can not be left empty ", msg1) 'J modified
        '---------------------------------------------
       MsgBox "Payment Term can not be left empty. "
        ssdcboDelivery.SetFocus: Exit Function 'M
        
    'Else  'M
      '  rsPO!po_termcode = ssdcboDelivery.Value 'M
    End If  'M
    
    
    
    'If Len(Trim$(ssOleDbPO.Text)) Then rsPO!po_ponumb = dcboPO
    
    'If Len(Trim$(SSOleDBToBeUsedFor)) Then rsPO!po_tbuf = SSOleDBToBeUsedFor.BoundText
    'If Len(Trim$(ssdcboDelivery)) Then rsPO!po_termcode = ssdcboDelivery.Value   'M
    'If Len(Trim$(ssdcboCondition)) Then rsPO!po_taccode = ssdcboCondition.Value   'M
    
    'Added by muzammil to Make sure the Po date < po requested date
    If CDate(DTPicker_poDate) > dtpRequestedDate.value Then 'Or DTPicker_poDate.Value = dtpRequestedDate.Value Then   'M  'J Modified
       MsgBox " Transaction Requested Date should be greater than Transaction Create Date by atleast one day." 'M
       dtpRequestedDate.SetFocus
       Exit Function  'M
    End If  'M
    
    CheckPoFields = True
    Exit Function
        
Handled:
    
    If Err Then Err.Clear
End Function
Private Function CheckLIFields() As Boolean
CheckLIFields = False
On Error GoTo Handler
Dim i As Long


    
        Call txt_Requested_Validate(False)
    
    If Len(Trim$(txt_Requested)) = 0 Then
    

    '    msg1 = translator.Trans("M00029") 'J added
     '   Call MsgBox(IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1))
        '---------------------------------------------
        MsgBox "Quantity Required does not contain a valid entry."
        txt_Requested.SetFocus
        Exit Function
    
    ElseIf Not IsNumeric(Trim$(txt_Requested)) Then
    

        'msg1 = translator.Trans("M00029") 'J added
        'MsgBox IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1)
        '---------------------------------------------
        MsgBox "Quantity Required does not contain a valid entry"
        txt_Requested.SetFocus: Exit Function
    Else
         If IsPrimQuantLessThanONE = False Then
                txt_Requested.SetFocus
                'Cancel = True
                'txt_Requested = 0
                Exit Function
        End If
        
    End If
    
    If Len(Trim$(txt_Price)) = 0 Then
    

        'msg1 = translator.Trans("M00030") 'J added
        'MsgBox IIf(msg1 = "", "Price cannot be left empty ", msg1) 'J modified
        '---------------------------------------------
        MsgBox "Unit Price cannot be left empty "
        txt_Price.SetFocus: Exit Function
        
    ElseIf Not (IsNumeric(txt_Price)) Then
    

        'msg1 = translator.Trans("M00031") 'J added
        'MsgBox IIf(msg1 = "", "Price does not have a valid entry", msg1) 'J modified
        '---------------------------------------------
        MsgBox "Unit Price does not have a valid entry"
        txt_Price.SetFocus: Exit Function
        

    End If
    

    
    If chk_FrmStkMst.value = vbChecked Then
    
        If Len(Trim$(ssdcboCommoditty.Text)) = 0 Then
        
            'Modified by Juan (9/13/2000) for Multilngual
            'msg1 = translator.Trans("M00025") 'J added
            'MsgBox IIf(msg1 = "", "Stock Number cannot be left empty", msg1) 'J modified
            '--------------------------------------------
            MsgBox "Stock Number cannot be left empty"
            ssdcboCommoditty.SetFocus: Exit Function
         End If

    End If
    If Len(Trim$(SSOleDBUnit)) = 0 Then
        MsgBox "Unit can not be Left Empty.", , "Imswin"
        SSOleDBUnit.SetFocus
        Exit Function
    End If
    
    If Len(ssdcboManNumber) > 30 Then
       MsgBox "Manufacturer text can not be greater than 30 charaters."
       ssdcboManNumber.SetFocus
       Exit Function
    End If
    
    If chk_FrmStkMst.value = 0 And Len(Trim$(txt_Descript)) = 0 Then
        MsgBox "Stock Description can not be left empty.", vbInformation, "Imswin"
        Call HighlightBackground(txt_Descript)
        txt_Descript.SetFocus
        Exit Function
    
    Else
     
        txt_Descript = FixTheFirstCarriageReturn(txt_Descript)
        
    End If
    
    If Len(Trim$(txt_remk)) > 0 Then
     
        txt_remk = FixTheFirstCarriageReturn(txt_remk)
        
    End If
    
    'this If Statement is in the case when the PO is an old one and there is no associated FQA for it
    If POFqa.Count > 0 Then
    
        If Len(Trim(SSOleDBToLocationFQA.Text)) = 0 Then
        
            MsgBox "Please fill the ToFQA location.", vbInformation, "Ims"
            Call HighlightBackground(SSOleDBToLocationFQA)
            SSOleDBToLocationFQA.SetFocus
            Exit Function
        
        
        ElseIf Len(Trim(SSOleDBtoUSChartFQA.Text)) = 0 Then
        
            MsgBox "Please fill the ToFQA Us Chart#.", vbInformation, "Ims"
            Call HighlightBackground(SSOleDBtoUSChartFQA)
            SSOleDBtoUSChartFQA.SetFocus
            Exit Function
        
        
        ElseIf Len(Trim(TxtToStocktypeFQA.Text)) = 0 Then
        
            MsgBox "Please fill the ToFQA Stocktype.", vbInformation, "Ims"
            Call HighlightBackground(TxtToStocktypeFQA)
            TxtToStocktypeFQA.SetFocus
            Exit Function
        
        ElseIf Len(Trim(SSOleDBToCamChartFQA.Text)) = 0 Then
        
            MsgBox "Please fill the ToFQA Cam. Chart#.", vbInformation, "Ims"
            Call HighlightBackground(SSOleDBToCamChartFQA)
            SSOleDBToCamChartFQA.SetFocus
            Exit Function
        
        End If
        
    End If
    
    If ShouldEccnControlsBeEnabled = True And chk_usexportLI.value = 1 Then
    
       If Len(Trim(SSoleEccnNo & "")) = 0 And Len(Trim(SSOleSourceofinfo & "")) = 0 Then
        MsgBox "Line item does not have Eccn# and Source of Info."
       ElseIf Len(Trim(SSoleEccnNo & "")) = 0 Then
        MsgBox "Line item does not have Eccn#."
       ElseIf Len(Trim(SSOleSourceofinfo & "")) = 0 Then
        MsgBox "Line item does not have Source of Info."
       End If
       
    End If
    
    CheckLIFields = True
    Exit Function
Handler:
    Err.Clear
End Function

Public Function InitializePOheaderRecordset()
 
Dim FNameSpace As String

'Dim RsBRQ As ADODB.Recordset

Dim DefSite As String
   FNameSpace = deIms.NameSpace

    Set rsDOCTYPE = GetDocumentType(False)
    Call deIms.Shipper(FNameSpace)
    Call deIms.Currency(FNameSpace)
    Call deIms.Priority(FNameSpace)
    Call deIms.TermDelivery(FNameSpace)
    'Call deIms.Supplier(fnamespace)
    Call deIms.ActiveSupplier(FNameSpace)
    Call deIms.TermCondition(FNameSpace)
    'Call deIms.INVENTORYLOCATION(Fnamespace, ponumb)
    Call deIms.Company(FNameSpace)
    Call deIms.GETSYSSITE(FNameSpace, DefSite)
    Call deIms.ActiveOriginator(FNameSpace)
    Call deIms.ActiveTbu(FNameSpace)
    Call deIms.SERVCODECAT(FNameSpace)
    Call deIms.ActiveShipTo(FNameSpace)
    Call deIms.ActiveCompany(FNameSpace)
    Call deIms.CompanyLocations(FNameSpace)
    mIsPoHeaderRsetsInit = True
    
End Function
Public Function ChangeMode(FMode As FormMode) As FormMode
On Error Resume Next
Dim bl As Boolean
Dim msg1 As String  '////
    'LockWindowUpdate (hWnd)
    
    If FMode = mdCreation Then
        lblStatus.ForeColor = vbRed
        
        'Modified by Juan (8/28/2000) for Multilingual
        'msg1 = translator.Trans("L00125") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '---------------------------------------------
        
    ElseIf FMode = mdModification Then
        lblStatus.ForeColor = vbBlue
                
        'Modified by Juan (8/28/2000) for Multilingual
        'msg1 = translator.Trans("L00126") 'J added
       lblStatus.Caption = IIf(msg1 = "", "Modification", msg1) 'J modified
        '---------------------------------------------
  
     ElseIf FMode = mdvisualization Then
        lblStatus.ForeColor = vbGreen
        
        'Modified by Juan (8/28/2000) for Multilingual
        'msg1 = translator.Trans("L00092") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
    
    End If
    
       
    FMode = FMode
    Call MakeReadOnly(FMode = mdvisualization)
    'Call ShowActiveRecords(False)
    
    'GetUnits ("")
   
    'LockWindowUpdate (0)
   ChangeMode = FMode
End Function
Private Sub MakeReadOnly(value As Boolean)
On Error Resume Next

    txtClause.locked = value
    txtRemarks.locked = value
    Text1.locked = value
    value = Not value
    cmd_Add.Enabled = value
    CmdAddSupEmail.Enabled = value
    cmd_Addterms.Enabled = value
    CmdcopyLI.Item(0).Enabled = value
    CmdcopyLI.Item(1).Enabled = value
    CmdcopyLI.Item(2).Enabled = value
    cmdRemove.Enabled = value
    fra_LineItem.Enabled = value
    Fra_ToFqa.Enabled = value
    Frm_FromFQA.Enabled = value
    dgRecepients.Enabled = value
    fra_Purchase.Enabled = value
    fra_FaxSelect.Enabled = value
    'dgRecipientList.Enabled = value 'JCG 2008/8/30
    SSOleDBDocType.Enabled = value
    Text1.Text = ""
    TxtToStocktypeFQA.Enabled = value
    
    If value = True Then
        If ShouldEccnControlsBeEnabled = True Then
        
            chk_USExportH.Enabled = True
            chk_usexportLI.Enabled = True
            Chk_license.Enabled = True
            SSoleEccnNo.Enabled = True
            SSOleSourceofinfo.Enabled = True
        Else
        
            chk_USExportH.Enabled = False
            chk_usexportLI.Enabled = False
            Chk_license.Enabled = False
            SSoleEccnNo.Enabled = False
            SSOleSourceofinfo.Enabled = False
                
        End If
        
    End If
    If Err Then Err.Clear
End Sub
Private Sub comsearch_Completed(Cancelled As Boolean, sStockNumber As String)
On Error Resume Next

    comsearch.Hide
    
    ssdcboCommoditty.Text = sStockNumber
    txt_Descript = comsearch.Description
    
 If objUnits Is Nothing Then Set objUnits = Mainpo.PoUnits
        
         objUnits.StockNumber = Trim$(sStockNumber)
         SSOleDBUnit.RemoveAll
         SSOleDBUnit.AddItem objUnits.PrimaryUnit
         SSOleDBUnit.AddItem objUnits.SecondaryUnit
         txt_Descript = objUnits.Description
'    Call GetUnits(rsPOITEM!poi_comm & "")
    If Err Then Err.Clear
End Sub
Private Sub comsearch_Unloading(Cancel As Integer)
On Error Resume Next
    Set comsearch = Nothing
End Sub
Private Sub WriteStatus(msg As String)
    Call MDI_IMS.WriteStatus(msg, 1)
End Sub
Private Sub BeforePrint()
On Error Resume Next

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath & "po.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("po.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Poheader.Ponumb + ";TRUE"
    End With
    
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::BeforePrint", Err.Description, Err)
    End If
End Sub


Public Sub SendEmailAndFax(Recipients As PoReceipients, FieldName As String, _
                           subject As String, Message As String, Attachment As String, _
                           Optional Orientation As OrientationConstants)
    Dim address() As String
    Dim str As String
    Dim i As Integer

    On Error Resume Next

    address = ToArrayFromRec(Recipients, FieldName, i, str)
    
    Dim faxAddresses() As String: faxAddresses = filterAddresses(address, True)
    If UBound(faxAddresses) > 0 Then
        Call sendFaxOnly(subject, faxAddresses, Attachment)
    End If
    
    Dim emailAddresses() As String: emailAddresses = filterAddresses(address, False)
    If UBound(emailAddresses) > 0 Then
        Call sendEmailOnly(subject, emailAddresses, Attachment)
    End If
    
    Kill Attachment

    If Not IsLoaded("MDI_IMS") Then End
    MDI_IMS.CrystalReport1.Reset
    
    If Err Then Err.Clear
End Sub
Public Function ToArrayFromRec(rs As PoReceipients, ByVal FieldName As String, Optional UpperBound As Integer, Optional ByVal Filter As String) As String()
Dim BK As Variant
Dim str() As String
Dim OldFilter As Variant

On Error GoTo ErrHandler
    ReDim str(0)
    UpperBound = -1
    If rs Is Nothing Then Exit Function
    
    'BK = rs.Bookmark
    
    
'''    If Len(Filter) Then
'''        OldFilter = rs.Filter
'''        rs.Filter = adFilterNone
'''        rs.Filter = Filter
'''    End If
    
    rs.MoveFirst
    Do While Not rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve str(UpperBound)
        str(UpperBound) = rs.Receipient
        rs.MoveNext
    Loop
    
    ToArrayFromRec = str
    
    'If Len(Filter) Then rs.Filter = OldFilter
    'rs.Bookmark = BK
    Exit Function
    
ErrHandler:
    'RaiseEvent Err.Description
    'If (Len(Filter) And Len(OldFilter)) Then rs.Filter = OldFilter
    Err.Raise Err.number, Err.Description
    Err.Clear
End Function

Public Function IsPrimQuantLessThanONE() As Boolean
 Dim PriQuantity As Double
 Dim unit1 As String
 Dim unit2 As String
   'SSOleDBUnit.MoveFirst
  'unit1 = SSOleDBUnit.Columns(0).text
 ' SSOleDBUnit.MoveNext
 ' unit2 = SSOleDBUnit.Columns(0).text
 'IsPrimQuantLessThanONE = False
 
 On Error GoTo Handler
  If Len(txt_Requested) > 0 And Len(SSOleDBUnit) > 0 Then
       If objUnits Is Nothing Then Set objUnits = Mainpo.PoUnits
       objUnits.StockNumber = Trim$(ssdcboCommoditty)
       If objUnits.PrimaryUnit <> objUnits.SecondaryUnit Then
             If objUnits.PrimaryUnit = Trim$(SSOleDBUnit) Then
                 If CDbl(txt_Requested) < 1 And CDbl(txt_Requested) > 0 Then
                    MsgBox "Quantity Can not be Less than 1", , "Imswin"
                    Set objUnits = Nothing
                    Exit Function
                 End If
            ElseIf objUnits.SecondaryUnit = Trim$(SSOleDBUnit) Then
                'Juan 2010-9-7 to add ratio functionality rather than computer factor
                'PriQuantity = CDbl(txt_Requested) * objUnits.ComputationFactor / 10000
                PriQuantity = CDbl(txt_Requested) / objUnits.ratioValue
                If PriQuantity < 1 And PriQuantity > 0 Then
                    MsgBox "The Quantity you entered is equal to " & PriQuantity & " " & objUnits.PrimaryUnit & " .It Can not be Less than 1 " & objUnits.PrimaryUnit, , "Imswin"
                    Set objUnits = Nothing
                    Exit Function
                 End If
            End If
         End If
   End If
   IsPrimQuantLessThanONE = True
   If Not objUnits Is Nothing Then Set objUnits = Nothing
    Exit Function
Handler:
   Err.Clear
End Function

Public Function PrefixFaxorEmail(RecipientName As String) As String
    PrefixFaxorEmail = RecipientName
    If Len(Trim$(RecipientName)) = 0 Then Exit Function
      RecipientName = UCase(RecipientName)
      If InStr(RecipientName, "") = 0 And InStr(RecipientName, "") = 0 Then
            If opt_Email.value = True Then RecipientName = ("" & RecipientName)
            If opt_FaxNum.value = True Then RecipientName = ("" & RecipientName)
            If opt_Email.value = False And opt_FaxNum.value = False Then RecipientName = ("" & RecipientName)
      End If
      
'''      If InStr(RecipientName, "@") > 0 Then
'''            If Not InStr(UCase(RecipientName), UCase("Internet!")) > 0 Then
'''              RecipientName = "Internet!" & RecipientName
'''            End If
'''
'''      Else
'''
'''            If Not InStr(UCase(RecipientName), UCase("Fax!")) > 0 Then
'''                RecipientName = "FAX!" & RecipientName
'''            End If
'''
'''      End If
      
    PrefixFaxorEmail = RecipientName
End Function

Public Sub FirstTimeAssignmentsHeader()

On Error GoTo ErrHand

mLoadMode = LoadingPOheader

If FormMode = mdModification And CInt(LblRevNumb) > 0 Then
         
    If UCase(Trim(Poheader.Docutype)) <> "R" Then
         
         SSOleDBDocType.Enabled = False
         SSoledbSupplier.Enabled = False
         SSOleDBCurrency.Enabled = False
         
        If Not Trim$(UCase(Poheader.StasINvt)) = "NI" Then
          
          SSOleDBcompany.Enabled = False
          SSOleDBInvLocation.Enabled = False
        
        Else
          
          SSOleDBcompany.Enabled = True
          SSOleDBInvLocation.Enabled = True
        
        End If
        
    ElseIf UCase(Trim(Poheader.Docutype)) = "R" Then
        
        SSOleDBDocType.Enabled = True
        SSoledbSupplier.Enabled = False
        SSOleDBCurrency.Enabled = True
        SSOleDBcompany.Enabled = True
        SSOleDBInvLocation.Enabled = True
          
   End If
        
        
 Else
       If FormMode = mdvisualization Then
       
           SSOleDBDocType.Enabled = False
        Else
           SSOleDBDocType.Enabled = True
        End If
        
        SSoledbSupplier.Enabled = True
        SSOleDBCurrency.Enabled = True
        SSOleDBcompany.Enabled = True
        SSOleDBInvLocation.Enabled = True
        
        
End If
    
If FormMode = mdModification Then
    
    chk_FrmStkMst.Enabled = False
    ssOleDbPO.Enabled = False
    TxtToStocktypeFQA.Enabled = True
    If POFqa Is Nothing Then Set POFqa = Mainpo.FQA
    
    If ConnInfo.Eccnactivate = Constyes Then chk_USExportH.value = IIf(ConnInfo.usexport = True, 1, 0)
    If ConnInfo.Eccnactivate = ConstOptional Then chk_USExportH.value = IIf(Poheader.usexport = True, 1, 0)  'IIf(ConnInfo.usexport = True, 1, 0)
    
    If POFqa.Count = 0 Then
    
        Fra_ToFqa.Enabled = False
        Frm_FromFQA.Enabled = False
        
        SSOleDBToCamChartFQA.BackColor = 16777152
        SSOleDBToLocationFQA.BackColor = 16777152
        TxtToStocktypeFQA.BackColor = 16777152
        SSOleDBtoUSChartFQA.BackColor = 16777152
        
    Else
    
        Fra_ToFqa.Enabled = True
        Frm_FromFQA.Enabled = True
        
        SSOleDBToCamChartFQA.BackColor = vbWhite
        SSOleDBToLocationFQA.BackColor = vbWhite
        TxtToStocktypeFQA.BackColor = vbWhite
        SSOleDBtoUSChartFQA.BackColor = vbWhite
        
    End If
    
ElseIf FormMode = mdvisualization Then
    
    chk_FrmStkMst.Enabled = False
    TxtToStocktypeFQA.Enabled = False
    'chk_USExportH.Enabled = False
    
ElseIf FormMode = mdCreation Then
    
    chk_FrmStkMst.Enabled = True
    ssOleDbPO.Enabled = True
    TxtToStocktypeFQA.Enabled = True
    'chk_USExportH.Value =
    
    Fra_ToFqa.Enabled = True
    Frm_FromFQA.Enabled = True
    
    'If ConnInfo.Eccnactivate = Constyes Or ConnInfo.Eccnactivate = ConstOptional Then chk_USExportH.Value = IIf(ConnInfo.usexport = True, 1, 0)
    If ConnInfo.Eccnactivate = Constyes Then chk_USExportH.value = IIf(ConnInfo.usexport = True, 1, 0)
    If ConnInfo.Eccnactivate = ConstOptional Then chk_USExportH.value = IIf(Poheader.usexport = True, 1, 0)  'IIf(ConnInfo.usexport = True, 1, 0)
    
End If

mLoadMode = NoLoadInProgress

Exit Sub
ErrHand:
mLoadMode = NoLoadInProgress

MsgBox Err.Description
Err.Clear

End Sub

Private Sub txt_SerialNum_GotFocus()
Call HighlightBackground(txt_SerialNum)
End Sub

Private Sub txt_SerialNum_LostFocus()
Call NormalBackground(txt_SerialNum)
End Sub

Private Sub Txt_supContaName_GotFocus()
Call HighlightBackground(Txt_supContaName)
End Sub

Private Sub Txt_supContaName_LostFocus()
Call NormalBackground(Txt_supContaName)
End Sub

Private Sub Txt_supContaPh_GotFocus()
Call HighlightBackground(Txt_supContaPh)
End Sub

Private Sub Txt_supContaPh_LostFocus()
Call NormalBackground(Txt_supContaPh)
End Sub

Private Sub txtClause_GotFocus()
Call HighlightBackground(txtClause)
End Sub

Private Sub txtClause_LostFocus()
Call NormalBackground(txtClause)
End Sub

Private Sub txtRemarks_GotFocus()
Call HighlightBackground(txtRemarks)
End Sub

Private Sub txtRemarks_LostFocus()
Call NormalBackground(txtRemarks)
End Sub

Public Sub HandleEdittingOfRemarks()
Dim x As Integer
 
 If lookups Is Nothing Then Set lookups = Mainpo.lookups
 x = lookups.CanUserDeleteRemark(Poheader.Ponumb, IIf(Poheader.revinumb = 0, 0, Poheader.Originalrevinumb - 1), PORemark.Linenumb)
 If Poheader.revinumb > 0 Then
          
        If Poheader.Originalrevinumb = 0 Then
                 
                 If PORemark.EditMode = 2 Then
                       txtRemarks.locked = False
                       'txtRemarks.Enabled = Not txtRemarks.locked
                 ElseIf PORemark.EditMode = 1 Or PORemark.EditMode = 0 Then
                       txtRemarks.locked = True
                       'txtRemarks.Enabled = Not txtRemarks.locked
                 End If
                 
        ElseIf Poheader.Originalrevinumb > 0 Then
        
        
                 
        
                 If x = 1 Then
                       txtRemarks.locked = True
                       'txtRemarks.Enabled = Not txtRemarks.locked
                 ElseIf x = 0 Then
                      If Poheader.Originalrevinumb = Poheader.revinumb Then
                           txtRemarks.locked = False
                           'txtRemarks.Enabled = Not txtRemarks.locked
                      Else
                          If PORemark.EditMode = 2 Then
                                 txtRemarks.locked = False
                                 'txtRemarks.Enabled = Not txtRemarks.locked
                           ElseIf PORemark.EditMode = 1 Or PORemark.EditMode = 0 Then
                                 txtRemarks.locked = True
                                 'txtRemarks.Enabled = Not txtRemarks.locked
                           End If
                      End If
                 End If
        End If

ElseIf Poheader.revinumb = 0 Then

         txtRemarks.locked = False
         'txtRemarks.Enabled = Not txtRemarks.locked
 
End If

If txtRemarks.locked = False Then
   Call HighlightBackground(txtRemarks)
Else
   Call NormalBackground(txtRemarks)
End If
End Sub

Public Sub HandleEdittingOfClause()
Dim x As Integer
 
 If lookups Is Nothing Then Set lookups = Mainpo.lookups
 x = lookups.CanUserDeleteClause(Poheader.Ponumb, IIf(Poheader.revinumb = 0, 0, Poheader.Originalrevinumb - 1), POClause.Linenumb)
 If Poheader.revinumb > 0 Then
          
        If Poheader.Originalrevinumb = 0 Then
                 
                 If POClause.EditMode = 2 Then
                       txtClause.locked = False
                       'txtClause.Enabled = Not txtClause.locked
                 ElseIf POClause.EditMode = 1 Or POClause.EditMode = 0 Then
                       txtClause.locked = True
                       'txtClause.Enabled = Not txtClause.locked
                 End If
                 
        ElseIf Poheader.Originalrevinumb > 0 Then
        
        
                 
        
                 If x = 1 Then
                       txtClause.locked = True
                       'txtClause.Enabled = Not txtClause.locked
                 ElseIf x = 0 Then
                      If Poheader.Originalrevinumb = Poheader.revinumb Then
                           txtClause.locked = False
                           'txtClause.Enabled = Not txtClause.locked
                      Else
                          If POClause.EditMode = 2 Then
                                 txtClause.locked = False
                                 'txtClause.Enabled = Not txtClause.locked
                           ElseIf POClause.EditMode = 1 Or POClause.EditMode = 0 Then
                                 txtClause.locked = True
                                 'txtClause.Enabled = Not txtClause.locked
                           End If
                      End If
                 End If
        End If

ElseIf Poheader.revinumb = 0 Then

         txtClause.locked = False
         'txtClause.Enabled = Not txtClause.locked
 
End If

If txtClause.locked = False Then
   Call HighlightBackground(txtClause)
Else
   Call NormalBackground(txtClause)
End If

End Sub

Public Function DoesDocTypeExist() As Boolean


If Len(Trim(SSOleDBDocType)) = 0 Then

   DoesDocTypeExist = False
   MsgBox "Please select a document type before selecting a supplier.", vbInformation, "Imswin"
   SSOleDBDocType.SetFocus
   Call HighlightBackground(SSOleDBDocType)
   Call NormalBackground(SSoledbSupplier)
Else

   DoesDocTypeExist = True
   
End If

End Function
Public Function sendOutlookEmailandFax()
Dim Params(1) As String
Dim i As Integer
Dim Attachments() As String
Dim subject As String
Dim reports(0) As String
Dim Recepients() As String
Dim attention As String

On Error GoTo errMESSAGE
     
     BeforePrint
     
    If PoReceipients Is Nothing Then Set PoReceipients = Mainpo.PoReceipients
    
    If Trim$(PoReceipients.Ponumb) <> Poheader.Ponumb Then PoReceipients.Move (Poheader.Ponumb)
    
    If PoReceipients.Count > 0 Then
      
        subject = getmessage
        reports(0) = Report_EmailFax_PO_name  ' "po.rpt"   MM : using the new report for email fax
                
        
        'JCG 2008/7/14
        'attention = "Attention Please "
        Dim poNum As String
        poNum = Poheader.Ponumb
        attention = "Please find here attached PO #" + poNum + " From Pecten Cameroon company"
                
        Dim ParamsForCrystalReports(1) As String
        ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";true"
        ParamsForCrystalReports(1) = "ponumb;" + Poheader.Ponumb + ";true"
        '---------
      
        'Send reports to it and creates the attachments and save them to a perticular FOLDER for AT&T
        
        'Attachments = generateattachments(reports) 'JCG 2008/7/31
        'Attachments = generateattachmentsPDF("po.rpt", "Purchase Order", ParamsForCrystalReports, MDI_IMS.CrystalReport1, poNum) 'JCG 2008/7/31
        Attachments = generateattachmentswithCR11(Report_EmailFax_PO_name, "Purchase Order", ParamsForCrystalReports, MDI_IMS.CrystalReport1)   'MM 030209 EFCR11
     
        Recepients = ToArrayFromRecO(PoReceipients)
        'Here we create the parameter FILE.
        'Send the attachments ,the subject and the recepients to be written in the Parameter file.
     
            Call WriteParameterFiles(Recepients, "", Attachments, subject, attention)
    Else
    
         MsgBox "No Recipients to Send", , "Imswin"
     
    End If
    
    
    
errMESSAGE:
    
    If Err.number <> 0 Then
        
        MsgBox Err.Description
    
    End If

End Function

'''Public Function CreateOutlookAttachment(Attachmentfilename As String)
'''
'''
'''
'''    Dim ifile As IMSFile
'''    Dim attachments(0) As String, str As String
'''
'''    Attachmentfilename = UserID & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
'''
'''    On Error Resume Next
'''
'''    Set ifile = New IMSFile
'''
'''    attachments(0) = "F:\OUTLOOK\OUT\" & Attachmentfilename
'''
'''    If Not FileExists(attachments(0)) Then MDI_IMS.SaveReport attachments(0), crptRTF
'''
'''    If Not ifile.FileExists(attachments(0)) Then
'''        MsgBox "Error Preparing Electronic Message"
'''        Exit Function
'''    End If
'''
'''    If ifile.FileExists(attachments(0)) Then ifile.DeleteFile (attachments(0))
'''
'''    Set ifile = Nothing
'''
'''End Function

Public Function WriteParameterFileFax(Attachments, Recipients, subject, sender, attention)
    On Error GoTo errMESSAGE
    
     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String
     Dim sql, companyNAME
     Dim datax As New ADODB.Recordset

     Filename = "Fax" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     'FileName = "c:\IMSRequests\IMSRequests\" & FileName
     
     Filename = ConnInfo.EmailParameterFolder & Filename

    For i = 0 To UBound(Recipients)
            recepientSTR = recepientSTR & Trim$(Recipients(i) & ";")
    Next

      i = 0

    For i = 0 To UBound(Attachments)
            reports = reports & Trim$(Attachments(i) & ";")
    Next

    sql = "SELECT po_ponumb, com_name FROM PO LEFT OUTER JOIN COMPANY ON " _
        & "po_compcode = com_compcode AND po_npecode = com_npecode WHERE " _
        & "po_ponumb = '" + ssOleDbPO.Text + "' AND po_npecode = '" + deIms.NameSpace + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, deIms.cnIms
    
    If datax.RecordCount > 0 Then
        companyNAME = datax!com_name
    Else
        companyNAME = ""
    End If
    

    Open Filename For Output As FileNumb

        Print #FileNumb, "[WINFAX]"
        Print #FileNumb, "Recipients=" & recepientSTR
        Print #FileNumb, "Reports=" & reports
        Print #FileNumb, "Subject=" & subject
        Print #FileNumb, "Sender=" & Trim(companyNAME)
        Print #FileNumb, "Attention=" & Trim$(attention)

    Close #FileNumb

errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If
End Function


Public Function WriteParameterFileEmail(Attachments() As String, Recipients() As String, subject As String, sender As String, attention As String) As String
On Error GoTo errMESSAGE
     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String

     Filename = "Email" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     'FileName = "c:\IMSRequests\IMSRequests\" & FileName
     
     Filename = ConnInfo.EmailParameterFolder & Filename

    For i = 0 To UBound(Recipients)
            recepientSTR = recepientSTR & Trim$(Recipients(i) & ";")
    Next

      i = 0

    For i = 0 To UBound(Attachments)
            reports = reports & Trim$(Attachments(i) & ";")
    Next

    
    
    Open Filename For Output As FileNumb

        Print #FileNumb, "[Email]"
        Print #FileNumb, "Recipients=" & recepientSTR
        Print #FileNumb, "Reports=" & reports
        Print #FileNumb, "Subject=" & subject
        Print #FileNumb, "Sender=" & Trim$("")
        Print #FileNumb, "Attention=" & Trim$(attention)

    Close #FileNumb

errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If
End Function
   
Function getmessage() As String

Dim messageSubject As String
     
     msg1 = translator.Trans("L00100")
     msg2 = msg1
     
     messageSubject = IIf(msg1 = "", "Purchase Order ", msg1 + " ") & Poheader.Ponumb
    
    
     If Len(LblRevNumb.Caption) > 0 And Not (LblRevNumb.Caption = "0") Then
        
        msg1 = translator.Trans("L00066")
        messageSubject = messageSubject & IIf(msg1 = "", "(revision No. ", msg1 + " ") & LblRevNumb.Caption & ")"
        
     Else
     
        msg1 = translator.Trans("M00090")
        messageSubject = messageSubject & IIf(msg1 = "", "(initial revision)", msg1)
        
     End If
      
          
     getmessage = messageSubject
End Function

Public Function generateattachments(reports() As String) As String()
  Dim l
  Dim Attachments(0) As String
  Dim IFile As IMSFile
  Dim Filename As String
  
  Set IFile = New IMSFile
  'l = UBound(reports)
On Error GoTo errMESSAGE
  
'  For i = 0 To l
      
      
    With MDI_IMS.CrystalReport1
        
        .ReportFileName = reportPath & "po.rpt"
        
        'Call translator.Translate_Reports("po.rpt")
        Call translator.Translate_Reports(reports(l))
        Call translator.Translate_SubReports
        
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Poheader.Ponumb + ";TRUE"
        
    End With
    
     Attachments(0) = "Report-" & "PO" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
     
    ' FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
    
     Filename = ConnInfo.EmailOutFolder & Attachments(0)
    
    If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
        
    If Not FileExists(Filename) Then MDI_IMS.SaveReport Filename, crptRTF
       
     generateattachments = Attachments
    
errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If

End Function

''''''Public Function WriteParameterFiles(Recepients() As String, sender As String, Attachments() As String, Subject As String, attention As String)
''''''
'''''' Dim l
'''''' Dim x
'''''' Dim y
'''''' Dim i
'''''' Dim email() As String
'''''' Dim fax() As String
''''''
''''''
''''''
'''''''Splitting the address into Emails and Faxes.
'''''' l = UBound(Recepients)
''''''
''''''    x = 0
''''''    y = 0
''''''
'''''' For i = 0 To l
''''''
''''''     If InStr(Recepients(i), "@") > 0 Then
''''''
''''''       ReDim Preserve email(x)
''''''       email(x) = Recepients(i)
''''''       x = x + 1
''''''
''''''    Else
''''''
''''''       ReDim Preserve fax(y)
''''''       fax(y) = Recepients(i)
''''''       y = y + 1
''''''
''''''    End If
''''''
''''''
''''''
'''''' Next i
''''''
''''''
''''''
'''''' Call WriteParameterFileEmail(Attachments, email, Subject, sender, attention)
''''''
'''''' Call WriteParameterFileFax(Attachments, fax, Subject, sender, attention)
''''''
''''''
''''''
''''''End Function


'The PO Creation Process generates an auto number internally  which it makes use when the user navigates to other PO tabs.
'In case before the user saves the PO , he wants to save the PO wiht his own PO number, then this _
functions set all the PO bojects wiht that perticular PO number.

'''Public Function SetPONUMBforAllPoObjects(PoNumber As String) As Boolean
'''
'''On Error GoTo ErrHandler
'''
'''SetPONUMBforAllPoObjects = False
'''
'''If Not Poheader Is Nothing Then Poheader.Ponumb = PoNumber
'''
'''If Not PoItem Is Nothing Then
'''
'''   Call PoItem.Replace("PONUMB", PoNumber, "ALL")
'''
'''End If
'''
'''If Not PoReceipients Is Nothing Then
'''
'''    Call PoReceipients.Replace("PONUMB", PoNumber, "ALL")
'''
'''End If
'''
'''If Not PORemark Is Nothing Then
'''
'''    Call PORemark.Replace("PONUMB", PoNumber, "ALL")
'''
'''End If
'''
'''If Not POClause Is Nothing Then
'''
'''    Call POClause.Replace("PONUMB", PoNumber, "ALL")
'''
'''
'''End If
'''
'''SetPONUMBforAllPoObjects = True
'''
'''Exit Function
'''
'''ErrHandler:
'''
''''n
'''
'''End Function


Public Sub SendAttEmailandFax()
On Error Resume Next

Dim i As RPTIFileInfo
Dim Params(1) As String
   
   
    With i
    
        
        '.Login = "sa"  'M
        .Login = ConnInfo.UId 'UserId 'M
        .Password = ConnInfo.Pwd ' DBPassword  M
        .ReportFileName = reportPath & "po.rpt"
                       
        Params(0) = "namespace=" & deIms.NameSpace
        Params(1) = "ponumb=" & Poheader.Ponumb & ""
        
        .parameters = Params
        
        
        Params(0) = ""
        Call WriteRPTIFile(i, Params(0))
'        .ReportFileName = ReportPath & "po2.rpt"
'        .ParameterFields(0) = "namespace=" + deIms.Namespace + ";TRUE"
'        .ParameterFields(1) = "ponumb;" + rsPO!po_ponumb + ";TRUE"

    BeforePrint
    
    'Modified by Juan (9/13/200) for Multilingual
    msg1 = translator.Trans("L00100") 'J added
    msg2 = msg1 'J added
    Dim messageSubject As String: messageSubject = IIf(msg1 = "", "Purchase Order ", msg1 + " ") & Poheader.Ponumb 'J modified
    If Len(LblRevNumb.Caption) > 0 And Not (LblRevNumb.Caption = "0") Then
        msg1 = translator.Trans("L00066") 'J added
        messageSubject = messageSubject & IIf(msg1 = "", "(revision No. ", msg1 + " ") & LblRevNumb.Caption & ")" 'J modified
    Else
        msg1 = translator.Trans("M00090") 'J added
        messageSubject = messageSubject & IIf(msg1 = "", "(initial revision)", msg1) 'J modified
    End If
    If PoReceipients Is Nothing Then Set PoReceipients = Mainpo.PoReceipients
    If Trim$(PoReceipients.Ponumb) <> Poheader.Ponumb Then PoReceipients.Move (Poheader.Ponumb)
        If PoReceipients.Count > 0 Then
          Call SendEmailAndFax(PoReceipients, "Receipient ", messageSubject, IIf(msg2 = "", "Purchase Order", msg2), "")  'J modified
       Else
         MsgBox "No Recipients to Send", , "Imswin"
        End If
    '-----------------------------------------

    End With
    
    

    If Err Then Err.Clear

End Sub

Private Sub SelectGatewayAndSendOutMails()

If ConnInfo.EmailClient = Outlook Then
    
    Call sendOutlookEmailandFax

ElseIf ConnInfo.EmailClient = ATT Then
    
    Call SendAttEmailandFax

ElseIf ConnInfo.EmailClient = Outlook Then
    
    MsgBox "Email is not set up properly. Please Configure the database for Emails.", vbInformation, "Imswin"

End If

End Sub

Public Function CanDocTypeBeRevised(DocCode As String) As Boolean

On Error GoTo ErrHandler

Dim rsDOCTYPE As ADODB.Recordset
Set rsDOCTYPE = New ADODB.Recordset
CanDocTypeBeRevised = False

rsDOCTYPE.Source = "select doc_reviflag from doctype where doc_code='" & Trim(DocCode) & "' and doc_npecode ='" & deIms.NameSpace & "'"

rsDOCTYPE.ActiveConnection = deIms.cnIms

rsDOCTYPE.Open

CanDocTypeBeRevised = rsDOCTYPE("doc_reviflag")

Exit Function

ErrHandler:

MsgBox "Errors occurred while trying to verify if the Document type could be revised.", vbCritical, "Imswin"

Err.Clear

End Function

Public Function GetServiceCode() As ADODB.Recordset

Dim rsSERVCODE As New ADODB.Recordset

On Error GoTo ErrHandler

rsSERVCODE.Source = "select srvc_code,srvc_desc from servcode where srvc_actvflag=1 AND srvc_npecode='" & deIms.NameSpace & "'"

rsSERVCODE.ActiveConnection = deIms.cnIms

rsSERVCODE.Open , , adOpenKeyset

Set GetServiceCode = rsSERVCODE

Exit Function

ErrHandler:

    MsgBox "Errors Occurred while trying to access the Service Codes.", vbInformation, "Imswin"
    
    Err.Clear

End Function

Public Function DeleteDefaultRecepientsForDoctype(DocType As String) As Integer

Dim i As Integer
Dim j As Integer
Dim Recepients() As String

On Error GoTo ErrHandler

DeleteDefaultRecepientsForDoctype = 1


If dgRecipientList.Rows > 0 And Len(Trim(DocType)) > 0 Then

       
    If lookups Is Nothing Then Set lookups = Mainpo.lookups
    
    If lookups.GetDefaultRecForDoctype(DocType, Recepients) = 1 Then
    
       MsgBox "Errors Occurred while Trying to Access the Distribution List for the Document type. Please Try again.", vbCritical, "Imswin"
       
       Exit Function
       
    End If

   If IsArrayLoaded(Recepients) Then

            If PoReceipients Is Nothing Then
                    Set PoReceipients = Mainpo.PoReceipients
                    PoReceipients.Move Poheader.Ponumb
                    LoadFromPOReceipients
            End If
        
            'dgRecipientList.MoveFirst
            
            For i = 0 To UBound(Recepients)
            
                dgRecipientList.MoveFirst
            
                For j = dgRecipientList.AddItemRowIndex(dgRecipientList.Bookmark) To dgRecipientList.Rows - 1
                
                    If Trim(Recepients(i)) = Trim(dgRecipientList.Columns(0).Text) Then
                    
                         PoReceipients.DeleteCurrentLI (dgRecipientList.Columns(0).Text)
                         
                         dgRecipientList.RemoveItem (dgRecipientList.AddItemRowIndex(dgRecipientList.Bookmark))
                        
                         Exit For
                        
                    Else
                         
                        dgRecipientList.MoveNext
                         
                    End If
                    
                Next j
            
            Next i
            
   End If

End If

DeleteDefaultRecepientsForDoctype = 0

Exit Function

ErrHandler:

    MsgBox "Errors Occurred While trying to delete the Default Recepients of the Document Type from the Recepients list. Error description :" & Err.Description, vbCritical, "Imswin"

    Err.Clear
    
End Function

Public Function ConvertRequisition(RequisitionNumber As String) As Boolean
Dim Error As Boolean
Dim TemporaryPo As String
Dim ErrorsReturned As String
Dim no As Integer
On Error GoTo ErrHandler
ConvertRequisition = False

Set PoItem = Mainpo.POITEMS
'Juan 2010-9-25 necessary to reset lineitem
Call PoItem.Move(RequisitionNumber)
If PoItem.Count > 0 Then
Else
    MsgBox "Errors Occurred while trying to Convert a Requsition. Error Description : " & Err.Description + ErrorsReturned
    IncrementProgressBar 10
    IncrementProgressBar 0
    Set PoItem = Nothing
    Set POClause = Nothing
    Set PORemark = Nothing
    Set PoReceipients = Nothing
    Err.Clear
    Exit Function
End If
With PoItem
    .PriQtydlvd = 0
    .PriQtyinvt = 0
    .PriQtyship = 0
    .PriQtytobedlvd = .Primreqdqty
    .Stasdlvy = "NR"
    .StasINvt = "NI"
    .Stasliit = "OP"
    .Stasship = "NS"
End With
'-------------------
Set POClause = Mainpo.POClauses
Set PoReceipients = Mainpo.PoReceipients

Set PORemark = Mainpo.POREMARKS
If POFqa Is Nothing Then Set POFqa = Mainpo.FQA
TemporaryPo = Poheader.Ponumb

frm_ConvertToPO.ProgressBar1.Max = 10

IncrementProgressBar 1
Error = Poheader.LoadFromRequsition(RequisitionNumber, TemporaryPo, CurrentUser, ErrorsReturned)
IncrementProgressBar 1
If Error = False Then Error = PoItem.LoadFromRequsition(RequisitionNumber, TemporaryPo, Poheader.fromstckmast, ErrorsReturned)
IncrementProgressBar 1
If Error = False Then Error = POClause.LoadFromRequsition(RequisitionNumber, TemporaryPo, ErrorsReturned)
IncrementProgressBar 1
If Error = False Then Error = PoReceipients.LoadFromRequsition(RequisitionNumber, TemporaryPo, ErrorsReturned)
IncrementProgressBar 1
If Error = False Then Error = PORemark.LoadFromRequsition(RequisitionNumber, TemporaryPo, ErrorsReturned)
IncrementProgressBar 1

If Error = False Then Error = CreateFqaForRequisition
    

If Error = True Then GoTo ErrHandler

LoadFromPOHEADER
IncrementProgressBar 1
If PoItem.Count > 0 Then PoItem.MoveFirst:   LoadFromPOITEM
IncrementProgressBar 1
If POClause.Count > 0 Then POClause.MoveFirst:  LoadFromPOClause
IncrementProgressBar 1
If PoReceipients.Count > 0 Then PoReceipients.MoveFirst:  LoadFromPOReceipients
IncrementProgressBar 1
If PORemark.Count > 0 Then PORemark.MoveFirst:  LoadFromPORemarks
IncrementProgressBar 1

ssOleDbPO = ""
IncrementProgressBar 10
IncrementProgressBar 0

Unload frm_ConvertToPO

Exit Function
ErrHandler:

MsgBox "Errors Occurred while trying to Convert a Requsition. Error Description : " & Err.Description + ErrorsReturned
IncrementProgressBar 10
IncrementProgressBar 0

Set PoItem = Nothing
Set POClause = Nothing
Set PORemark = Nothing
Set PoReceipients = Nothing

Err.Clear
End Function

Public Sub IncrementProgressBar(value As Integer)

If value > 0 And value < 10 Then

If frm_ConvertToPO.ProgressBar1.value = frm_ConvertToPO.ProgressBar1.Max Then

Else

frm_ConvertToPO.ProgressBar1.value = frm_ConvertToPO.ProgressBar1.value + value

End If

ElseIf value = 0 Then

frm_ConvertToPO.ProgressBar1.value = 0

ElseIf value = 10 Then

frm_ConvertToPO.ProgressBar1.value = frm_ConvertToPO.ProgressBar1.Max

End If

End Sub
Public Function PopulateCombosWithFQA(CompanyCode As String, Optional LocationCode As String) As Boolean

Dim rsCOMPANY As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUc As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset
Dim nAMEsP As String
On Error GoTo ErrHand

PopulateCombosWithFQA = False

'Get Company FQA

nAMEsP = deIms.NameSpace

rsCOMPANY.Source = "select FQA from FQA where Namespace ='" & nAMEsP & "' and Companycode ='" & Trim(CompanyCode) & "' and Level ='C'"

rsCOMPANY.Open , deIms.cnIms

''Do While Not rsCOMPANY.EOF
''
''    SSOleCompany.AddItem rsCOMPANY("FQA")
''    rsCOMPANY.MoveNext
''
''Loop

If rsCOMPANY.EOF = False Then TxtToCompanyFQA = rsCOMPANY("FQA")

'Get Location FQA

RsLocation.Source = "select FQA from FQA where Namespace ='" & nAMEsP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='LB' OR LEVEL ='LS'"

RsLocation.Open , deIms.cnIms

Do While Not RsLocation.EOF

    SSOleDBToLocationFQA.AddItem RsLocation("FQA")
    RsLocation.MoveNext
    
Loop

'Get US Chart FQA

RsUc.Source = "select FQA from FQA where Namespace ='" & nAMEsP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='UC'"

RsUc.Open , deIms.cnIms


Do While Not RsUc.EOF

    SSOleDBtoUSChartFQA.AddItem RsUc("FQA")
    RsUc.MoveNext
    
Loop

'Get Cam Chart FQA

RsCC.Source = "select FQA from FQA where Namespace ='" & nAMEsP & "' and Companycode ='" & Trim(CompanyCode) & "' and Locationcode='" & Trim(LocationCode) & "' and Level ='CC'"

RsCC.Open , deIms.cnIms


Do While Not RsCC.EOF

    SSOleDBToCamChartFQA.AddItem RsCC("FQA")
    RsCC.MoveNext
    
Loop

Set rsCOMPANY = Nothing
Set RsLocation = Nothing
Set RsUc = Nothing
Set RsCC = Nothing

PopulateCombosWithFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes." & Err.Description, vbCritical, "Ims"

Err.Clear

End Function




Public Function LoadFromFROMFQA()

CleanFROMFQA

Call POFqa.MoveLineTo(0)

TxtFromCamChart = POFqa.FromCamChart
TxtFromCompany = POFqa.FromCompany
TxtFromLocation = POFqa.FromLocation
TxtFromType = POFqa.FromStockType
TxtFromUsChart = POFqa.FromUSChart


End Function


Public Function LoadFromTOFQA()

CleanToFQAControls

SSOleDBToCamChartFQA = POFqa.ToCamChart
TxtToCompanyFQA = POFqa.ToCompany
SSOleDBToLocationFQA = POFqa.Tolocation
TxtToStocktypeFQA = POFqa.ToStockType
SSOleDBtoUSChartFQA = POFqa.ToUSChart

End Function

Public Function SavetoFROMFQA()

If POFqa.Count = 0 Then Exit Function

POFqa.Ponumb = Poheader.Ponumb
POFqa.FromCamChart = TxtFromCamChart
POFqa.FromCompany = TxtFromCompany
POFqa.FromLocation = TxtFromLocation
POFqa.FromStockType = TxtFromType
POFqa.FromUSChart = TxtFromUsChart

End Function


Public Function SaveToTOFQA()



On Error GoTo ErrHandler

'This may be the case of OLD Pos which does not have an associated FQA

If POFqa.Count = 0 Then Exit Function

POFqa.ToCamChart = SSOleDBToCamChartFQA
POFqa.ToCompany = TxtToCompanyFQA
POFqa.Tolocation = SSOleDBToLocationFQA
POFqa.ToStockType = TxtToStocktypeFQA
POFqa.ToUSChart = SSOleDBtoUSChartFQA
'pofqa.User = cre

Exit Function
ErrHandler:

MsgBox "Errors Occured while saving the FQA.", vbCritical, Err.Description

Err.Clear
End Function

Public Function CleanFROMFQA()

TxtFromCamChart = ""
TxtFromCompany = ""
TxtFromLocation = ""
TxtFromType = ""
TxtFromUsChart = ""

End Function

Public Function CleanToFQAControls()

SSOleDBToCamChartFQA = ""
TxtToCompanyFQA = ""
SSOleDBToLocationFQA = ""
TxtToStocktypeFQA = ""
SSOleDBtoUSChartFQA = ""

End Function

Public Function InitializeNewTOFQARecord(CompanyCode As String, LocationCode As String) As Boolean
Dim Company As String
Dim CamChart As String
Dim stocktype As String
Dim Location As String
Dim UsChart As String
'Dim X As Integer
    If lookups Is Nothing Then Set lookups = Mainpo.lookups
    
    Call lookups.LoadFQAFromLocation(Trim(CompanyCode), Trim(LocationCode), Company, Location, UsChart, CamChart, stocktype)
    
    If POFqa.Count = 0 Then Call POFqa.AddNew
                
    'If POFqa.count > 1 Then X = POFqa.count - 1
                
    POFqa.ItemNo = POFqa.Count - 1
    POFqa.Ponumb = Poheader.Ponumb
    
    POFqa.FromCamChart = TxtFromCamChart
    POFqa.FromCompany = TxtFromCompany
    POFqa.FromUSChart = TxtFromUsChart
    POFqa.FromLocation = TxtFromLocation
    POFqa.FromStockType = TxtFromType
    
    POFqa.ToCamChart = CamChart
    POFqa.ToCompany = Company
    POFqa.creauser = "DBo"
    POFqa.ToUSChart = UsChart
    POFqa.Tolocation = Location
    POFqa.ToStockType = stocktype
    
    Call LoadFromTOFQA
    
End Function

'this Function is not being used ,since To Location has all open fields now.
Public Function ModifyToFQAWithNewLocation() As Boolean

Dim Company As String
Dim CamChart As String
Dim stocktype As String
Dim Location As String
Dim UsChart As String
Dim LineNo As Integer
Dim x

    If POFqa Is Nothing Then Set POFqa = Mainpo.FQA

''    If Trim(SSOleDBInvLocation.Tag) = "DCH" Then
''
''    SSOleDBToCamChartFQA.Enabled = False
''    SSOleDBToLocationFQA.Enabled = False
''    SSOleDBtoUSChartFQA.Enabled = False
''
''    Else
''
''    SSOleDBToCamChartFQA.Enabled = False
''    SSOleDBToLocationFQA.Enabled = False
''    SSOleDBtoUSChartFQA.Enabled = False
''
''    End If

    If POFqa.Count = 0 Then Exit Function

    If lookups Is Nothing Then Set lookups = Mainpo.lookups
    Call lookups.LoadFQAFromLocation(Trim(SSOleDBcompany.Tag), Trim(SSOleDBInvLocation.Tag), Company, Location, UsChart, CamChart, stocktype)
    
    POFqa.MoveFirst
    
    Do While Not POFqa.EOF
    
        POFqa.Ponumb = Poheader.Ponumb
        POFqa.ToCamChart = CamChart
        POFqa.ToCompany = Company
        POFqa.creauser = "DBo"
        POFqa.ToUSChart = UsChart
        POFqa.Tolocation = Location
        POFqa.ToStockType = stocktype
        If POFqa.MoveNext = False Then Exit Do
        
    Loop
    
    If PoItem Is Nothing Then
    
    Else
        POFqa.MoveLineTo (PoItem.Linenumb)
        LoadFromTOFQA
        
   End If
End Function

Public Function FillFromFQAControls(CompanyCode As String, LocationCode As String) As Boolean

Dim Company As String
Dim CamChart As String
Dim stocktype As String
Dim Location As String
Dim UsChart As String

    If lookups Is Nothing Then Set lookups = Mainpo.lookups
    Call lookups.LoadFQAFromLocation(Trim(CompanyCode), Trim(LocationCode), Company, Location, UsChart, CamChart, stocktype)
    
    TxtFromCamChart = CamChart
    TxtFromCompany = Company
    TxtFromLocation = Location
    TxtFromType = "0000"
    TxtFromUsChart = UsChart
    
End Function

Public Function CreateFqaForRequisition() As Boolean
Dim no As Integer
Dim i As Integer
CreateFqaForRequisition = True
On Error GoTo ErrHand

    
    no = PoItem.Count
    
    If no > 0 Then
    
        For i = 1 To no
    
           If POFqa.AddNew = True Then
            
                Call InitializeNewTOFQARecord(Trim(Poheader.CompCode), Trim(Poheader.invloca))
            
            Else: GoTo ErrHand
           
           End If
            
        Next i

    End If
CreateFqaForRequisition = False
Exit Function

ErrHand:

MsgBox "Errors occurred while trying to generate FQAs for the line items. Please close the form and start again." & Err.Description, vbCritical, "Ims"

Err.Clear

End Function

Public Function LoadToFQACombos() As Boolean

Dim RsUc As ADODB.Recordset
Dim RsCC As ADODB.Recordset
Dim RsLocations As ADODB.Recordset

LoadToFQACombos = False

On Error GoTo Err

If FormMode <> mdvisualization Then
    
    If GToFQAComboLoaded = False Then
            
            If lookups Is Nothing Then Set lookups = Mainpo.lookups
            
            SSOleDBtoUSChartFQA.RemoveAll
            
            Set RsUc = lookups.GetAllUSAccounts()
            
            Do While Not RsUc.EOF
            
                SSOleDBtoUSChartFQA.AddItem Trim(RsUc("fqa"))
                RsUc.MoveNext
            
            Loop
            
            Set RsCC = lookups.GetAllCCAccounts()
            
            SSOleDBToCamChartFQA.RemoveAll
            
            Do While Not RsCC.EOF
            
                SSOleDBToCamChartFQA.AddItem Trim(RsCC("fqa"))
                RsCC.MoveNext
            
            Loop
            
            Set RsLocations = lookups.GetAllLocations()
            
            SSOleDBToLocationFQA.RemoveAll
            
            Do While Not RsLocations.EOF
            
                SSOleDBToLocationFQA.AddItem Trim(RsLocations("fqa"))
                RsLocations.MoveNext
            
            Loop
            
            
            GToFQAComboLoaded = True
            
            
    End If
    
End If

LoadToFQACombos = True

Exit Function

Err:

MsgBox "Errors Occurred while trying to populate the combos. Close the form and try it again." & Err.Description, vbCritical, "Ims"

Err.Clear

End Function



Public Function FillEccnCombos(lookups As IMSPODLL.lookups)
Dim RsEccnNo As New ADODB.Recordset
On Error GoTo ErrHand
        
        If lookups Is Nothing Then Set lookups = Mainpo.lookups
        
         Set RsEccnNo = lookups.GetListofEccns(1)
               
         'if SSoleEccnno SSoleEccnno.Enabled = True
              
              Set SSoleEccnNo.DataSourceList = RsEccnNo  ' RSStockNos
              SSoleEccnNo.DataFieldToDisplay = "eccn_no"
              SSoleEccnNo.DataFieldList = "eccnid"
              SSoleEccnNo.Columns(0).Visible = False

               SSoleEccnNo.Columns(2).Width = 6000
               SSoleEccnNo.RowHeight = 500

       Set GRsEccnNo = lookups.GetListofEccns(0)
        
Exit Function
ErrHand:
MsgBox Err.Description
End Function
Public Function FillSourceOfinfoCombos(lookups As IMSPODLL.lookups)
Dim RsSourceOfInfo As New ADODB.Recordset
On Error GoTo ErrHand
        
        If lookups Is Nothing Then Set lookups = Mainpo.lookups
        
         Set RsSourceOfInfo = lookups.GetListofEccnSource(1)
               
         'if SSoleEccnno SSoleEccnno.Enabled = True
              
              Set SSOleSourceofinfo.DataSourceList = RsSourceOfInfo  ' RSStockNos
              SSOleSourceofinfo.DataFieldToDisplay = "source"
              SSOleSourceofinfo.DataFieldList = "sourceid"
              SSOleSourceofinfo.Columns(0).Visible = False

               'SSoleSourceofInfo.Columns(1).Width = 6000
               

       Set GRSSourceOfInfo = lookups.GetListofEccnSource(0)
        
Exit Function
ErrHand:
MsgBox Err.Description
End Function
Public Function GetEccnForSelectedStock(StockNumber As String, ByRef Eccnid As Integer, ByRef Eccnno As String, ByRef EccnLicense As Boolean, ByRef Sourceid As Integer, ByRef Sourceno As String) As Boolean

On Error GoTo ErrHand

Dim Rseccn As New ADODB.Recordset

    Rseccn.Source = "select isnull(stk_eccnid,0) stk_eccnid,isnull(stk_eccnsourceid,0) stk_eccnsourceid, isnull(stk_eccnlicsreq,0) stk_eccnlicsreq , isnull(eccn_no,'') eccn_no, isnull(source,'') source  from stockmaster s "
    Rseccn.Source = Rseccn.Source & " inner join eccn e on e.eccnid =s.stk_eccnid"
    Rseccn.Source = Rseccn.Source & " left outer join picklist p on p.sourceid =s.stk_eccnsourceid"
    Rseccn.Source = Rseccn.Source & " where s.stk_npecode ='" & deIms.NameSpace & "' and s.stk_stcknumb='" & StockNumber & "'"
    Rseccn.ActiveConnection = deIms.cnIms
    Rseccn.Open , , 3, 3
    
If Rseccn.RecordCount > 0 Then

    Eccnid = Rseccn!stk_eccnid
    Eccnno = Rseccn!eccn_no
    EccnLicense = Rseccn!stk_eccnlicsreq
    Sourceid = Rseccn!stk_eccnsourceid
    Sourceno = Rseccn!Source
    
End If
    
GetEccnForSelectedStock = True

Exit Function

ErrHand:
    
    MsgBox Err.Description

End Function

Public Function ShouldEccnControlsBeEnabled() As Boolean
On Error GoTo ErrHand

    If ConnInfo.Eccnactivate = Constno Then
        ShouldEccnControlsBeEnabled = False
    ElseIf ConnInfo.Eccnactivate = Constyes Or ConnInfo.Eccnactivate = ConstOptional Then
        ShouldEccnControlsBeEnabled = True
    End If

Exit Function
ErrHand:

MsgBox Err.Description

End Function
