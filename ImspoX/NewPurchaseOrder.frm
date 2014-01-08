VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.0#0"; "LRNAVI~1.OCX"
Begin VB.Form frm_NewPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Order"
   ClientHeight    =   7065
   ClientLeft      =   2520
   ClientTop       =   2190
   ClientWidth     =   9060
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9060
   Tag             =   "02020100"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   1200
      TabIndex        =   52
      Top             =   6480
      Width           =   4455
      _ExtentX        =   6800
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailVisible    =   -1  'True
      FirstEnabled    =   0   'False
      LastEnabled     =   0   'False
      NewEnabled      =   -1  'True
      NextEnabled     =   0   'False
      PreviousEnabled =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DeleteToolTipText=   ""
   End
   Begin TabDlg.SSTab sst_PO 
      Height          =   6285
      Left            =   120
      TabIndex        =   130
      Top             =   120
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   11086
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   758
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Transaction Order"
      TabPicture(0)   =   "NewPurchaseOrder.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_PO"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Purchase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "NewPurchaseOrder.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1"
      Tab(1).Control(1)=   "lbl_New"
      Tab(1).Control(2)=   "lbl_Recipients"
      Tab(1).Control(3)=   "dgRecipientList"
      Tab(1).Control(4)=   "fra_FaxSelect"
      Tab(1).Control(5)=   "cmd_Add"
      Tab(1).Control(6)=   "txt_Recipient"
      Tab(1).Control(7)=   "dgRecepients"
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Line Items"
      TabPicture(2)   =   "NewPurchaseOrder.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_LineItem"
      Tab(2).Control(1)=   "fra_LI"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Remarks"
      TabPicture(3)   =   "NewPurchaseOrder.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtRemarks"
      Tab(3).Control(1)=   "CmdcopyLI(1)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Notes/Instructions"
      TabPicture(4)   =   "NewPurchaseOrder.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Addterms"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtClause"
      Tab(4).Control(2)=   "CmdcopyLI(2)"
      Tab(4).ControlCount=   3
      Begin VB.Frame fra_Purchase 
         ClipControls    =   0   'False
         Height          =   5220
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   8295
         Begin VB.CheckBox chk_FreightFard 
            Caption         =   "FFR Mandatory"
            Height          =   195
            Left            =   4635
            TabIndex        =   22
            Top             =   160
            Width           =   1575
         End
         Begin VB.TextBox Txt_supContaName 
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   9
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Txt_supContaPh 
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   10
            Top             =   3450
            Width           =   2295
         End
         Begin VB.CheckBox chk_ConfirmingOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirming Order"
            DataField       =   "po_confordr"
            DataMember      =   "PO"
            Height          =   288
            Left            =   6360
            TabIndex        =   21
            Top             =   827
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_poDate 
            Bindings        =   "NewPurchaseOrder.frx":008C
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   6480
            TabIndex        =   7
            Top             =   2760
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   76283905
            CurrentDate     =   36850
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboDelivery 
            Bindings        =   "NewPurchaseOrder.frx":0097
            Height          =   315
            Left            =   6480
            TabIndex        =   17
            Top             =   4710
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":00C3
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":00DF
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
            Bindings        =   "NewPurchaseOrder.frx":00FB
            Height          =   315
            Left            =   1920
            TabIndex        =   8
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
            Columns.Count   =   4
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
            TabIndex        =   127
            Top             =   4130
            Width           =   1665
         End
         Begin VB.CheckBox chk_Forwarder 
            Caption         =   "Forwarder"
            Height          =   288
            Left            =   4635
            TabIndex        =   20
            Top             =   827
            Width           =   1140
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
            TabIndex        =   105
            Top             =   1470
            Width           =   2295
         End
         Begin VB.CheckBox chk_Requ 
            Caption         =   "Print Required date for each LI ? Y/N"
            Height          =   288
            Left            =   4635
            TabIndex        =   19
            Top             =   492
            Width           =   3225
         End
         Begin VB.Frame fra_Stat 
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   1620
            Left            =   4560
            TabIndex        =   55
            Top             =   1080
            Width           =   3600
            Begin VB.Label LblStatus7 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1200
               TabIndex        =   134
               Top             =   1200
               Width           =   2250
            End
            Begin VB.Label LblStatus6 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1200
               TabIndex        =   133
               Top             =   880
               Width           =   2250
            End
            Begin VB.Label LblStatus5 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1200
               TabIndex        =   132
               Top             =   560
               Width           =   2250
            End
            Begin VB.Label LblStatus4 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1200
               TabIndex        =   131
               Top             =   240
               Width           =   2250
            End
            Begin VB.Label lbl_Shipping 
               BackColor       =   &H8000000A&
               Caption         =   "Shipping"
               Height          =   225
               Left            =   105
               TabIndex        =   59
               Top             =   885
               Width           =   1200
            End
            Begin VB.Label lbl_Delivery 
               BackColor       =   &H8000000A&
               Caption         =   "Delivery"
               Height          =   225
               Left            =   105
               TabIndex        =   58
               Top             =   585
               Width           =   1200
            End
            Begin VB.Label lbl_Status 
               BackColor       =   &H8000000A&
               Caption         =   "PO"
               Height          =   225
               Left            =   120
               TabIndex        =   57
               Top             =   300
               Width           =   1200
            End
            Begin VB.Label lbl_Inventory 
               BackColor       =   &H8000000A&
               Caption         =   "Inventory"
               Height          =   225
               Left            =   105
               TabIndex        =   56
               Top             =   1215
               Width           =   1200
            End
         End
         Begin MSComCtl2.DTPicker dtpRequestedDate 
            Bindings        =   "NewPurchaseOrder.frx":0106
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
            TabIndex        =   12
            Top             =   3435
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   76283907
            CurrentDate     =   36402
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboShipper 
            Bindings        =   "NewPurchaseOrder.frx":012C
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0162
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":017E
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
            Bindings        =   "NewPurchaseOrder.frx":019A
            Height          =   315
            Left            =   1920
            TabIndex        =   16
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":01C6
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":01E2
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
         Begin VB.CheckBox chk_FrmStkMst 
            Alignment       =   1  'Right Justify
            Caption         =   "From Stock Master"
            Height          =   285
            Left            =   4560
            TabIndex        =   18
            Top             =   4440
            Width           =   2100
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBPriority 
            Bindings        =   "NewPurchaseOrder.frx":01FE
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":022A
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0246
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
            Bindings        =   "NewPurchaseOrder.frx":0262
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   1800
            Width           =   2295
            DataFieldList   =   "Column 0"
            _Version        =   196617
            DataMode        =   2
            Cols            =   1
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":028E
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":02AA
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
            Bindings        =   "NewPurchaseOrder.frx":02C6
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":02F2
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":030E
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
            Bindings        =   "NewPurchaseOrder.frx":032A
            Height          =   315
            Left            =   1920
            TabIndex        =   11
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0356
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0372
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
            Bindings        =   "NewPurchaseOrder.frx":038E
            Height          =   315
            Left            =   1920
            TabIndex        =   13
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":03BA
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":03D6
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
            Bindings        =   "NewPurchaseOrder.frx":03F2
            Height          =   315
            Left            =   1920
            TabIndex        =   15
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":041E
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":043A
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
            Bindings        =   "NewPurchaseOrder.frx":0456
            Height          =   315
            Left            =   5760
            TabIndex        =   14
            Top             =   3780
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0482
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":049E
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
         Begin VB.Label Label1 
            Caption         =   "Contact Ph"
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Contact Name"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Company"
            Height          =   225
            Left            =   90
            TabIndex        =   129
            Top             =   4125
            Width           =   1350
         End
         Begin VB.Label lbl_InvLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Location"
            Height          =   225
            Left            =   90
            TabIndex        =   128
            Top             =   4455
            Width           =   1350
         End
         Begin VB.Label lbl_Revision 
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Number"
            Height          =   225
            Left            =   90
            TabIndex        =   126
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
            TabIndex        =   125
            Top             =   135
            Width           =   615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Term"
            Height          =   225
            Left            =   4560
            TabIndex        =   120
            Top             =   4800
            Width           =   1365
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "T && C"
            Height          =   225
            Left            =   60
            TabIndex        =   119
            Top             =   4800
            Width           =   1605
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
            TabIndex        =   116
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
            TabIndex        =   115
            Top             =   3120
            Width           =   1275
         End
         Begin VB.Label LblAppBy 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   114
            Top             =   2130
            Width           =   2295
         End
         Begin VB.Label lbl_Supplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Code"
            Height          =   225
            Left            =   90
            TabIndex        =   74
            Top             =   2745
            Width           =   1125
         End
         Begin VB.Label lbl_ToBe 
            BackStyle       =   0  'Transparent
            Caption         =   "To Be Used For"
            Height          =   225
            Left            =   90
            TabIndex        =   73
            Top             =   2415
            Width           =   1605
         End
         Begin VB.Label lbl_Shipper 
            BackStyle       =   0  'Transparent
            Caption         =   "Shipper"
            Height          =   225
            Left            =   90
            TabIndex        =   72
            Top             =   465
            Width           =   1245
         End
         Begin VB.Label lbl_Currency 
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            Height          =   225
            Left            =   90
            TabIndex        =   71
            Top             =   3795
            Width           =   1665
         End
         Begin VB.Label lbl_DelivDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Required"
            Height          =   225
            Left            =   4560
            TabIndex        =   70
            Top             =   3435
            Width           =   1845
         End
         Begin VB.Label lbl_RevisionDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Date"
            Height          =   225
            Left            =   2160
            TabIndex        =   69
            Top             =   135
            Width           =   1080
         End
         Begin VB.Label lbl_RequDate 
            BackStyle       =   0  'Transparent
            Caption         =   "PO Creation Date"
            Height          =   225
            Left            =   4560
            TabIndex        =   68
            Top             =   2775
            Width           =   1815
         End
         Begin VB.Label lbl_ShipTo 
            BackStyle       =   0  'Transparent
            Caption         =   "Ship To"
            Height          =   225
            Left            =   4560
            TabIndex        =   67
            Top             =   3765
            Width           =   1095
         End
         Begin VB.Label lbl_ChargeTo 
            BackStyle       =   0  'Transparent
            Caption         =   "Charge To"
            Height          =   285
            Left            =   90
            TabIndex        =   66
            Top             =   795
            Width           =   1605
         End
         Begin VB.Label lbl_Priority 
            BackStyle       =   0  'Transparent
            Caption         =   "Shipping Mode"
            Height          =   225
            Left            =   90
            TabIndex        =   65
            Top             =   1125
            Width           =   1605
         End
         Begin VB.Label lbl_Buyer 
            BackStyle       =   0  'Transparent
            Caption         =   "Buyer"
            Height          =   225
            Left            =   90
            TabIndex        =   64
            Top             =   1455
            Width           =   1605
         End
         Begin VB.Label lbl_Originator 
            BackStyle       =   0  'Transparent
            Caption         =   "Originator"
            Height          =   225
            Left            =   90
            TabIndex        =   63
            Top             =   1785
            Width           =   1725
         End
         Begin VB.Label lbl_DateSent 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Sent"
            Height          =   225
            Left            =   4560
            TabIndex        =   62
            Top             =   3105
            Width           =   1815
         End
         Begin VB.Label lbl_Site 
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
            Height          =   225
            Left            =   4560
            TabIndex        =   61
            Top             =   4080
            Width           =   330
         End
         Begin VB.Label lbl_ApprovedBy 
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
            Height          =   225
            Left            =   90
            TabIndex        =   60
            Top             =   2115
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   2
         Left            =   -73035
         TabIndex        =   38
         Top             =   528
         Width           =   1695
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   1
         Left            =   -74760
         TabIndex        =   40
         Top             =   550
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74640
         TabIndex        =   46
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame fra_LineItem 
         BorderStyle     =   0  'None
         Height          =   5280
         Left            =   -74880
         TabIndex        =   78
         Top             =   960
         Width           =   8520
         Begin VB.CommandButton CmdcopyLI 
            Caption         =   "Copy From ...."
            Height          =   305
            Index           =   0
            Left            =   360
            TabIndex        =   32
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txt_SerialNum 
            DataField       =   "poi_serlnumb"
            DataMember      =   "POITEM"
            Height          =   285
            Left            =   1920
            TabIndex        =   28
            Top             =   1880
            Width           =   1836
         End
         Begin VB.TextBox txt_remk 
            DataField       =   "poi_remk"
            DataMember      =   "POITEM"
            DataSource      =   "deIms"
            Height          =   675
            Left            =   2040
            MaxLength       =   16
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   4440
            Width           =   6420
         End
         Begin VB.TextBox txt_Descript 
            DataField       =   "poi_desc"
            DataMember      =   "POITEM"
            Height          =   675
            Left            =   2040
            MaxLength       =   400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   3720
            Width           =   6420
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboRequisition 
            Bindings        =   "NewPurchaseOrder.frx":04BA
            Height          =   315
            Left            =   5760
            TabIndex        =   24
            Top             =   180
            Width           =   1575
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":04C5
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":04E1
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
            DataFieldToDisplay=   "Column 0"
         End
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
            Left            =   7080
            TabIndex        =   31
            Top             =   2640
            Width           =   1275
         End
         Begin VB.TextBox txt_LI 
            BackColor       =   &H00FFFFC0&
            DataField       =   "poi_liitnumb"
            DataMember      =   "POITEM"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   113
            Top             =   180
            Width           =   435
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   3240
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTP_Required 
            Bindings        =   "NewPurchaseOrder.frx":04FD
            DataField       =   "poi_liitreqddate"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   5760
            TabIndex        =   35
            Top             =   510
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   76283907
            CurrentDate     =   36405
         End
         Begin VB.TextBox txt_TotalLIs 
            BackColor       =   &H00FFFFC0&
            DataField       =   "LCount"
            DataMember      =   "LineItemCount"
            Enabled         =   0   'False
            Height          =   315
            Left            =   2670
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   180
            Width           =   420
         End
         Begin VB.TextBox txt_AFE 
            DataField       =   "poi_afe"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   1200
            Width           =   2310
         End
         Begin VB.Frame fra_Status 
            Caption         =   "Status"
            Enabled         =   0   'False
            Height          =   1530
            Left            =   3960
            TabIndex        =   85
            Top             =   780
            Width           =   4530
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "NewPurchaseOrder.frx":052E
               Height          =   315
               Index           =   0
               Left            =   1320
               TabIndex        =   106
               Top             =   150
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
               Bindings        =   "NewPurchaseOrder.frx":0555
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   107
               Top             =   480
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
               Bindings        =   "NewPurchaseOrder.frx":057C
               Height          =   315
               Index           =   2
               Left            =   1320
               TabIndex        =   108
               Top             =   810
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
               Bindings        =   "NewPurchaseOrder.frx":05A3
               Height          =   315
               Index           =   3
               Left            =   1320
               TabIndex        =   109
               Top             =   1140
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
               Caption         =   "Inventory"
               Height          =   225
               Left            =   105
               TabIndex        =   89
               Top             =   1125
               Width           =   1200
            End
            Begin VB.Label lbl_StatShipping 
               Caption         =   "Shipping"
               Height          =   225
               Left            =   105
               TabIndex        =   88
               Top             =   810
               Width           =   1200
            End
            Begin VB.Label lbl_StatDelivery 
               Caption         =   "Delivery"
               Height          =   225
               Left            =   105
               TabIndex        =   87
               Top             =   495
               Width           =   1200
            End
            Begin VB.Label lbl_StatItem 
               Caption         =   "Item"
               Height          =   225
               Left            =   105
               TabIndex        =   86
               Top             =   195
               Width           =   1200
            End
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCommoditty 
            Bindings        =   "NewPurchaseOrder.frx":05CA
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Top             =   510
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":05F6
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0612
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
            Columns(0).Width=   5292
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Frame fra_Quantity 
            Height          =   1320
            Left            =   150
            TabIndex        =   50
            Top             =   2280
            Width           =   5745
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
               TabIndex        =   29
               Top             =   360
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
               TabIndex        =   112
               Top             =   840
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
               TabIndex        =   111
               Top             =   840
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
               TabIndex        =   118
               Top             =   840
               Width           =   720
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBUnit 
               Bindings        =   "NewPurchaseOrder.frx":062E
               Height          =   315
               Left            =   4320
               TabIndex        =   30
               Top             =   360
               Width           =   1215
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
               _Version        =   196617
               DataMode        =   2
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
               stylesets(0).Picture=   "NewPurchaseOrder.frx":065A
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
               stylesets(1).Picture=   "NewPurchaseOrder.frx":0676
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
            Begin VB.Label lbl_Delivered 
               Caption         =   "Delivered"
               Height          =   225
               Left            =   120
               TabIndex        =   83
               Top             =   840
               Width           =   690
            End
            Begin VB.Label lbl_Shipped 
               Caption         =   "Shipped"
               Height          =   225
               Left            =   2040
               TabIndex        =   82
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lbl_Requested 
               Caption         =   "Qtantity Required"
               Height          =   225
               Left            =   120
               TabIndex        =   81
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lbl_Unit 
               Caption         =   "Unit"
               Height          =   195
               Left            =   3480
               TabIndex        =   80
               Top             =   360
               Width           =   360
            End
            Begin VB.Label lbl_Inventory2 
               Caption         =   "Inventory"
               Height          =   225
               Left            =   3960
               TabIndex        =   79
               Top             =   840
               Width           =   735
            End
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboManNumber 
            Bindings        =   "NewPurchaseOrder.frx":0692
            Height          =   315
            Left            =   1920
            TabIndex        =   25
            Top             =   855
            Width           =   1830
            DataFieldList   =   "Column 0"
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":069D
            stylesets(0).AlignmentText=   0
            stylesets(1).Name=   "ColHeader"
            stylesets(1).HasFont=   -1  'True
            BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(1).Picture=   "NewPurchaseOrder.frx":06B9
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   3440
            Columns(0).Caption=   "Part Number"
            Columns(0).Name =   "Part Number"
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   4683
            Columns(1).Caption=   "Manufacturer"
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
            Bindings        =   "NewPurchaseOrder.frx":06D5
            Height          =   315
            Left            =   1920
            TabIndex        =   27
            Top             =   1560
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0701
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":071D
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
         Begin VB.Label lbl_PartNum 
            Caption         =   "Manufacturer P/N"
            Height          =   225
            Left            =   120
            TabIndex        =   124
            Top             =   855
            Width           =   1815
         End
         Begin VB.Label lbl_SerialNum 
            Caption         =   "Serial Number"
            Height          =   225
            Left            =   120
            TabIndex        =   123
            Top             =   1875
            Width           =   1785
         End
         Begin VB.Label Label8 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   1080
            TabIndex        =   122
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label lblReqLineitem 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "poi_requliitnumb"
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7440
            TabIndex        =   121
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lbl_Of 
            Caption         =   "of"
            Height          =   225
            Left            =   2430
            TabIndex        =   99
            Top             =   180
            Width           =   150
         End
         Begin VB.Label lbl_Cost 
            Caption         =   "Unit Price"
            Height          =   225
            Left            =   6000
            TabIndex        =   98
            Top             =   2640
            Width           =   825
         End
         Begin VB.Label lbl_Description 
            Caption         =   "Description"
            Height          =   225
            Left            =   1080
            TabIndex        =   97
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label lbl_Item 
            Caption         =   "Item"
            Height          =   225
            Left            =   120
            TabIndex        =   96
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label lbl_Commodity 
            Caption         =   "Commodity"
            Height          =   225
            Left            =   120
            TabIndex        =   95
            Top             =   510
            Width           =   1740
         End
         Begin VB.Label lbl_AFE 
            Caption         =   "A.F.E"
            Height          =   225
            Left            =   120
            TabIndex        =   94
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lbl_Custom 
            Caption         =   "Customs"
            Height          =   225
            Left            =   120
            TabIndex        =   93
            Top             =   1530
            Width           =   1290
         End
         Begin VB.Label lbl_Requisition 
            AutoSize        =   -1  'True
            Caption         =   "From R/Q/B#"
            Height          =   195
            Left            =   3960
            TabIndex        =   92
            Top             =   180
            Width           =   1800
         End
         Begin VB.Label lbl_Total 
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
            Left            =   6000
            TabIndex        =   91
            Top             =   3240
            Width           =   1065
         End
         Begin VB.Label lbl_RequDate2 
            Caption         =   "Date Required"
            Height          =   225
            Left            =   3960
            TabIndex        =   90
            Top             =   540
            Width           =   1800
         End
      End
      Begin VB.TextBox txtClause 
         DataField       =   "poc_clau"
         DataMember      =   "POCLAUSE"
         Height          =   4875
         Left            =   -74760
         MaxLength       =   3500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   960
         Width           =   8300
      End
      Begin VB.TextBox txtRemarks 
         DataField       =   "por_remk"
         DataMember      =   "POREM"
         Height          =   5175
         Left            =   -74760
         MaxLength       =   7000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   1020
         Width           =   8295
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   2055
         Left            =   -72840
         TabIndex        =   48
         Top             =   3840
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
         TabIndex        =   44
         Top             =   3360
         Width           =   6144
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74640
         TabIndex        =   45
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Addterms 
         Caption         =   "Add Clause"
         Height          =   288
         Left            =   -74730
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   528
         Width           =   1695
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   1650
         Left            =   -74700
         TabIndex        =   49
         Top             =   3735
         Width           =   1635
         Begin VB.OptionButton opt_SupFax 
            Caption         =   "Supplier's"
            Height          =   288
            Left            =   60
            TabIndex        =   41
            Top             =   336
            Width           =   1440
         End
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   60
            TabIndex        =   42
            Top             =   768
            Width           =   1515
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   60
            TabIndex        =   43
            Top             =   1260
            Width           =   1515
         End
      End
      Begin VB.Frame fra_PO 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   225
         TabIndex        =   100
         Top             =   450
         Width           =   8430
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssOleDbPO 
            Bindings        =   "NewPurchaseOrder.frx":0739
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   120
            Width           =   2295
            DataFieldList   =   "Column 0"
            ListAutoValidate=   0   'False
            MinDropDownItems=   8
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":0765
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":0781
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
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBDocType 
            Bindings        =   "NewPurchaseOrder.frx":079D
            Height          =   315
            Left            =   5520
            TabIndex        =   1
            Top             =   120
            Width           =   2775
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
            stylesets(0).Picture=   "NewPurchaseOrder.frx":07C9
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
            stylesets(1).Picture=   "NewPurchaseOrder.frx":07E5
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
            _ExtentX        =   4895
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "Column 1"
         End
         Begin VB.Label lbl_DocumentType 
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   3960
            TabIndex        =   102
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lbl_Purchase 
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.Frame fra_LI 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   -74865
         TabIndex        =   75
         Top             =   480
         Width           =   8520
         Begin VB.Label LblPOI_Doctype 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            TabIndex        =   117
            Top             =   180
            Width           =   2535
         End
         Begin VB.Label LblPOi_PONUMB 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "poi_ponumb"
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   51
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label lbl_DocType 
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
            TabIndex        =   77
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label lbl_PO2 
            Caption         =   "Purchase Order#"
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
            Left            =   105
            TabIndex        =   76
            Top             =   225
            Width           =   1665
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
         Height          =   2325
         Left            =   -72720
         TabIndex        =   47
         Top             =   660
         Width           =   6015
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
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
         stylesets(0).Picture=   "NewPurchaseOrder.frx":0801
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
         stylesets(1).Picture=   "NewPurchaseOrder.frx":081D
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowAddNew     =   -1  'True
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
         Columns(0).Width=   5292
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         TabNavigation   =   1
         _ExtentX        =   10610
         _ExtentY        =   4101
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
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74715
         TabIndex        =   104
         Top             =   570
         Width           =   1500
      End
      Begin VB.Label lbl_New 
         Caption         =   "New"
         Height          =   300
         Left            =   -74715
         TabIndex        =   103
         Top             =   3420
         Width           =   1380
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -66720
         Y1              =   3210
         Y2              =   3210
      End
   End
   Begin VB.Label LblCompanyCode 
      Caption         =   "Company Code"
      Height          =   375
      Left            =   5040
      TabIndex        =   137
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   5205
      TabIndex        =   53
      Top             =   6480
      Width           =   3660
   End
End
Attribute VB_Name = "frm_NewPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MainPO As New IMSPODLL.MainPO
Dim Poheader As IMSPODLL.Poheader
Dim PoItem As IMSPODLL.POITEMS
Dim PORemark As IMSPODLL.POREMARKS
Dim POClause As IMSPODLL.POClauses
Dim PoReceipients As IMSPODLL.PoReceipients
Dim CheckLoad As Boolean
Dim CheckErrors As Boolean
Dim CheckIfCombosLoaded As Boolean
Dim IntiClass As InitialValuesPOheader
Dim FNamespace  As String
Dim FormMode As FormMode
Dim mIsPoheaderCombosLoaded As Boolean
Dim mIsPoNumbComboLoaded As Boolean
Dim mIsPoItemsComboLoaded As Boolean
Dim mSaveToPoRevision As Boolean
Dim mIsPoItemCombosLoaded As Boolean
Dim mIsDocTypeLoaded As Boolean
Dim mIsInvLocationLoaded As Boolean
Dim Lookups As IMSPODLL.Lookups
Dim GRsDoctype As ADODB.Recordset
Dim RsUNits As ADODB.Recordset
Dim objUnits As IMSPODLL.PoUnits
Dim GPOnumb As String
Dim mLoadMode As LoadMode
Dim rsDOCTYPE As ADODB.Recordset
Dim IsThisADifferentPO As Boolean
Dim mCheckPoFields As Boolean
Dim mCheckLIFields As Boolean
Dim mIsPoHeaderRsetsInit As Boolean
Dim msg1 As String
Dim msg2 As String
'Dim FNamespace As String
Dim WithEvents st As frm_ShipTerms
Attribute st.VB_VarHelpID = -1
Dim WithEvents comsearch As frm_StockSearch
Attribute comsearch.VB_VarHelpID = -1

Private Sub NavBar1_OnCloseClick()
Unload Me
End Sub

Private Sub NavBar1_OnEMailClick()
On Error Resume Next

Dim i As RPTIFileInfo
Dim Params(1) As String
   
   
    With i
    
        
        '.Login = "sa"  'M
        .Login = ConnInfo.UId 'UserId 'M
        .Password = ConnInfo.Pwd ' DBPassword  M
        .ReportFileName = ReportPath & "po.rpt"
                       
        Params(0) = "namespace=" & deIms.NameSpace
        Params(1) = "ponumb=" & Poheader.Ponumb & ""
        
        .Parameters = Params
        
        
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
    If PoReceipients Is Nothing Then Set PoReceipients = MainPO.PoReceipients
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

Private Sub NavBar1_OnPrintClick()

On Error GoTo Errhandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = ReportPath + "po.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Poheader.Ponumb + ";true"
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("M00392") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction", msg1) 'J modified
        Call translator.Translate_Reports("po.rpt") 'J added
        msg1 = translator.Trans("M00091") 'J added
        msg2 = translator.Trans("M00093") 'J added
        Dim curr
        curr = " : $ "
        .Formulas(99) = "gttext = ' " + msg1 + " ' + {DOCTYPE.doc_desc} + ' " + msg2 + " ' + {CURRENCY.curr_desc} + ' " + curr + "' + totext({PO.po_totacost})" 'J modified
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
    Exit Sub
    
Errhandler:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)
    End If
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

Private Sub ssdcboCommoditty_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboCommoditty.DroppedDown Then ssdcboCommoditty.DroppedDown = True
End Sub

Private Sub ssdcboCommoditty_KeyPress(KeyAscii As Integer)
ssdcboCommoditty.MoveNext
End Sub

Private Sub ssdcboCondition_Click()
ssdcboCondition.SelStart = 0
ssdcboCondition.SelLength = 0
End Sub

Private Sub ssdcboCondition_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboCondition.DroppedDown Then ssdcboCondition.DroppedDown = True
End Sub

Private Sub ssdcboCondition_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboCondition.text)) > 0 And Not ssdcboCondition.IsItemInList Then
  Cancel = True
   ssdcboCondition.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub ssdcboDelivery_Click()
ssdcboDelivery.SelStart = 0
ssdcboDelivery.SelLength = 0
End Sub

Private Sub ssdcboDelivery_DropDown()
ssdcboDelivery.RemoveAll
If deIms.rsTermDelivery.State = 1 Then deIms.rsTermDelivery.Close
Call deIms.TermDelivery(FNamespace)
 deIms.rsTermDelivery.Filter = "tod_actvflag<>0"
Do While Not deIms.rsTermDelivery.EOF
       ssdcboDelivery.AddItem deIms.rsTermDelivery!tod_termcode & ";" & deIms.rsTermDelivery!tod_desc
       deIms.rsTermDelivery.MoveNext
Loop

deIms.rsTermDelivery.Filter = ""

End Sub

Private Sub ssdcboDelivery_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboDelivery.DroppedDown Then ssdcboDelivery.DroppedDown = True
End Sub

Private Sub ssdcboDelivery_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboDelivery.text)) > 0 And Not ssdcboDelivery.IsItemInList Then
  Cancel = True
   ssdcboDelivery.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub ssdcboManNumber_Click()
ssdcboManNumber.SelStart = 0
ssdcboManNumber.SelLength = 0
End Sub

Private Sub ssdcboManNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboManNumber.DroppedDown Then ssdcboManNumber.DroppedDown = True
End Sub

Private Sub ssdcboManNumber_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboManNumber.text)) > 0 Then
   If ssdcboManNumber.IsItemInList = False Then
        Cancel = True
        ssdcboManNumber.SetFocus
        MsgBox "Invalid Value For Company Code.", , "Imswin"
    End If
End If
End Sub

Private Sub ssdcboShipper_DropDown()
ssdcboShipper.RemoveAll
If deIms.rsSHIPPER.State = 1 Then deIms.rsSHIPPER.Close
Call deIms.SHIPPER(FNamespace)
Do While Not deIms.rsSHIPPER.EOF
       ssdcboShipper.AddItem deIms.rsSHIPPER!shi_code & ";" & deIms.rsSHIPPER!shi_name
       deIms.rsSHIPPER.MoveNext
   Loop
End Sub

Private Sub ssdcboShipper_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboShipper.DroppedDown Then ssdcboShipper.DroppedDown = True
End Sub

Private Sub ssdcboShipper_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboShipper.text)) > 0 And Not ssdcboShipper.IsItemInList Then
  Cancel = True
   ssdcboShipper.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBCompany_DropDown()
SSOleDBCompany.RemoveAll
If deIms.rsActiveCompany.State = 1 Then deIms.rsActiveCompany.Close
Call deIms.ActiveCompany(FNamespace)

Do While Not deIms.rsActiveCompany.EOF
       SSOleDBCompany.AddItem deIms.rsActiveCompany!com_compcode & ";" & deIms.rsActiveCompany!com_name
       deIms.rsActiveCompany.MoveNext
       
   Loop
End Sub

Private Sub SSOleDBCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCompany.DroppedDown Then SSOleDBCompany.DroppedDown = True
End Sub

Private Sub SSOleDBCompany_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCompany.text)) > 0 Then
   If SSOleDBCompany.IsItemInList = False Then
        Cancel = True
        SSOleDBCompany.SetFocus
        MsgBox "Invalid Value For Company Code.", , "Imswin"
   
     
    End If
End If
End Sub

Private Sub SSOleDBCurrency_DropDown()
SSOleDBCurrency.RemoveAll
If deIms.rsCURRENCY.State = 1 Then deIms.rsCURRENCY.Close
Call deIms.Currency(FNamespace)

 Do While Not deIms.rsCURRENCY.EOF
       SSOleDBCurrency.AddItem deIms.rsCURRENCY!curr_code & ";" & deIms.rsCURRENCY!curr_desc
       deIms.rsCURRENCY.MoveNext
   Loop
End Sub

Private Sub SSOleDBCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCurrency.DroppedDown Then SSOleDBCurrency.DroppedDown = True
End Sub

Private Sub SSOleDBCurrency_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCurrency.text)) > 0 Then
    If Not SSOleDBCurrency.IsItemInList Then
         Cancel = True
          SSOleDBCurrency.SetFocus
        MsgBox "Invalid Value", , "Imswin"
    Else
      If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
      If Lookups.CurrencyDetlExist(SSOleDBCurrency.Columns(0).text) = False Then
         MsgBox "No Currency Detail for today.Please Update Currency Table"
         SSOleDBCompany.text = ""
         SSOleDBCompany.SetFocus
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
If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
   
   Set rsCUSTOM = Lookups.GetCustom
   Do While Not rsCUSTOM.EOF
     SSOleDBCustCategory.AddItem rsCUSTOM!cust_cate
     rsCUSTOM.MoveNext
   Loop
   rsCUSTOM.Close
 Set rsCUSTOM = Nothing
 
End Sub

Private Sub SSOleDBCustCategory_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCustCategory.DroppedDown Then SSOleDBCustCategory.DroppedDown = True
End Sub

Private Sub SSOleDBCustCategory_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCustCategory.text)) > 0 Then
   If SSOleDBCustCategory.IsItemInList = False Then
        Cancel = True
        SSOleDBCustCategory.SetFocus
        MsgBox "Invalid Value For Company Code.", , "Imswin"
    End If
End If
End Sub

Private Sub SSOleDBDocType_Click()
SSOleDBDocType.SelLength = 0
SSOleDBDocType.SelStart = 0
End Sub

Private Sub SSOleDBDocType_DropDown()
'If mIsDocTypeLoaded = True Then
      SSOleDBDocType.RemoveAll
      If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
        Dim GRsDoctype As ADODB.Recordset
        Set GRsDoctype = Lookups.GetDoctypeForUser(CurrentUser)

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

Private Sub SSOleDBDocType_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBDocType.DroppedDown Then SSOleDBDocType.DroppedDown = True
End Sub

Private Sub SSOleDBDocType_KeyPress(KeyAscii As Integer)
SSOleDBDocType.MoveNext
End Sub

Private Sub SSOleDBDocType_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBDocType.text)) > 0 And Not SSOleDBDocType.IsItemInList Then
  Cancel = True
   SSOleDBDocType.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBInvLocation_Click()
SSOleDBInvLocation.Tag = SSOleDBInvLocation.Columns(0).text
SSOleDBInvLocation.SelLength = 0
SSOleDBInvLocation.SelStart = 0
End Sub

Private Sub SSOleDBInvLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBInvLocation.DroppedDown Then SSOleDBInvLocation.DroppedDown = True
End Sub

Private Sub SSOleDBInvLocation_KeyPress(KeyAscii As Integer)
SSOleDBInvLocation.MoveNext
End Sub

Private Sub SSOleDBInvLocation_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBInvLocation.text)) > 0 And Not SSOleDBInvLocation.IsItemInList Then
  Cancel = True
   SSOleDBInvLocation.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBOriginator_Click()
SSOleDBOriginator.SelLength = 0
SSOleDBOriginator.SelStart = 0
End Sub

Private Sub SSOleDBOriginator_DropDown()
SSOleDBOriginator.RemoveAll
If deIms.rsActiveOriginator.State = 1 Then deIms.rsActiveOriginator.Close
Call deIms.ActiveOriginator(FNamespace)
Do While Not deIms.rsActiveOriginator.EOF
       SSOleDBOriginator.AddItem deIms.rsActiveOriginator!ori_code '& ";" & deIms.rsActiveOriginator!ori_code
       deIms.rsActiveOriginator.MoveNext
   Loop
End Sub

Private Sub SSOleDBOriginator_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBOriginator.DroppedDown Then SSOleDBOriginator.DroppedDown = True
End Sub

Private Sub SSOleDBOriginator_KeyPress(KeyAscii As Integer)
SSOleDBOriginator.MoveNext
End Sub

Private Sub SSOleDBOriginator_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBOriginator.text)) > 0 And Not SSOleDBOriginator.IsItemInList Then
  Cancel = True
   SSOleDBOriginator.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBPO_DropDown()
If mIsPoNumbComboLoaded = False Then
  If deIms.rsPonumb.State = 1 Then
     deIms.rsPonumb.Close
  End If
   Call deIms.Ponumb(deIms.NameSpace)
    Do While Not deIms.rsPonumb.EOF
       ssOleDbPO.AddItem deIms.rsPonumb!po_ponumb
       deIms.rsPonumb.MoveNext
    Loop
    mIsPoNumbComboLoaded = True

End If
End Sub

Private Sub SSOleDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
 If Not FormMode = mdCreation Then
    If Not ssOleDbPO.DroppedDown Then ssOleDbPO.DroppedDown = True
 End If
End Sub

Private Sub SSOleDBPO_KeyPress(KeyAscii As Integer)
If Not FormMode = mdCreation Then ssOleDbPO.MoveNext
End Sub

Private Sub SSOleDBPriority_Click()
SSOleDBPriority.SelLength = 0
SSOleDBPriority.SelStart = 0
End Sub

Private Sub SSOleDBPriority_DropDown()
SSOleDBPriority.RemoveAll
If deIms.rsPRIORITY.State = 1 Then deIms.rsPRIORITY.Close
Call deIms.PRIORITY(FNamespace)
Do While Not deIms.rsPRIORITY.EOF
       SSOleDBPriority.AddItem deIms.rsPRIORITY!pri_code & ";" & deIms.rsPRIORITY!pri_desc
       deIms.rsPRIORITY.MoveNext
   Loop
End Sub

Private Sub SSOleDBPriority_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBPriority.DroppedDown Then SSOleDBPriority.DroppedDown = True
End Sub

Private Sub SSOleDBPriority_KeyPress(KeyAscii As Integer)
SSOleDBPriority.MoveNext
End Sub

Private Sub SSOleDBPriority_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBPriority.text)) > 0 And Not SSOleDBPriority.IsItemInList Then
  Cancel = True
   SSOleDBPriority.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBShipTo_Click()
SSOleDBShipTo.SelLength = 0
SSOleDBShipTo.SelStart = 0
End Sub

Private Sub SSOleDBShipTo_DropDown()
SSOleDBShipTo.RemoveAll

If deIms.rsActiveShipTo.State = 1 Then deIms.rsActiveShipTo.Close
Call deIms.ActiveShipTo(FNamespace)
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

Private Sub SSOleDBShipTo_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBShipTo.text)) > 0 And Not SSOleDBShipTo.IsItemInList Then
  Cancel = True
   SSOleDBShipTo.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSoledbSupplier_DropDown()
'If mIsSupplierComboLoaded = False Then

Dim rsSUPPLIER As ADODB.Recordset

SSoledbSupplier.RemoveAll
If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
If Lookups.GetUserMenuLevel(CurrentUser) = 5 Then
  Set rsSUPPLIER = Lookups.GetLocalSuppliers
Else
  If deIms.rsActiveSupplier.State = 1 Then
     deIms.rsActiveSupplier.Close
     Call deIms.ActiveSupplier(deIms.NameSpace)
  End If
  Set rsSUPPLIER = deIms.rsActiveSupplier
End If

If rsSUPPLIER.RecordCount > 0 Then
    rsSUPPLIER.MoveFirst
    Do While Not rsSUPPLIER.EOF
       SSoledbSupplier.AddItem rsSUPPLIER!sup_code & ";" & rsSUPPLIER!sup_name & ";" & rsSUPPLIER!sup_city & ";" & rsSUPPLIER!sup_phonnumb
       rsSUPPLIER.MoveNext
    Loop
End If

End Sub

Private Sub SSoledbSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSoledbSupplier.DroppedDown Then SSoledbSupplier.DroppedDown = True
End Sub

Private Sub SSoledbSupplier_Validate(Cancel As Boolean)
If Len(Trim$(SSoledbSupplier.text)) > 0 And Not SSoledbSupplier.IsItemInList Then
  Cancel = True
   SSoledbSupplier.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBToBeUsedFor_Click()
SSOleDBToBeUsedFor.SelLength = 0
SSOleDBToBeUsedFor.SelStart = 0
End Sub

Private Sub SSOleDBToBeUsedFor_DropDown()
SSOleDBToBeUsedFor.RemoveAll
If deIms.rsActiveTbu.State = 1 Then deIms.rsActiveTbu.Close
Call deIms.ActiveTbu(FNamespace)
Do While Not deIms.rsActiveTbu.EOF
       SSOleDBToBeUsedFor.AddItem deIms.rsActiveTbu!tbu_name '& ";" & deIms.rsActiveOriginator!tbu_name
       deIms.rsActiveTbu.MoveNext
   Loop
End Sub

Private Sub SSOleDBToBeUsedFor_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBToBeUsedFor.DroppedDown Then SSOleDBToBeUsedFor.DroppedDown = True
End Sub

Private Sub SSOleDBToBeUsedFor_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBToBeUsedFor.text)) > 0 And Not SSOleDBToBeUsedFor.IsItemInList Then
  Cancel = True
   SSOleDBToBeUsedFor.SetFocus
 MsgBox "Invalid Value", , "Imswin"
End If
End Sub

Private Sub SSOleDBUnit_DropDown()
If Len(ssdcboCommoditty.text) > 0 And chk_FrmStkMst.Value = 1 Then
        If objUnits Is Nothing Then Set objUnits = MainPO.PoUnits
       
        objUnits.StockNumber = Trim$(ssdcboCommoditty.text)
       
        SSOleDBUnit.RemoveAll
        
        If RsUNits Is Nothing Then
          If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
          Set RsUNits = Lookups.GetAllUnits
        End If
        RsUNits.MoveFirst
        RsUNits.Find ("uni_code='" & Trim$(objUnits.PrimaryUnit) & "'")
        SSOleDBUnit.AddItem objUnits.PrimaryUnit & ";" & RsUNits("uni_desc")
        
        RsUNits.MoveFirst
        RsUNits.Find ("uni_code='" & Trim$(objUnits.SecondaryUnit) & "'")
        SSOleDBUnit.AddItem objUnits.SecondaryUnit & ";" & RsUNits("uni_desc")


        

        
End If
End Sub

Private Sub SSOleDBUnit_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBUnit.text)) > 0 Then
   
   'If SSOleDBUnit.IsItemInList Then
    If Len(Trim$(ssdcboCommoditty.text)) > 0 And Len(txt_Requested) > 0 Then
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
        Cancel = True
        SSOleDBUnit.SetFocus
        MsgBox "Invalid Value", , "Imswin"
   End If
   
'End If
End Sub

Private Sub st_Completed(Cancelled As Boolean, Terms As String)
On Error Resume Next

    If Not Cancelled Then
        txtClause.SelText = Terms
        
        Terms = txtClause.text
        POClause.Clause = Terms
    End If
    
    Set st = Nothing
End Sub

  
      



Private Sub chk_FrmStkMst_Click()
mIsPoItemsComboLoaded = False
End Sub

Private Sub Command1_Click()
'deIms.rsActiveStockMasterLooKUP.Close
'Set deIms.rsActiveStockMasterLooKUP = Nothing
'Call deIms.ActiveStockMasterLooKUP(deims.namespace)
'Set ssdcboCommoditty.DataSourceList = deIms.rsActiveStockMasterLooKUP
'    ssdcboCommoditty.DataFieldToDisplay = "stk_stcknumb"
'    ssdcboCommoditty.DataFieldList = "stk_desc"
End Sub

Private Sub cmd_Add_Click()
On Error Resume Next

    If Len(Trim$(txt_Recipient)) Then
       
       If InStr(1, txt_Recipient, "@") Then
           txt_Recipient = UCase(txt_Recipient)
           If InStr(1, txt_Recipient, "INTERNET!") = 0 Then txt_Recipient = ("INTERNET!" & txt_Recipient)
       
       Else
           txt_Recipient = ("FAX!" & txt_Recipient)
       End If
       
        Call AddRecepient(txt_Recipient)
        txt_Recipient = ""
    Else
        dgRecepients_DblClick
    End If
End Sub

Private Sub cmd_Addterms_Click()
On Error Resume Next
    Set st = New frm_ShipTerms
    st.Show
    If Err Then Err.Clear
End Sub

Private Sub CmdcopyLI_Click(Index As Integer)
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

End Sub

Private Sub cmdremove_Click()
If FormMode = mdCreation Then
    dgRecipientList.RemoveItem dgRecipientList.SelectTypeRow
    'PoReceipients.
End If
End Sub

Private Sub dgRecepients_DblClick()
On Error Resume Next
    Call AddRecepient(dgRecepients.Columns(1).Value)
    
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsDOCTYPE = Nothing
    If deIms.rsPonumb.State = 1 Then Call deIms.rsPonumb.Close
    If deIms.rsSHIPPER.State = 1 Then Call deIms.rsSHIPPER.Close
    If deIms.rsCURRENCY.State = 1 Then Call deIms.rsCURRENCY.Close
    If deIms.rsPRIORITY.State = 1 Then Call deIms.rsPRIORITY.Close
    If deIms.rsTermDelivery.State = 1 Then Call deIms.rsTermDelivery.Close
     'Call deIms.Supplier(fnamespace)
    If deIms.rsActiveSupplier.State = 1 Then Call deIms.rsActiveSupplier.Close
    If deIms.rsTermCondition.State = 1 Then Call deIms.rsTermCondition.Close
    'Call deIms.INVENTORYLOCATION(Fnamespace, ponumb)
    If deIms.rsCOMPANY.State = 1 Then Call deIms.rsCOMPANY.Close
    'Call deIms.rsGETSYSSITE.Close
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
    If Not MainPO Is Nothing Then Set MainPO = Nothing
    
    mIsPoheaderCombosLoaded = False
    mIsDocTypeLoaded = False
    mIsPoNumbComboLoaded = False
    'Set deIms.rsSHIPPER = Nothing
''    Set deIms.rsCURRENCY = Nothing
''     Set deIms.rsPRIORITY = Nothing
''     Set deIms.rsTermDelivery = Nothing
''    ' deIms.Supplier(fnamespace)
''     Set deIms.rsActiveSupplier = Nothing
''     Set deIms.rsTermCondition = Nothing
''    ' deIms.INVENTORYLOCATION(Fnamespace, ponumb)
''     Set deIms.rsCOMPANY = Nothing
''     'deIms.rsGETSYSSITE = Nothing
''     Set deIms.rsActiveOriginator = Nothing
''     Set deIms.rsActiveTbu = Nothing
''     Set deIms.rsSERVCODECAT = Nothing
''     Set deIms.rsActiveShipTo = Nothing
''     Set deIms.rsActiveCompany = Nothing
''     Set deIms.rsCompanyLocations = Nothing
    'mIsPoHeaderRsetsInit = True
    
    
    
End Sub

Private Sub NavBar1_BeforeSaveClick()
Dim mpo As String


Select Case sst_PO.Tab
        Case 0
            Me.MousePointer = vbHourglass
            If CheckPoFields Then
                     mpo = ssOleDbPO.text
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
                     
                     
                     'De1.Cn1.BeginTrans
                     deIms.cnIms.Errors.Clear
                     
                     deIms.cnIms.BeginTrans
                     
                     If mSaveToPoRevision Then InsertPoRevision (Poheader.Ponumb)
                     mSaveToPoRevision = False
                     'CheckErrors = Poheader.SAVE
                      PoHeaderErrors = Poheader.SAVE
                      If Not PoItem Is Nothing Then
                           POITEMErrors = PoItem.Update
                          If POITEMErrors = True Then
                            WriteStatus ("Poitems Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Poitems")
                          End If
                       End If
                       
                      If Not PoReceipients Is Nothing Then
                            poRecepientsErrors = PoReceipients.Update
                         If poRecepientsErrors = True Then
                           WriteStatus ("Recipients Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Recipients")
                          End If
                      End If
                      
                      If Not PORemark Is Nothing Then
                             PoremarksErrors = PORemark.Update
                             If PoremarksErrors = True Then
                           WriteStatus ("Remarks Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Remarks")
                          End If
                      End If
                             
                      If Not POClause Is Nothing Then
                            poClauseErrors = POClause.Update
                          If poClauseErrors = True Then
                           WriteStatus ("Clause Saved Successfully")
                          Else
                            WriteStatus ("Errors Occurred While Trying to Save Clause")
                          End If
                      End If
                      
                     If PoHeaderErrors = True And POITEMErrors = True And poRecepientsErrors = True And PoremarksErrors = True And poClauseErrors = True Then
                         deIms.cnIms.CommitTrans
                         
                           Poheader.Requery
                           LoadFromPOHEADER
                           mIsPoNumbComboLoaded = False
                           MsgBox "Transaction Order # " & mpo & " saved successfully"
                           WriteStatus ("")
                       Else
                         deIms.cnIms.RollbackTrans
                          MsgBox "Errors Occured.Could Not Save The Transaction Order."
                         WriteStatus ("Rolling Back the Transaction")
                         Poheader.CancelUpdate: LoadFromPOHEADER
                       
                        WriteStatus ("")
                       End If
                       
                      
                          Set PoItem = Nothing
                          Set PoReceipients = Nothing
                          Set PORemark = Nothing
                          Set POClause = Nothing
                          
''''''''                     If CheckErrors = True Then
''''''''                       'This case may Arise when the User Clicks on save without
''''''''                       ' even going to the POITEM tab.
''''''''                       If Not PoReceipients Is Nothing Then
''''''''                          CheckErrors = PoReceipients.Update
''''''''                           Set PoReceipients = Nothing
''''''''                       End If
''''''''
''''''''                        If Not PoItem Is Nothing Then
''''''''                           SaveToPOITEM
''''''''
''''''''
''''''''                           CheckErrors = PoItem.Update
''''''''                           Set PoItem = Nothing
''''''''                        End If
''''''''                      End If
''''''''
''''''''                     'If CheckErrors = True Then CheckErrors = PORemark.Update
''''''''                     'poClauseErrors = POclause.
''''''''
''''''''                     'if de1.Cn1.Errors.Count >0 or CheckErrors
''''''''
''''''''                     If CheckErrors = True Then
''''''''                         If Not PORemark Is Nothing Then
''''''''                             'savetoPORemarks
''''''''                             CheckErrors = PORemark.Update
''''''''                             Set PORemark = Nothing
''''''''                         End If
''''''''                     End If
''''''''
''''''''                     If CheckErrors = True Then
''''''''                         If Not POClause Is Nothing Then
''''''''                             'savetoPOclause
''''''''                             CheckErrors = POClause.Update
''''''''                             Set POClause = Nothing
''''''''                         End If
''''''''                     End If
''''''''
                     
'''                     If CheckErrors = False Then
'''                            MsgBox "Errors Occured.Could Not Save The Po."
'''                            Poheader.CancelUpdate
'''                     Else
                           
                        'Getting the Saved Data to the User by Requering the Database.
''                           Poheader.Requery
''                           CheckErrors = LoadFromPOHEADER
''
''                           mIsPoNumbComboLoaded = False
                           
                         '  If Not PoReceipients Is Nothing Then Set PoReceipients = Nothing
                           
                           
                          ' If Not PoItem Is Nothing Then Set PoItem = Nothing
                           
                           'If CheckErrors = True Then MsgBox "PO saved successfully"
''                     End If
                      
                      
                    CheckErrors = True
                    
                    FormMode = ChangeMode(mdVisualization)
                    
                    If FormMode = mdVisualization Then
                       
                       NavBar1.NewEnabled = True
                       NavBar1.NextEnabled = True
                       NavBar1.PreviousEnabled = True
                       NavBar1.LastEnabled = True
                       NavBar1.FirstEnabled = True
                       NavBar1.SaveEnabled = False
                       NavBar1.EditEnabled = True
                       NavBar1.CancelEnabled = False
                       ssOleDbPO.Enabled = True
                    End If
          End If
 End Select
   Me.MousePointer = vbArrow

End Sub

Private Sub NavBar1_OnCancelClick()

If FormMode = mdModification Or FormMode = mdCreation Then

  
   'Cancelling all the Changes made by the user
        Select Case sst_PO.Tab
             
             Case 0
             ssOleDbPO.Enabled = True
              If FormMode = mdModification Then
                  Call LoadFromPOHEADER
              ElseIf FormMode = mdCreation Then
                  Poheader.CancelUpdate
                  Poheader.MoveFirst
                  Call LoadFromPOHEADER
                  'Poheader.Move Trim$(ssOleDbPO)
              End If
              FormMode = ChangeMode(mdVisualization)
             mSaveToPoRevision = False
             
             If FormMode = mdVisualization Then
                       
                       NavBar1.NewEnabled = True
                       NavBar1.NextEnabled = True
                       NavBar1.PreviousEnabled = True
                       NavBar1.LastEnabled = True
                       NavBar1.FirstEnabled = True
                       NavBar1.SaveEnabled = False
                       NavBar1.EditEnabled = True
                       NavBar1.CancelEnabled = False
                       ssOleDbPO.Enabled = True
                    End If
             
             
             Case 1
             Case 2
              
              If FormMode = mdModification And PoItem.editmode = 0 Then
                  Call LoadFromPOITEM
              ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And PoItem.editmode = 2) Then
              'CAncels only the Latest REcord
                  PoItem.CancelUpdate
                  If PoItem.Count > 0 Then
                      Call LoadFromPOITEM
                  Else
                    ClearAllPoLineItems
                  End If
              End If
              
             Case 3
               If FormMode = mdModification And PORemark.editmode = 0 Then
                  Call LoadFromPORemarks
               ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And PORemark.editmode = 2) Then
                     PORemark.CancelUpdate
                   Call LoadFromPORemarks
               End If
             Case 4
               If FormMode = mdModification And POClause.editmode = 0 Then
                  Call LoadFromPOClause
                  
               ElseIf (FormMode = mdCreation) Or (FormMode = mdModification And POClause.editmode = 2) Then
                   POClause.CancelUpdate
                   Call LoadFromPOClause
               End If
         End Select
 End If
End Sub

Private Sub NavBar1_OnEditClick()

    If Trim$(Poheader.stas) = "CA" Or Trim$(Poheader.stas) = "CL" Then
         MsgBox " Can not Edit this Document.It is Cancelled"
         GoTo CANNOTEDIT
    End If
     
    If deIms.CanUserEditDocType(CurrentUser, Trim$(Poheader.docutype)) Then
     
       If Len(Trim$(LblRevNumb)) > 0 Then
         If CInt(LblRevNumb) > 0 Then
                If MsgBox(" You will Create a new Revision.Do You want to Continue ?", vbYesNo) = vbYes Then
                    
                    LblRevNumb.Caption = IIf(Len(LblRevNumb.Caption) = 0, 0, CInt(LblRevNumb.Caption) + 1)
                    LblRevDate = Format(Now(), "MM/DD/YY")
                    mSaveToPoRevision = True
                 Else
                    GoTo CANNOTEDIT
                 End If
           End If
        Else
             GoTo CANNOTEDIT
         
         End If
        
           Me.MousePointer = vbHourglass
           
           
           FormMode = ChangeMode(mdModification)
        
           
           Select Case sst_PO.Tab
             Case 0
                  If mIsDocTypeLoaded = False Then LoadDocType
                  If mIsPoheaderCombosLoaded = False Then
                    CheckErrors = LoadPoHeaderCombos
                    'FillUPCOMBOS
                    'Disabling the PO so that the user can not Navigate in Create Mode
                    ssOleDbPO.Enabled = False
                  End If
                  
           Me.MousePointer = vbArrow
       End Select
             
             'If FormMode = mdVisualization Then
                       
                       NavBar1.NewEnabled = False
                       NavBar1.NextEnabled = False
                       NavBar1.PreviousEnabled = False
                       NavBar1.LastEnabled = False
                       NavBar1.FirstEnabled = False
                       NavBar1.SaveEnabled = True
                       NavBar1.EditEnabled = False
                       NavBar1.CancelEnabled = True
                       
              'End If
     
             
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
        Case 1
        Case 2
          If Not FormMode = mdVisualization Then
            If CheckLIFields = True Then
                       SaveToPOITEM
                    Else
                       Exit Sub
                    End If
          End If
               If PoItem.MoveFirst Then LoadFromPOITEM
          
        Case 3
            If PORemark.MoveFirst Then LoadFromPORemarks
        Case 4
            If POClause.MoveFirst Then LoadFromPOClause
 End Select
End Sub

Private Sub NavBar1_OnLastClick()
 
 Select Case sst_PO.Tab
        Case 0
            If Poheader.MoveLast Then LoadFromPOHEADER
        Case 1
        Case 2
            
               If FormMode <> mdVisualization Then
                    If CheckLIFields = True Then
                       SaveToPOITEM
                    Else
                       Exit Sub
                    End If
                    
                End If
                If PoItem.MoveLast Then LoadFromPOITEM
            
        Case 3
            If PORemark.MoveLast Then LoadFromPORemarks
        Case 4
            If POClause.MoveLast Then LoadFromPOClause
 End Select

 
End Sub

Private Sub NavBar1_OnNewClick()
  
  Select Case (sst_PO.Tab)
   
   Case 0
        ssOleDbPO.SetFocus
        FormMode = ChangeMode(mdCreation)
          If mIsDocTypeLoaded = False Then
              SSOleDBDocType.text = ""
              LoadDocType
           End If
         If mIsPoheaderCombosLoaded = False Then
           CheckErrors = LoadPoHeaderCombos
         End If
        
        If mIsPoheaderCombosLoaded = True Then
           CheckErrors = Poheader.AddNew
           If CheckErrors = False Then
           MsgBox "Can't Add a PO"
           Else
           SetInitialVAluesPoHeader
            ToggleNavButtons (mdCreation)
           'Disabling the PO so that the user can not Navigate in Create Mode
            'ssOleDbPO = False
           End If
        End If
  
  Case 1
        
        
        
  Case 2
         
        If mIsPoItemsComboLoaded = False Then
            CheckErrors = LoadPoItemCombos
        End If
      
         If mIsPoItemsComboLoaded = True Then
           
            If PoItem.Count > 0 Then SaveToPOITEM
          ' CheckErrors = PoItem.Update
           CheckErrors = PoItem.AddNew
           
           
           If CheckErrors = False Then
              MsgBox "Error In Adding A POITEM"
           Else
              ClearAllPoLineItems
              SetInitialVAluesPOITEM
           End If
           
        End If
         
  Case 3
        If PORemark.Count > 0 Then savetoPORemarks
        
        CheckErrors = PORemark.AddNew
        If CheckErrors = False Then
             MsgBox "Error In Adding A Remarks"
        Else
             ClearPoRemarks
        End If
  Case 4
  
         If POClause.Count > 0 Then savetoPOclause
          
        CheckErrors = POClause.AddNew
        If CheckErrors = False Then
             MsgBox "Error In Adding A Notes/Clause"
        Else
             ClearPoclause
        End If
        
        
 End Select
End Sub

Private Sub NavBar1_OnNextClick()

 Select Case sst_PO.Tab
        Case 0
            
            If Poheader.MoveNext Then LoadFromPOHEADER
        Case 1
        Case 2
             If Not FormMode = mdVisualization Then
              If CheckLIFields = True Then
                       SaveToPOITEM
                    Else
                       Exit Sub
                    End If
              End If
            
               If PoItem.MoveNext Then LoadFromPOITEM
            
        Case 3
            If PORemark.MoveNext Then LoadFromPORemarks
        Case 4
             If POClause.MoveNext Then LoadFromPOClause
 End Select


End Sub

Private Sub NavBar1_OnPreviousClick()


 Select Case sst_PO.Tab
        Case 0
            If Poheader.MovePrevious Then LoadFromPOHEADER
        Case 1
        Case 2
            If Not FormMode = mdVisualization Then
              If CheckLIFields = True Then
                       SaveToPOITEM
                    Else
                       Exit Sub
                    End If
             End If
            If PoItem.MovePrevious Then LoadFromPOITEM
        Case 3
            If PORemark.MovePrevious Then LoadFromPORemarks
        Case 4
            If POClause.MovePrevious Then LoadFromPOClause
 End Select
   

End Sub

Private Sub NavBar1_OnSaveClick()
'''''''Need to Assign Values to PO Class.Create a class for that purpose.
''''''Select Case sst_PO.Tab
''''''        Case 0
''''''            SaveToPOHEADER
''''''            CheckErrors = Poheader.SAVE
''''''            If CheckErrors = False Then
''''''                   MsgBox "Errors Occured"
''''''            Else
''''''                  CheckErrors = LoadFromPOHEADER
''''''                  If CheckErrors = False Then MsgBox "PO saved successfully"
''''''            End If
''''''             CheckErrors = True
''''''        Case 1
''''''        Case 2
''''''           CheckErrors = PoItem.Update
''''''           If CheckErrors = False Then MsgBox "Errors Occurred"
''''''           CheckErrors = True
''''''
''''''        Case 3
''''''            CheckErrors = PORemark.Update
''''''            If CheckErrors = False Then MsgBox "Errors Occurred"
''''''            CheckErrors = True
''''''        Case 4
'''''' End Select
''''''


  
End Sub

Private Sub opt_Email_Click()
On Error GoTo handler
Dim co As MSDataGridLib.column

If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
'''''''''   'Set dgRecepients.DataSource = Nothing
    Set co = dgRecepients.Columns(1)
'''''''''
'''''''''    'Modified by Juan (8/28/2000) for Multilingual
'''''''''    'msg1 = translator.Trans("L00121") 'J added
   ' co.Caption = IIf(msg1 = "", "Email Address", msg1) 'J modified

    co.Caption = "Email Address"
'''''''''    '---------------------------------------------
'''''''''
    co.DataField = "phd_mail"
'''''''''    dgRecepients.Columns(1).Caption = "Email Address"
'''''''''    dgRecepients.Columns(1).DataField = "phd_mail"

    dgRecepients.Columns(0).DataField = "phd_name"
    'Set dgRecepients.DataSource = GetAddresses("ATEMAIL")

     Set dgRecepients.DataSource = Lookups.GetAddresses("ATEMAIL")
     
     
''''''     On Error GoTo Handler
''''''If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
''''''
''''''    Set co = dgRecepients.Columns(1)
''''''
''''''    'Modified by Juan (8/28/2000) for Multilingual
''''''    'msg1 = translator.Trans("L00122") 'J added
''''''    'co.Caption = IIf(msg1 = "", "Fax Number", "") 'J modified
''''''    '---------------------------------------------
''''''
''''''    co.DataField = "phd_faxnumb"
''''''
''''''    dgRecepients.Columns(0).DataField = "phd_name"
''''''
'''''''    Set dgRecepients.DataSource = GetAddresses(deIms.Namespace, deIms.cnIms, adLockReadOnly, atFax)
''''''     Set dgRecepients.DataSource = Lookups.GetAddresses("ATFAX")

   Exit Sub
        
     
handler:
     Err.Raise Err.number, , Err.Description
     Err.Clear
     
End Sub

Private Sub opt_FaxNum_Click()
Dim co As MSDataGridLib.column

On Error GoTo handler
If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
    'msg1 = translator.Trans("L00122") 'J added
    'co.Caption = IIf(msg1 = "", "Fax Number", "") 'J modified
    '---------------------------------------------
    
    co.DataField = "phd_faxnumb"
    
    dgRecepients.Columns(0).DataField = "phd_name"
     
'    Set dgRecepients.DataSource = GetAddresses(deIms.Namespace, deIms.cnIms, adLockReadOnly, atFax)
     Set dgRecepients.DataSource = Lookups.GetAddresses("ATFAX")

Exit Sub
handler:

  Err.Raise Err.number, , Err.Description
  Err.Clear
End Sub

Private Sub opt_SupFax_Click()
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim co As MSDataGridLib.column
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
   ' msg1 = translator.Trans("L00124") 'J added
    'co.Caption = IIf(msg1, "Supplier Email", msg1) 'J modified
    '---------------------------------------------
    co.Caption = "Supplier Email"
    co.DataField = "sup_mail"
    
    dgRecepients.Columns(0).DataField = "sup_name"
    
    Set Rs = New ADODB.Recordset
    
    With Rs
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        Set .ActiveConnection = deIms.cnIms
        .Open ("select sup_name, sup_mail from SUPPLIER where sup_npecode = '" & deIms.NameSpace & "' and sup_mail IS NOT NULL and len(sup_mail) > 3 order by 1")
        Set dgRecepients.DataSource = .DataSource
    End With
End Sub

''''Private Sub ssdcboCategoryCode_Click()
''''If CheckIfCombosLoaded = False Then FillUPCOMBOS
''''End Sub

Private Sub ssdcboCommoditty_Click()

  If Len(Trim$(ssdcboCommoditty.text)) > 0 Then
    If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
     Dim rsMANUFACTURER As ADODB.Recordset
    
        
        Set rsMANUFACTURER = Lookups.GetManuFActurer(Trim$(ssdcboCommoditty))
        
        If Not rsMANUFACTURER.EOF Then
        
            Do While Not rsMANUFACTURER.EOF
                  ssdcboManNumber.AddItem rsMANUFACTURER!stm_manucode & ";" & rsMANUFACTURER!stm_partnumb & ";" & rsMANUFACTURER!stm_estmpric
                  rsMANUFACTURER.MoveNext
            Loop
            
        End If
           
        Set rsMANUFACTURER = Nothing
        'This is a Global Variable which Stores the info about the Stock number.
        'Can use the Public Type "StockDesc" instead.
        
        If objUnits Is Nothing Then Set objUnits = MainPO.PoUnits
        'Set RsUNits = Lookups.GetUnitForTheStckNo(Trim$(ssdcboCommoditty))
        objUnits.StockNumber = Trim$(ssdcboCommoditty.text)
        'If Not RsUNits.EOF Then
            'Do While Not RsUNits.EOF
                  SSOleDBUnit.RemoveAll
                  
                  'NON-STOCK.Append a "N".In SaveToPOitem ,we checkif it is in-stock or non-stock.
                  
                       'SSOleDBUnit.AddItem RsUNits!stk_primuon & ";" & "N"
                       'SSOleDBUnit.AddItem RsUNits!stk_secouom & ";" & "N"
                       
                       If RsUNits Is Nothing Then
                          If Lookups Is Nothing Then Lookups = MainPO.Lookups
                          Set RsUNits = Lookups.GetAllUnits
                        End If
                        
                        RsUNits.MoveFirst
                        RsUNits.Find ("uni_code='" & Trim$(objUnits.PrimaryUnit) & "'")
                        SSOleDBUnit.AddItem objUnits.PrimaryUnit & ";" & RsUNits("uni_desc")
                        
                        RsUNits.MoveFirst
                        RsUNits.Find ("uni_code='" & Trim$(objUnits.SecondaryUnit) & "'")
                        SSOleDBUnit.AddItem objUnits.SecondaryUnit & ";" & RsUNits("uni_desc")
                        SSOleDBUnit = ""
                       
                       'SSOleDBUnit.AddItem objUnits.PrimaryUnit
                       'SSOleDBUnit.AddItem objUnits.SECONDARYUNIT
                  
                       txt_Descript = objUnits.Description
                      
                      
                   '    SSOleDBUnit.Columns(1).Visible = False
                  
                  
            'Loop
        'End If
        
       ' Set RsUNits = Nothing
        Set Lookups = Nothing
        
        
  End If
End Sub

Private Sub ssdcboManNumber_Change()
If Not PoItem.editmode <> 2 Or PoItem.editmode <> -1 And mLoadMode = NoLoadInProgress Then PoItem.Manupartnumb = ssdcboManNumber.text
End Sub


Private Sub ssdcboRequisition_Click()
Dim RsReqPO As ADODB.Recordset
'Dim Lookups As IMSPODLL.Lookups
Set Lookups = MainPO.Lookups
lblReqLineitem = Trim$(ssdcboRequisition.Columns(2).text)
Set RsReqPO = Lookups.GetReqisitionLineItem(lblReqLineitem, Trim$(ssdcboRequisition.text))
If RsReqPO.RecordCount > 0 Then Call LoadPOLINEFromRequsition(RsReqPO)
ssdcboRequisition.SelLength = 0
ssdcboRequisition.SelStart = 0
End Sub

Private Sub ssdcboShipper_Click()
If CheckIfCombosLoaded = False Then FillUPCOMBOS
End Sub

Private Sub SSOleDBCompany_Click()
 SSOleDBInvLocation = ""
 SSOleDBInvLocation.RemoveAll
 If CheckIfCombosLoaded = False Then FillUPCOMBOS

    Dim Value As String

   ' Value = Trim(SSOleDBCompany.Columns(0).Text)
    deIms.rsCompanyLocations.Filter = ""
    deIms.rsCompanyLocations.Filter = "loc_compcode='" & Trim$(SSOleDBCompany.Columns(0).text) & "'"

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
SSOleDBCompany.SelStart = 0
SSOleDBCompany.SelLength = 0
End Sub

Private Sub SSOleDBCurrency_Click()
If CheckIfCombosLoaded = False Then FillUPCOMBOS
SSOleDBCurrency.SelLength = 0
SSOleDBCurrency.SelStart = 0
End Sub

Private Sub SSOleDBCustCategory_Change()
'If PoItem.editmode <> 2 And PoItem.editmode <> "-1" And mLoadMode = NoLoadInProgress Then PoItem.Custcate = Trim$(SSOleDBCustCategory.text)
End Sub

Private Sub Form_Load()
FNamespace = deIms.NameSpace
Dim mLoadForm As Boolean
Dim x As String
Dim Count As Integer

NavBar1.EditEnabled = True
mDidUserOpenStkMasterForm = False
MainPO.Configure deIms.NameSpace, deIms.cnIms

Set Poheader = MainPO.Poheader

   InitializePOheaderRecordset
  'LoadPoHeaderCombos
  If Poheader.EOF = False Then LoadFromPOHEADER
       PoReceipeintsInit
      
sst_PO.Tab = 0
 
 mCheckLIFields = True
 mCheckPoFields = True
 Call DisableButtons(frm_NewPurchase, NavBar1)
'''''    NavBar1.EditEnabled = True
'''''    NavBar1.EditVisible = True
'''''    NavBar1.CloseEnabled = True
'''''    NavBar1.CloseVisible = True
'''''    NavBar1.NewEnabled = True
    NavBar1.PreviousEnabled = True
    NavBar1.LastEnabled = True
    NavBar1.FirstEnabled = True
    NavBar1.NextEnabled = True
    FormMode = ChangeMode(mdVisualization)
    If NavBar1.EditEnabled = False Then
         NavBar1.EditVisible = False
    Else
         NavBar1.EditVisible = True
    End If
    
    If NavBar1.NewEnabled = False Then
           NavBar1.NewVisible = False
    Else
            NavBar1.NewVisible = True
    End If
End Sub


Public Function LoadFromPOHEADER() As Boolean
Dim RsStatus As New ADODB.Recordset
Dim rsDOCTYPE As New ADODB.Recordset

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
On Error GoTo handler


'LblRevDate = Poheader.re


'deIms.rsDOCTYPE.MoveFirst
'rsDOCTYPE.Find "doc_code='" & Poheader.docutype & "'"
'SSOleDBDocType.Text = rsDOCTYPE!doc_desc
rsDOCTYPE.MoveFirst
rsDOCTYPE.Find "doc_code='" & Poheader.docutype & "'"

'SSOleDBDocType.Columns(0).Text = rsDOCTYPE!doc_code
SSOleDBDocType.Columns(0).text = Poheader.docutype

'If Not Len(SSOleDBDocType.Text) > 0 Then
    SSOleDBDocType.text = rsDOCTYPE!doc_desc
    ssOleDbPO = Poheader.Ponumb
    LblRevNumb = Poheader.revinumb
    LblRevDate = Format(Poheader.daterevi, "MM/DD/YY")
'Else
 '  SSOleDBDocType = ""
'End If

ssdcboShipper.Columns(0).text = Poheader.shipcode
If Len(Poheader.shipcode) > 0 Then
    deIms.rsSHIPPER.MoveFirst
    deIms.rsSHIPPER.Find ("shi_code='" & Poheader.shipcode & "'")
    ssdcboShipper.text = deIms.rsSHIPPER!shi_name
Else
    ssdcboShipper.text = ""
End If

txt_ChargeTo = Poheader.chrgto

SSOleDBPriority.Columns(0).text = Poheader.priocode
deIms.rsPRIORITY.MoveFirst
deIms.rsPRIORITY.Find ("pri_code='" & Poheader.priocode & "'")
SSOleDBPriority.text = deIms.rsPRIORITY!pri_desc

txt_Buyer = Poheader.buyr

SSOleDBOriginator = Poheader.orig

LblAppBy = Poheader.apprby
'ssdcboCategoryCode = Poheader.catecode

SSOleDBToBeUsedFor.Columns(0).text = Poheader.tbuf
'deims.rsTOBEUSEDFOR.Find (tbu_name
'SSOleDBToBeUsedFor.Text =

SSoledbSupplier.Columns(0).text = Poheader.suppcode

deIms.rsActiveSupplier.MoveFirst
deIms.rsActiveSupplier.Find ("sup_code='" & Poheader.suppcode & "'")
SSoledbSupplier.text = deIms.rsActiveSupplier!sup_name

Txt_supContaName = Poheader.SuppContactName
Txt_supContaPh = Poheader.SuppContaPH


SSOleDBCurrency.Columns(0).text = Poheader.Currcode
deIms.rsCURRENCY.MoveFirst
deIms.rsCURRENCY.Find ("curr_code='" & Poheader.Currcode & "'")
SSOleDBCurrency.text = deIms.rsCURRENCY!curr_desc

LblCompanyCode.Caption = Poheader.CompCode
deIms.rsActiveCompany.MoveFirst
deIms.rsActiveCompany.Find ("com_compcode='" & Poheader.CompCode & "'")
SSOleDBCompany.text = deIms.rsActiveCompany!com_name

'SSOleDBInvLocation = Poheader.invloca
If Len(Trim$(Poheader.invloca)) > 0 Then
    deIms.rsCompanyLocations.MoveFirst
    deIms.rsCompanyLocations.Find ("loc_locacode='" & Poheader.invloca & "'")
    If Not deIms.rsCompanyLocations.EOF Then
    SSOleDBInvLocation.Tag = Poheader.invloca
    SSOleDBInvLocation.text = deIms.rsCompanyLocations!loc_name
    Else
    SSOleDBInvLocation.text = Poheader.invloca
        SSOleDBInvLocation.Tag = Poheader.invloca
    End If
Else
 SSOleDBInvLocation = ""
End If


chk_ConfirmingOrder = IIf(Poheader.confordr = True, 1, 0)

'ssdcboCondition = Poheader.taccode
If Len(Poheader.taccode) > 0 Then
deIms.rsTermCondition.MoveFirst
'deIms.rsTermCondition.Find ("cond_condcode='" & Poheader.taccode & "'")
deIms.rsTermCondition.Find ("tac_taccode='" & Poheader.taccode & "'")
ssdcboCondition.Tag = Poheader.taccode
ssdcboCondition.text = deIms.rsTermCondition!tac_desc
Else
 ssdcboCondition.text = ""
End If

ssdcboDelivery.Columns(0).text = Poheader.termcode
deIms.rsTermDelivery.MoveFirst
deIms.rsTermDelivery.Find ("tod_termcode='" & Poheader.termcode & "'")
ssdcboDelivery.text = deIms.rsTermDelivery!tod_desc
ssdcboDelivery.Tag = Poheader.termcode


SSOleDBShipTo.Columns(0).text = Poheader.shipto

If Len(Poheader.shipto) > 0 Then
  If deIms.rsActiveShipTo.State = 0 Then Call deIms.ActiveShipTo(deIms.NameSpace)
  deIms.rsActiveShipTo.MoveFirst
    deIms.rsActiveShipTo.Find ("sht_code='" & Poheader.shipto & "'")
    SSOleDBShipTo.text = deIms.rsActiveShipTo!sht_name
    SSOleDBShipTo.Tag = Poheader.shipto
Else
    SSOleDBShipTo.text = ""
End If

chk_FrmStkMst = IIf(Poheader.fromstckmast = True, 1, 0)
txtSite = Poheader.site



dtpRequestedDate = Poheader.reqddelvdate
LblDateSent = Format(Poheader.datesent, "mm/dd/yy")
DTPicker_poDate = Format(Poheader.Createdate, "mm/dd/yy")


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
handler:
   MsgBox Err.Description
   Err.Clear
   mLoadMode = NoLoadInProgress
End Function
Public Function LoadFromPOITEM() As Boolean

On Error GoTo handler
LoadFromPOITEM = False

Dim RsStatus As New ADODB.Recordset

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


dcbostatus(0).text = Stasliit
dcbostatus(1) = Stasdlvy
dcbostatus(2) = Stasship
dcbostatus(3) = StasINvt



LblPOi_PONUMB = PoItem.Ponumb
txt_LI = PoItem.Linenumb
txt_TotalLIs = PoItem.Count
ssdcboCommoditty.text = PoItem.Comm
ssdcboManNumber = PoItem.Manupartnumb
txt_AFE = PoItem.Afe
SSOleDBCustCategory = PoItem.Custcate
txt_SerialNum = PoItem.Serlnumb
txt_Requested = FormatNumber$(PoItem.Primreqdqty, 2)
'txtSecRequested = PoItem.Secoreqdqty

'SSOleDBSecUnit = PoItem.Secouom  'Row Member - SECONDARYUNIT,ListField - uni_desc , BoundColumns - uni_code
txt_Delivered = FormatNumber$(PoItem.PriQtydlvd)
txt_Shipped = FormatNumber$(PoItem.PriQtyship)
txt_Inventory2 = FormatNumber$(PoItem.PriQtyinvt)

txt_Descript = PoItem.Description
txt_remk = PoItem.remk

 
'If the PO was created in Primary mode
If Len(PoItem.UnitOfPurch) = 0 Or Trim$(PoItem.UnitOfPurch) = "P" Then
   
    txt_Requested = FormatNumber$(PoItem.Primreqdqty, 2)
    SSOleDBUnit = PoItem.Primuom     'w MemBer -GET_UNIT_OF_MEASURE ,ListField-uni_desc,BoundColumns-Uni_code
    txt_Price = FormatNumber$(PoItem.PrimUnitprice)
    txt_Total = FormatNumber$(PoItem.PriTotaprice)
    
'else If the PO was created in Secondary mode
ElseIf Trim$(PoItem.UnitOfPurch) = "S" Then

    txt_Requested = FormatNumber$(PoItem.Secoreqdqty, 2)
    SSOleDBUnit = PoItem.Secouom      'w MemBer -GET_UNIT_OF_MEASURE ,ListField-uni_desc,BoundColumns-Uni_code
    txt_Price = FormatNumber$(PoItem.SecUnitPrice)
    txt_Total = FormatNumber$(PoItem.SecTotaprice)
    
End If

'ssdcboRequisition =
lblReqLineitem = PoItem.Requliitnumb
'LblPOI_Doctype = PoItem.
 LoadFromPOITEM = True
 mLoadMode = NoLoadInProgress
 Exit Function
handler:
  Err.Clear
  mLoadMode = NoLoadInProgress
End Function

Public Function LoadPoHeaderCombos() As Boolean

Dim FNamespace As String * 5

'Dim RsBRQ As ADODB.Recordset

'Dim defsite As String

On Error GoTo handler
  
   LoadPoHeaderCombos = False
   
   FNamespace = deIms.NameSpace


  '  Set RsBRQ = GetRequisitions(Fnamespace, deIms.cnIms)

    'Call deIms.PONUMB(fnamespace)
    
''''''    Set rsDOCTYPE = GetDocumentType(False)
''''''    Call deIms.Shipper(fnamespace)
''''''    Call deIms.Currency(fnamespace)
''''''    Call deIms.Priority(fnamespace)
''''''    Call deIms.TermDelivery(fnamespace)
''''''    'Call deIms.Supplier(fnamespace)
''''''    Call deIms.ActiveSupplier(fnamespace)
''''''    Call deIms.TermCondition(fnamespace)
''''''    'Call deIms.INVENTORYLOCATION(Fnamespace, ponumb)
''''''    Call deIms.Company(fnamespace)
''''''    Call deIms.GETSYSSITE(fnamespace, defsite)
''''''    Call deIms.ActiveOriginator(fnamespace)
''''''    Call deIms.ActiveTbu(fnamespace)
''''''    Call deIms.servcodecat(fnamespace)
''''''    Call deIms.ActiveSHIPTO(fnamespace)
''''''    Call deIms.ActiveCompany(fnamespace)
''''''    Call deIms.CompanyLocations(fnamespace)
    If Not mIsPoHeaderRsetsInit = True Then InitializePOheaderRecordset
    LoadPoHeaderCombos = FillUPCOMBOS
    mIsPoheaderCombosLoaded = LoadPoHeaderCombos



''''''    Call FillFromREcordset(SSOleDBDocType, rsDOCTYPE, "Doc_Code")
''''''    Call FillFromREcordset(ssOleDbPO, deIms.rsPOnumb, "po_ponumb")
''''''    Call FillFromREcordset(ssdcboDelivery, deIms.rsTermDelivery, "tod_desc")
''''''    Call FillFromREcordset(ssdcboCondition, deIms.rsTermCondition, "tac_desc")
''''''    Call FillFromREcordset(ssdcboCategoryCode, deIms.rsSERVCODECAT, "scs_desc")
''''''    'Call FillFromREcordset(ssdcboRequisition, RsBRQ, "")
''''''    Call FillFromREcordset(SSOleDBCurrency, deIms.rsCURRENCY, "curr_desc")
''''''    Call FillFromREcordset(SSOleDBPriority, deIms.rsPRIORITY, "pri_desc")
''''''    Call FillFromREcordset(SSoledbSupplier, deIms.rsSUPPLIER, "sup_desc")
''''''    Call FillFromREcordset(SSOleDBInvLocation, deIms.rsINVENTORYLOCATION, "loc_name")
''''''    Call FillFromREcordset(ssdcboShipper, deIms.rsSHIPPER, "shi_name")
''''''    Call FillFromREcordset(SSOleDBOriginator, deIms.rsActiveOriginator, "ori_code")
''''''    Call FillFromREcordset(SSOleDBToBeUsedFor, deIms.rsActiveTbu, "tbu_name")
''''''    Call FillFromREcordset(SSOleDBShipTo, deIms.rsActiveShipTo, "sht_name")


    'ssoledbCustomCategory

    'POITEM TAb
    'Call deims.UNIT(FNameSpace)
    'Call deIms.Custom(FNamespace)
    'Not NEEDED
    'Call deIms.POSTATUS(FNamespace)
    'Call GetUnits("")
    'Call GetActiveStockNumbers(False)
    'Set ssdcboCommoditty.DataSourceList = deIms
    
    'Set rsPOREM = deIms.rsPOREM
    'Set rsPOITEM = deIms.rsPOITEM
    'Set rsrecepList = deIms.rsPOREC
    'Set rsPOCLAUSE = deIms.rsPOCLAUSE

    'Call ChangeMode(mdVisualization)
    
    
    sst_PO.Tab = 0





    NavBar1.NextEnabled = sst_PO.Tab <> 0
    NavBar1.LastEnabled = sst_PO.Tab <> 0
    NavBar1.FirstEnabled = sst_PO.Tab <> 0
    NavBar1.PreviousEnabled = sst_PO.Tab <> 0
    
    Exit Function
    
handler:
    MsgBox Err.Description
    Err.Clear
End Function


''''''Private Function GetDocumentType(All As Boolean) As ADODB.Recordset
''''''Dim x As Integer
''''''On Error GoTo handler
''''''
''''''
''''''    With deims
''''''
''''''
''''''        If All Then
''''''            Set GetDocumentType = .UserDocumentType("", x)
''''''        Else
''''''            Set GetDocumentType = .UserDocumentType(CurrentUser, x)
''''''        End If
''''''
''''''    End With
''''''
''''''     Exit Function
''''''handler:
''''''
''''''   Err.Clear
''''''End Function


''''Public Sub FillFromREcordset(Control As Control, Rs As adodb.Recordset, FieldToDisplay As String)
''''Dim FieldsARR()
''''Dim Field As adodb.Field
''''Dim count As Integer
''''Dim RowOfData As String
''''Dim I As Integer
''''
''''  count = 0
''''
''''  For Each Field In Rs.Fields
''''
''''    ReDim FieldsARR(count)
''''    FieldsARR(count) = Field.Name
''''    count = count + 1
''''  Next Field
''''
'''''  Control.Cols = count - 1
''''  count = Control.Cols
''''
''''  Do While Not Rs.EOF
''''     ' Control.AddItem Rs & "(" & FieldsARR(0) & ")" & ";" & Rs & "(" & FieldsARR(1) & ")"
''''      Rs.MoveNext
''''  Loop
''''
''''End Sub
Public Function FillUPCOMBOS() As Boolean
 Dim rsSUPPLIER As ADODB.Recordset
 Dim Count As Integer

If Lookups Is Nothing Then Set Lookups = MainPO.Lookups

If Lookups.GetUserMenuLevel(CurrentUser) = 5 Then
 Set rsSUPPLIER = Lookups.GetLocalSuppliers
Else
 Set rsSUPPLIER = deIms.rsActiveSupplier
End If
'   Set rsDOCTYPE = GetDocumentType(False)
 
 FillUPCOMBOS = False
 
 On Error GoTo handler
 
     Set IntiClass = New InitialValuesPOheader
 
   If Not deIms.rsSHIPPER.EOF Then
        deIms.rsSHIPPER.MoveFirst
        IntiClass.InitShipperCode = Trim$(deIms.rsSHIPPER!shi_code)
        IntiClass.InitShipperName = Trim$(deIms.rsSHIPPER!shi_name)
   End If
   
   Do While Not deIms.rsSHIPPER.EOF
       ssdcboShipper.AddItem deIms.rsSHIPPER!shi_code & ";" & deIms.rsSHIPPER!shi_name
       deIms.rsSHIPPER.MoveNext
   Loop
       
   If Not deIms.rsTermDelivery.EOF Then deIms.rsTermDelivery.Filter = "tod_actvflag<>0"
   If Not deIms.rsTermDelivery.EOF Then
        deIms.rsTermDelivery.MoveFirst
        IntiClass.InitDelivery = Trim$(deIms.rsTermDelivery!tod_desc)
   End If
   
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
   
   Do While Not deIms.rsTermCondition.EOF
       ssdcboCondition.AddItem deIms.rsTermCondition!tac_taccode & ";" & deIms.rsTermCondition!tac_desc
       deIms.rsTermCondition.MoveNext
   Loop
   deIms.rsTermCondition.Filter = ""
       
       
'''   If Not deIms.rsSERVCODECAT.EOF Then
'''    deIms.rsSERVCODECAT.MoveFirst
'''    IntiClass.InitCategoryCode = Trim$(deIms.rsSERVCODECAT!scs_desc)
'''   End If
'''
'''   Do While Not deIms.rsSERVCODECAT.EOF
'''       ssdcboCategoryCode.AddItem deIms.rsSERVCODECAT!scs_code & ";" & deIms.rsSERVCODECAT!scs_desc
'''       deIms.rsSERVCODECAT.MoveNext
'''   Loop
   
   If Not deIms.rsCURRENCY.EOF Then
     deIms.rsCURRENCY.MoveFirst
     IntiClass.InitCurrency = Trim$(deIms.rsCURRENCY!curr_desc)
   End If
   Do While Not deIms.rsCURRENCY.EOF
       SSOleDBCurrency.AddItem deIms.rsCURRENCY!curr_code & ";" & deIms.rsCURRENCY!curr_desc
       deIms.rsCURRENCY.MoveNext
   Loop
   
   If Not deIms.rsPRIORITY.EOF Then deIms.rsPRIORITY.Filter = "pri_actvflag <>0"
   If Not deIms.rsPRIORITY.EOF Then
     
     deIms.rsPRIORITY.MoveFirst
     IntiClass.InitPriority = Trim$(deIms.rsPRIORITY!pri_desc)
   End If
   Do While Not deIms.rsPRIORITY.EOF
       SSOleDBPriority.AddItem deIms.rsPRIORITY!pri_code & ";" & deIms.rsPRIORITY!pri_desc
       deIms.rsPRIORITY.MoveNext
   Loop
   deIms.rsPRIORITY.Filter = ""
   
   
''''''   If Not deIms.rsSUPPLIER.EOF Then
''''''      deIms.rsSUPPLIER.MoveFirst
''''''      IntiClass.InitSupplier = Trim$(deIms.rsSUPPLIER!sup_name)
''''''   End If
''''''   Do While Not deIms.rsSUPPLIER.EOF
''''''       SSoledbSupplier.AddItem deIms.rsSUPPLIER!sup_code & ";" & deIms.rsSUPPLIER!sup_name & ";" & deIms.rsSUPPLIER!sup_city & ";" & deIms.rsSUPPLIER!sup_phonnumb
''''''       deIms.rsSUPPLIER.MoveNext
''''''   Loop
   
   
   If Not rsSUPPLIER.EOF Then
      rsSUPPLIER.MoveFirst
      IntiClass.InitSupplier = Trim$(rsSUPPLIER!sup_name)
   End If
   Do While Not rsSUPPLIER.EOF
       SSoledbSupplier.AddItem rsSUPPLIER!sup_code & ";" & rsSUPPLIER!sup_name & ";" & rsSUPPLIER!sup_city & ";" & rsSUPPLIER!sup_phonnumb
       rsSUPPLIER.MoveNext
   Loop
   
   
   
   'LOad The Company FIRST
'''''   Do While Not deIms.rsINVENTORYLOCATION.EOF
'''''       SSOleDBInvLocation.AddItem deIms.rsINVENTORYLOCATION & ";" & deIms.rsINVENTORYLOCATION!shi_name
'''''       deIms.rsINVENTORYLOCATION.MoveNext
'''''   Loop
'''''
   
   If Not deIms.rsActiveOriginator.EOF Then
     deIms.rsActiveOriginator.MoveFirst
     IntiClass.InitOriginator = Trim$(deIms.rsActiveOriginator!ori_code)
   End If
   Do While Not deIms.rsActiveOriginator.EOF
       SSOleDBOriginator.AddItem deIms.rsActiveOriginator!ori_code '& ";" & deIms.rsActiveOriginator!ori_code
       deIms.rsActiveOriginator.MoveNext
   Loop
   
   If Not deIms.rsActiveTbu.EOF Then
      deIms.rsActiveTbu.MoveFirst
      IntiClass.InitToBeUsedFor = Trim$(deIms.rsActiveTbu!tbu_name)
   End If
   Do While Not deIms.rsActiveTbu.EOF
       SSOleDBToBeUsedFor.AddItem deIms.rsActiveTbu!tbu_name '& ";" & deIms.rsActiveOriginator!tbu_name
       deIms.rsActiveTbu.MoveNext
   Loop
   
   If Not deIms.rsActiveShipTo.EOF Then
       deIms.rsActiveShipTo.MoveFirst
       IntiClass.InitShipTo = Trim$(deIms.rsActiveShipTo!sht_name)
   End If
   Do While Not deIms.rsActiveShipTo.EOF
       SSOleDBShipTo.AddItem deIms.rsActiveShipTo!sht_code & ";" & deIms.rsActiveShipTo!sht_name
       deIms.rsActiveShipTo.MoveNext
   Loop
   
   If Not deIms.rsActiveCompany.EOF Then
       deIms.rsActiveCompany.MoveFirst
       IntiClass.InitCompanyCode = Trim$(deIms.rsActiveCompany!com_compcode)
       IntiClass.InitCompanyName = Trim$(deIms.rsActiveCompany!com_name)
   End If
   Do While Not deIms.rsActiveCompany.EOF
       SSOleDBCompany.AddItem deIms.rsActiveCompany!com_compcode & ";" & deIms.rsActiveCompany!com_name
       deIms.rsActiveCompany.MoveNext
       
   Loop
   
   IntiClass.InitpoDate = Format(Now(), "mm/dd/yy")
   IntiClass.InitBuyer = CurrentUser
''   Do While Not deIms.rsActiveOriginator.EOF
''       SSOleDBOriginator.AddItem deIms.rsActiveOriginator!ori_code '& ";" & deIms.rsActiveOriginator!ori_code
''       deIms.rsActiveOriginator.MoveNext
''   Loop

''''''''    If Not rsDOCTYPE Is Nothing Then
''''''''            If rsDOCTYPE.RecordCount > 0 Then
''''''''                  SSOleDBDocType.Enabled = True
''''''''                  rsDOCTYPE.MoveFirst
''''''''                  Do While Not rsDOCTYPE.EOF
''''''''                      SSOleDBDocType.AddItem rsDOCTYPE!doc_code & ";" & rsDOCTYPE!doc_desc
''''''''                      rsDOCTYPE.MoveNext
''''''''                  Loop
''''''''             Else
''''''''             '   SSOleDBDocType.Enabled = False
''''''''
''''''''            End If
''''''''
''''''''    Else
''''''''            '  SSOleDBDocType.Enabled = False
''''''''    End If
    Set rsSUPPLIER = Nothing
    CheckIfCombosLoaded = True
   FillUPCOMBOS = True
   Exit Function
handler:
   
       
    Err.Clear
End Function

Private Sub SSOleDBInvLocation_DropDown()

'If Not mIsInvLocationLoaded = False Then
        SSOleDBInvLocation = ""
         SSOleDBInvLocation.RemoveAll
         If CheckIfCombosLoaded = False Then FillUPCOMBOS
        
            Dim Value As String
        '
        LblCompanyCode.Caption = Trim$(SSOleDBCompany.Columns(0).text)
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

    Dim Value As String
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
End Sub

Private Sub SSOleDBPO_Click()
Poheader.Move Trim$(ssOleDbPO)
Call LoadFromPOHEADER
End Sub

Public Function SaveToPOHEADER() As Boolean
On Error GoTo handler
SaveToPOHEADER = False

'LblRevDate = Poheader.re
Poheader.Npecode = deIms.NameSpace
 Poheader.docutype = SSOleDBDocType.Columns(0).text
 Poheader.Ponumb = ssOleDbPO
 Poheader.revinumb = LblRevNumb
 If Len(LblRevDate.Caption) > 0 Then Poheader.daterevi = LblRevDate
 Poheader.shipcode = Trim$(ssdcboShipper.Columns(0).text)
 Poheader.chrgto = txt_ChargeTo
 Poheader.priocode = Trim$(SSOleDBPriority.Columns(0).text)
 Poheader.buyr = txt_Buyer
 Poheader.orig = SSOleDBOriginator
 'Poheader.apprby = LblAppBy
' Poheader.catecode = Trim$(ssdcboCategoryCode.Columns(0).Text)
  Poheader.tbuf = SSOleDBToBeUsedFor
  Poheader.suppcode = Trim$(SSoledbSupplier.Columns(0).text)
  Poheader.SuppContactName = Txt_supContaName
  Poheader.SuppContaPH = Txt_supContaPh
  Poheader.Currcode = Trim$(SSOleDBCurrency.Columns(0).text)
  Poheader.CompCode = Trim$(SSOleDBCompany.Columns(0).text)
  'Poheader.invloca = Trim$(SSOleDBInvLocation.Columns(0).text)
   Poheader.invloca = Trim$(SSOleDBInvLocation.Tag)
  Poheader.confordr = IIf(chk_ConfirmingOrder = 1, True, False)
  Poheader.taccode = Trim$(ssdcboCondition.Columns(0).text)
  Poheader.termcode = Trim$(ssdcboDelivery.Columns(0).text)
  Poheader.fromstckmast = chk_FrmStkMst
  Poheader.site = txtSite
  Poheader.shipto = Trim$(SSOleDBShipTo.Columns(0).text)
  Poheader.reqddelvdate = dtpRequestedDate
  Poheader.datesent = LblDateSent
  Poheader.Createdate = DTPicker_poDate
  'Poheader.Stasinvt = dcbostatus(7)
  'Poheader.Stasship = dcbostatus(6)
  'Poheader.stasdelv = dcbostatus(5)
  'Poheader.stas = dcbostatus(4)
  Poheader.forwr = chk_Forwarder
  Poheader.freigforwr = chk_FreightFard
  Poheader.reqddelvflag = chk_Requ
  
  'If this is a POREVISION
  If mSaveToPoRevision = True Then
     Poheader.apprby = ""
     Poheader.stas = "OH"
  Else
     Poheader.apprby = LblAppBy
  End If
  
  
  SaveToPOHEADER = True
  
  Exit Function
handler:
   MsgBox Err.number
   Err.Clear
   
End Function
Private Function GetDocumentType(All As Boolean) As ADODB.Recordset
On Error Resume Next
Dim Rs As ADODB.Recordset
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

Private Sub ssOleDbPO_GotFocus()
'''If mIsPoNumbComboLoaded = False Then
'''  If deIms.rsPonumb.State = 1 Then
'''     deIms.rsPonumb.Close
'''  End If
'''   Call deIms.Ponumb(deIms.NameSpace)
'''    Do While Not deIms.rsPonumb.EOF
'''       SSOleDBPO.AddItem deIms.rsPonumb!PO_PONUMB
'''       deIms.rsPonumb.MoveNext
'''    Loop
'''    mIsPoNumbComboLoaded = True
'''
'''End If
End Sub

Public Function SetInitialVAluesPoHeader()
Dim RsStatus As New ADODB.Recordset



If Lookups Is Nothing Then Set Lookups = MainPO.Lookups

 SetInitialVAluesPoHeader = False
  SSOleDBDocType = IntiClass.InitDocType
  ssOleDbPO = IntiClass.InitPO
  LblRevNumb = 0
  LblRevDate = ""
  'lbl_ApprovedBy = ""
  LblAppBy = ""
  ssdcboShipper.Columns(0).text = IntiClass.InitShipperCode
  ssdcboShipper.text = IntiClass.InitShipperName
  txt_ChargeTo = IntiClass.InitChargeTo
  SSOleDBPriority = IntiClass.InitPriority
  txt_Buyer = IntiClass.InitBuyer
  SSOleDBOriginator = IntiClass.InitOriginator
  LblDateSent.Caption = ""
  'ssdcboCategoryCode = IntiClass.InitCategoryCode
  SSOleDBToBeUsedFor = IntiClass.InitToBeUsedFor
  SSoledbSupplier = IntiClass.InitSupplier
  SSOleDBCurrency = IntiClass.InitCurrency
  LblCompanyCode.Caption = Trim$(IntiClass.InitCompanyCode)
  SSOleDBCompany.text = IntiClass.InitCompanyName
  SSOleDBInvLocation = IntiClass.InitInvLocation

  ssdcboCondition = IntiClass.InitCondition
  ssdcboDelivery = IntiClass.InitDelivery
  
  'txtSite = IntiClass.InitSite
  txtSite = Lookups.GetMYSite
  SSOleDBShipTo = IntiClass.InitShipTo
  dtpRequestedDate = Now() + 1
  DTPicker_poDate = IntiClass.InitpoDate
  SetInitialVAluesPoHeader = True
  
  
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


Private Sub ssOleDbPO_LostFocus()
'''''If FormMode = mdCreation And Len(ssOleDbPO.Text) Then
'''''      If ssOleDbPO.IsItemInList Then
'''''          ssOleDbPO.SetFocus
'''''          MsgBox "PO Number Already Exists"
'''''          ssOleDbPO.Text = ""
'''''
'''''      End If
''''' End If


ssOleDbPO_Validate (False)
End Sub

Private Sub ssOleDbPO_Validate(Cancel As Boolean)
If FormMode = mdCreation And Len(ssOleDbPO.text) > 0 Then
     If deIms.rsPonumb.State = 0 Then Call deIms.Ponumb(deIms.NameSpace)
        deIms.rsPonumb.MoveFirst
        deIms.rsPonumb.Find "PO_PONUMB='" & Trim$(ssOleDbPO.text) & "'"
        If Not deIms.rsPonumb.EOF Then
          ssOleDbPO.SetFocus
          MsgBox "PO Number Already Exists"
          ssOleDbPO.text = ""
          Cancel = True
      End If
 End If
End Sub

Private Sub SSoledbSupplier_Click()

If Len(SSoledbSupplier.text) > 0 Then
    
    deIms.rsActiveSupplier.MoveFirst
    deIms.rsActiveSupplier.Find ("sup_code='" & SSoledbSupplier.Columns(0).text & "'")
    Txt_supContaName.text = IIf(IsNull(deIms.rsActiveSupplier!sup_contaname), "", deIms.rsActiveSupplier!sup_contaname)
    Txt_supContaPh.text = IIf(IsNull(deIms.rsActiveSupplier!sup_contaph), "", deIms.rsActiveSupplier!sup_contaph)
    
    If Not IsNull(deIms.rsActiveSupplier!sup_mail) And Not Len(deIms.rsActiveSupplier!sup_mail) = 0 Then
        AddRecepient (deIms.rsActiveSupplier!sup_mail)
      'dgRecipientList.AddItem Trim$(deIms.rsActiveSupplier!sup_mail)
    Else
       If Not IsNull(deIms.rsActiveSupplier!sup_faxnumb) Then AddRecepient (deIms.rsActiveSupplier!sup_faxnumb)
    End If
    
    SSoledbSupplier.SelLength = 0
    SSoledbSupplier.SelStart = 0
End If

End Sub

Private Sub SSOleDBUnit_Click()
''If Len(Trim$(ssdcboCommoditty.Text)) > 0 Then
     Dim Lookups As IMSPODLL.Lookups
        
'''        If Not RsUNits Is Nothing Then Set RsUNits = Nothing
'''        Set Lookups = MainPO.Lookups
'''
'''        'This is a Global Variable which Stores the info about the Stock number.
'''        'Can use the Public Type "StockDesc" instead.
'''        Set RsUNits = Lookups.GetUnitForTheStckNo(Trim$(ssdcboCommoditty))
'''
'''        If Not RsUNits.EOF Then
'''            'Do While Not RsUNits.EOF
'''                  SSOleDBUnit.RemoveAll
'''
'''                  'NON-STOCK.Append a "N".In SaveToPOitem ,we checkif it is in-stock or non-stock.
'''                  If chk_FrmStkMst = 0 Then
'''                       SSOleDBUnit.AddItem RsUNits!stk_primuon & ";" & "N"
'''                       SSOleDBUnit.AddItem RsUNits!stk_secouom & ";" & "N"
'''                  Else
'''                  'IN-STOCK
'''                       'But Same Units
'''                        If Trim$(RsUNits!stk_primuon) = Trim$(RsUNits!stk_secouom) Then
'''                           SSOleDBUnit.AddItem RsUNits!stk_primuon & ";" & "N"
'''                           SSOleDBUnit.AddItem RsUNits!stk_secouom & ";" & "N"
'''                        Else
'''                        'Different Units
'''                          SSOleDBUnit.AddItem RsUNits!stk_primuon & ";" & "P"
'''                          SSOleDBUnit.AddItem RsUNits!stk_secouom & ";" & "S"
'''                    End If
'''
'''                  End If
'''
'''                  txt_Descript = RsUNits!stk_desc
'''
'''
'''                  SSOleDBUnit.Columns(1).Visible = False
'''
'''
'''            'Loop
'''        End If
'''
'''       ' Set RsUNits = Nothing
'''        Set Lookups = Nothing
'''
'''
'''  End If
End Sub

Private Sub SSOleDBUnit_GotFocus()
''''If Len(ssdcboCommoditty.text) > 0 Then
''''        If objUnits Is Nothing Then Set objUnits = MainPO.PoUnits
''''
''''        objUnits.StockNumber = Trim$(ssdcboCommoditty.text)
''''
''''        SSOleDBUnit.RemoveAll
''''
''''        If RsUNits Is Nothing Then Set RsUNits = Lookups.GetAllUnits
''''
''''        RsUNits.MoveFirst
''''        RsUNits.Find ("uni_code='" & Trim$(objUnits.PrimaryUnit) & "'")
''''        SSOleDBUnit.AddItem objUnits.PrimaryUnit & ";" & RsUNits("uni_desc")
''''
''''        RsUNits.MoveFirst
''''        RsUNits.Find ("uni_code='" & Trim$(objUnits.SECONDARYUNIT) & "'")
''''        SSOleDBUnit.AddItem objUnits.SECONDARYUNIT & ";" & RsUNits("uni_desc")
''''
''''
''''
''''
''''
''''End If
End Sub

Private Sub sst_PO_Click(PreviousTab As Integer)
    On Error Resume Next

Dim editmode(1) As Long, STR As String



   ' If ((fm = mdCreation) Or (fm = mdModification)) Then
        
   '     ValidateControls
        
        Select Case PreviousTab
            
            Case 0
                 If FormMode <> mdVisualization Then
                    If mCheckLIFields = True Then
                        If FormMode <> mdVisualization Then mCheckPoFields = CheckPoFields
                        
                        
                        If FormMode <> mdVisualization And mCheckPoFields = True Then
                              SaveToPOHEADER
                           Else
                             sst_PO.Tab = 0
                           End If
                     End If
                  End If
            
                
                
   '   If Not (CheckPoFields) Then sst_PO.Tab = 0
               
                ' sst_PO.Tab = 0
            Case 1
                
                
                
            Case 2
                
            
               'Save the POItem back to  POITEMs Object
               If FormMode <> mdVisualization Then
                If mCheckPoFields = True Then
                   If PoItem.Count > 0 And FormMode <> mdVisualization Then
                      If CheckLIFields Then
                        SaveToPOITEM
                        Else
                        sst_PO.Tab = 2
                       End If
                    End If
                End If
               End If
               
               
            Case 3
            
            If FormMode <> mdVisualization Then
               txtRemarks.SetFocus
              If mCheckPoFields = True And mCheckLIFields = True Then
               If PORemark.Count > 0 And FormMode <> mdVisualization Then savetoPORemarks
              End If
            End If
            
            Case 4
              '  Call txtClause_Validate(False)
              
              If FormMode <> mdVisualization Then
                  txtClause.SetFocus
                    If mCheckPoFields = True And mCheckLIFields = True Then
                         If POClause.Count > 0 And FormMode <> mdVisualization Then savetoPOclause
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
            
            'Set NavBar1.Recordset = rsPO
   '         Call BindControls("PO")
            
            NavBar1.NextEnabled = True
            NavBar1.LastEnabled = True
            NavBar1.FirstEnabled = True
            NavBar1.PreviousEnabled = True
            
      If FormMode = mdVisualization Then
                    
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
            
            
            
        Case 1
        
             If PoReceipients Is Nothing Then Set PoReceipients = MainPO.PoReceipients
             FirstTimeAssignments
             If Trim$(PoReceipients.Ponumb) <> Poheader.Ponumb Then
            
            GPOnumb = Poheader.Ponumb
                    
                         If PoReceipients.Move(GPOnumb) Then
                          'PORemark.Ponumb = GPOnumb
                    
                        'If Not (PORemark.EOF = True And PORemark.BOF = True) Then
                           LoadFromPOReceipients
                        Else
                          'This means that there are no Line Items Corresponding to This PO
                           ClearPoReceipients
                        End If
              
             End If
             
             
             If FormMode = mdModification Then
                   NavBar1.NewEnabled = True
                   NavBar1.EditEnabled = False
                   NavBar1.CancelEnabled = True
                   
             ElseIf FormMode = mdVisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  
            End If
            
            
            
            
            
            
            'If POExist Then
                'PORECChange
               ' NavBar1.NewEnabled = True
               ' NavBar1.SaveEnabled = True
               ' NavBar1.CancelEnabled = True
    '            Set NavBar1.Recordset = rsrecepList
                
     '           dgRecepients.Enabled = True And Editting
     '           dgRecipientList.Enabled = dgRecepients.Enabled
            'End If
        Case 2
        
            
            
            FirstTimeAssignmentsPOITEM
        
             If PoItem Is Nothing Then Set PoItem = MainPO.POITEMS
            If Trim$(PoItem.Ponumb) <> Poheader.Ponumb Then
                'This is When the User Selects a New PO on POHEADER and Click POITEMS
                'the First Time.
                   'IsThisADifferentPO = True
                    'GPOnumb = ssOleDbPO
                   GPOnumb = Poheader.Ponumb
                    If PoItem.Move(GPOnumb) Then
                        LoadFromPOITEM
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
                  
                  
                If FormMode = mdVisualization Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = False
                    NavBar1.CancelEnabled = False
                    
                 ElseIf FormMode = mdCreation Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = True
                    
                    'This case Arises When the user Clicks on POITEM tab in Creation mode
                    'or when he clicks on it in MODIFICATION mode when there are no line items .
                    If FormMode = mdCreation And PoItem.Count = 0 Then
                       NavBar1_OnNewClick
                    End If
                    txt_AFE = txt_ChargeTo
                 ElseIf FormMode = mdModification Then
                    
                    NavBar1.EditEnabled = False
                    NavBar1.NewEnabled = True
                    NavBar1.CancelEnabled = True
                    
                    If FormMode = mdCreation And PoItem.Count = 0 Then
                       NavBar1_OnNewClick
                    End If
                    txt_AFE = txt_ChargeTo
                 End If
                 
                 NavBar1.SaveEnabled = False
        
            
        Case 3

             If PORemark Is Nothing Then Set PORemark = MainPO.POREMARKS
             
             If PORemark.Ponumb <> Poheader.Ponumb Then
                'This is When the User Selects a New PO on POHEADER and Click POClause
                'the First Time.
         
                    GPOnumb = Poheader.Ponumb
                    
                         If PORemark.Move(GPOnumb) Then
                          'PORemark.Ponumb = GPOnumb
                    
                        'If Not (PORemark.EOF = True And PORemark.BOF = True) Then
                      
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
                   txtRemarks.SetFocus
                   CmdcopyLI(1).Enabled = True
             ElseIf FormMode = mdVisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
                  CmdcopyLI(1).Enabled = False
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  txtRemarks.SetFocus
                  CmdcopyLI(1).Enabled = True
                 If PORemark.Count = 0 Then
                       NavBar1_OnNewClick
                 End If
            End If
            
        Case 4
        
             
             If POClause Is Nothing Then Set POClause = MainPO.POClauses
             
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
                   txtClause.SetFocus
             ElseIf FormMode = mdVisualization Then
             
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = False
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = False
             ElseIf FormMode = mdCreation Then
               
                  NavBar1.EditEnabled = False
                  NavBar1.NewEnabled = True
                  NavBar1.SaveEnabled = False
                  NavBar1.CancelEnabled = True
                  txtClause.SetFocus
                  If POClause.Count = 0 Then
                       NavBar1_OnNewClick
                 End If
                  
            End If
                   
    End Select

End Sub

Public Function SetInitialVAluesPOITEM() As String
txt_LI = PoItem.Count
txt_TotalLIs = PoItem.Count
DTP_Required = Format(Now(), "MM/DD/YY")
End Function

Public Function LoadPoItemCombos() As Boolean
Dim Lookups As Lookups
Dim RSStockNos As ADODB.Recordset
Dim RsRequsition As ADODB.Recordset
'Dim RsUNits As ADODB.Recordset

On Error GoTo handler

mIsPoItemsComboLoaded = False
LoadPoItemCombos = False
Set Lookups = MainPO.Lookups


     'Stockmaster
     If chk_FrmStkMst.Value = 1 And mIsPoItemsComboLoaded = False Then
             
             ssdcboCommoditty.Enabled = True
             SSOleDBUnit.RemoveAll
             
             
             'Set RSStockNos = Lookups.GetStockNUmbers
             'Do While Not RSStockNos.EOF
                        
              '     ssdcboCommoditty.AddItem RSStockNos!stk_stcknumb & ";" & RSStockNos!stk_desc
                '   RSStockNos.MoveNext
               '    DoEvents
              Set ssdcboCommoditty.DataSourceList = deIms.rsActiveStockmasterLookup  ' RSStockNos
              ssdcboCommoditty.DataFieldToDisplay = "stk_stcknumb"
              ssdcboCommoditty.DataFieldList = "stk_desc"
              
             'Loop
             
          mIsPoItemsComboLoaded = True
          'Set RSStockNos = Nothing
        ''Manufacturer
     
     ElseIf chk_FrmStkMst.Value = 0 And mIsPoItemsComboLoaded = False Then
              ssdcboCommoditty = ""
              ssdcboCommoditty.Enabled = False
            ''Load AllUnits
             Set RsUNits = Lookups.GetAllUnits
             Do While Not RsUNits.EOF
                SSOleDBUnit.AddItem RsUNits!uni_code & ";" & RsUNits!uni_desc
                RsUNits.MoveNext
             Loop
            
              mIsPoItemsComboLoaded = True
        ''    Load ALL Units of Stockmaster
    End If
         
    ' Requisition Number
    
         Set RsRequsition = Lookups.GetRequisitions(deIms.NameSpace)
               
              If Not RsRequsition.EOF = True Then
                       ssdcboRequisition.Enabled = True
                   
                   Do While Not RsRequsition.EOF
                      ssdcboRequisition.AddItem RsRequsition!po_ponumb & ";" & RsRequsition!doc_desc & ";" & RsRequsition!poi_liitnumb & ";" & RsRequsition!poi_desc & ";" & RsRequsition!poi_primreqdqty
                      RsRequsition.MoveNext
                   Loop
                   
              Else
                    ssdcboRequisition.text = ""
                    ssdcboRequisition.Enabled = False
              End If
               
          Set Lookups = Nothing
          
    LoadPoItemCombos = True
    mIsPoItemsComboLoaded = True
    Exit Function
          
handler:
   MsgBox "Coud Not load all the Poitem Combos.  " & Err.Description
  Err.Clear
End Function


Public Function ClearAllPoLineItems() As Boolean
On Error GoTo handler
ClearAllPoLineItems = False

LblPOi_PONUMB = Trim$(ssOleDbPO)
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
'LblPOI_Doctype = PoItem.
 ClearAllPoLineItems = True
 
 Exit Function
handler:
  Err.Clear
End Function


Public Function LoadFromPORemarks() As Boolean
mLoadMode = loadingPoRemark
txtRemarks.text = PORemark.remarks
mLoadMode = NoLoadInProgress
End Function

Public Function ClearPoRemarks() As Boolean
txtRemarks.text = ""
End Function

Public Function ClearPoclause() As Boolean
txtClause.text = ""
End Function
Public Function SaveToPOITEM() As Boolean
PoItem.NameSpace = deIms.NameSpace
PoItem.Ponumb = LblPOi_PONUMB
PoItem.Linenumb = txt_LI
'poItem.Namespace = fnamespace
'Incase this is a NON-STOCK case.

    If PoItem.editmode = 2 Then
        If chk_FrmStkMst.Value = 0 Then
              PoItem.Comm = Trim$(ssOleDbPO.text) & "/" & Trim$(txt_LI)
        Else
              PoItem.Comm = Trim$(ssdcboCommoditty)
        End If
    End If
    
PoItem.Manupartnumb = ssdcboManNumber
PoItem.Afe = txt_AFE
PoItem.Custcate = SSOleDBCustCategory
PoItem.Serlnumb = txt_SerialNum

''''''''If SSOleDBUnit.Columns(1).Text = "P" Then
''''''''
''''''''   PoItem.Primreqdqty = CDbl(txt_Requested)
''''''''   PoItem.Primuom = SSOleDBUnit.Columns(0).Text
''''''''   PoItem.PrimUnitprice = txt_Price
''''''''   PoItem.UnitOfPurch = SSOleDBUnit.Columns(1).Text
''''''''   PoItem.Secoreqdqty = CDbl(txt_Requested) / RsUNits!stk_compfctr * 10000
''''''''   PoItem.Secouom = RsUNits!stk_secouom
''''''''   PoItem.SecUnitPrice = txt_Price * RsUNits!stk_compfctr / 10000
''''''''
''''''''
''''''''ElseIf SSOleDBUnit.Columns(1).Text = "S" Then
''''''''
''''''''
''''''''   PoItem.Secoreqdqty = txt_Requested
''''''''   PoItem.Secouom = SSOleDBUnit.Columns(0).Text
''''''''   PoItem.SecUnitPrice = FormatNumber(CDbl(txt_Price), 2)
''''''''
''''''''   PoItem.UnitOfPurch = SSOleDBUnit.Columns(1).Text
''''''''
''''''''   PoItem.Primreqdqty = CDbl(txt_Requested) * RsUNits!stk_compfctr / 10000
''''''''   PoItem.Primuom = RsUNits!stk_primuon
''''''''   PoItem.PrimUnitprice = CDbl(txt_Price) / RsUNits!stk_compfctr * 10000
''''''''
''''''''
''''''''Else  'This means it is Non-Stock.The ssoledbunit combo will have "A" appended to it
''''''''
''''''''   PoItem.Secoreqdqty = FormatNumber(txt_Requested, 2)
''''''''   PoItem.Secouom = SSOleDBUnit.Columns(0).Text
''''''''   PoItem.UnitOfPurch = "P"
''''''''   PoItem.Primreqdqty = FormatNumber(CDbl(txt_Requested), 2)
''''''''   PoItem.Primuom = SSOleDBUnit.Columns(0).Text
''''''''   PoItem.PrimUnitprice = CDbl(txt_Price)
''''''''   PoItem.SecUnitPrice = CDbl(txt_Price)
''''''''End If

Call SaveUnitsToPoItem
 
 

'PoItem = SSOleDBUnit
PoItem.Description = txt_Descript
PoItem.remk = txt_remk
'PoItem.PriTotaprice = CDbl(txt_Total)

PoItem.Liitreqddate = DTP_Required
PoItem.Requnumb = Trim$(ssdcboRequisition)
PoItem.Requliitnumb = (lblReqLineitem)

If Not objUnits Is Nothing Then Set objUnits = Nothing
End Function

Public Sub FirstTimeAssignmentsPOITEM()
LblPOI_Doctype.Caption = SSOleDBDocType.text
'ssdcboCommoditty.ColumnHeaders = True
'ssdcboCommoditty.Columns(0).Caption = "Code"
'ssdcboCommoditty.Columns(1).Caption = "Description"
End Sub

Private Sub txt_Price_Change()
'If PoItem.editmode <> 2 And PoItem.editmode <> -1 And Len(Trim$(txt_Price)) > 0 And mLoadMode = NoLoadInProgress Then Call SaveUnitsToPoItem
End Sub

Private Sub txt_Price_Validate(Cancel As Boolean)
If Len(txt_Price) > 0 Then
     txt_Price = FormatNumber(txt_Price, 2)
     
   If Len(txt_Requested) > 0 Then
        txt_Total = FormatNumber(CDbl(txt_Price) * txt_Requested, 2)
   End If
   
End If
End Sub

Private Sub txt_Requested_Change()

'If PoItem.editmode <> 2 And PoItem.editmode <> -1 And Len(Trim$(txt_Requested)) > 0 And mLoadMode = NoLoadInProgress Then Call SaveUnitsToPoItem

End Sub

Private Sub txt_Requested_Validate(Cancel As Boolean)
If Len(txt_Requested) > 0 Then
     
    If IsPrimQuantLessThanONE = False Then
        txt_Requested.SetFocus
        Cancel = True
        'txt_Requested = 0
        Exit Sub
    End If
     
     
'''''''    If txt_Requested < 1 And txt_Requested > 0 Then
'''''''       MsgBox "Quantity can not be Less than 0", , "Imswin"
'''''''       Cancel = True
'''''''       txt_Requested.SetFocus
'''''''
    

       txt_Requested = FormatNumber(txt_Requested, 2)

       If Len(txt_Price) > 0 Then
          txt_Total = FormatNumber(txt_Price * txt_Requested, 2)
       End If
     End If

End Sub


Public Function LoadPOLINEFromRequsition(Rs As ADODB.Recordset) As Boolean

'''''''        !poi_afe =
'''''''        !poi_remk = Rs("poi_remk") & ""
'''''''        !poi_comm = Rs("poi_comm") & ""
'''''''        !poi_desc =
'''''''        !poi_requnumb = Rs("po_ponumb") & ""
'''''''        !poi_primuom =
'''''''        !poi_serlnumb =
'''''''        !poi_custcate =
'''''''        !poi_requnumb = Rs("poi_ponumb") & ""
'''''''        !poi_unitprice = Rs("poi_unitprice") & ""
'''''''
'''''''        !poi_endrentdate = Rs("poi_endrentdate") & ""
'''''''        !poi_liitreqddate = Rs("poi_liitreqddate") & ""
'''''''        !poi_liitrelsdate = Rs("poi_liitrelsdate") & ""
'''''''        !poi_starrentdate = Rs("poi_starrentdate") & ""
'''''''
'''''''
'''''''        !poi_requliitnumb = Rs("poi_liitnumb") & ""
'''''''        !poi_secoreqdqty = Rs("poi_secoreqdqty") & ""
'''''''        !poi_primreqdqty = Rs("poi_primreqdqty") & ""
'''''''
'''''''
'''''''        txt_Price =
'''''''
'''''''        dcboUnit.BoundText = Rs("poi_primuom") & ""
'''''''
'''''''    End With
    
    
    
    

ssdcboCommoditty.text = Rs("poi_comm") & ""
ssdcboManNumber = Rs("poi_manupartnumb") & ""
txt_AFE = Rs("poi_afe") & ""
SSOleDBCustCategory = Rs("poi_custcate") & ""
txt_SerialNum = Rs("poi_serlnumb") & ""

'txtSecRequested = ""



txt_Delivered = ""
txt_Shipped = ""
txt_Inventory2 = ""


txt_Price = Rs("poi_unitprice") & ""

SSOleDBUnit.RemoveAll
SSOleDBUnit = Rs("poi_primuom")
txt_Total = FormatNumber(Rs("poi_unitprice") * Rs("poi_primreqdqty"), 2)
txt_Requested = Rs("poi_primreqdqty") & ""

If RsUNits Is Nothing Then
  If Lookups Is Nothing Then Lookups = MainPO.Lookups
  Set RsUNits = Lookups.GetAllUnits
End If
RsUNits.MoveFirst
RsUNits.Find ("uni_code='" & Trim$(Rs("poi_primuom")) & "'")
SSOleDBUnit.AddItem Rs("poi_primuom") & ";" & RsUNits("uni_desc")

RsUNits.MoveFirst
RsUNits.Find ("uni_code='" & Trim$(Rs("poi_primuom")) & "'")
SSOleDBUnit.AddItem Rs("poi_secouom") & ";" & RsUNits("uni_desc")


txt_Descript = Rs("poi_desc") & ""
txt_remk = ""
dcbostatus(1) = ""
dcbostatus(2) = ""
dcbostatus(3) = ""
dcbostatus(0) = ""
'ssdcboRequisition =
'lblReqLineitem = ""
End Function

Private Sub txt_SerialNum_Change()
'If PoItem.editmode <> 2 Or PoItem.editmode <> -1 And mLoadMode = NoLoadInProgress Then PoItem.Serlnumb = txt_SerialNum
End Sub



Public Function SaveUnitsToPoItem() As Boolean
SaveUnitsToPoItem = False
On Error GoTo handler

If chk_FrmStkMst.Value = 1 Then
 
   'This means the PO is In-Stock
        If objUnits Is Nothing Then
           Set objUnits = MainPO.PoUnits
           objUnits.StockNumber = ssdcboCommoditty.text
        End If
        
        If objUnits.SecondaryUnit = Trim$(SSOleDBUnit.text) And objUnits.SecondaryUnit <> objUnits.PrimaryUnit Then
           'It means it is in Seconday mode
                 PoItem.Secoreqdqty = txt_Requested
                 PoItem.Secouom = SSOleDBUnit.text
                 PoItem.SecUnitPrice = FormatNumber(CDbl(txt_Price), 2)
                 PoItem.SecTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
                 
                 PoItem.UnitOfPurch = "S"
                 
                 PoItem.Primreqdqty = CDbl(txt_Requested) * objUnits.ComPutationFactor / 10000
                 PoItem.Primuom = objUnits.PrimaryUnit
                 PoItem.PrimUnitprice = FormatNumber(CDbl(txt_Price) / objUnits.ComPutationFactor * 10000, 2)
                 PoItem.PriTotaprice = FormatNumber(CDbl(PoItem.Primreqdqty) * CDbl(PoItem.PrimUnitprice), 2)
                 
        ElseIf objUnits.PrimaryUnit = Trim$(SSOleDBUnit.text) And objUnits.SecondaryUnit <> objUnits.PrimaryUnit Then
                 'It is in Primary Mode
          
                 PoItem.Primreqdqty = CDbl(txt_Requested)
                 PoItem.Primuom = SSOleDBUnit.text
                 PoItem.PrimUnitprice = FormatNumber(CDbl(txt_Price), 2)
                 PoItem.PriTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
                 
                 PoItem.UnitOfPurch = "P"
                 
                 PoItem.Secoreqdqty = CDbl(txt_Requested) / objUnits.ComPutationFactor * 10000
                 PoItem.Secouom = objUnits.SecondaryUnit ' RsUNits!stk_secouom
                 PoItem.SecUnitPrice = FormatNumber(CDbl(txt_Price) * objUnits.ComPutationFactor / 10000, 2)
                 PoItem.PriTotaprice = FormatNumber(CDbl(PoItem.Secoreqdqty) * CDbl(PoItem.SecUnitPrice), 2)
        
        ElseIf objUnits.SecondaryUnit = objUnits.PrimaryUnit And objUnits.ComPutationFactor = 0 Then
        If Len(txt_Requested) > 0 Then
             PoItem.Secoreqdqty = FormatNumber(txt_Requested, 2)
        End If
        PoItem.Secouom = SSOleDBUnit.text
        If Len(txt_Price) > 0 Then
        PoItem.SecUnitPrice = IIf(Len(txt_Price) > 0, FormatNumber(CDbl(txt_Price), 2), "")
        End If
        If Len(txt_Requested) > 0 Then
        PoItem.SecTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
        End If
        PoItem.UnitOfPurch = "P"
        
        If Len(txt_Requested) > 0 Then
           PoItem.Primreqdqty = FormatNumber(CDbl(txt_Requested), 2)
        End If
        PoItem.Primuom = SSOleDBUnit.text
        If Len(txt_Price) > 0 Then
           PoItem.PrimUnitprice = CDbl(txt_Price)
        End If
        
        If Len(txt_Price) And Len(txt_Requested) > 0 Then
        PoItem.PriTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
        End If
        End If
   
   Else
   
    'This means the po is Non-Stock
        If Len(txt_Requested) > 0 Then
         PoItem.Secoreqdqty = FormatNumber(txt_Requested, 2)
        End If
        
        
        PoItem.Secouom = SSOleDBUnit.text
        If Len(txt_Price) > 0 Then
           PoItem.SecUnitPrice = IIf(Len(txt_Price) > 0, FormatNumber(txt_Price, 2), 0)
        End If
        
        If Len(txt_Price) And Len(txt_Requested) > 0 Then
          PoItem.SecTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
        End If
        
        PoItem.UnitOfPurch = "P"
        If Len(txt_Requested) > 0 Then
          PoItem.Primreqdqty = FormatNumber(CDbl(txt_Requested), 2)
        End If
        PoItem.Primuom = SSOleDBUnit.text
        If Len(txt_Price) > 0 Then
        PoItem.PrimUnitprice = CDbl(txt_Price)
        End If
        
        If Len(txt_Price) And Len(txt_Requested) > 0 Then
            PoItem.PriTotaprice = FormatNumber(CDbl(txt_Requested) * CDbl(txt_Price), 2)
        End If
   End If
    SaveUnitsToPoItem = True
  Exit Function
handler:
   MsgBox Err.Description
   Err.Clear


End Function

Public Function ToggleNavButtons(FMode As FormMode) As Boolean


 
        If FormMode = mdVisualization Then
                    
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
txtClause.text = POClause.Clause
mLoadMode = NoLoadInProgress

End Function

Public Function savetoPOclause() As Boolean

 If POClause.editmode = 2 Then
    POClause.NameSpace = deIms.NameSpace
    POClause.Linenumb = POClause.Count
    POClause.Ponumb = Poheader.Ponumb
 End If
 
 POClause.Clause = Trim(txtClause.text)

End Function

Public Function savetoPORemarks() As Boolean

If PORemark.editmode = 2 Then
    PORemark.NameSpace = deIms.NameSpace
    PORemark.Linenumb = PORemark.Count
    PORemark.Ponumb = Poheader.Ponumb
End If

PORemark.remarks = Trim(txtRemarks.text)

End Function

Public Function LoadFromPOReceipients() As Boolean
mLoadMode = loadingPoRemark
 Do While Not PoReceipients.EOF
    dgRecipientList.AddItem PoReceipients.Receipient
    PoReceipients.MoveNext
 Loop
mLoadMode = NoLoadInProgress
 
End Function

Public Function ClearPoReceipients()
dgRecipientList.RemoveAll
End Function

'Code after this Line Exist in PoOrder in Imswin
'------------------------------------------------------------------------

Private Sub AddRecepient(RecipientName As String, Optional ShowMessage As Boolean = True)
On Error Resume Next
Dim retval As Long

    
    If PoReceipients Is Nothing Then
            Set PoReceipients = MainPO.PoReceipients
            PoReceipients.Move Poheader.Ponumb
            LoadFromPOReceipients
    End If
    
    If Len(Trim$(RecipientName)) = 0 Then Exit Sub
      If InStr(RecipientName, "@") > 0 Then
        If Not InStr(UCase(RecipientName), UCase("Internet!")) > 0 Then
          RecipientName = "Internet!" & RecipientName
        End If
    Else
      If Not InStr(UCase(RecipientName), UCase("Fax!")) > 0 Then
          RecipientName = "FAX!" & RecipientName
       End If
    End If
    If IsRecipientInList(RecipientName, ShowMessage) Then Exit Sub
    If ((opt_FaxNum) And (InStr(1, RecipientName, "FAX!", vbTextCompare) = 0)) Then _
        RecipientName = FixFaxNumber(RecipientName)
    

    With PoReceipients
        .AddNew
        .Receipient = RecipientName
        .NameSpace = deIms.NameSpace
        .Linenumb = PoReceipients.Count
        
         If Len(Poheader.Ponumb) = 0 Then
             PoReceipients.Ponumb = Trim$(ssOleDbPO.text)
         Else
             PoReceipients.Ponumb = Poheader.Ponumb
         End If
         
    End With
    
    
    dgRecipientList.AddItem RecipientName
    

End Sub

Public Sub SortGrid(Rs As ADODB.Recordset, Grid As DataGrid, col As Integer)
On Error Resume Next
    Dim SortOrder As String
    Dim BK As Variant

    BK = Rs.Bookmark
    SortOrder = Grid.Tag
    SortOrder = IIf(UCase(SortOrder) = "ASC", "ASC", "DESC")
    Grid.Tag = IIf(UCase(SortOrder) = "ASC", "DESC", "ASC")

    Rs.Sort = ""
    Rs.Sort = ((Grid.Columns(col).DataField) + " " + SortOrder)
    Rs.Bookmark = BK
    If Err Then Err.Clear
End Sub
Private Function IsRecipientInList(RecepientName As String, Optional ShowMessage As Boolean = True) As Boolean
IsRecipientInList = False
On Error GoTo handler




 PoReceipients.MoveFirst
Do While Not PoReceipients.EOF
  
   If PoReceipients.Receipient = RecepientName Then
              IsRecipientInList = True
              Exit Do
   End If
   PoReceipients.MoveNext
Loop

   Exit Function
handler:
   Err.Raise Err.number, , Err.Description
   Err.Clear
   




''''''''On Error Resume Next
''''''''Dim BK As Variant
''''''''
''''''''
''''''''    rsrecepList.MoveFirst
''''''''    If Not (rsrecepList.EOF Or rsrecepList.BOF) Then BK = rsrecepList.Bookmark
''''''''
''''''''    rsrecepList.Filter = "porc_rec = '" & RecepientName & "'"
''''''''
''''''''    If Not (rsrecepList.EOF) Then
''''''''
''''''''        If ((Not (rsrecepList.RecordCount = 0))) Then
''''''''
''''''''            #If DBUG Then
''''''''                If Err Then Stop
''''''''            #Else
''''''''                On Error Resume Next
''''''''            #End If
''''''''
''''''''            If ShowMessage Then
''''''''                If opt_Email Then
''''''''
''''''''                    'Modified by Juan (8/29/2000) for Multilingual
''''''''                    msg1 = translator.Trans("M00076") 'J added
''''''''                    MsgBox IIf(msg1 = "", "Email Address Already in list", msg1) 'J modified
''''''''                    '------------------------------------------
''''''''
''''''''                ElseIf opt_FaxNum Then
''''''''
''''''''                    'Modified by Juan (8/29/2000) for Multilingual
''''''''                    msg1 = translator.Trans("M00077") 'J added
''''''''                    MsgBox IIf(msg1 = "", "Fax Number Already in list", msg1) 'J modified
''''''''                    '---------------------------------------------
''''''''
''''''''                End If
''''''''            End If
''''''''            IsRecipientInList = True
''''''''        End If
''''''''    End If
''''''''
''''''''     rsrecepList.Filter = adFilterNone
''''''''     If IsRecipientInList Then Call rsrecepList.Find("porc_rec = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
''''''''
''''''''     If rsrecepList.RecordCount = 0 Then IsRecipientInList = 0
''''''''
''''''''     'If rsrecepList.EOF Then rsrecepList.MoveFirst
''''''''    If Err Then Err.Clear
''''''''
End Function

Private Function FixFaxNumber(Faxnumber As String) As String
On Error Resume Next

    If Len(Faxnumber) < 7 Then Exit Function

    If Left$(Faxnumber, 1) = "+" Then
        Faxnumber = Right$(Faxnumber, Len(Faxnumber) - 1)
    End If
    
    If Mid$(Faxnumber, 1, 4) <> "FAX!" Then _
        FixFaxNumber = "FAX!" & Faxnumber

    'Modified by Juan (9/14/2000) for Multilingual
  '  msg1 = translator.Trans("M00078") 'J added
  '  If Err Then Err.Clear: MsgBox IIf(msg1 = "", "err occured", msg1) 'J modified
    '---------------------------------------------

End Function


Public Sub PoReceipeintsInit()
dgRecipientList.Columns(0).Width = dgRecipientList.Width
End Sub

Public Sub FirstTimeAssignments()

If FormMode = mdCreation Then
   cmdRemove.Visible = True
Else
   cmdRemove.Visible = False
End If
dgRecipientList.Caption = "Recipients"
End Sub




Public Function LoadDocType()
If mIsDocTypeLoaded = False Then
      If Lookups Is Nothing Then Set Lookups = MainPO.Lookups
     Dim GRsDoctype As ADODB.Recordset
     Set GRsDoctype = Lookups.GetDoctypeForUser(CurrentUser)
        
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
    

    
    If Len(Trim$(SSOleDBDocType.text)) = 0 Then
        'Call MsgBox(LoadResString(101)): dcboDocumentType.SetFocus: Exit Function
        MsgBox "Document Type can not be Left Empty."
        SSOleDBDocType.SetFocus
        Exit Function
    End If
        
    If Len(Trim$(ssdcboShipper.text)) = 0 Then
        'Call MsgBox(LoadResString(102)): ssdcboShipper.SetFocus: Exit Function
        MsgBox "Shipper can not be Left Empty."
        ssdcboShipper.SetFocus
        Exit Function
    End If
    'Else
    
   '     rsPO!po_shipcode = ssdcboShipper.Value
   ' End If
        
    
    If Len(Trim$(SSOleDBPriority.text)) = 0 Then
        'Call MsgBox(LoadResString(103)): dcboPriority.SetFocus: Exit Function
         MsgBox "Priority can not be Left Empty."
         SSOleDBPriority.SetFocus: Exit Function
    'Else
      '  rsPO!po_priocode = dcboPriority.BoundText
    End If
    
    If Len(Trim$(SSOleDBCurrency.text)) = 0 Then
        'Call MsgBox(LoadResString(104)):
        MsgBox "Currency can not Be left Empty"
        SSOleDBCurrency.SetFocus: Exit Function
        
   ' Else
    '    rsPO!po_currcode = dcboCurrency.BoundText
    End If
    
    
    If Len(Trim$(SSOleDBOriginator.text)) = 0 Then
       ' Call MsgBox(LoadResString(105)): dcboOriginator.SetFocus: Exit Function
       MsgBox "Originator can not be Left Empty"
       SSOleDBCurrency.SetFocus
       Exit Function
    'Else
     '   rsPO!po_orig = dcboOriginator.BoundText
    End If
    
        
    If Len(Trim$(SSOleDBShipTo.text)) = 0 Then
        'Call MsgBox(LoadResString(106)):
        MsgBox "Ship to can not be left empty."
        SSOleDBShipTo.SetFocus: Exit Function
        
    'Else
     '   rsPO!po_shipto = dcboShipto.BoundText
        
    End If
    
    
    If Len(Trim$(SSOleDBCompany.text)) = 0 Then  'M
    
        'Modified by Juan (9/13/2000) for Multilingual
       ' msg1 = translator.Trans("M00023") 'J added
        'MsgBox IIf(msg1 = "", "Company Can not be left empty", msg1), , "Imswin" 'J modified
        '---------------------------------------------
        MsgBox "Company can not be left Empty."
      SSOleDBCompany.SetFocus
      Exit Function 'M
        
    'Else  'M
     '  rsPO!po_compcode = dcboCompany.BoundText 'M
    End If 'M
    
    

    
    If Len(Trim$(SSOleDBInvLocation.text)) = 0 Then
      '  Call MsgBox(LoadResString(107)): SSOleDBInvLocation.SetFocus: Exit Function
        MsgBox "Inventory Location Can not be left Empty."
        SSOleDBInvLocation.SetFocus: Exit Function
    'Else
     '   rsPO!po_invloca = dcboInvLocation.BoundText
    End If
    
    If Len(Trim$(SSoledbSupplier.text)) = 0 Then
       ' Call MsgBox(LoadResString(108)): dcboSupplier.SetFocus: Exit Function
        MsgBox "Supplier Can not be Left Empty"
        SSoledbSupplier.SetFocus: Exit Function
    'Else
     '   rsPO!po_suppcode = dcboSupplier.Value
    End If
    
    'Modified by Muzammil 08/14/00
    'Reason - Should scream at the user when left empty and the user tries clicking some
    'other tab.
    
    
    If Len(Trim$(ssdcboCondition.text)) = 0 Then          'M
    
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
    
    If Len(Trim$(ssdcboDelivery.text)) = 0 Then  'M
    
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
    If DTPicker_poDate.Value > dtpRequestedDate.Value Then 'Or DTPicker_poDate.Value = dtpRequestedDate.Value Then   'M  'J Modified
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
On Error GoTo handler
Dim i As Long


    
        Call txt_Requested_Validate(False)
    
    If Len(Trim$(txt_Requested)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
    '    msg1 = translator.Trans("M00029") 'J added
     '   Call MsgBox(IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1))
        '---------------------------------------------
        MsgBox "Requested amount does not contain a valid entry"
        txt_Requested.SetFocus
        Exit Function
    
    ElseIf Not IsNumeric(Trim$(txt_Requested)) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        'msg1 = translator.Trans("M00029") 'J added
        'MsgBox IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1)
        '---------------------------------------------
        MsgBox "Requested amount does not contain a valid entry"
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
    
        'Modified by Juan (9/13/2000) for Multilingual
        'msg1 = translator.Trans("M00030") 'J added
        'MsgBox IIf(msg1 = "", "Price cannot be left empty ", msg1) 'J modified
        '---------------------------------------------
        MsgBox "Price cannot be left empty "
        txt_Price.SetFocus: Exit Function
        
    ElseIf Not (IsNumeric(txt_Price)) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        'msg1 = translator.Trans("M00031") 'J added
        'MsgBox IIf(msg1 = "", "Price does not have a valid entry", msg1) 'J modified
        '---------------------------------------------
        MsgBox "Price does not have a valid entry"
        txt_Price.SetFocus: Exit Function
        
    'Else
     '   rsPOITEM!poi_unitprice = CDbl(txt_Price)
    End If
    
'    If Len(Trim$(dcboCustomCategory)) = 0 Then
'        MsgBox "Custom Category canot be left empty"
'        dcboCustomCategory.SetFocus: Exit Function
'    End If
    
    'If Len(Trim$(rsPOITEM!poi_primuom & "")) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        'msg1 = translator.Trans("M00032") 'J added
        'MsgBox IIf(msg1 = "", "Unit cannot be left empty", msg1) 'J modified
      '  MsgBox "Unit cannot be left empty"
        '---------------------------------------------
     '   SSOleDBUnit.SetFocus: Exit Function
   ' End If
    
    If chk_FrmStkMst.Value = vbChecked Then
    
        If Len(Trim$(ssdcboCommoditty.text)) = 0 Then
        
            'Modified by Juan (9/13/2000) for Multilngual
            'msg1 = translator.Trans("M00025") 'J added
            'MsgBox IIf(msg1 = "", "Stock Number cannot be left empty", msg1) 'J modified
            '--------------------------------------------
            MsgBox "Stock Number cannot be left empty"
            ssdcboCommoditty.SetFocus: Exit Function
         End If
''''''        Else
''''''            Dim STR As String, OldNum As String
''''''
''''''
''''''            STR = LCase(Trim$(rsPOITEM!poi_comm & ""))
''''''            OldNum = LCase(Trim$(rsPOITEM!poi_comm.OriginalValue & ""))
''''''
''''''            If OldNum <> STR Then
''''''
''''''                If Not deIms.StockNumberExist(STR, True) Then
''''''
''''''                    'Modified by Juan (9/13/2000) for multilingual
''''''                    msg1 = translator.Trans("L00119") 'J added
''''''                    msg2 = translator.Trans("M00026") 'J added
''''''                    MsgBox IIf(msg1 = "", "Stock number ", msg1 + " ") & STR & IIf(msg2 = "", " does not exist", " " + msg2) 'J modified
''''''                    '---------------------------------------------
''''''
''''''                    ssdcboCommoditty.SetFocus: Exit Function
''''''
''''''                End If
''''''
''''''            End If
''''''
''''''        End If
    End If
    If Len(Trim$(SSOleDBUnit)) = 0 Then
        MsgBox "Unit can not be Left Empty.", , "Imswin"
        SSOleDBUnit.SetFocus
        Exit Function
    End If
    
    CheckLIFields = True
    Exit Function
handler:
    Err.Clear
End Function

Public Function InitializePOheaderRecordset()
 
Dim FNamespace As String * 5

'Dim RsBRQ As ADODB.Recordset

Dim DefSite As String
   FNamespace = deIms.NameSpace

    Set rsDOCTYPE = GetDocumentType(False)
    Call deIms.SHIPPER(FNamespace)
    Call deIms.Currency(FNamespace)
    Call deIms.PRIORITY(FNamespace)
    Call deIms.TermDelivery(FNamespace)
    'Call deIms.Supplier(fnamespace)
    Call deIms.ActiveSupplier(FNamespace)
    Call deIms.TermCondition(FNamespace)
    'Call deIms.INVENTORYLOCATION(Fnamespace, ponumb)
    Call deIms.Company(FNamespace)
    Call deIms.GETSYSSITE(FNamespace, DefSite)
    Call deIms.ActiveOriginator(FNamespace)
    Call deIms.ActiveTbu(FNamespace)
    Call deIms.servcodecat(FNamespace)
    Call deIms.ActiveShipTo(FNamespace)
    Call deIms.ActiveCompany(FNamespace)
    Call deIms.CompanyLocations(FNamespace)
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
  
     ElseIf FMode = mdVisualization Then
        lblStatus.ForeColor = vbGreen
        
        'Modified by Juan (8/28/2000) for Multilingual
        'msg1 = translator.Trans("L00092") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
    
    End If
    
       
    FMode = FMode
    Call MakeReadOnly(FMode = mdVisualization)
    'Call ShowActiveRecords(False)
    
    'GetUnits ("")
   
    'LockWindowUpdate (0)
   ChangeMode = FMode
End Function
Private Sub MakeReadOnly(Value As Boolean)
On Error Resume Next

    txtClause.locked = Value
    txtRemarks.locked = Value
    
    Value = Not Value
    
    cmd_Add.Enabled = Value
    
    'txtClause.Enabled = Value
    'TxtRemarks.Enabled = Value
    
    
    cmd_Addterms.Enabled = Value
    
    CmdcopyLI.Item(0).Enabled = Value  'M
    CmdcopyLI.Item(1).Enabled = Value  'M
    CmdcopyLI.Item(2).Enabled = Value  'M
    
    fra_LineItem.Enabled = Value
    dgRecepients.Enabled = Value
    'PO.dgRecepients.AllowRowSizing
    fra_Purchase.Enabled = Value
    fra_FaxSelect.Enabled = Value
    dgRecipientList.Enabled = Value
    SSOleDBDocType.Enabled = Value
    'RefreshAll
    If Err Then Err.Clear
End Sub
Private Sub comsearch_Completed(Cancelled As Boolean, sStockNumber As String)
On Error Resume Next

    comsearch.Hide
    
    ssdcboCommoditty.text = sStockNumber
    txt_Descript = comsearch.Description
    
 If objUnits Is Nothing Then Set objUnits = MainPO.PoUnits
        
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
Private Sub WriteStatus(Msg As String)
    Call MDI_IMS.WriteStatus(Msg, 1)
End Sub
Private Sub BeforePrint()
On Error Resume Next

    With MDI_IMS.CrystalReport1
        .ReportFileName = ReportPath & "po.rpt"
        
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
                           Subject As String, Message As String, Attachment As String, _
                           Optional Orientation As OrientationConstants)
    Dim address() As String
    Dim STR As String
    Dim i As Integer

    On Error Resume Next

    address = ToArrayFromRec(Recipients, FieldName, i, STR)
    
    Dim faxAddresses() As String: faxAddresses = filterAddresses(address, True)
    If UBound(faxAddresses) > 0 Then
        Call sendFaxOnly(Subject, faxAddresses, Attachment)
    End If
    
    Dim emailAddresses() As String: emailAddresses = filterAddresses(address, False)
    If UBound(emailAddresses) > 0 Then
        Call sendEmailOnly(Subject, emailAddresses, Attachment)
    End If
    
    Kill Attachment

    If Not IsLoaded("MDI_IMS") Then End
    MDI_IMS.CrystalReport1.Reset
    
    If Err Then Err.Clear
End Sub
Public Function ToArrayFromRec(Rs As PoReceipients, ByVal FieldName As String, Optional UpperBound As Integer, Optional ByVal Filter As String) As String()
Dim BK As Variant
Dim STR() As String
Dim OldFilter As Variant

On Error GoTo Errhandler
    ReDim STR(0)
    UpperBound = -1
    If Rs Is Nothing Then Exit Function
    
    'BK = rs.Bookmark
    
    
'''    If Len(Filter) Then
'''        OldFilter = rs.Filter
'''        rs.Filter = adFilterNone
'''        rs.Filter = Filter
'''    End If
    
    Rs.MoveFirst
    Do While Not Rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve STR(UpperBound)
        STR(UpperBound) = Rs.Receipient
        Rs.MoveNext
    Loop
    
    ToArrayFromRec = STR
    
    'If Len(Filter) Then rs.Filter = OldFilter
    'rs.Bookmark = BK
    Exit Function
    
Errhandler:
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
 
 On Error GoTo handler
  If Len(txt_Requested) > 0 And Len(SSOleDBUnit) > 0 Then
       If objUnits Is Nothing Then Set objUnits = MainPO.PoUnits
       objUnits.StockNumber = Trim$(ssdcboCommoditty)
       If objUnits.PrimaryUnit <> objUnits.SecondaryUnit Then
             If objUnits.PrimaryUnit = Trim$(SSOleDBUnit) Then
                 If CDbl(txt_Requested) < 1 And CDbl(txt_Requested) > 0 Then
                    MsgBox "Quantity Can not be Less than 1", , "Imswin"
                    Set objUnits = Nothing
                    Exit Function
                 End If
            ElseIf objUnits.SecondaryUnit = Trim$(SSOleDBUnit) Then
                PriQuantity = CDbl(txt_Requested) * objUnits.ComPutationFactor / 10000
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
handler:
   Err.Clear
End Function
