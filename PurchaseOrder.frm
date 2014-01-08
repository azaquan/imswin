VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#8.0#0"; "LRNavigators.ocx"
Begin VB.Form frm_Purchase 
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
      Left            =   1080
      TabIndex        =   51
      Top             =   6540
      Width           =   3855
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
      TabIndex        =   137
      Top             =   120
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   11086
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   758
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Transaction Order"
      TabPicture(0)   =   "PurchaseOrder.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_PO"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Purchase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "PurchaseOrder.frx":001C
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
      TabPicture(2)   =   "PurchaseOrder.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fra_LI"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fra_LineItem"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Remarks"
      TabPicture(3)   =   "PurchaseOrder.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdcopyLI(1)"
      Tab(3).Control(1)=   "txtRemarks"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Notes/Instructions"
      TabPicture(4)   =   "PurchaseOrder.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Addterms"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtClause"
      Tab(4).Control(2)=   "CmdcopyLI(2)"
      Tab(4).ControlCount=   3
      Begin VB.Frame fra_Purchase 
         ClipControls    =   0   'False
         Height          =   5100
         Left            =   -74760
         TabIndex        =   53
         Top             =   960
         Width           =   8295
         Begin MSComCtl2.DTPicker DTPicker_poDate 
            Bindings        =   "PurchaseOrder.frx":008C
            DataField       =   "po_date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataMember      =   "po"
            Height          =   315
            Left            =   6480
            TabIndex        =   7
            Top             =   2760
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   25231361
            CurrentDate     =   36850
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboDelivery 
            Bindings        =   "PurchaseOrder.frx":0099
            DataField       =   "po_termcode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   6450
            TabIndex        =   17
            Top             =   4680
            Width           =   1665
            DataFieldList   =   "tod_termcode"
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
            stylesets(0).Picture=   "PurchaseOrder.frx":00C7
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
            stylesets(1).Picture=   "PurchaseOrder.frx":00E3
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
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "tod_desc"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "tod_termcode"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Object.DataMember      =   "TermDelivery"
            DataFieldToDisplay=   "tod_desc"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo dcboSupplier 
            Bindings        =   "PurchaseOrder.frx":00FF
            DataField       =   "po_suppcode"
            DataMember      =   "PO"
            DataSource      =   "deIms"
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   3105
            Width           =   2295
            DataFieldList   =   "sup_code"
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
            ForeColorEven   =   8388608
            BackColorOdd    =   16771818
            RowHeight       =   423
            Columns.Count   =   4
            Columns(0).Width=   3200
            Columns(0).Caption=   "Supplier"
            Columns(0).Name =   "sup_name"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "sup_name"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Caption=   "City"
            Columns(1).Name =   "sup_city"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "sup_city"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Caption=   "sup_code"
            Columns(2).Name =   "sup_code"
            Columns(2).CaptionAlignment=   0
            Columns(2).DataField=   "sup_code"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Caption=   "Phone Number"
            Columns(3).Name =   "sup_phonnumb"
            Columns(3).CaptionAlignment=   0
            Columns(3).DataField=   "sup_phonnumb"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Object.DataMember      =   "SUPPLIER"
            DataFieldToDisplay=   "sup_code"
         End
         Begin VB.TextBox txtSite 
            BackColor       =   &H00FFFFC0&
            DataField       =   "po_site"
            DataMember      =   "PO"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6450
            TabIndex        =   134
            Top             =   4100
            Width           =   1665
         End
         Begin MSDataListLib.DataCombo dcboToBeUsedFor 
            Bindings        =   "PurchaseOrder.frx":0116
            DataField       =   "po_tbuf"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   2775
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ForeColor       =   -2147483640
            ListField       =   "tbu_name"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dcboCurrency 
            Bindings        =   "PurchaseOrder.frx":014A
            DataField       =   "po_currcode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Top             =   3435
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "curr_desc"
            BoundColumn     =   "curr_code"
            Text            =   ""
            Object.DataMember      =   "CURRENCY"
         End
         Begin VB.CheckBox chk_Forwarder 
            Caption         =   "Forwarder"
            DataField       =   "po_forwr"
            DataMember      =   "PO"
            Height          =   288
            Left            =   4635
            TabIndex        =   19
            Top             =   827
            Width           =   3300
         End
         Begin VB.TextBox txt_ChargeTo 
            DataField       =   "po_chrgto"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   3
            Top             =   800
            Width           =   2295
         End
         Begin VB.TextBox txt_Buyer 
            BackColor       =   &H00FFFFC0&
            CausesValidation=   0   'False
            DataField       =   "po_buyr"
            DataMember      =   "PO"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   1460
            Width           =   2295
         End
         Begin VB.CheckBox chk_Requ 
            Caption         =   "Print Required date for each LI ? Y/N"
            DataField       =   "po_reqddelvflag"
            DataMember      =   "PO"
            Height          =   288
            Left            =   4635
            TabIndex        =   18
            Top             =   492
            Width           =   3225
         End
         Begin VB.Frame fra_Stat 
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   1620
            Left            =   4560
            TabIndex        =   54
            Top             =   1020
            Width           =   3600
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":0182
               DataField       =   "po_stas"
               DataMember      =   "PO"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   4
               Left            =   1260
               TabIndex        =   111
               Top             =   180
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               MatchEntry      =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":0199
               DataField       =   "po_stasdelv"
               DataMember      =   "PO"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   5
               Left            =   1260
               TabIndex        =   112
               Top             =   510
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               MatchEntry      =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":01B0
               DataField       =   "po_stasship"
               DataMember      =   "PO"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   6
               Left            =   1260
               TabIndex        =   113
               Top             =   840
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               MatchEntry      =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":01C7
               DataField       =   "po_stasinvt"
               DataMember      =   "PO"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   7
               Left            =   1260
               TabIndex        =   114
               Top             =   1170
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               MatchEntry      =   -1  'True
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin VB.Label lbl_Shipping 
               BackColor       =   &H8000000A&
               Caption         =   "Shipping"
               Height          =   225
               Left            =   105
               TabIndex        =   58
               Top             =   885
               Width           =   1200
            End
            Begin VB.Label lbl_Delivery 
               BackColor       =   &H8000000A&
               Caption         =   "Delivery"
               Height          =   225
               Left            =   105
               TabIndex        =   57
               Top             =   585
               Width           =   1200
            End
            Begin VB.Label lbl_Status 
               BackColor       =   &H8000000A&
               Caption         =   "PO"
               Height          =   225
               Left            =   105
               TabIndex        =   56
               Top             =   300
               Width           =   1200
            End
            Begin VB.Label lbl_Inventory 
               BackColor       =   &H8000000A&
               Caption         =   "Inventory"
               Height          =   225
               Left            =   105
               TabIndex        =   55
               Top             =   1215
               Width           =   1200
            End
         End
         Begin MSComCtl2.DTPicker dtpRequestedDate 
            Bindings        =   "PurchaseOrder.frx":01DE
            DataField       =   "po_reqddelvdate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataMember      =   "PO"
            Height          =   315
            Left            =   6450
            TabIndex        =   10
            Top             =   3435
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   25231363
            CurrentDate     =   36402
         End
         Begin MSDataListLib.DataCombo dcboPriority 
            Bindings        =   "PurchaseOrder.frx":0206
            DataField       =   "po_priocode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   4
            Top             =   1125
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "pri_desc"
            BoundColumn     =   "pri_code"
            Text            =   ""
            Object.DataMember      =   "Priority"
         End
         Begin MSDataListLib.DataCombo dcboOriginator 
            Bindings        =   "PurchaseOrder.frx":024B
            DataField       =   "po_orig"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   1800
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "ori_code"
            BoundColumn     =   "ori_code"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboShipper 
            Bindings        =   "PurchaseOrder.frx":0262
            CausesValidation=   0   'False
            DataField       =   "po_shipcode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   480
            Width           =   3015
            DataFieldList   =   "shi_code"
            ListAutoValidate=   0   'False
            AutoRestore     =   0   'False
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
            stylesets(0).Picture=   "PurchaseOrder.frx":029A
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
            stylesets(1).Picture=   "PurchaseOrder.frx":02B6
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
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "shi_name"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "shi_code"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   5318
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Object.DataMember      =   "Shipper"
            DataFieldToDisplay=   "shi_name"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCategoryCode 
            Bindings        =   "PurchaseOrder.frx":02D2
            DataField       =   "po_catecode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   106
            Top             =   2445
            Width           =   2295
            DataFieldList   =   "scs_desc"
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
            stylesets(0).Picture=   "PurchaseOrder.frx":0311
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
            stylesets(1).Picture=   "PurchaseOrder.frx":032D
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   4180
            Columns(0).Caption=   "Designation"
            Columns(0).Name =   "Designation"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "scs_desc"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1191
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "scs_code"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   16777152
            Enabled         =   0   'False
            Object.DataMember      =   "SERVCODECAT"
            DataFieldToDisplay=   "scs_desc"
         End
         Begin MSDataListLib.DataCombo dcboShipto 
            Bindings        =   "PurchaseOrder.frx":0349
            DataField       =   "po_shipto"
            DataMember      =   "PO"
            Height          =   315
            Left            =   5760
            TabIndex        =   12
            Top             =   3765
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "sht_name"
            BoundColumn     =   "sht_code"
            Text            =   ""
            Object.DataMember      =   "ActiveShipTo"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
            Bindings        =   "PurchaseOrder.frx":037F
            DataField       =   "po_taccode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1920
            TabIndex        =   16
            Top             =   4680
            Width           =   2295
            DataFieldList   =   "tac_taccode"
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
            stylesets(0).Picture=   "PurchaseOrder.frx":03AD
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
            stylesets(1).Picture=   "PurchaseOrder.frx":03C9
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
            Columns(0).Caption=   "Description"
            Columns(0).Name =   "Name"
            Columns(0).DataField=   "tac_desc"
            Columns(0).FieldLen=   256
            Columns(1).Width=   5292
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "Code"
            Columns(1).Name =   "Code"
            Columns(1).DataField=   "tac_taccode"
            Columns(1).FieldLen=   256
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Object.DataMember      =   "TermCondition"
            DataFieldToDisplay=   "tac_desc"
         End
         Begin MSDataListLib.DataCombo dcboInvLocation 
            Bindings        =   "PurchaseOrder.frx":03E5
            DataField       =   "po_invloca"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   4095
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "loc_name"
            BoundColumn     =   "loc_locacode"
            Text            =   ""
            Object.DataMember      =   "INVENTORYLOCATION"
         End
         Begin MSDataListLib.DataCombo dcboCompany 
            Bindings        =   "PurchaseOrder.frx":0412
            DataField       =   "po_compcode"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   3765
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "com_name"
            BoundColumn     =   "com_compcode"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin VB.CheckBox chk_FrmStkMst 
            Alignment       =   1  'Right Justify
            Caption         =   "From Stock Master"
            DataField       =   "po_fromstckmast"
            DataMember      =   "PO"
            Height          =   405
            Left            =   4560
            TabIndex        =   15
            Top             =   4320
            Width           =   2340
         End
         Begin VB.CheckBox chk_ConfirmingOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirming Order"
            DataField       =   "po_confordr"
            DataMember      =   "PO"
            Height          =   288
            Left            =   60
            TabIndex        =   14
            Top             =   4380
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Company"
            Height          =   225
            Left            =   90
            TabIndex        =   136
            Top             =   3765
            Width           =   1350
         End
         Begin VB.Label lbl_InvLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Invt. Location"
            Height          =   225
            Left            =   90
            TabIndex        =   135
            Top             =   4100
            Width           =   1350
         End
         Begin VB.Label lbl_Revision 
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Number"
            Height          =   225
            Left            =   90
            TabIndex        =   133
            Top             =   130
            Width           =   1245
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "po_revinumb"
            DataMember      =   "PO"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            TabIndex        =   132
            Top             =   135
            Width           =   615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Term"
            Height          =   225
            Left            =   4560
            TabIndex        =   125
            Top             =   4725
            Width           =   1845
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "T && C"
            Height          =   225
            Left            =   60
            TabIndex        =   124
            Top             =   4725
            Width           =   1605
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "po_daterevi"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataMember      =   "PO"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6300
            TabIndex        =   121
            Top             =   135
            Width           =   1035
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "po_datesent"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataMember      =   "PO"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6450
            TabIndex        =   120
            Top             =   3105
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "po_apprby"
            DataMember      =   "PO"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   119
            Top             =   2115
            Width           =   2295
         End
         Begin VB.Label lbl_Supplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Code"
            Height          =   225
            Left            =   90
            TabIndex        =   74
            Top             =   3105
            Width           =   1725
         End
         Begin VB.Label lbl_ToBe 
            BackStyle       =   0  'Transparent
            Caption         =   "To Be Used For"
            Height          =   225
            Left            =   90
            TabIndex        =   73
            Top             =   2775
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
            Top             =   3435
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
            Left            =   4635
            TabIndex        =   69
            Top             =   225
            Width           =   1680
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
            Top             =   4095
            Width           =   1530
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
         Begin VB.Label lbl_CatCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Category Code"
            Height          =   225
            Left            =   90
            TabIndex        =   59
            Top             =   2445
            Width           =   1650
         End
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   2
         Left            =   -73040
         TabIndex        =   47
         Top             =   528
         Width           =   1695
      End
      Begin VB.CommandButton CmdcopyLI 
         Caption         =   "Copy From ...."
         Height          =   288
         Index           =   1
         Left            =   -74760
         TabIndex        =   44
         Top             =   550
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74640
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame fra_LineItem 
         BorderStyle     =   0  'None
         Height          =   5280
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   8520
         Begin VB.CommandButton CmdcopyLI 
            Caption         =   "Copy From ...."
            Height          =   305
            Index           =   0
            Left            =   360
            TabIndex        =   42
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txt_SerialNum 
            DataField       =   "poi_serlnumb"
            DataMember      =   "POITEM"
            Height          =   285
            Left            =   1920
            TabIndex        =   33
            Top             =   1880
            Width           =   1836
         End
         Begin VB.TextBox txtdesc 
            DataField       =   "poi_remk"
            DataMember      =   "POITEM"
            DataSource      =   "deIms"
            Height          =   675
            Left            =   2040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
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
            TabIndex        =   43
            Top             =   3720
            Width           =   6420
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboRequisition 
            Bindings        =   "PurchaseOrder.frx":043F
            DataField       =   "poi_requnumb"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   5760
            TabIndex        =   34
            Top             =   180
            Width           =   1575
            DataFieldList   =   "po_ponumb"
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
            stylesets(0).Picture=   "PurchaseOrder.frx":0450
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
            stylesets(1).Picture=   "PurchaseOrder.frx":046C
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
            Columns(0).DataField=   "po_ponumb"
            Columns(0).FieldLen=   256
            Columns(1).Width=   3836
            Columns(1).Caption=   "Type"
            Columns(1).Name =   "Type"
            Columns(1).DataField=   "doc_desc"
            Columns(1).FieldLen=   256
            Columns(2).Width=   1005
            Columns(2).Caption=   "Item"
            Columns(2).Name =   "Item"
            Columns(2).DataField=   "poi_liitnumb"
            Columns(2).FieldLen=   256
            Columns(3).Width=   5292
            Columns(3).Caption=   "Description"
            Columns(3).Name =   "Description"
            Columns(3).DataField=   "poi_desc"
            Columns(3).FieldLen=   256
            Columns(4).Width=   1693
            Columns(4).Caption=   "Qty"
            Columns(4).Name =   "Qty"
            Columns(4).DataField=   "poi_primreqdqty"
            Columns(4).FieldLen=   256
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "po_ponumb"
         End
         Begin VB.TextBox txt_Price 
            DataField       =   "poi_unitprice"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   7080
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1275
         End
         Begin VB.TextBox txt_Item 
            BackColor       =   &H00FFFFC0&
            DataField       =   "poi_liitnumb"
            DataMember      =   "POITEM"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   118
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
            TabIndex        =   115
            Top             =   3240
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Bindings        =   "PurchaseOrder.frx":0488
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
            Format          =   25231363
            CurrentDate     =   36405
         End
         Begin VB.TextBox txt_linumber 
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
            TabIndex        =   31
            Top             =   1200
            Width           =   2310
         End
         Begin MSDataListLib.DataCombo dcboCustomCategory 
            Bindings        =   "PurchaseOrder.frx":04B9
            DataField       =   "poi_custcate"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   1440
            TabIndex        =   32
            Top             =   1530
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "cust_cate"
            Text            =   ""
            Object.DataMember      =   "CUSTOM"
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
               Bindings        =   "PurchaseOrder.frx":0504
               DataField       =   "poi_stasliit"
               DataMember      =   "POITEM"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   0
               Left            =   1320
               TabIndex        =   107
               Top             =   150
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               Style           =   2
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":053B
               DataField       =   "poi_stasdlvy"
               DataMember      =   "POITEM"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   108
               Top             =   480
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               Style           =   2
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":0572
               DataField       =   "poi_stasship"
               DataMember      =   "POITEM"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   2
               Left            =   1320
               TabIndex        =   109
               Top             =   810
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               Style           =   2
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
            End
            Begin MSDataListLib.DataCombo dcbostatus 
               Bindings        =   "PurchaseOrder.frx":05A9
               DataField       =   "poi_stasinvt"
               DataMember      =   "POITEM"
               DataSource      =   "deIms"
               Height          =   315
               Index           =   3
               Left            =   1320
               TabIndex        =   110
               Top             =   1140
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Locked          =   -1  'True
               Style           =   2
               BackColor       =   16777152
               ForeColor       =   16711680
               ListField       =   "sts_name"
               BoundColumn     =   "sts_code"
               Text            =   ""
               Object.DataMember      =   "POSTATUS"
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
            Bindings        =   "PurchaseOrder.frx":05E0
            DataField       =   "poi_comm"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   1920
            TabIndex        =   29
            Top             =   510
            Width           =   1830
            DataFieldList   =   "stk_stcknumb"
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
            stylesets(0).Picture=   "PurchaseOrder.frx":0612
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
            stylesets(1).Picture=   "PurchaseOrder.frx":062E
            stylesets(1).AlignmentText=   1
            HeadFont3D      =   4
            DefColWidth     =   5292
            BeveColorScheme =   1
            ForeColorEven   =   8388608
            BackColorEven   =   16771818
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   291
            Columns.Count   =   2
            Columns(0).Width=   2566
            Columns(0).Caption=   "Code"
            Columns(0).Name =   "Code"
            Columns(0).DataField=   "stk_stcknumb"
            Columns(0).FieldLen=   256
            Columns(1).Width=   9578
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "stk_desc"
            Columns(1).FieldLen=   256
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Object.DataMember      =   "StockMasterLookup"
            DataFieldToDisplay=   "stk_stcknumb"
         End
         Begin VB.Frame fra_Quantity 
            Height          =   1320
            Left            =   150
            TabIndex        =   36
            Top             =   2280
            Width           =   6705
            Begin VB.TextBox txtSecRequested 
               DataField       =   "poi_secoreqdqty"
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
               Left            =   60
               TabIndex        =   39
               Top             =   960
               Width           =   945
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
               Left            =   60
               TabIndex        =   37
               Top             =   390
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
               Left            =   5640
               TabIndex        =   117
               Top             =   360
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
               Left            =   4320
               TabIndex        =   116
               Top             =   360
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
               Left            =   3000
               TabIndex        =   123
               Top             =   360
               Width           =   720
            End
            Begin MSDataListLib.DataCombo dcboSecUnit 
               Bindings        =   "PurchaseOrder.frx":064A
               DataField       =   "poi_secouom"
               DataMember      =   "POITEM"
               Height          =   315
               Left            =   1035
               TabIndex        =   40
               Top             =   960
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "uni_desc"
               BoundColumn     =   "uni_code"
               Text            =   ""
               Object.DataMember      =   "SECONDARYUNIT"
            End
            Begin MSDataListLib.DataCombo dcboUnit 
               Bindings        =   "PurchaseOrder.frx":065B
               DataField       =   "poi_primuom"
               DataMember      =   "POITEM"
               Height          =   315
               Left            =   1035
               TabIndex        =   38
               Top             =   390
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "uni_desc"
               BoundColumn     =   "uni_code"
               Text            =   ""
               Object.DataMember      =   "GET_UNIT_OF_MEASURE"
            End
            Begin VB.Label Label12 
               Caption         =   "Secondary Unit"
               Height          =   195
               Left            =   1080
               TabIndex        =   129
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label10 
               Caption         =   "Qty. Rqd. "
               Height          =   195
               Left            =   90
               TabIndex        =   128
               Top             =   720
               Width           =   960
            End
            Begin VB.Label lbl_Delivered 
               Caption         =   "Delivered"
               Height          =   225
               Left            =   3000
               TabIndex        =   83
               Top             =   120
               Width           =   690
            End
            Begin VB.Label lbl_Shipped 
               Caption         =   "Shipped"
               Height          =   225
               Left            =   4320
               TabIndex        =   82
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lbl_Requested 
               Caption         =   "Qty. Rqd. "
               Height          =   225
               Left            =   90
               TabIndex        =   81
               Top             =   180
               Width           =   870
            End
            Begin VB.Label lbl_Unit 
               Caption         =   "Primary Unit"
               Height          =   195
               Left            =   1080
               TabIndex        =   80
               Top             =   180
               Width           =   1560
            End
            Begin VB.Label lbl_Inventory2 
               Caption         =   "Inventory"
               Height          =   225
               Left            =   5640
               TabIndex        =   79
               Top             =   120
               Width           =   735
            End
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboManNumber 
            Bindings        =   "PurchaseOrder.frx":066C
            DataField       =   "poi_manupartnumb"
            DataMember      =   "POITEM"
            Height          =   315
            Left            =   1920
            TabIndex        =   30
            Top             =   855
            Width           =   1830
            DataFieldList   =   "stm_partnumb"
            _Version        =   196617
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
            stylesets(0).Picture=   "PurchaseOrder.frx":067D
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
            stylesets(1).Picture=   "PurchaseOrder.frx":0699
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
            Columns(0).DataField=   "stm_partnumb"
            Columns(0).FieldLen=   256
            Columns(1).Width=   4683
            Columns(1).Caption=   "Manufacturer"
            Columns(1).Name =   "Manufacturer"
            Columns(1).DataField=   "stm_manucode"
            Columns(1).FieldLen=   256
            _ExtentX        =   3238
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            DataFieldToDisplay=   "stm_partnumb"
         End
         Begin VB.Label lbl_PartNum 
            Caption         =   "Manufacturer P/N"
            Height          =   225
            Left            =   120
            TabIndex        =   131
            Top             =   855
            Width           =   1815
         End
         Begin VB.Label lbl_SerialNum 
            Caption         =   "Serial Number"
            Height          =   225
            Left            =   120
            TabIndex        =   130
            Top             =   1875
            Width           =   1785
         End
         Begin VB.Label Label8 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   1080
            TabIndex        =   127
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
            TabIndex        =   126
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
            Caption         =   "Price"
            Height          =   225
            Left            =   7080
            TabIndex        =   98
            Top             =   2400
            Width           =   1185
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
            Left            =   7080
            TabIndex        =   91
            Top             =   3000
            Width           =   1320
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
         TabIndex        =   48
         Top             =   960
         Width           =   8300
      End
      Begin VB.TextBox txtRemarks 
         DataField       =   "por_remk"
         DataMember      =   "POREM"
         Height          =   5175
         Left            =   -74760
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   1020
         Width           =   8295
      End
      Begin MSDataGridLib.DataGrid dgRecepients 
         Height          =   2055
         Left            =   -72840
         TabIndex        =   28
         Top             =   3840
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   3625
         _Version        =   393216
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
         TabIndex        =   23
         Top             =   3360
         Width           =   6144
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74640
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Addterms 
         Caption         =   "Add Clause"
         Height          =   288
         Left            =   -74730
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   528
         Width           =   1695
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   1650
         Left            =   -74700
         TabIndex        =   24
         Top             =   3735
         Width           =   1635
         Begin VB.OptionButton opt_SupFax 
            Caption         =   "Supplier's"
            Height          =   288
            Left            =   60
            TabIndex        =   25
            Top             =   336
            Width           =   1440
         End
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   60
            TabIndex        =   26
            Top             =   768
            Width           =   1515
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   60
            TabIndex        =   27
            Top             =   1260
            Width           =   1515
         End
      End
      Begin VB.Frame fra_PO 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   -74775
         TabIndex        =   100
         Top             =   450
         Width           =   8430
         Begin MSDataListLib.DataCombo dcboPO 
            Bindings        =   "PurchaseOrder.frx":06B5
            DataField       =   "po_ponumb"
            DataMember      =   "PO"
            Height          =   315
            Left            =   1560
            TabIndex        =   0
            Top             =   180
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "po_ponumb"
            BoundColumn     =   "po_ponumb"
            Text            =   ""
            Object.DataMember      =   "PO"
         End
         Begin MSDataListLib.DataCombo dcboDocumentType 
            Bindings        =   "PurchaseOrder.frx":06CC
            DataField       =   "po_docutype"
            DataMember      =   "PO"
            Height          =   315
            Left            =   5760
            TabIndex        =   1
            Top             =   165
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "doc_desc"
            BoundColumn     =   "doc_code"
            Text            =   ""
            Object.DataMember      =   "UserDocumentType"
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
            Left            =   3840
            TabIndex        =   102
            Top             =   240
            Width           =   1875
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
            Height          =   228
            Left            =   60
            TabIndex        =   101
            Top             =   240
            Width           =   1488
         End
      End
      Begin VB.Frame fra_LI 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   135
         TabIndex        =   75
         Top             =   480
         Width           =   8520
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            TabIndex        =   122
            Top             =   180
            Width           =   2535
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "poi_ponumb"
            DataMember      =   "POITEM"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   50
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
         TabIndex        =   22
         Top             =   660
         Width           =   6015
         _Version        =   196617
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
         stylesets(0).Picture=   "PurchaseOrder.frx":06D9
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
         stylesets(1).Picture=   "PurchaseOrder.frx":06F5
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
         Columns(0).Width=   5292
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         TabNavigation   =   1
         _ExtentX        =   10610
         _ExtentY        =   4101
         _StockProps     =   79
         DataMember      =   "POREC"
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
      TabIndex        =   52
      Top             =   6480
      Width           =   3660
   End
End
Attribute VB_Name = "frm_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fm As FormMode
Dim SysUom As String
Dim Ponumb As String
Dim DefSite As String
Dim ComFactor As Double
Dim FNamespace As String
Dim Requery(3) As Boolean
Dim AddingRecord As Boolean
Private vPKValues() As Variant
Dim DeleteInProgress As Boolean
Dim WithEvents st As frm_ShipTerms
Attribute st.VB_VarHelpID = -1
Dim WithEvents rsPO As ADODB.Recordset
Attribute rsPO.VB_VarHelpID = -1
Dim WithEvents rsPOREM As ADODB.Recordset
Attribute rsPOREM.VB_VarHelpID = -1
Dim WithEvents rsPOITEM As ADODB.Recordset
Attribute rsPOITEM.VB_VarHelpID = -1
Dim WithEvents rsPOCLAUSE As ADODB.Recordset
Attribute rsPOCLAUSE.VB_VarHelpID = -1
Dim WithEvents comsearch As frm_StockSearch
Attribute comsearch.VB_VarHelpID = -1
Dim WithEvents rsrecepList As ADODB.Recordset
Attribute rsrecepList.VB_VarHelpID = -1
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'set ship terms form

Private Sub cmd_Addterms_Click()
On Error Resume Next
    Set st = New frm_ShipTerms
    
    st.Show
    If Err Then Err.Clear
End Sub

'click remove delete current recepient record

Private Sub cmd_Remove_Click()
On Error Resume Next
    rsrecepList.Delete
    If Err Then Err.Clear
End Sub


Private Sub CmdcopyLI_Click(Index As Integer)
' Do it only if the user is either modifiying or Adding a record
 
 Select Case (Index)
   
   Case 0
    
    If rsPOITEM.State <> adStateClosed Then
    
    If rsPOITEM.RecordCount > 0 Then
     
      If fm = mdCreation Or fm = mdModification Then
    
       Load FrmCopyPOItems
       FrmCopyPOItems.Show
      End If
    
     End If
     
    End If
   
   Case 1
       
       
   If rsPOREM.State <> adStateClosed Then
    
    If rsPOREM.RecordCount > 0 Then
     
      If fm = mdCreation Or fm = mdModification Then
        Load FrmCopyPORemarks
        FrmCopyPORemarks.Show
      End If
      
     End If
     
   End If
    
   Case 2
   
      If rsPOCLAUSE.State <> adStateClosed Then
    
    If rsPOCLAUSE.RecordCount > 0 Then
     
      If fm = mdCreation Or fm = mdModification Then
    
       Load FrmCopyPOClause
       FrmCopyPOClause.Show
      End If
    
     End If
     
    End If
   
 End Select
End Sub

'in creattion mode, click remove delete current recepient record

Private Sub cmdremove_Click()
On Error Resume Next

    If fm = mdCreation Then rsrecepList.Delete
    
    If Err Then Err.Clear
End Sub

'assign value to poitem fields

Private Sub comsearch_Completed(Cancelled As Boolean, sStockNumber As String)
On Error Resume Next

    comsearch.Hide
    rsPOITEM!poi_comm = sStockNumber
    ssdcboCommoditty.text = sStockNumber
    rsPOITEM!poi_desc = comsearch.Description
    
    Call GetUnits(rsPOITEM!poi_comm & "")
    If Err Then Err.Clear
End Sub

'set memory free

Private Sub comsearch_Unloading(Cancel As Integer)
On Error Resume Next
    Set comsearch = Nothing
End Sub

'assign value to po company code

Private Sub dcboCompany_Click(Area As Integer)
On Error Resume Next

    If Area = 2 Then
    
        Call GetLocations(dcboCompany.BoundText)
        If Editting Then rsPO!po_compcode = dcboCompany.BoundText
        dcboInvLocation.text = ""
        dcboCompany.SelStart = 0
        dcboCompany.SelLength = 0

        
    ElseIf Area = 0 Then
    
        Call GetActiveCompanies(True)
    End If
End Sub

'call function

Private Sub dcboCompany_LostFocus()
On Error Resume Next
    Call dcboCompany_Click(2)
End Sub

'assign value to po company code

Private Sub dcboCompany_Validate(Cancel As Boolean)
'On Error Resume Next

 '   If rsPO.editmode <> adEditNone Then
  '      rsPO!po_compcode = dcboCompany.BoundText
   ' End If
   
    Cancel = False
    With dcboCompany
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
   
End Sub

'assign value to po currency code

Private Sub dcboCurrency_Click(Area As Integer)
On Error Resume Next
    If Area = 2 Then _
        If Editting Then rsPO!po_currcode = dcboCurrency.BoundText
    If Err Then Err.Clear
End Sub

'call function

Private Sub dcboCurrency_LostFocus()
On Error Resume Next
    Call dcboCurrency_Click(2)
End Sub

Private Sub dcboCurrency_Validate(Cancel As Boolean)
    Cancel = False
    With dcboCurrency
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub

'assige value to poitem custom category

Private Sub dcboCustomCategory_Click(Area As Integer)
On Error Resume Next
    If Area = 2 Then rsPOITEM!poi_custcate = dcboCustomCategory.BoundText
    If Err Then Err.Clear
End Sub

'assign value to po document type

Private Sub dcboDocumentType_Click(Area As Integer)
    If Area = 2 Then
        On Error Resume Next
        If Editting Then rsPO!po_docutype = dcboDocumentType.BoundText
        
        If Err Then Err.Clear
    
        If Area = 2 And fm = mdCreation Then Call GetDistributors("USER")
    End If
End Sub

Private Sub dcboDocumentType_LostFocus()
    With dcboDocumentType
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub

'assign value to inventory location field

Private Sub dcboInvLocation_Click(Area As Integer)
On Error Resume Next
  'Debug.Print "area is " & Area
    If Area = 2 Then
    
        If Editting Then _
            If rsPO!po_invloca <> dcboInvLocation.BoundText Then _
                rsPO!po_invloca = dcboInvLocation.BoundText
       
        If dcboInvLocation.SelLength > 0 Then
            dcboInvLocation.SelStart = 0
            dcboInvLocation.SelLength = 0
        End If
    ElseIf Area = 0 Then
    
        Set dcboInvLocation.RowSource = Nothing
        Call deIms.LOCATION(FNamespace)
        
        
        rsPO!po_invloca = dcboInvLocation.BoundText
        Set dcboInvLocation.RowSource = deIms
'        rsPO!po_invloca = dcboInvLocation.BoundText
        
    End If
    
    
    If Err Then Err.Clear
End Sub

'call function

Private Sub dcboInvLocation_LostFocus()
On Error Resume Next
    Call dcboInvLocation_Click(2)
End Sub

Private Sub dcboInvLocation_Validate(Cancel As Boolean)
    Dim matcher As New ADODB.Recordset
    Cancel = False
    With dcboInvLocation
        If .text <> "" Then
            Set matcher = deIms.Recordsets("INVENTORYLOCATION").Clone
            matcher.MoveFirst
            matcher.Find "loc_name = '" + .text + "'"
            If matcher.EOF Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
           End If
        End If
    End With
End Sub

'assign value to po originator field

Private Sub dcboOriginator_Click(Area As Integer)
'On Error Resume Next

    'FG on 8/15/00 Fix the combo TBUF listing all records when clicked in the field
    'Combo was ok
    'If Area = 2 Then
     '  If Editting Then rsPO!po_orig = dcboOriginator.BoundText
       
    'ElseIf Area = 0 Then
    
        'deIms.rsActiveOriginator.Requery
        'Set dcboOriginator.RowSource = Nothing
        'Call deIms.ActiveOriginator(FNamespace)
        'dcboOriginator.RowMember = "ActiveOriginator"
        
        'Set dcboOriginator.RowSource = deIms
        
    'End If
    
    'If Err Then Err.Clear
    dcboOriginator.SelStart = 0
    dcboOriginator.SelLength = 0
End Sub

'call function

Private Sub dcboOriginator_LostFocus()
'On Error Resume Next
    'Call dcboOriginator_Click(2)
    

    With dcboOriginator
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
    
End Sub

'set data to po combo

Public Sub dcboPO_Click(Area As Integer)
On Error Resume Next
'Dim str As String

    Select Case Area
        Case 0
            dcboPO.Refresh
        Case 2
            Err.Clear
            
            deIms.rsPO.CancelBatch
            Call rsPO.CancelBatch(adAffectCurrent)
            Call rsPO.Find("po_ponumb = '" & dcboPO & "'", 0, adSearchForward, adBookmarkFirst)
            Call rsPO.Find("po_ponumb = '" & dcboPO & "'", 0, adSearchForward, adBookmarkFirst)
            
            If rsPO.EOF Then
                rsPO.MoveLast
            Else
                Call rsPO.Move(0)
            End If
            
            Ponumb = dcboPO
            
            Requery(3) = True
            Requery(1) = True
            Requery(0) = True
            Requery(2) = True
            
            ToggleNavButtons
            Set NavBar1.Recordset = rsPO
    End Select
    If Err Then Err.Clear
End Sub

Private Sub dcboPO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub dcboPO_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
End Sub

'check enter po number exist or not show message

Private Sub dcboPO_Validate(Cancel As Boolean)
Dim cmd As ADODB.Command

On Error Resume Next

    'rsPO!po_ponumb = Trim$(dcboPO)


    If fm = mdCreation Then
        Set cmd = deIms.Commands("POEXIST")
        cmd.Parameters("NAMESPACE") = FNamespace
        
        cmd.Parameters("PONUM") = dcboPO
        Call cmd.Execute(0, 0, adExecuteNoRecords)
        
        If cmd.Parameters(0) > 0 Then
            Cancel = True
            dcboPO.SetFocus
            
            'Modified by Juan (9/27/2000) for Multilingual
            msg1 = translator.Trans("M00079") 'J added
            MsgBox IIf(msg1 = "", "Po Number Already exist", msg1) 'J modified
            '---------------------------------------------
            
        Else
            rsPO!po_ponumb = dcboPO
        End If
    End If
    
End Sub

'assign value to po priority field

Private Sub dcboPriority_Click(Area As Integer)
'On Error Resume Next

    'If Area = 2 Then _
        If Editting Then _
            'rsPO!po_priocode = dcboPriority.BoundText
        
 '   If Err Then Err.Clear
End Sub

'call function

Private Sub dcboPriority_LostFocus()
On Error Resume Next
    'Call dcboPriority_Click(2)
End Sub

Private Sub dcboPriority_Validate(Cancel As Boolean)
    Cancel = False
    With dcboPriority
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub

'assign value to poitem primary quantity field

Private Sub dcboSecUnit_Click(Area As Integer)
On Error Resume Next

    If Area = 2 Then
        rsPOITEM!poi_primuom = dcboSecUnit.BoundText
    End If
        
    If Area = 0 Then 'J
        rsPOITEM!poi_primuom = dcboSecUnit.BoundText 'J
    End If 'J
    
    
    If Err Then Call LogErr(Name & "::dcboSecUnit_Click", Err.Description, Err, True)
End Sub

'assign value to poitem secondary unit field

Private Sub dcboSecUnit_Validate(Cancel As Boolean)
On Error Resume Next

    With dcboSecUnit
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
            Else
                .SelStart = 0
                .SelLength = 0
                rsPOITEM!poi_secouom = dcboSecUnit.BoundText
            End If
        End If
    End With


    If Err Then Err.Clear
End Sub

'get value for ship to combo

Private Sub dcboShipto_Click(Area As Integer)
On Error Resume Next

    If Area = 0 Then
    
        Set dcboShipto.RowSource = Nothing
        
        Call deIms.rsActiveShipTo.Requery
        Call deIms.ActiveShipTo(FNamespace)
        dcboShipto.RowMember = "ActiveShipTo"
        
        Set dcboShipto.RowSource = deIms
        
        If Err Then Err.Clear
    End If

    dcboShipto.SelStart = 0
    dcboShipto.SelLength = 0
End Sub

Private Sub dcboShipto_Validate(Cancel As Boolean)
    Cancel = False
    With dcboShipto
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub


'assign value to po supplier code and get recepient mail number
'assign to recepient list

Private Sub dcboSupplier_Click()
On Error Resume Next

Dim STR As String

    STR = Trim$(rsPO!po_suppcode & "")
    If Editting Then rsPO!po_suppcode = dcboSupplier.Value
    
    STR = FixFaxNumber(GetSupplierEmailForPO(FNamespace, STR, deIms.cnIms))
    
    
    
    
    If Len(STR) Then
    
        rsrecepList.Filter = adFilterNone
        rsrecepList.Filter = "porc_rec = '" & STR & "'"
        If Not rsrecepList.EOF Then rsrecepList.Delete
    
        
        rsrecepList.Update
        rsrecepList.Filter = adFilterNone
    End If
    
   GetSupplierFax
   
   
    dcboSupplier.SelStart = 0
    dcboSupplier.SelLength = 0
    
    

End Sub

'set value to supplier data combo

Private Sub dcboSupplier_DropDown()
On Error Resume Next


    Set dcboSupplier.DataSourceList = Nothing
    
    deIms.rsSuppLookup.Close
    Call deIms.SuppLookup(FNamespace)
    dcboSupplier.DataMemberList = "SuppLookup"
    
    Set dcboSupplier.DataSourceList = deIms
    
    If Err Then Call LogErr(Name & "::dcboSupplier_DropDown", Err.Description, Err, True)
End Sub

'load supplier combo

Private Sub dcboSupplier_InitColumnProps()
On Error Resume Next

    With dcboSupplier
        .Columns.RemoveAll
        Call .Columns.Add(0)
        Call .Columns.Add(1)
        Call .Columns.Add(2)
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .Columns(0).Caption = IIf(msg1 = "", "Supplier", msg1) 'J modified
        '---------------------------------------------
        
        .Columns(0).DataField = "sup_name"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .Columns(1).Caption = IIf(msg1 = "", "City", msg1) 'J modified
        '---------------------------------------------
            
        .Columns(1).DataField = "sup_city"
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .Columns(2).Caption = IIf(msg1 = "", "Phone Number", msg1) 'J modified
        '---------------------------------------------
            
        .Columns(2).DataField = "sup_phonnumb"
        'modified by muzammil /08/04/00
        'Reason - During creating an EXE the sheridan combo box was not displaying the value
        'when executing   (combo name).value,Since it was supposed to pass the value of sup_code.
        'So added sup_code to the combo box.
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .Columns(3).Caption = IIf(msg1 = "", "supplier code", msg1) 'J modified
        '---------------------------------------------
        
        .Columns(3).DataField = "sup_code"
        
    End With
        
        
End Sub

Private Sub dcboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not dcboSupplier.DroppedDown Then dcboSupplier.DroppedDown = True
End Sub

Private Sub dcboSupplier_KeyPress(KeyAscii As Integer)
    With dcboSupplier
        If KeyAscii = 13 Then
            If Not .IsItemInList Then
                Call dcboSupplier_Validate(True)
            End If
        Else
            If .text <> Left(.SelBookmarks(.row), Len(.text)) Then
                .MoveNext
            End If
        End If
    End With
End Sub

Private Sub dcboSupplier_Validate(Cancel As Boolean)
    Dim i
    Cancel = False
    With dcboSupplier
        If .text <> "" Then
            .MoveFirst
            For i = 0 To .Rows
                If Trim(.text) = Trim(.Columns(3).text) Then
                    .SelStart = 0
                    .SelLength = 0
                    Exit Sub
                End If
                .MoveNext
            Next
            msg1 = translator.Trans("M00699")
            MsgBox IIf(msg1 = "", "Invalid Value", msg1)
            .text = ""
            Cancel = True
            .SetFocus
        End If
    End With
End Sub

Private Sub dcboToBeUsedFor_Change()

'rsPO!po_tbuf = dcboToBeUsedFor.BoundText 'M 11/20/00
End Sub

'get to bo user for value to combo

Private Sub dcboToBeUsedFor_Click(Area As Integer)
On Error Resume Next

    'FG on 8/15/00 Fix the combo TBUF listing all records when clicked in the field
    'Combo was ok
    'If Area = 0 Then
    
        Set dcboToBeUsedFor.RowSource = Nothing
        
        deIms.rsActiveTbu.Requery
        If Err Then Err.Clear
        Call deIms.ActiveTbu(FNamespace)
        dcboToBeUsedFor.RowMember = "ActiveTbu"
        Set dcboToBeUsedFor.RowSource = deIms
        
        
        If Err Then Err.Clear
    'ElseIf Area = 2 Then
        
        Debug.Print rsPO!po_tbuf
    'End If
      
    If Err Then Err.Clear
    If dcboToBeUsedFor.SelLength > 0 Then
        dcboToBeUsedFor.SelStart = 0
        dcboToBeUsedFor.SelLength = 0
    End If
    If dcboToBeUsedFor.text <> "" Then
        Call dcboToBeUsedFor_KeyPress(13)
    End If
End Sub

Public Sub dcboToBeUsedFor_KeyPress(KeyAscii As Integer)
    
End Sub


Private Sub dcboToBeUsedFor_Validate(Cancel As Boolean)
    Dim matcher As New ADODB.Recordset
    Cancel = False

    With dcboToBeUsedFor
        If .text <> "" Then
            Set matcher = deIms.Recordsets("ActiveTbu").Clone
            matcher.Find "tbu_name = '" + .text + "'"
            If matcher.EOF Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub

'get value for secondary unit combo

Private Sub dcboUnit_Click(Area As Integer)
On Error Resume Next

    If Area = 2 Then
        If chk_FrmStkMst.Value = 0 Then
            rsPOITEM!poi_secouom = dcboUnit.BoundText
            dcboSecUnit.BoundText = dcboUnit.BoundText
        End If
    End If
    
    If Area = 0 Then 'J added
         If chk_FrmStkMst.Value = 0 Then 'J added
            rsPOITEM!poi_secouom = dcboUnit.BoundText 'J added
            dcboSecUnit.BoundText = dcboUnit.BoundText 'J added
        End If 'J added
    End If 'J added
    

    If Err Then Call LogErr(Name + "::dcboUnit_Click", Err.Description, Err.number, True)
End Sub

'get value for poitem primary unit

Private Sub dcboUnit_Validate(Cancel As Boolean)
On Error Resume Next
    
    
    With dcboUnit
        If .text <> "" Then
            If Not .MatchedWithList Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
            Else
                .SelStart = 0
                .SelLength = 0
                rsPOITEM!poi_primuom = dcboUnit.BoundText
            End If
        End If
    End With
    
End Sub

'call function sortgrid

Private Sub dgRecepients_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    Call SortGrid(dgRecepients.DataSource, dgRecepients, ColIndex)
End Sub

'cancel recepient list update

Private Sub dgRecipientList_AfterColUpdate(ByVal ColIndex As Integer)
Dim STR As String


    If fm = mdCreation Then
        STR = dgRecipientList.Columns(0).text
        rsrecepList("porc_rec") = STR
        
        dgRecipientList.CancelUpdate
    End If

End Sub

Private Sub dgRecipientList_AfterUpdate(RtnDispErrMsg As Integer)
    RtnDispErrMsg = 0
End Sub

'set size and load recepient list combo

Private Sub dgRecipientList_InitColumnProps()


    With dgRecipientList
        .Columns.RemoveAll
        
        Call .Columns.Add(0)
        .Columns(0).Width = 5790
        .Columns(0).DataField = "porc_rec"
        
    End With
    'dgRecipientList.Redraw = True
    
   
End Sub

Private Sub DTPicker_poDate_Change()
On Error Resume Next
    rsPO!PO_Date = DTPicker_poDate.Value
    If Err Then Err.Clear
End Sub


'assign value to po requecy deliver

Private Sub dtpRequestedDate_Change()
On Error Resume Next
    rsPO!po_reqddelvdate = dtpRequestedDate.Value
    If Err Then Err.Clear
End Sub

'unlock custom category data combo

Private Sub dcboCustomCategory_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    dcboCustomCategory.locked = False
End Sub

'lock custom category data combo

Private Sub dcboCustomCategory_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    dcboCustomCategory.locked = Not NavBar1.NewEnabled
End Sub

'unlock po combo

Private Sub dcboPO_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    dcboPO.locked = False
End Sub

'unload form, close recordset and set memory free

Private Sub Form_Unload(Cancel As Integer)
Dim ctl
On Error Resume Next
    
    Hide
    With deIms
        .rsActiveTbu.Close
        .rsActiveShipTo.Close
        .rsActiveOriginator.Close
        .rsStockMasterLookup.Close
    End With
        
    For Each ctl In Controls
        Set ctl.DataSource = Nothing
    Next
    
    Set rsPO = Nothing
    Set rsPOREM = Nothing
    Set rsPOITEM = Nothing
    Set rsrecepList = Nothing
    
    If Err Then Err.Clear
    Set frm_Purchase = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub


'get email parameters and call send email and fax function

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
        Params(1) = "ponumb=" & rsPO!po_ponumb & ""
        
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
    Dim messageSubject As String: messageSubject = IIf(msg1 = "", "Purchase Order ", msg1 + " ") & Ponumb 'J modified
    If Len(Label5.Caption) > 0 And Not (Label5.Caption = "0") Then
        msg1 = translator.Trans("L00066") 'J added
        messageSubject = messageSubject & IIf(msg1 = "", "(revision No. ", msg1 + " ") & Label5.Caption & ")" 'J modified
    Else
        msg1 = translator.Trans("M00090") 'J added
        messageSubject = messageSubject & IIf(msg1 = "", "(initial revision)", msg1) 'J modified
    End If
    Call SendEmailAndFax(rsrecepList, "porc_rec", messageSubject, IIf(msg2 = "", "Purchase Order", msg2), "") 'J modified
    '-----------------------------------------
    
    End With
    
    

    If Err Then Err.Clear
End Sub

'get recordsets for po form, date format, set navbar button

Private Sub Form_Load()
Dim Rs As ADODB.Recordset

'Added by Juan Gonzalez (8/28/200) for Multilingual
'Set translator.ActiveForm = ims.
Call translator.Translate_Forms("frm_Purchase")
'-----------------------------------------------------

On Error Resume Next
    
    Ponumb = ""
    GetUnitOfMeasurement
    FNamespace = deIms.NameSpace
    Call DisableButtons(Me, NavBar1)
    DTPicker1.CustomFormat = "MM'/'dd'/'yyyy"
    dtpRequestedDate.CustomFormat = "MM'/'dd'/'yyyy"
    
    
    Call deIms.Unit(FNamespace)
    Call deIms.Custom(FNamespace)
          
    If Err Then Err.Clear
    Call deIms.Shipper(FNamespace)
    Call deIms.Currency(FNamespace)
    Call deIms.Priority(FNamespace)
    Call deIms.POSTATUS(FNamespace)
    Call deIms.TermDelivery(FNamespace)
    
    If Err Then Err.Clear
    Call deIms.Supplier(FNamespace)
    Call deIms.TermCondition(FNamespace)

    
    'Call deIms.StockMasterLookup(FNamespace)
    Call GetCompany(FNamespace, "PE", deIms.cnIms, Ponumb)
    Call deIms.INVENTORYLOCATION(FNamespace, Ponumb)
    
    Ponumb = ""
    If Err Then Err.Clear

    Set dcboCurrency.RowSource = deIms
    Set dcboPriority.RowSource = deIms
    Set dcboSupplier.DataSourceList = deIms
    
    Set dcbostatus(0).RowSource = deIms
    Set dcbostatus(1).RowSource = deIms
    Set dcbostatus(2).RowSource = deIms
    Set dcbostatus(3).RowSource = deIms
    
    Set dcbostatus(4).RowSource = deIms
    Set dcbostatus(5).RowSource = deIms
    Set dcbostatus(6).RowSource = deIms
    Set dcbostatus(7).RowSource = deIms
    
    Set dcboInvLocation.RowSource = deIms
    'Set ssdcboCommoditty.DataSourceList = deIms
    
    Set dcboCustomCategory.RowSource = deIms
    Set ssdcboShipper.DataSourceList = deIms
    
    Set ssdcboDelivery.DataSourceList = deIms
    Set ssdcboCondition.DataSourceList = deIms
    Set ssdcboCategoryCode.DataSourceList = deIms
    Set ssdcboRequisition.DataSourceList = GetRequisitions(FNamespace, deIms.cnIms)
    If deIms.rsPO.State And adStateOpen = adStateOpen Then deIms.rsPO.Close
        
    If Err Then Err.Clear
    
    Call GetActiveCompanies(False)
    Call deIms.PO(FNamespace)
    
    Set rsPO = deIms.rsPO
    Ponumb = deIms.rsPO!po_ponumb
  
     
    If Err Then Err.Clear
    
    rsPO.CancelUpdate
    
    sst_PO.Tab = 0
    'AddingRecord = True
    ReDim vPKValues(2, 0)
    Set NavBar1.Recordset = rsPO
    Call deIms.GETSYSSITE(FNamespace, DefSite)
    If Err Then Err.Clear
    
    Set rsPOREM = deIms.rsPOREM
    Set rsPOITEM = deIms.rsPOITEM
    Set rsrecepList = deIms.rsPOREC
    Set rsPOCLAUSE = deIms.rsPOCLAUSE
    'NavBar1.Updatable = CBool(GetDocumentType(False))
    
    Call GetUnits("")
    Call GetDocumentType(True)
    dcboDocumentType.DataMember = "PO"
    
    Visible = False
    If Err Then Err.Clear
    
    Call BindAll
    NavBar1.EditEnabled = True
    NavBar1.EditVisible = True
    NavBar1.AllowAddNew = NavBar1.NewEnabled
    
    NavBar1.Width = 0
    NavBar1.Height = 0
    Set dcboPO.RowSource = deIms
    Set dcboPO.DataSource = deIms
    
    Call ChangeMode(mdVisualization)
    
    Call GetActiveStockNumbers(False)
    
    AddingRecord = False
    
    If SysUom = "seco" Then
        txt_Requested.Enabled = False
        
    ElseIf SysUom = "prim" Then
        txtSecRequested.Enabled = False
        
    Else
    
        txt_Requested.Enabled = True
        txtSecRequested.Enabled = True
    End If
    
        
    Set NavBar1.Recordset = rsPO
    NavBar1.NextEnabled = sst_PO.Tab <> 0
    NavBar1.LastEnabled = sst_PO.Tab <> 0
    NavBar1.FirstEnabled = sst_PO.Tab <> 0
    NavBar1.PreviousEnabled = sst_PO.Tab <> 0
    frm_Purchase.Caption = frm_Purchase.Caption + " - " + frm_Purchase.Tag
        
'Changed de position by Juan from dcboOriginator_click
        deIms.rsActiveOriginator.Requery
        Set dcboOriginator.RowSource = Nothing
        Call deIms.ActiveOriginator(FNamespace)
        dcboOriginator.RowMember = "ActiveOriginator"
        
        Set dcboOriginator.RowSource = deIms
'------------------------------------------------------
    
    
'Changed de position by Juan from dcboToBeUsedFor_click
        Set dcboToBeUsedFor.RowSource = Nothing
        
        deIms.rsActiveTbu.Requery
        If Err Then Err.Clear
        Call deIms.ActiveTbu(FNamespace)
        dcboToBeUsedFor.RowMember = "ActiveTbu"
        Set dcboToBeUsedFor.RowSource = deIms
        
        
        If Err Then Err.Clear
        
        'Debug.Print rsPO!po_tbuf
'-------------------------------------------------------
    
End Sub

'set recordset location,lock type, cursor type

Private Function GetRecset() As ADODB.Recordset
On Error Resume Next

    Set GetRecset = New ADODB.Recordset
    
    With GetRecset
            
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .ActiveConnection = deIms.cnIms
        .LockType = adLockBatchOptimistic
    End With
        
End Function

'set form control

Private Sub BindControls(sDataMember As String)
On Error Resume Next
Dim ctl As Control
Dim STR As String

    sDataMember = UCase(sDataMember)
    
    For Each ctl In Me.Controls
        STR = UCase(ctl.DataMember)
        
        If Err Then Err.Clear
        
        If Len(STR) Then
            If STR = sDataMember Then
                Set ctl.DataSource = Nothing
                Set ctl.DataSource = deIms
                
                ctl.DataMember = sDataMember
            End If
        End If
        
        If Err Then Err.Clear
    Next ctl

    Set ssdcboCommoditty.DataSource = Nothing
    Set ssdcboCommoditty.DataSourceList = Nothing

    Set ssdcboCommoditty.DataSource = deIms
    Set ssdcboCommoditty.DataSourceList = deIms
End Sub

'set recordsets update to cancel and close recordsets

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

    NavBar1.CancelUpdate
    Call rsPO.CancelUpdate
    Call rsPOREM.CancelUpdate
    Call rsPOITEM.CancelUpdate
    Call rsPOCLAUSE.CancelUpdate
    Call rsrecepList.CancelUpdate
    

    If Err Then Err.Clear
'    If fm = mdModification Or fm = mdCreation Then
'
'        If CheckPoFields Then
'
'            Select Case MsgBox("Save Changes", vbYesNoCancel)
'
'                Case vbYes
'
'                    If ValidatePOData Then
'                        rsPO.UpdateBatch
'
'                    Else
'                        Cancel = True: MsgBox "error saving po": Exit Sub
'                    End If
'
'                    If ValidatePOREMData Then
'                        rsPOREM.UpdateBatch
'
'                    Else
'                        Cancel = True: MsgBox "error saving Remarks": Exit Sub
'                    End If
'
'                    If ValidatePORECData Then
'                        rsrecepList.UpdateBatch
'                    Else
'                        Cancel = True: MsgBox "error saving Recipients ": Exit Sub
'                    End If
'
'                    If ValidatePOClauseData Then
'                        rsPOCLAUSE.UpdateBatch
'                    Else
'                        Cancel = True: MsgBox "error saving Trms and condition": Exit Sub
'                    End If
'
'                Case vbNo
'
'                    rsPO.CancelBatch
'                    rsPOREM.CancelBatch
'                    rsPOITEM.CancelBatch
'                    rsPOCLAUSE.CancelBatch
'                    rsrecepList.CancelBatch
'
'                Case vbCancel
'                    Cancel = True: Exit Sub
'
'            End Select
'
'        End If
'
'    End If
        
    deIms.cnIms.RollbackTrans
    
    With deIms
        .rsPO.Close
        .rsPOREC.Close
        .rsPOREM.Close
        .rsPOITEM.Close
        .rsPOCLAUSE.Close
    End With
    
    If Err Then Err.Clear
End Sub

'before save records validate data format,if po number was not entered
'get auto po number,set form mode

Private Sub NavBar1_BeforeSaveClick()
On Error Resume Next

Dim STR As String

    Call ValidateControls
    
    Call WriteStatus("Preparing to save PO")
    
    If fm = mdVisualization Then Exit Sub
    
    If sst_PO.Tab = 0 Then
        NavBar1.AllowUpdate = True
        If ((IsNull(rsPO!po_ponumb)) Or (IsEmpty(rsPO!po_ponumb)) Or (Len(Trim$(rsPO!po_ponumb)) = 0)) Then
            Call deIms.GetAutoNumber(rsPO!po_docutype, FNamespace, Ponumb)
            
            Ponumb = Trim$(Ponumb)
            Call WriteStatus("Retrieving Auto Number")
            
            If Ponumb = "" Then
            
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00017") 'J added
                MsgBox IIf(msg1 = "", "Auto Numbering is not set for this document Type", msg1) 'J modified
                '---------------------------------------------
                
                NavBar1.AllowUpdate = False: Exit Sub
            End If
            
            If Len(Ponumb) Then
                rsPO!po_ponumb = Ponumb
                
                'Modified by Juan (9/13/2000) for Multilingual
                msg1 = translator.Trans("M00018") 'J added
                MsgBox IIf(msg1 = "", "Please note that your transaction number will be ", msg1 + " ") & Ponumb
                '---------------------------------------------
                
            Else
                'MsgBox "Error Retrieving Auto number"
                'NavBar1.AllowUpdate = False: Exit Sub
            End If
    
        
        End If
    End If

    DoEvents
    If Err Then Err.Clear
    If sst_PO.Tab = 0 Then
        Call WriteStatus("Checking PO Data")
        NavBar1.AllowUpdate = CheckPoFields
        
    ElseIf sst_PO.Tab = 2 Then
        
        NavBar1.AllowUpdate = CheckLIFields
    End If
        If NavBar1.AllowUpdate = True Then NavBar1.AllowUpdate = fm <> mdVisualization
    
    DoEvents
    If NavBar1.AllowUpdate And sst_PO.Tab = 0 Then Call InsertPoRevision(Ponumb)
End Sub

'on cancel click, cancel updatebatch

Private Sub NavBar1_OnCancelClick()

On Error Resume Next

    If sst_PO.Tab = 0 Then
        rsPOREM.CancelBatch
        rsPOITEM.CancelBatch
        rsrecepList.CancelBatch
        rsPO.CancelBatch adAffectCurrent
        Call ChangeMode(mdVisualization)
        
        Requery(3) = True
        Requery(1) = True
        Requery(0) = True
        Requery(2) = True
        
    End If
    
    Call NavBar1.Recordset.CancelBatch(adAffectCurrent)
    
  'kin li modity for po combo visualization
  
      If fm = mdVisualization Then
        dcboPO.Enabled = True
        
    End If
    If Err Then Err.Clear
End Sub

'call function sendmessage

Private Sub NavBar1_OnCloseClick()
On Error Resume Next

    Call SendMessage(HWND, WM_CLOSE, 0, 0)
    
    If Err Then Err.Clear
End Sub

'generate a new revision number, check po number status
'check document type, check po approve status, show message


Private Sub NavBar1_OnEditClick()
On Error Resume Next
Dim Numb As Integer
Dim Msg As String, Style As String

        Numb = Label5.Caption + 1


'Modified by Juan (9/13/2000) for Multilingual
msg1 = translator.Trans("M00019") 'J added
msg2 = translator.Trans("M00020") 'J added
Msg = IIf(msg1 = "", "You will generate a new revision # ", msg1 + " ") & Numb & IIf(msg2 = "", ". Do you want to continue ?", msg2) 'J modified
'---------------------------------------------

Style = vbYesNo + vbCritical + vbDefaultButton2


    If rsPO!po_stas & "" = "CA" Or rsPO!po_stas & "" = "CL" Then _
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00021") 'J added
        MsgBox IIf(msg1 = "", "This transaction is closed. It cannot be modified.", msg1): Exit Sub
        '----------------------------------------------
        
    End If
    If deIms.CanUserEditDocType(CurrentUser, dcboDocumentType.BoundText) Then
        Call ChangeMode(mdModification)
         
        If Not Len(Trim(Label1.Caption)) = 0 Then
            If MsgBox(Msg, Style) = vbNo Then
                Call ChangeMode(mdVisualization)
            Else
                Label5.Caption = Numb
            End If
        End If
        
        If ((IsPoApprove(rsPO!po_ponumb & "")) Or _
            (rsPO!po_revinumb & "" > 0)) Then
            
            dcboDocumentType.Enabled = False
        End If
    Else
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00022") 'J added
        MsgBox IIf(msg1 = "", "You are not authorized to edit this transaction.", msg1) 'J modified
        '---------------------------------------------
        
        Call ChangeMode(mdVisualization)
    End If
    
    
    
    'FIX KL 8/8/00 to fix PO NON STOCK where commodity# could be modified
    If chk_FrmStkMst.Value = vbChecked Then
'        txt_Descript.locked = True
         ssdcboCommoditty.Enabled = True
    Else
       
       ssdcboCommoditty.Enabled = False
    End If
    
    If fm = mdModification Then
        dcboPO.Enabled = False
    Else
        dcboPO.Enabled = True
    End If
    
    
    'Modified by muzammil 08/14/00
    'Reason - Enables the user to add a new record to that perticular tab on
    'which the user is  when he clicks on the modify button
      If sst_PO.Tab = 2 Then
         If rsPOITEM.RecordCount = 0 Then
         rsPOITEM.AddNew
         AddLIDef
         End If
     ElseIf sst_PO.Tab = 3 Then
     If rsPOREM.RecordCount = 0 Then
         rsPOREM.AddNew
         AddRemDef
         End If
      ElseIf sst_PO.Tab = 4 Then
      If rsPOCLAUSE.RecordCount = 0 Then  'M
         rsPOCLAUSE.AddNew
         AddClauseDef
         End If
       End If
     
    If Err Then Call LogErr(Name & "::NavBar1_OnEditClick", Err.Description, Err.number, True)
End Sub

'get crystol report parameter and application path


Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = ReportPath + "po.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + Ponumb + ";true"
        
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
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)
    End If
End Sub

'before save get po number, validate data format, set recordsets update
'if update successfully, show message, set form mode


Private Sub NavBar1_OnSaveClick()
Dim STR As String
On Error Resume Next

    If sst_PO.Tab = 0 Then
    
    
        If Ponumb = "" Then Ponumb = dcboPO
        
        STR = Ponumb
        deIms.cnIms.CommitTrans
        
        
        DoEvents
        ValidateControls
        AddingRecord = False
        If Err Then Err.Clear: deIms.cnIms.Errors.Clear
        
        rsPO.Update
        rsPOREM.Update
        rsPOITEM.Update
        'rsPOITEM.UpdateBatch adAffectAllChapters
        'rsPOITEM.UpdateBatch
        rsPOCLAUSE.Update
        rsrecepList.Update
        
       
        
        DoEvents
        'Call rsPO.Move(0)
        Call WriteStatus("Validating Po Data")
        If Not ValidatePOData Then Exit Sub
        
        'rsPO.UpdateBatch
        
        DoEvents
        PutPODataInsert
        Call WriteStatus("Updating PO")
        
        Set dgRecipientList.DataSource = Nothing
        If POExist Then rsrecepList.UpdateBatch
        
        
        If fm = mdCreation Then
            Call GetDistributors("SYS")
            Call WriteStatus("Getting Auto Distribution list")
        End If
            
        DoEvents
        On Error Resume Next
        Set rsPO = deIms.rsPO
        
        DoEvents
        Ponumb = STR
    
        Call WriteStatus("Saving Line items")
    
        If Not SaveClause Then Exit Sub
        If Not SaveRemarks Then Exit Sub
        If Not SaveLineItems Then Exit Sub
        If Not SaveRecipients Then Exit Sub
        Call UpdatePoTotalCost(Ponumb, FNamespace, deIms.cnIms)
    
        
        If Err = 0 Then
            

            STR = Ponumb
            deIms.rsPO.Close
            Call deIms.PO(FNamespace)
            
            Set rsPO = deIms.rsPO
            Set NavBar1.Recordset = rsPO
            
            BindAll
            'Call rsPO.MoveFirst
            
            'Modified by Juan (9/13/2000) for Multilingual
            msg1 = translator.Trans("L00065") 'J added
            msg2 = translator.Trans("M00024") 'J added
            MsgBox IIf(msg1 = "", "Transaction # ", msg1 + " ") & STR & IIf(msg2 = "", " saved successfully.", " " + msg2) 'J modified
            '---------------------------------------------
            
            Call ChangeMode(mdVisualization)
            Call rsPO.MoveLast
            Call rsPO.Find("po_ponumb = '" & STR & "'", 0, adSearchBackward, adBookmarkLast)
            
            POIChange
            PORemChange
            PORECChange
            POCLAUSEChange
            
            Requery(3) = True
            Requery(1) = True
            Requery(2) = True
            Requery(0) = True
        End If
            
            
    End If
    Call WriteStatus("")
    
    
    If fm = mdVisualization Then
        dcboPO.Enabled = True
        
    End If
    
    
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err, True)
End Sub

'check  recordset po number, set navbar equal to po recordset
'get location data, reset navbar button


Private Sub rsPO_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Dim STR As String

    
    STR = CStr(rsPO!po_ponumb)
    
    If sst_PO.Tab = 0 Then
    
        If ((Ponumb <> STR) Or (adReason = adRsnAddNew)) Then
            
            Ponumb = STR
            Requery(3) = True
            Requery(1) = True
            Requery(0) = True
            Requery(2) = True
                        
        End If
    End If
    
    
    Set NavBar1.Recordset = rsPO
    Call GetLocations(rsPO!po_compcode & "")
    
    
    dcboInvLocation.ReFill
    
    ToggleNavButtons
        
End Sub

'bind all data to controls

Private Sub BindAll()
On Error Resume Next
Dim ctl As Control
Dim STR As String

    For Each ctl In Me.Controls
        STR = ctl.DataMember
    
        If Len(STR) Then
          
'          If ctl.Name = "DTPicker_poDate" Then
'              MsgBox ctl.Name
'           End If
            
            'ctl.DataMember = str
            Set ctl.DataSource = Nothing
            Set ctl.DataSource = deIms
            
            ctl.Refresh
        End If
        
        If Err Then Err.Clear
    Next ctl
    
    Set dcboPO.RowSource = deIms
    Set dcboPO.DataSource = deIms
End Sub

'set keyboad function

Private Sub cboPurchase_KeyPress(KeyAscii As Integer)
On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'set primary unit and secondary unit


Private Sub chk_FrmStkMst_Click()
On Error Resume Next

    If chk_FrmStkMst.Value = vbChecked Then
        txt_Descript.locked = True
        ssdcboCommoditty.Enabled = True
        
    Else
    
        If LCase(SysUom) = "seco" Then
            dcboUnit.Enabled = False
            dcboSecUnit.Enabled = True
            
        ElseIf LCase(SysUom = "prim") Then
        
            dcboUnit.Enabled = True
            dcboSecUnit.Enabled = False
        Else
            dcboUnit.Enabled = True
            dcboSecUnit.Enabled = True
        End If
    
        txt_Descript.locked = False
        
        ssdcboCommoditty.Enabled = False
    End If
    

        
    If rsPO!po_fromstckmast <> chk_FrmStkMst Then
       rsPO!po_fromstckmast = chk_FrmStkMst
    End If
End Sub

'call function to add a recepient to recepient list

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

'call function to add a recepient to recepient list

Private Sub dgRecepients_DblClick()
On Error Resume Next
    Call AddRecepient(dgRecepients.Columns(1).Value)
    
    If Err Then Err.Clear
End Sub

Private Sub dgRecipientList_DblClick()
On Error Resume Next
    'If fm = mdCreation Then rsrecepList.Delete
End Sub

'depend on tab status, set allow add new, check data fields format
'validate text remark, and check text clause


Private Sub NavBar1_BeforeNewClick()
On Error Resume Next


    NavBar1.AllowAddNew = False
    If ((Not (Editting)) And (sst_PO.Tab <> 0)) Then Exit Sub
    
    If sst_PO.Tab = 0 Then
        NavBar1.AllowAddNew = False
        
        rsPO.CancelUpdate
        Call deIms.rsPO.CancelBatch
    
        If Err Then Err.Clear
        deIms.cnIms.Errors.Clear
        
        'deIms.cnIms.BeginTrans
        NavBar1.AllowAddNew = True
    
    ElseIf sst_PO.Tab = 2 Then
        Dim i As Integer
        
        i = rsPOITEM.RecordCount
        If i Then If CheckLIFields = False Then Exit Sub

    ElseIf sst_PO.Tab = 3 Then
        Call txtRemarks_Validate(False)
        
    ElseIf sst_PO.Tab = 4 Then
        Call txtClause_Validate(False)
    End If
    
    DoEvents
    NavBar1.AllowAddNew = ((sst_PO.Tab = 0) Or (fm = mdCreation) Or (fm = mdModification))
    
    If (NavBar1.AllowAddNew) And (sst_PO.Tab = 0) Then
            
        i = GetDocumentType(False)
        
        If i Then
            NavBar1.AllowAddNew = True
            NavBar1.AllowUpdate = True
        Else
            NavBar1.AllowAddNew = True
            NavBar1.AllowUpdate = True
        End If
    
    End If
        
            
    AddingRecord = True
    
    NavBar1.Recordset.Update
    
    DoEvents
    If sst_PO.Tab = 1 Then NavBar1.CancelUpdate
    If Err Then Err.Clear
    
    NavBar1.Width = 0
    
End Sub

'set data values to po recordset

Private Sub NavBar1_OnNewClick()
On Error Resume Next

Dim Rs As New ADODB.Recordset
            
    AddingRecord = False
    Select Case sst_PO.Tab
    
        Case 0
            With rsPO
                
                !po_forwr = 0
                !PO_Date = Date
                !po_ponumb = ""
                !po_suppcode = ""
                !po_site = DefSite
                !po_datesent = Null
                !po_revinumb = "0"
                !po_buyr = CurrentUser
                !po_daterevi = Null
              'Modified by muzammil.Ticket No-41.11/09/00
'                !po_fromstckmast = 1 'M
                !po_fromstckmast = 0  'M
                
                !po_stas = "OH"
                !po_stasdelv = "NR"
                !po_stasship = "NS"
                !po_stasinvt = "NI"
                !po_confordr = False
                !po_reqddelvdate = Date + 1
                !po_npecode = FNamespace
                !po_reqddelvflag = False
                !po_currcode = "USD"
                '!po_docutype = "P"
                dcboPO = ""
                dcboSupplier = ""
                dcboSupplier.BoundText = ""
           End With
            Call ChangeMode(mdCreation)
            
            
        Case 1

        Case 2
        
            AddLIDef
            dcboSecUnit.text = "EACH"
            Call dcboSecUnit_Click(2)
            
        Case 3
            AddRemDef
            
        Case 4
            AddClauseDef
    End Select
    
    NavBar1.Recordset.Update
    'NavBar1.Recordset.MoveLast
    AddingRecord = False
    
    If sst_PO.Tab <> 0 And fm <> mdCreation Then Call ChangeMode(mdModification)
    If Err Then Err.Clear
End Sub

'set data grid, and get data grid information

Private Sub opt_Email_Click()
On Error Resume Next
Dim co As MSDataGridLib.column

    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
    msg1 = translator.Trans("L00121") 'J added
    co.Caption = IIf(msg1 = "", "Email Address", msg1) 'J modified
    '---------------------------------------------
    
    co.DataField = "phd_mail"
    
    dgRecepients.Columns(0).DataField = "phd_name"
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atEmail)
End Sub

'set data grid, and get data grid information

Private Sub opt_FaxNum_Click()
On Error Resume Next
Dim co As MSDataGridLib.column
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
    msg1 = translator.Trans("L00122") 'J added
    co.Caption = IIf(msg1 = "", "Fax Number", "") 'J modified
    '---------------------------------------------
    
    co.DataField = "phd_faxnumb"
    
    dgRecepients.Columns(0).DataField = "phd_name"
     
    Set dgRecepients.DataSource = GetAddresses(deIms.NameSpace, deIms.cnIms, adLockReadOnly, atFax)
End Sub

'set data grid, and get data grid information

Private Sub opt_SupFax_Click()
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim co As MSDataGridLib.column
    
    Set co = dgRecepients.Columns(1)
    
    'Modified by Juan (8/28/2000) for Multilingual
    msg1 = translator.Trans("L00124") 'J added
    co.Caption = IIf(msg1, "Supplier Email", msg1) 'J modified
    '---------------------------------------------
    
    co.DataField = "sup_mail"
    
    dgRecepients.Columns(0).DataField = "sup_name"
    
    Set Rs = New ADODB.Recordset
    
    With Rs
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        Set .ActiveConnection = deIms.cnIms
        .Open ("select sup_name, sup_mail from SUPPLIER where sup_npecode = '" & FNamespace & "' and sup_mail IS NOT NULL and len(sup_mail) > 3 order by 1")
        Set dgRecepients.DataSource = .DataSource
    End With
End Sub

'get units information

Private Sub rsPOITEM_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next


    If Not ((rsPOITEM.BOF) Or (rsPOITEM.EOF)) Then
        Call GetUnits(rsPOITEM!poi_comm & "")
        If Err Then Err.Clear
    End If
End Sub

'check insert record to po remark table

Private Sub rsporem_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo errWillChangeRecord

    If (adReason <> adRsnFirstChange) And (adReason <> adRsnUndoAddNew) And Not AddingRecord Then
        
        If IsEmpty(rsPOREM(0).OriginalValue) Then
        
            If Not PutPOREMDataInsert Then
                adStatus = adStatusCancel
            End If
            
        Else
        
            Select Case adReason
            
                Case adRsnUpdate
                
                    If Not DeleteInProgress Then
                    
                        If Not PutPOREMDataUpdate Then
                            adStatus = adStatusCancel
                        End If
                        
                    End If
                    
                Case adRsnAddNew
                
                    If Not PutPOREMDataInsert Then
                        adStatus = adStatusCancel
                    End If
                    
                Case adRsnDelete
                    DeleteInProgress = True
            End Select
        End If
    End If

    Exit Sub

errWillChangeRecord:

End Sub

'call function to get active stovk number

Private Sub ssdcboCommoditty_DropDown()
On Error Resume Next
    Call GetActiveStockNumbers(True)
End Sub

'set stock number combo datafield and caption

Private Sub ssdcboCommoditty_InitColumnProps()
On Error Resume Next

    With ssdcboCommoditty.Columns
        .RemoveAll
        Call .Add(0)
        Call .Add(1)
    End With
    
    
    With ssdcboCommoditty.Columns(0)
    
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00123") 'J added
        .Caption = IIf(msg1 = "", "Stock No.", msg1) 'J modified
        '---------------------------------------------
        
        .DataField = "stk_stcknumb"
    End With
    
    With ssdcboCommoditty.Columns(1)
    
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00029") 'J added
        .Caption = IIf(msg1 = "", "Description", "") 'J modified
        '---------------------------------------------
        
        .DataField = "stk_desc"
    End With
    
'   With ssdcboCommoditty.Columns(2)
'        .Caption = "Active"
'        .DataField = "stk_flag"
'    End With
    
    ssdcboCommoditty.AllowInput = True
    ssdcboCommoditty.ColumnHeaders = True
    
End Sub

Private Sub ssdcboCommoditty_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
End Sub

Private Sub ssdcboCommoditty_KeyPress(KeyAscii As Integer)
On Error Resume Next

End Sub

Private Sub ssdcboCommoditty_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
End Sub

'call function to validate stock number exist on datadase or not
Private Sub ssdcboCommoditty_Validate(Cancel As Boolean)
'''On Error Resume Next
'''Dim STR As String, OldNum As String
'''
'''
'''    STR = LCase(ssdcboCommoditty.text)
'''    OldNum = LCase(Trim$(rsPOITEM!poi_comm.OriginalValue & ""))
'''
'''    If OldNum <> STR Then
'''
'''        If Not deIms.StockNumberExist(STR, True) Then
'''            Cancel = True
'''
'''            'Modified by Juan (9/27/2000) for Multilingual
'''            msg1 = translator.Trans("L00119") 'J added
'''            msg2 = translator.Trans("M00026") 'J added
'''            MsgBox IIf(msg1 = "", "Stock number", msg1) + " " & STR & " " + IIf(msg2 = "", "does not exist", msg2) 'J modified
'''            '---------------------------------------------
'''
'''        End If
'''
'''    End If
    
End Sub

'assign value to po recordset

Private Sub ssdcboCondition_Click()
On Error Resume Next
    If Editting Then rsPO!po_taccode = ssdcboCondition.Value
End Sub

Private Sub ssdcboCondition_Validate(Cancel As Boolean)
    Cancel = False
    With ssdcboCondition
        If .text <> "" Then
            If Not .IsItemInList Then
                msg1 = translator.Trans("M00699")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub


'assign value to po recordset

Private Sub ssdcboDelivery_Click()
On Error Resume Next
   If Editting Then rsPO!po_termcode = ssdcboDelivery.Value
End Sub

Private Sub ssdcboDelivery_Validate(Cancel As Boolean)
    Cancel = False
    With ssdcboDelivery
        If .text <> "" Then
            If Not .IsItemInList Then
                msg1 = translator.Trans("M00699")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                .text = ""
                Cancel = True
                .SetFocus
            Else
                .SelStart = 0
                .SelLength = 0
            End If
        End If
    End With
End Sub


'call function to assign value to po recordset

Private Sub ssdcboManNumber_GotFocus()
On Error Resume Next
    Call GetUnits(rsPOITEM!poi_comm & "")
End Sub

'assign value to po recordset

Private Sub ssdcboManNumber_Validate(Cancel As Boolean)
On Error Resume Next
    If Len(ssdcboManNumber) Then _
        rsPOITEM!poi_manupartnumb = ssdcboManNumber
        
    If Err Then Err.Clear
End Sub

'SQL statement get po line item information for require

Private Sub ssdcboRequisition_Click()
On Error Resume Next
Dim cmd As ADODB.Command
Dim Rs As ADODB.Recordset
Dim STR As String
Dim StockNumber As String


    deIms.rsSECONDARYUNIT.Close
    deIms.rsGET_UNIT_OF_MEASURE.Close
    
    Set dcboUnit.RowSource = Nothing
    Set dcboUnit.DataSource = Nothing
    Set dcboSecUnit.RowSource = Nothing
    Set dcboSecUnit.RowSource = Nothing

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    lblReqLineitem = ssdcboRequisition.Columns("Item").text
    
    cmd.CommandText = "SELECT * FROM POITEM WHERE poi_npecode = '" & FNamespace & "'"
    cmd.CommandText = cmd.CommandText & " AND poi_liitnumb = '" & lblReqLineitem & "'"
    cmd.CommandText = cmd.CommandText & " and poi_ponumb = ?"
    
    Set Rs = cmd.Execute(0, Array(ssdcboRequisition.text))
    
    'Modified by Juan (9/13/2000) for Multilingual
    msg1 = translator.Trans("M00027") 'J added
    If Rs.EOF And Rs.BOF Then MsgBox IIf(msg1 = "", "Error Retrieving B/R/Q Line", msg1): Exit Sub 'J modified
    If Rs.RecordCount = 0 Then MsgBox IIf(msg1 = "", "Error Retrieving B/R/Q Line", msg1): Exit Sub 'J modified
    '---------------------------------------------
    
    With rsPOITEM
        !poi_afe = Rs("poi_afe") & ""
        !poi_remk = Rs("poi_remk") & ""
        !poi_comm = Rs("poi_comm") & ""
        !poi_desc = Rs("poi_desc") & ""
        !poi_requnumb = Rs("po_ponumb") & ""
        !poi_primuom = Rs("poi_primuom") & ""
        !poi_secouom = Rs("poi_secouom") & ""
        !poi_serlnumb = Rs("poi_serlnumb") & ""
        !poi_custcate = Rs("poi_custcate") & ""
        !poi_requnumb = Rs("poi_ponumb") & ""
        !poi_unitprice = Rs("poi_unitprice") & ""
        
        !poi_endrentdate = Rs("poi_endrentdate") & ""
        !poi_liitreqddate = Rs("poi_liitreqddate") & ""
        !poi_liitrelsdate = Rs("poi_liitrelsdate") & ""
        !poi_starrentdate = Rs("poi_starrentdate") & ""

        
        !poi_requliitnumb = Rs("poi_liitnumb") & ""
        !poi_secoreqdqty = Rs("poi_secoreqdqty") & ""
        !poi_primreqdqty = Rs("poi_primreqdqty") & ""
        !poi_manupartnumb = Rs("poi_manupartnumb") & ""
        
        txt_Price = Rs("poi_unitprice") & ""
        Set ssdcboCommoditty.DataSource = deIms
        dcboUnit.BoundText = Rs("poi_primuom") & ""
        ssdcboCommoditty.text = Rs("poi_comm") & ""
    End With
    
    Set dcboSecUnit.RowSource = deIms
    Set dcboSecUnit.DataSource = deIms
    StockNumber = Rs("poi_comm") & ""
    
    Call GetUnits(rsPOITEM!poi_comm & "")
    ssdcboRequisition.SelStart = 0
    ssdcboRequisition.SelLength = 0
    
End Sub

Private Sub ssdcboRequisition_Validate(Cancel As Boolean)
''''    Cancel = False
''''    With ssdcboRequisition
''''        If .text <> "" Then
''''            If Not .IsItemInList Then
''''                msg1 = translator.Trans("M00699")
''''                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
''''                .text = ""
''''                Cancel = True
''''                .SetFocus
''''            Else
''''                .SelStart = 0
''''                .SelLength = 0
''''            End If
''''        End If
''''    End With
End Sub


Private Sub ssdcboShipper_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not ssdcboShipper.DroppedDown Then ssdcboShipper.DroppedDown = True
End Sub

Private Sub ssdcboShipper_KeyPress(KeyAscii As Integer)
    With ssdcboShipper
        If KeyAscii = 13 Then
            If Not .IsItemInList Then
                Call ssdcboShipper_Validate(True)
            End If
        Else
            If .text <> Left(.SelBookmarks(.row), Len(.text)) Then
                .MoveNext
            End If
        End If
    End With
End Sub

'call shipper combo

Private Sub ssdcboShipper_LostFocus()
On Error Resume Next
    ssdcboShipper_Click
End Sub



Public Sub ssdcboShipper_Validate(Cancel As Boolean)
    Cancel = False
    With ssdcboShipper
        If Not .IsItemInList Then
            msg1 = translator.Trans("M00699")
            MsgBox IIf(msg1 = "", "Shipper doesn't exist", msg1)
            .text = ""
            Cancel = True
            .SetFocus
        Else
            .SelStart = 0
            .SelLength = 0
        End If
    End With
End Sub


Private Sub txt_ChargeTo_Validate(Cancel As Boolean)
'''If Len(txt_ChargeTo) > 25 Then
'''MsgBox "Charge To can not have more than 25 characters", , "Imswin"
'''txt_ChargeTo.SetFocus
'''End If
End Sub

'validate price text box, convert unit price, total price to double data format

Private Sub txt_Price_Validate(Cancel As Boolean)
On Error Resume Next

Dim i As Long
Dim Rs As ADODB.Recordset

    If Not Editting Then Exit Sub
    
    If Len(txt_Price) = 0 Then Exit Sub
    If fm = mdVisualization Then Exit Sub
   
    
    Set Rs = rsPOITEM
    
    If Not IsNumeric(txt_Price) Then
        Cancel = True
        
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00028") 'J added
        MsgBox IIf(msg1 = "", "Price has an invalid value", msg1): Exit Sub 'J modified
        '---------------------------------------------
        
    End If
    
    
    'Modified on  08/07/00 by  / Muzammil.Modified or added lines have 'M' on the right side
    
    
''    If SysUom = "seco" Then                                                           'M
''     rsPOITEM!poi_unitprice = FormatNumber(CDbl(txt_Price) * ComFactor / 10000, 4)    'M
''    Else                                                                              'M
    rsPOITEM!poi_unitprice = CDbl(txt_Price)
''    End If                                                                            'M
    
    
    If (Rs.editmode = adEditAdd) Or (Rs.editmode = adEditInProgress) Then
        
        If Len(txt_Requested) = 0 Or Len(txt_Price) = 0 Then
            txt_Total = 0
        
        Else
        
            If SysUom = "seco" Then
                Rs!poi_totaprice = CDbl((txtSecRequested) * CDbl(txt_Price))
            Else
                Rs!poi_totaprice = CDbl((txt_Requested) * CDbl(txt_Price))
            End If
        
        End If
    End If
            
          '/////
           
           
                                                                    'M
    
    
            
            
          '//////
            
            
    If Err Then Err.Clear
            
    'rs!poi_totaprice = CDbl((txt_Requested) * CDbl(txt_Price))
    
'    If SysUom = "seco" Then
'
'        If Len(Trim$(txtSecRequested)) > 0 Then
'
'            If ComFactor = 0 Then
'                i = txt_Price
'            Else
'                If fm = mdCreation Or txt_Price.Tag = "" Then
'                    txt_Price = FormatNumber$(CDbl(txt_Price) * 10000 / ComFactor, 2)
'                    txt_Price.Tag = txt_Price
'                End If
'            End If
'
'        End If
'
'    End If
        
    rsPOITEM.Update
    If Err Then Err.Clear
End Sub

'validate requested text box and calculate secondary requested unit

Private Sub txt_Requested_Validate(Cancel As Boolean)
On Error Resume Next


    If Len(Trim$(txt_Requested)) = 0 Then Exit Sub
    
    If chk_FrmStkMst.Value = 0 Then
        rsPOITEM!poi_secoreqdqty = txt_Requested
    ElseIf Not IsStringEqual(rsPOITEM!poi_primuom & "", rsPOITEM!poi_secouom & "") Then
    
        If SysUom = "prim" Or SysUom = "both" Then
            txtSecRequested.text = FormatNumber(CDbl(txt_Requested) * 10000 / ComFactor, 4)
        End If
    Else
         txtSecRequested.text = txt_Requested
    End If
        
    If Err Then Err.Clear
    
End Sub

'assign value to po clause recordset

Private Sub txtClause_Validate(Cancel As Boolean)
On Error Resume Next
Dim STR As String, x As Long


    x = txtClause.SelStart
    STR = Trim$(txtClause.text)
    
    'Modified by Muzammil 08/10/00
    'Reason -To eliminate VBCRLF before any text which was giving
    'problems during Email Generation
    
    Do While InStr(1, Trim$(STR), vbCrLf) = 1  'M
    STR = Mid(STR, 3, Len(STR))                'M
    Loop                                        'M
    
    
    rsPOCLAUSE!poc_clau = IIf(Len(STR), STR, "")
    
    txtClause.SelStart = x
    
    rsPOCLAUSE.Update
    If Err Then Err.Clear
End Sub

'assign values to po line item description field and stock number field

Private Sub ssdcboCommoditty_Click()
On Error Resume Next

    If (Len(Trim$(ssdcboCommoditty.Columns(0).text)) <> 0) Then
        rsPOITEM!poi_desc = ssdcboCommoditty.Columns(1).text
        rsPOITEM!poi_comm = ssdcboCommoditty.Columns(0).text
    End If
    
   
    ssdcboManNumber = ""
    rsPOITEM!poi_manupartnumb = ""
    Call GetUnits(rsPOITEM!poi_comm & "")
    If Not ssdcboCommoditty.DroppedDown Then Call GetActiveStockNumbers(False)
End Sub

'execute comsearch function

Private Sub ssdcboCommoditty_DblClick()
On Error Resume Next
    Set comsearch = New frm_StockSearch
    
    comsearch.Execute
End Sub

'assign value to po line item recordset primary unit field

Private Sub ssdbcboUnit_Click()
On Error Resume Next
    rsPOITEM!poi_primuom = dcboUnit.BoundText
    If Err Then Err.Clear
End Sub

'assign value to po recordset shipcode

Private Sub ssdcboShipper_Click()
On Error Resume Next

    If Editting Then rsPO!po_shipcode = ssdcboShipper.Value
    If Err Then Err.Clear
End Sub

'on tab clicks, depend on conditions set navbar buttoms

Private Sub sst_PO_Click(PreviousTab As Integer)
    On Error Resume Next

Dim editmode(1) As Long, STR As String


    dgRecepients.Enabled = False
    dgRecipientList.Enabled = False

    If ((fm = mdCreation) Or (fm = mdModification)) Then
        
        ValidateControls
        
        Select Case PreviousTab
            
            Case 0
                If Not (CheckPoFields) Then sst_PO.Tab = 0
                
            Case 2
                
               'Modified by Muzammil 08/14/00
                'Reason - If Line item misses required info then user should
                'be thrown back to the line item tab again irrespective of where he
                'is trying to go.
                'If (sst_PO.Tab <> 0) Then _

                    If ((Not CheckLIFields)) Then sst_PO.Tab = 2
                
            Case 4
                Call txtClause_Validate(False)
                
        End Select
        
    End If
    
    Call ToggleNavButtons
    editmode(0) = CLng(rsPO.editmode)
    
    
'    PONumb = dcboPO.Text
    If (PreviousTab = 0) Then Ponumb = rsPO!po_ponumb
    dcboPO = Ponumb
    
    If Err Then Err.Clear
    NavBar1.NextEnabled = sst_PO.Tab <> 0
    NavBar1.LastEnabled = sst_PO.Tab <> 0
    NavBar1.FirstEnabled = sst_PO.Tab <> 0
    NavBar1.PreviousEnabled = sst_PO.Tab <> 0
    
    Select Case sst_PO.Tab
    
        Case 0
            
            Set NavBar1.Recordset = rsPO
            Call BindControls("PO")
            
            NavBar1.NextEnabled = False
            NavBar1.LastEnabled = False
            NavBar1.FirstEnabled = False
            NavBar1.PreviousEnabled = False
        Case 1
        
        
            'If POExist Then
                PORECChange
                NavBar1.NewEnabled = False
                NavBar1.SaveEnabled = False
                NavBar1.CancelEnabled = False
                Set NavBar1.Recordset = rsrecepList
                
                dgRecepients.Enabled = True And Editting
                dgRecipientList.Enabled = dgRecepients.Enabled
            'End If
        Case 2
            '11/20/00
            'It the user wants to create a non-stock PO then Commodity Combo
            'should remain Disable
            
            ssdcboCommoditty.Enabled = chk_FrmStkMst.Value 'M
            
            
            POIChange
            Set rsPOITEM = deIms.rsPOITEM
            Set NavBar1.Recordset = deIms.rsPOITEM
            
            Call BindControls("POITEM")
            Label7.Caption = dcboDocumentType.text
            dcboSecUnit.text = "EACH"

            
            
            
            
        Case 3
        
            PORemChange
            Set rsPOREM = deIms.rsPOREM
            Set NavBar1.Recordset = deIms.rsPOREM
            
        Case 4
        
            POCLAUSEChange
            Set rsPOCLAUSE = deIms.rsPOCLAUSE
            Set NavBar1.Recordset = rsPOCLAUSE
            Call BindControls("POCLAUSE")
    End Select
    
    Set dcboPO.RowSource = Nothing
    Set dcboPO.DataSource = Nothing
    
    Set dcboPO.RowSource = deIms
    Set dcboPO.DataSource = deIms
    
End Sub

'assign value to po clause recordset

Private Sub st_Completed(Cancelled As Boolean, Terms As String)
On Error Resume Next

    If Not Cancelled Then
        txtClause.SelText = Terms
        
        Terms = txtClause.text
        rsPOCLAUSE!poc_clau = Terms
    End If
    
    Set st = Nothing
End Sub


Private Sub txt_Requested_Change()
On Error Resume Next
Dim Rs As ADODB.Recordset

    Set Rs = rsPOITEM
    
    If Not Editting Then Exit Sub
    If SysUom = "seco" Then Exit Sub
    If (Rs.editmode = adEditAdd) Or (Rs.editmode = adEditInProgress) Then

        If Len(txt_Requested) = 0 Or Len(txt_Price) = 0 Then
            txt_Total = 0
        
        Else
            Rs!poi_totaprice = CDbl((txt_Requested) * CDbl(txt_Price))
        
        End If
        
    End If
End Sub

Private Function CheckLIFields() As Boolean
On Error Resume Next
Dim i As Long

    CheckLIFields = True
    i = rsPOITEM.editmode
    If i = adEditNone Then Exit Function
    
    If Err Then Err.Clear
    i = rsPOITEM.RecordCount
    
    If i = 0 Then Exit Function
    
    If Err Then Err.Clear
    CheckLIFields = False
    
    If SysUom = "seco" Then
        Call txtSecRequested_Validate(False)
   Else
        Call txt_Requested_Validate(False)
    End If
    
    If Len(Trim$(txt_Requested)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00029") 'J added
        Call MsgBox(IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1))
        '---------------------------------------------
        
        txt_Requested.SetFocus: Exit Function
    
    ElseIf Not IsNumeric(Trim$(txt_Requested)) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00029") 'J added
        MsgBox IIf(msg1 = "", "Requested amount does not contain a valid entry", msg1)
        '---------------------------------------------
        
        txt_Requested.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txt_Price)) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00030") 'J added
        MsgBox IIf(msg1 = "", "Price cannot be left empty ", msg1) 'J modified
        '---------------------------------------------
        
        txt_Price.SetFocus: Exit Function
        
    ElseIf Not (IsNumeric(txt_Price)) Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00031") 'J added
        MsgBox IIf(msg1 = "", "Price does not have a valid entry", msg1) 'J modified
        '---------------------------------------------
        
        txt_Price.SetFocus: Exit Function
        
    Else
        rsPOITEM!poi_unitprice = CDbl(txt_Price)
    End If
    
'    If Len(Trim$(dcboCustomCategory)) = 0 Then
'        MsgBox "Custom Category canot be left empty"
'        dcboCustomCategory.SetFocus: Exit Function
'    End If
    
    If Len(Trim$(rsPOITEM!poi_primuom & "")) = 0 Then
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00032") 'J added
        MsgBox IIf(msg1 = "", "Unit cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        dcboUnit.SetFocus: Exit Function
    End If
    
    If chk_FrmStkMst.Value = vbChecked Then
    
        If Len(Trim$(rsPOITEM!poi_comm & "")) = 0 Then
        
            'Modified by Juan (9/13/2000) for Multilngual
            msg1 = translator.Trans("M00025") 'J added
            MsgBox IIf(msg1 = "", "Stock Number cannot be left empty", msg1) 'J modified
            '--------------------------------------------
            
            ssdcboCommoditty.SetFocus: Exit Function
        Else
            Dim STR As String, OldNum As String
            
            
'''''''            STR = LCase(Trim$(rsPOITEM!poi_comm & ""))
'''''''            OldNum = LCase(Trim$(rsPOITEM!poi_comm.OriginalValue & ""))
'''''''
'''''''            If OldNum <> STR Then
'''''''
'''''''                If Not deIms.StockNumberExist(STR, True) Then
'''''''
'''''''                    'Modified by Juan (9/13/2000) for multilingual
'''''''                    msg1 = translator.Trans("L00119") 'J added
'''''''                    msg2 = translator.Trans("M00026") 'J added
'''''''                    MsgBox IIf(msg1 = "", "Stock number ", msg1 + " ") & STR & IIf(msg2 = "", " does not exist", " " + msg2) 'J modified
'''''''                    '---------------------------------------------
'''''''
'''''''                    ssdcboCommoditty.SetFocus: Exit Function
'''''''
'''''''                End If
                
'''''''            End If
            
        End If
    End If
    
    
    CheckLIFields = True: Err.Clear
End Function

Private Function CheckPoFields() As Boolean
On Error GoTo Handled
Dim i As Long

    i = rsPO.editmode
    If i = adEditNone Then CheckPoFields = True: Exit Function
    
    CheckPoFields = False
    
    If Len(Trim$(dcboDocumentType.BoundText)) = 0 Then
        Call MsgBox(LoadResString(101)): dcboDocumentType.SetFocus: Exit Function
        
    Else
        rsPO!po_docutype = dcboDocumentType.BoundText
    End If
   

        
    If Len(Trim$(ssdcboShipper.text)) = 0 Then
        Call MsgBox(LoadResString(102)): ssdcboShipper.SetFocus: Exit Function
        
    Else
    
        rsPO!po_shipcode = ssdcboShipper.Value
    End If
        
    
    If Len(Trim$(dcboPriority.text)) = 0 Then
        Call MsgBox(LoadResString(103)): dcboPriority.SetFocus: Exit Function
        
    Else
        rsPO!po_priocode = dcboPriority.BoundText
    End If
    
    If Len(Trim$(dcboCurrency.text)) = 0 Then
        Call MsgBox(LoadResString(104)): dcboCurrency.SetFocus: Exit Function
        
    Else
        rsPO!po_currcode = dcboCurrency.BoundText
    End If
    
    
    If Len(Trim$(dcboOriginator.text)) = 0 Then
        Call MsgBox(LoadResString(105)): dcboOriginator.SetFocus: Exit Function
    Else
        rsPO!po_orig = dcboOriginator.BoundText
    End If
    
        
    If Len(Trim$(dcboShipto.text)) = 0 Then
        Call MsgBox(LoadResString(106)): dcboShipto.SetFocus: Exit Function
        
    Else
        rsPO!po_shipto = dcboShipto.BoundText
        
    End If
    
    
    If Len(Trim$(dcboCompany.text)) = 0 Then  'M
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00023") 'J added
        MsgBox IIf(msg1 = "", "Company Can not be left empty", msg1), , "Imswin" 'J modified
        '---------------------------------------------
        
      dcboCompany.SetFocus
      Exit Function 'M
        
    Else  'M
       rsPO!po_compcode = dcboCompany.BoundText 'M
    End If 'M
    
    

    
    If Len(Trim$(rsPO!po_invloca & "")) = 0 Then
        Call MsgBox(LoadResString(107)): dcboInvLocation.SetFocus: Exit Function
    
    Else
        rsPO!po_invloca = dcboInvLocation.BoundText
    End If
    
    If Len(Trim$(dcboSupplier)) = 0 Then
        Call MsgBox(LoadResString(108)): dcboSupplier.SetFocus: Exit Function
        
    Else
        rsPO!po_suppcode = dcboSupplier.Value
    End If
    
    'Modified by Muzammil 08/14/00
    'Reason - Should scream at the user when left empty and the user tries clicking some
    'other tab.
    
    
    If Len(Trim$(ssdcboCondition.text)) = 0 Then          'M
    
        'Modified by Juan (9/13/2000) for Multilingual
        msg1 = translator.Trans("M00034") 'J added
        MsgBox IIf(msg1 = "", "T & C can not be left empty ", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboCondition.SetFocus
        Exit Function 'M
    Else   'M
        rsPO!po_taccode = ssdcboCondition.Value  'M
    End If  'M
    
    If Len(Trim$(ssdcboDelivery.text)) = 0 Then  'M
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00035") 'J added
        MsgBox IIf(msg1 = "", "Payment Term can not be left empty ", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboDelivery.SetFocus: Exit Function 'M
        
    Else  'M
        rsPO!po_termcode = ssdcboDelivery.Value 'M
    End If  'M
    
    
    
    If Len(Trim$(dcboPO)) Then rsPO!po_ponumb = dcboPO
    
    If Len(Trim$(dcboToBeUsedFor)) Then rsPO!po_tbuf = dcboToBeUsedFor.BoundText
    'If Len(Trim$(ssdcboDelivery)) Then rsPO!po_termcode = ssdcboDelivery.Value   'M
    'If Len(Trim$(ssdcboCondition)) Then rsPO!po_taccode = ssdcboCondition.Value   'M
    
    'Added by muzammil to Make sure the Po date < po requested date
    If DTPicker_poDate.Value > dtpRequestedDate.Value Then 'Or DTPicker_poDate.Value = dtpRequestedDate.Value Then   'M  'J Modified
       MsgBox " Transaction Requested Date should be greater than Transaction Create Date by atleast one day." 'M
       dtpRequestedDate.SetFocus
       Exit Function  'M
    End If  'M
    
    CheckPoFields = True: Err.Clear
        
Handled:
    If Err Then Err.Clear
End Function


Private Sub AddRecepient(RecipientName As String, Optional ShowMessage As Boolean = True)
On Error Resume Next
Dim retval As Long

    If Not (deIms.rsPOREC.State And adStateOpen) = adStateOpen Then
    
       Call deIms.POREC(Ponumb, FNamespace)
       
    ElseIf Requery(0) Then
        PORECChange
        
    End If
    
    Set rsrecepList = deIms.rsPOREC
    If Len(Trim$(RecipientName)) = 0 Then Exit Sub
    If ((opt_FaxNum) And (InStr(1, RecipientName, "FAX!", vbTextCompare) = 0)) Then _
        RecipientName = FixFaxNumber(RecipientName)
        
        
        
        
    If IsRecipientInList(RecipientName, ShowMessage) Then Exit Sub
        AddingRecord = True
    
    
    
    
    With rsrecepList
        .AddNew
                    
        !porc_rec = RecipientName
        !porc_npecode = FNamespace
        !porc_recpnumb = GetRecpNumb
        !porc_ponumb = rsPO!po_ponumb
        !porc_recpnumb = IIf(!porc_recpnumb < .RecordCount, .RecordCount, !porc_recpnumb)
        .Update
    End With
    
    cmdRemove.Visible = fm = mdCreation And CBool(rsrecepList.RecordCount)
    Set dgRecipientList.DataSource = deIms
    

End Sub

Private Sub AssignPoNumb()
On Error Resume Next
Dim l As Long, i As Long

    
    Exit Sub
    With rsPOREM
    
        .Filter = adFilterAffectedRecords
        l = .RecordCount - 1
        
        .MoveFirst
        
        While (Not (.EOF))
            !por_ponumb = Ponumb
            .MoveNext
        Wend
        
    End With
            
End Sub

Private Function GetRecpNumb() As Long
On Error Resume Next
Dim cmd As ADODB.Command
Dim Rs As ADODB.Recordset
Dim pm(2) As ADODB.Parameter

    Set cmd = New ADODB.Command
    Set Rs = New ADODB.Recordset
    Set pm(0) = New ADODB.Parameter
    Set pm(1) = New ADODB.Parameter
    Set pm(2) = New ADODB.Parameter
    
    With cmd
        .CommandText = "PORecepNum"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
    End With
    
    With pm(0)
        .Size = 15
        .Value = Ponumb
        .Type = adVarChar
        .direction = adParamInput
    End With
    
    With pm(1)
        .Size = 5
        .Type = adVarChar
        .Value = FNamespace
        .direction = adParamInput
    End With
    
    With pm(2)
        .Value = 0
        .Type = adInteger
        .direction = adParamOutput
    End With
    
    Call cmd.Parameters.Append(pm(0))
    Call cmd.Parameters.Append(pm(1))
    Call cmd.Parameters.Append(pm(2))
    
    Set Rs = cmd.Execute
    
    cmd.Parameters.Delete (0)
    cmd.Parameters.Delete (0)
    cmd.Parameters.Delete (0)
    
    Set Rs = Nothing
    Set cmd = Nothing
    Set pm(0) = Nothing
    Set pm(1) = Nothing
    Set pm(2) = Nothing
    

End Function

Private Sub POIChange()
On Error Resume Next
Dim editmode(1) As Long, STR As String
Dim cmd As ADODB.Command
    
    
    editmode(0) = CLng(rsPO.editmode)
    Set NavBar1.Recordset = rsPOITEM
    
    With rsPOITEM
    
        If Not rsPOITEM.State = adStateClosed Then
    
            Set NavBar1.Recordset = rsPOITEM
            editmode(1) = .editmode
            
            If Err Then Err.Clear
            
            STR = CStr(!POI_PONUMB)
        End If
        
        If Ponumb = "" Then STR = "@"
        If STR = Ponumb Then
        
            If Err Then Err.Clear
            
            If editmode(0) = adEditAdd Then _
                If editmode(1) <> adEditAdd Then NavBar1.AddNew
                
        Else
            
            If (Not (Requery(1)) And editmode(0) = adEditAdd) Then
                rsPOITEM!POI_PONUMB = Ponumb
                Exit Sub
            End If
            
            If editmode(1) <> adEditNone Then .Update
            
            deIms.rsPOITEM.Close
            
            If Err Then Err.Clear
            deIms.cnIms.Errors.Clear
            
            Call deIms.PoItem(Ponumb, FNamespace)
            
            If deIms.cnIms.Errors.Count Then
                deIms.cnIms.Errors.Clear
            End If
            
            Set rsPOITEM = deIms.rsPOITEM
            
            'Modified by Muzammil 08/14/00
            'Reason - when no items where there and the user modifies the PO ,he would need
            'to click on "+".Now when he clicks on poitem tab a new record is added to the
            'recorset ,This way the user does not have to click on "+" and can straight away
            'start creating line item.
            If rsPOITEM.RecordCount = 0 And fm = mdModification Then   'M
            rsPOITEM.AddNew                                            'M
            AddLIDef                                                   'M
             End If                                                    'M
            Set NavBar1.Recordset = rsPOITEM
            
            Call BindControls("POITEM")
            
            If Err Then Err.Clear
            Set NavBar1.Recordset = rsPOITEM
            If editmode(0) = adEditAdd Then NavBar1.AddNew
            Requery(1) = False
        End If
            
    End With
    
    On Error Resume Next
    
        
    txt_linumber = rsPOITEM.RecordCount
    
    'Line added by Muzammil 09/14/00
'Problem - Line items when added to a Po and saved in one shot
'are saved succesfully,but if tried saving new lineItems or modified old ones WITHOUT
'reopening the whole form they simply would not save.
'reason- The variable 'Requery(1)' if found to be 'true' in the procedure
' "SAvelineitems" would  exit the function without saving the new or
'modified records.

'Added the Following line
Requery(1) = False   'M
    If Err Then Err.Clear
End Sub

Private Sub PORemChange()
On Error Resume Next
Dim editmode(1) As Long, STR As String

    
    editmode(0) = CLng(rsPO.editmode)
    Set NavBar1.Recordset = rsPOREM
    
    With rsPOREM
    
        Set NavBar1.Recordset = rsPOREM
        editmode(1) = .editmode
        
        If Err Then Err.Clear
        
        STR = CStr(!por_ponumb)
        
        If Ponumb = "" Then STR = "@"
        If STR = Ponumb Then
        
            If Err Then Err.Clear
            
            If editmode(0) = adEditAdd Then
                If editmode(1) <> adEditAdd Then
                    rsPOREM.AddNew
                    NavBar1_OnNewClick
                End If
            End If

                
        Else
            
            If (Not (Requery(2))) Then
                rsPOREM!por_ponumb = Ponumb
                
                'Modified by Muzammil 08/14/00
            'Reason - when no items where there and the user modifies the PO ,he would need
            'to click on "+".Now when he clicks on poitem tab a new record is added to the
            'recorset ,This way the user does not have to click on "+" and can straight away
            'start creating line item.
            If rsPOREM.RecordCount = 0 And fm = mdModification Then   'M
            rsPOREM.AddNew                                            'M
            AddRemDef                                                   'M
             End If                                                       'M
            
                
                
                Exit Sub
            End If
            
            If editmode(1) <> adEditNone Then .Update
            
            deIms.rsPOREM.Close
            Call deIms.porem(Ponumb, FNamespace)
            
            Set rsPOREM = deIms.rsPOREM
            Set NavBar1.Recordset = rsPOREM
            Set txtRemarks.DataSource = Nothing
             
            txtRemarks.DataMember = "POREM"
            Set txtRemarks.DataSource = deIms
            
            Requery(2) = False
            If Err Then Err.Clear
            Set NavBar1.Recordset = rsPOREM
            If editmode(0) = adEditAdd Then NavBar1.AddNew
        End If
            
    End With
    
    Requery(2) = False
End Sub

Private Sub POCLAUSEChange()
On Error Resume Next
Dim editmode(1) As Long, STR As String

    
    editmode(0) = CLng(rsPO.editmode)
    Set NavBar1.Recordset = rsPOCLAUSE
    
    With rsPOCLAUSE
    
        Set NavBar1.Recordset = rsPOCLAUSE
        editmode(1) = .editmode
        
        If Err Then Err.Clear
        
        STR = CStr(!poc_ponumb)
        If Ponumb = "" Then STR = "@"
        If STR = Ponumb Then
        
            If Err Then Err.Clear
            
            If editmode(0) = adEditAdd Then
                If editmode(1) <> adEditAdd Then
                    rsPOCLAUSE.AddNew
                    NavBar1_OnNewClick
                End If
            End If
                
        Else
            
            If (Not (Requery(3))) Then
                rsPOCLAUSE!poc_ponumb = Ponumb
                
                'Modified by Muzammil 08/14/00
            'Reason - when no items where there and the user modifies the PO ,he would need
            'to click on "+".Now when he clicks on poitem tab a new record is added to the
            'recorset ,This way the user does not have to click on "+" and can straight away
            'start creating line item.
            If rsPOCLAUSE.RecordCount = 0 And fm = mdModification Then   'M
            rsPOCLAUSE.AddNew                                            'M
            AddClauseDef                                                   'M
             End If                                                       'M
            
                
                
                Exit Sub
            End If
            
            If editmode(1) <> adEditNone Then .Update
            
            deIms.rsPOCLAUSE.Close
            
            Call deIms.POClause(Ponumb, FNamespace)
            
            Set rsPOCLAUSE = deIms.rsPOCLAUSE
            
            If Err Then Err.Clear
            
            Set txtClause.DataSource = Nothing
            
            txtClause.DataMember = "POCLAUSE"
            Set txtClause.DataSource = deIms
            
''''
''''            'Modified by Muzammil 08/14/00
''''            'Reason - when no items where there and the user modifies the PO ,he would need
''''            'to click on "+".Now when he clicks on poitem tab a new record is added to the
''''            'recorset ,This way the user does not have to click on "+" and can straight away
''''            'start creating line item.
''''            If rsPOCLAUSE.RecordCount = 0 And fm = mdModification Then   'M
''''            rsPOCLAUSE.AddNew                                            'M
''''            AddClauseDef                                                   'M
''''             End If                                                       'M
''''
            
            
            Set NavBar1.Recordset = rsPOCLAUSE
            If editmode(0) = adEditAdd Then NavBar1.AddNew
        End If
            
    End With
    
    Set txtClause.DataSource = Nothing
    Set txtClause.DataSource = deIms
    Requery(3) = False
End Sub

Private Function PutPODataInsert() As Boolean
Dim cmd As Command
Dim retval As Long

    On Error GoTo errPutPODataInsert

    PutPODataInsert = False

    Set cmd = deIms.Commands("POINSERT_SP")

'    If deIms.POExist(rsPO!po_ponumb, FNameSpace) Then
'        PutPODataInsert = True
'        Exit Function
'    End If
    
    On Error Resume Next
    If Trim$(UCase$(rsPO!po_stas & "")) = "OP" Then _
        rsPO!po_daterevi = Now()
        
        On Error GoTo errPutPODataInsert
    
    'Check for valid data.
    If Not ValidatePOData() Then
        Err.Clear
        Exit Function
    End If

    Dim i As Long, x As Long
    
    DoEvents
    'Set the parameter values for the command to be executed.
    rsPO!po_stas = "OH"
    rsPO!po_sendby = Null
    rsPO!po_apprby = Null
    cmd.Parameters("@po_stas") = "OH"
    cmd.Parameters("@User") = CurrentUser
    cmd.Parameters("@po_ponumb") = Ponumb
    cmd.Parameters("@po_npecode") = GetPKValue(rsPO, rsPO.Bookmark, "po_npecode")
    cmd.Parameters("@po_buyr") = GetPKValue(rsPO, rsPO.Bookmark, "po_buyr")
    cmd.Parameters("@po_date") = GetPKValue(rsPO, rsPO.Bookmark, "po_date")
    cmd.Parameters("@po_apprby") = GetPKValue(rsPO, rsPO.Bookmark, "po_apprby")
    cmd.Parameters("@po_totacost") = GetPKValue(rsPO, rsPO.Bookmark, "po_totacost")
    cmd.Parameters("@po_tbuf") = GetPKValue(rsPO, rsPO.Bookmark, "po_tbuf")
    cmd.Parameters("@po_suppcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_suppcode")
    cmd.Parameters("@po_docutype") = GetPKValue(rsPO, rsPO.Bookmark, "po_docutype")
    cmd.Parameters("@po_priocode") = GetPKValue(rsPO, rsPO.Bookmark, "po_priocode")
    cmd.Parameters("@po_currcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_currcode")
    cmd.Parameters("@po_reqddelvdate") = GetPKValue(rsPO, rsPO.Bookmark, "po_reqddelvdate")
    cmd.Parameters("@po_shipcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_shipcode")
    cmd.Parameters("@po_datesent") = GetPKValue(rsPO, rsPO.Bookmark, "po_datesent")
    cmd.Parameters("@po_orig") = GetPKValue(rsPO, rsPO.Bookmark, "po_orig")
    cmd.Parameters("@po_site") = GetPKValue(rsPO, rsPO.Bookmark, "po_site")
    cmd.Parameters("@po_chrgto") = GetPKValue(rsPO, rsPO.Bookmark, "po_chrgto")
    cmd.Parameters("@po_sendby") = GetPKValue(rsPO, rsPO.Bookmark, "po_sendby")
    cmd.Parameters("@po_confordr") = GetPKValue(rsPO, rsPO.Bookmark, "po_confordr")
    cmd.Parameters("@po_quotnumb") = GetPKValue(rsPO, rsPO.Bookmark, "po_quotnumb")
    cmd.Parameters("@po_forwr") = GetPKValue(rsPO, rsPO.Bookmark, "po_forwr")
    cmd.Parameters("@po_catecode") = GetPKValue(rsPO, rsPO.Bookmark, "po_catecode")
    cmd.Parameters("@po_shipto") = GetPKValue(rsPO, rsPO.Bookmark, "po_shipto")
    cmd.Parameters("@po_stasdelv") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasdelv")
    cmd.Parameters("@po_stasship") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasship")
    cmd.Parameters("@po_stasinvt") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasinvt")
    cmd.Parameters("@po_revinumb") = GetPKValue(rsPO, rsPO.Bookmark, "po_revinumb")
    cmd.Parameters("@po_srvccode") = GetPKValue(rsPO, rsPO.Bookmark, "po_srvccode")
    cmd.Parameters("@po_invloca") = GetPKValue(rsPO, rsPO.Bookmark, "po_invloca")
    cmd.Parameters("@po_daterevi") = GetPKValue(rsPO, rsPO.Bookmark, "po_daterevi")
    cmd.Parameters("@po_taccode") = GetPKValue(rsPO, rsPO.Bookmark, "po_taccode")
    cmd.Parameters("@po_termcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_termcode")
    cmd.Parameters("@po_fromstckmast") = GetPKValue(rsPO, rsPO.Bookmark, "po_fromstckmast")
    cmd.Parameters("@po_reqddelvflag") = GetPKValue(rsPO, rsPO.Bookmark, "po_reqddelvflag")
    cmd.Parameters("@compcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_compcode")

    
    'Execute the command.
    DoEvents
    Call cmd.Execute(Options:=adExecuteNoRecords)
    PutPODataInsert = True

    DoEvents
    Exit Function

errPutPODataInsert:
    If rsPO.EOF Or rsPO.BOF Then Err.Clear
    If Err.number = -2147217900 Then Err.Clear: PutPODataInsert = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Function PutPODataUpdate() As Boolean
Dim cmd As Command

    On Error GoTo errPutPODataUpdate

    PutPODataUpdate = False

    Set cmd = deIms.Commands("POUPDATE_SP")


    'Check for valid data.
    If Not ValidatePOData() Then
        'Raise the ClassError event to detect invalid data.
        MsgBox Err.Description: Err.Clear
        Exit Function
    End If


    DoEvents
    'Set the parameter values for the command to be executed.
    cmd.Parameters("User") = CurrentUser
    cmd.Parameters("@po_ponumb") = Ponumb
    cmd.Parameters("@po_npecode") = GetPKValue(rsPO, rsPO.Bookmark, "po_npecode")
    cmd.Parameters("@po_buyr") = GetPKValue(rsPO, rsPO.Bookmark, "po_buyr")
    cmd.Parameters("@po_date") = GetPKValue(rsPO, rsPO.Bookmark, "po_date")
    cmd.Parameters("@po_apprby") = GetPKValue(rsPO, rsPO.Bookmark, "po_apprby")
    cmd.Parameters("@po_totacost") = GetPKValue(rsPO, rsPO.Bookmark, "po_totacost")
    cmd.Parameters("@po_tbuf") = GetPKValue(rsPO, rsPO.Bookmark, "po_tbuf")
    cmd.Parameters("@po_suppcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_suppcode")
    cmd.Parameters("@po_docutype") = GetPKValue(rsPO, rsPO.Bookmark, "po_docutype")
    cmd.Parameters("po_priocode") = GetPKValue(rsPO, rsPO.Bookmark, "po_priocode")
    cmd.Parameters("@po_currcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_currcode")
    cmd.Parameters("@po_reqddelvdate") = GetPKValue(rsPO, rsPO.Bookmark, "po_reqddelvdate")
    cmd.Parameters("@po_shipcode") = GetPKValue(rsPO, rsPO.Bookmark, "po_shipcode")
    cmd.Parameters("@po_datesent") = GetPKValue(rsPO, rsPO.Bookmark, "po_datesent")
    cmd.Parameters("@po_stas") = GetPKValue(rsPO, rsPO.Bookmark, "po_stas")
    cmd.Parameters("@po_orig") = GetPKValue(rsPO, rsPO.Bookmark, "po_orig")
    cmd.Parameters("@po_site") = GetPKValue(rsPO, rsPO.Bookmark, "po_site")
    cmd.Parameters("@po_chrgto") = GetPKValue(rsPO, rsPO.Bookmark, "po_chrgto")
    cmd.Parameters("@po_sendby") = GetPKValue(rsPO, rsPO.Bookmark, "po_sendby")
    cmd.Parameters("@po_confordr") = GetPKValue(rsPO, rsPO.Bookmark, "po_confordr")
    cmd.Parameters("@po_quotnumb") = GetPKValue(rsPO, rsPO.Bookmark, "po_quotnumb")
    cmd.Parameters("@po_forwr") = GetPKValue(rsPO, rsPO.Bookmark, "po_forwr")
    cmd.Parameters("@po_catecode") = GetPKValue(rsPO, rsPO.Bookmark, "po_catecode")
    cmd.Parameters("@po_shipto") = GetPKValue(rsPO, rsPO.Bookmark, "po_shipto")
    cmd.Parameters("@po_stasdelv") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasdelv")
    cmd.Parameters("@po_stasship") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasship")
    cmd.Parameters("@po_stasinvt") = GetPKValue(rsPO, rsPO.Bookmark, "po_stasinvt")
    cmd.Parameters("@po_revinumb") = GetPKValue(rsPO, rsPO.Bookmark, "po_revinumb")
    cmd.Parameters("@po_reqddelvflag") = GetPKValue(rsPO, rsPO.Bookmark, "po_reqddelvflag")
    cmd.Parameters("@po_srvccode") = GetPKValue(rsPO, rsPO.Bookmark, "po_srvccode")
    cmd.Parameters("@po_invloca") = GetPKValue(rsPO, rsPO.Bookmark, "po_invloca")
    cmd.Parameters("@po_fromstckmast") = GetPKValue(rsPO, rsPO.Bookmark, "po_fromstckmast")
    cmd.Parameters("@po_daterevi") = GetPKValue(rsPO, rsPO.Bookmark, "po_daterevi")

    DoEvents
    'Execute the command.
    cmd.Execute

    DoEvents
    PutPODataUpdate = True
    
    Exit Function

errPutPODataUpdate:
    If rsPO.EOF And rsPO.BOF Then Err.Clear
    If Err = -2147217900 Then Err.Clear: PutPODataUpdate = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Function GetPKValue(Rs As ADODB.Recordset, vBookMark As Variant, sColName As String) As Variant
On Error Resume Next
Dim i As Integer

    GetPKValue = Rs(sColName)
    If Not ((IsNull(Rs(sColName))) Or (IsEmpty(Rs(sColName)))) Then _
        GetPKValue = Trim$(GetPKValue & "")
    



    If IsEmpty(GetPKValue) Then
    
        GetPKValue = Null
    
    ElseIf Len(Trim$(GetPKValue & "")) = 0 Then
        GetPKValue = Null
        
    End If
    
    For i = 1 To UBound(vPKValues, 2)
        If vPKValues(0, i) = vBookMark And LCase(vPKValues(1, i)) = LCase(sColName) Then
            GetPKValue = vPKValues(2, i)
            Exit Function
        End If
    Next i
    
    If Err Then Err.Clear
End Function

Private Function ValidatePOData() As Boolean
Dim i As Long


    ValidatePOData = False
        
    rsPO!po_npecode = FNamespace
    For i = 0 To rsPO.Fields.Count - 1

        Select Case LCase(rsPO.Fields(i).Name)
                
            Case "po_buyr", "po_date", "po_apprby", "po_totacost", "po_tbuf", "po_suppcode", "po_docutype", "po_priocode", "po_currcode", "po_reqddelvdate", "po_shipcode", "po_datesent", "po_stas", "po_orig", "po_site", "po_chrgto", "po_sendby", "po_confordr", "po_quotnumb", "po_forwr", "po_catecode", "po_shipto", "po_stasdelv", "po_stasship", "po_stasinvt", "po_revinumb", "po_reqddelvflag", "po_srvccode", "po_invloca", "po_fromstckmast", "po_daterevi"
                
                If IsEmpty(rsPO(i)) And Not rsPO(i).Type = adBoolean Then
                    MsgBox rsPO(i).Name & " error."
                    Exit Function
                End If
                
        End Select
        
    Next i

    DoEvents
    'Verify the field is not null.
    If IsNull(rsPO("po_ponumb")) Then
        Call deIms.GetAutoNumber(rsPO!po_docutype, FNamespace, rsPO!po_ponumb)
        'Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_ponumb")) Then
        If Len(Trim(rsPO("po_ponumb"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00037") 'J added
            MsgBox IIf(msg1 = "", "The field ' PO Number ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    rsPO("po_npecode") = FNamespace


    'Verify the text field contains text.
    If Not IsNull(rsPO("po_apprby")) Then
        If Len(Trim(rsPO("po_apprby"))) = 0 Then
            rsPO("po_apprby") = Null
        End If
    End If

    'Verify the decimal field contains a valid value.
    If Not IsNull(rsPO("po_totacost")) Then
        If Not IsNumeric(rsPO("po_totacost")) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00038") 'J added
            MsgBox IIf(msg1 = "", "The field ' Total Cost ' does not contain a valid numeric value.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_tbuf")) Then
        If Len(Trim(rsPO("po_tbuf"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00039") 'J added
            MsgBox IIf(msg1 = "", "The field ' To be used for ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_suppcode")) Then
    
        'Modified by Juan (9/14/2000) for Multilngual
        msg1 = translator.Trans("M00040") 'J added
        MsgBox IIf(msg1 = "", "The field ' Supplier Code ' cannot be null.", msg1) 'J modified
        '--------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_suppcode")) Then
        If Len(Trim(rsPO("po_suppcode"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00041") 'j added
            MsgBox IIf(msg1 = "", "The field ' Supplier Code ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    DoEvents
    'Verify the field is not null.
    If IsNull(rsPO("po_docutype")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00042")  'J added
        MsgBox IIf(msg1 = "", "The field ' Document type ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_docutype")) Then
        If Len(Trim(rsPO("po_docutype"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00043") 'J added
            MsgBox IIf(msg1 = "", "The field '  Document type ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_priocode")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00044") 'J added
        MsgBox IIf(msg1 = "", "The field ' Transport Mode ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_priocode")) Then
        If Len(Trim(rsPO("po_priocode"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00045") 'J added
            MsgBox IIf(msg1 = "", "The field ' Transport Mode ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_currcode")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00046") 'J added
        MsgBox IIf(msg1 = "", "The field ' Currency ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_currcode")) Then
        If Len(Trim(rsPO("po_currcode"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00047") 'J added
            MsgBox IIf(msg1 = "", "The field ' Currency ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_reqddelvdate")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00048") 'J added
        MsgBox IIf(msg1 = "", "The field ' Requested Delivery Date ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the date field contains a valid date.
    If Not IsNull(rsPO("po_reqddelvdate")) Then
        If Not IsDate(rsPO("po_reqddelvdate")) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00049") 'J added
            MsgBox IIf(msg1 = "", "The field ' Requested Delivery Date ' does not contain a valid date.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_shipcode")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00050") 'J added
        MsgBox IIf(msg1 = "", "The field ' Ship Code ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    DoEvents
    'Verify the text field contains text.
    If Not IsNull(rsPO("po_shipcode")) Then
        If Len(Trim(rsPO("po_shipcode"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00051")  'J added
            MsgBox IIf(msg1 = "", "The field ' Ship Code ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the date field contains a valid date.
    If Not IsNull(rsPO("po_datesent")) Then
        If Not IsDate(rsPO("po_datesent")) Then
                rsPO("po_datesent") = Null
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_stas")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00052") 'J added
        MsgBox IIf(msg1 = "", "The field ' Date Sent ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_stas")) Then
        If Len(Trim(rsPO("po_stas"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00053") 'J added
            MsgBox IIf(msg1 = "", "The field ' PO Status ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_orig")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00054")  'J added
        MsgBox IIf(msg1 = "", "The field ' Originator ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_orig")) Then
        If Len(Trim(rsPO("po_orig"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00055") 'J added
            MsgBox IIf(msg1 = "", "The field ' Originator ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_site")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00056") 'J added
        MsgBox IIf(msg1 = "", "The field ' Site ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_site")) Then
        If Len(Trim(rsPO("po_site"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00057") 'J added
            MsgBox IIf(msg1 = "", "The field ' Site ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    DoEvents
    'Verify the text field contains text.
    If Not IsNull(rsPO("po_chrgto")) Then
        If Len(Trim(rsPO("po_chrgto"))) = 0 Then
            rsPO("po_chrgto") = Null
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_sendby")) Then
        If Len(Trim(rsPO("po_sendby"))) = 0 Then
            rsPO("po_sendby") = Null
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_quotnumb")) Then
        If Len(Trim(rsPO("po_quotnumb"))) = 0 Then
            rsPO("po_quotnumb") = Null
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_catecode")) Then
        If Len(Trim(rsPO("po_catecode"))) = 0 Then
            rsPO("po_catecode") = Null
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_shipto")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00050") 'J added
        MsgBox IIf(msg1 = "", "The field ' Ship To ' cannot be null.", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_shipto")) Then
        If Len(Trim(rsPO("po_shipto"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00059") 'J added
            MsgBox IIf(msg1 = "", "The field ' Ship To ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_stasdelv")) Then
        If Len(Trim(rsPO("po_stasdelv"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00060") 'J added
            MsgBox IIf(msg1 = "", "The field ' Delivery Status ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_stasship")) Then
        If Len(Trim(rsPO("po_stasship"))) = 0 Then
            
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00061") 'J added
            MsgBox IIf(msg1 = "", "The field ' Ship Status ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_stasinvt")) Then
        If Len(Trim(rsPO("po_stasinvt"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00062") 'J added
            MsgBox IIf(msg1 = "", "The field ' Inventory Status ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    DoEvents
    'Verify the integer field contains a valid value.
    If Not IsNull(rsPO("po_revinumb")) Then
        If Not IsNumeric(rsPO("po_revinumb")) _
            And InStr(rsPO("po_revinumb"), ".") = 0 Then
            
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00063") 'J added
            MsgBox IIf(msg1 = "", "The field ' Revision Number ' does not contain a valid number.", msg1) 'J modified
            '---------------------------------------------

        Exit Function
        End If
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_srvccode")) Then
        If Len(Trim(rsPO("po_srvccode"))) = 0 Then
            rsPO("po_srvccode") = Null
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_invloca")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00064") 'J added
        MsgBox IIf(msg1 = "", "The field ' Inventory Location ' cannot be null.", msg1) 'J modified
        '---------------------------------------------

        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPO("po_invloca")) Then
        If Len(Trim(rsPO("po_invloca"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00065") 'J added
            MsgBox IIf(msg1 = "", "The field ' Inventory Location ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------
        
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rsPO("po_fromstckmast")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00066") 'J added
        MsgBox IIf(msg1 = "", "The field ' From Stock Master ' cannot be null.", msg1) 'J modified
        '---------------------------------------------

        Exit Function
    End If

    'Verify the date field contains a valid date.
    If Not IsNull(rsPO("po_daterevi")) Then
        If Not IsDate(rsPO("po_daterevi")) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00067") 'J added
            MsgBox IIf(msg1 = "", "The field ' Date Revised ' does not contain a valid date.", msg1) 'J modified
            '---------------------------------------------
            
            Exit Function
        End If
    End If

    DoEvents
    ValidatePOData = True

End Function

Private Function PutPORECDataDelete() As Boolean

    Dim cmd As Command

    On Error GoTo errPutPORECDataDelete

    PutPORECDataDelete = False

    Set cmd = deIms.Commands("PORECDELETE_SP")


    'Set the parameter values for the command to be executed.
    cmd.Parameters("@USER") = CurrentUser
    cmd.Parameters("@porc_ponumb") = Ponumb
    cmd.Parameters("@porc_npecode") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_npecode")
    cmd.Parameters("@porc_rec") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_rec")

    'Execute the command.
    cmd.Execute
    
    DoEvents
    PutPORECDataDelete = True

    Exit Function

errPutPORECDataDelete:
    MsgBox Err.Description: Err.Clear
End Function

Private Function ValidatePORECData() As Boolean
On Error Resume Next

    Dim i As Long

    ValidatePORECData = False

    rsrecepList!porc_npecode = FNamespace
    rsrecepList!porc_ponumb = rsPO!po_ponumb
    For i = 0 To rsrecepList.Fields.Count - 1
    
        Select Case LCase(rsrecepList.Fields(i).Name)
        
            Case "porc_ponumb", "porc_npecode", "porc_rec", "porc_recpnumb"
                
                If IsEmpty(rsrecepList(i)) And Not rsrecepList(i).Type = adBoolean Then
                    MsgBox rsrecepList(i).Name & " error."
                    Exit Function
                End If
                
        End Select
        
    Next i

    DoEvents
    'Verify the field is not null.
    If IsNull(rsrecepList("porc_rec")) Then
        MsgBox "The field ' porc_rec ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsrecepList("porc_rec")) Then
        If Len(Trim(rsrecepList("porc_rec"))) = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00070") 'J added
            MsgBox IIf(msg1 = "", "The field ' Recepient ' does not contain valid text.", msg1) 'J modified
            '---------------------------------------------

            Exit Function
        End If
    End If


    ValidatePORECData = True

End Function

Private Function PutPORECDataInsert() As Boolean

    Dim cmd As Command

    On Error GoTo errPutPORECDataInsert

    PutPORECDataInsert = False
    
    rsrecepList!porc_ponumb = rsPO!po_ponumb
    Set cmd = deIms.Commands("PORECINSERT_SP")

    'Check for valid data.
    If Not ValidatePORECData() Then Exit Function

    'Set the parameter values for the command to be executed.
    cmd.Parameters("@User") = CurrentUser
    cmd.Parameters("@porc_ponumb") = Ponumb
    cmd.Parameters("@porc_rec") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_rec")
    cmd.Parameters("@porc_npecode") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_npecode")
    cmd.Parameters("@porc_recpnumb") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_recpnumb")
    'Execute the command.
    cmd.Execute

    DoEvents
    PutPORECDataInsert = True
    
    If Err Then Err.Clear
    Exit Function

errPutPORECDataInsert:
    If rsrecepList.EOF Or rsrecepList.BOF Then Err.Clear
    If Err = -2147217900 Then Err.Clear: PutPORECDataInsert = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Function PutPORECDataUpdate() As Boolean

    Dim cmd As Command

    On Error GoTo errPutPORECDataUpdate

    PutPORECDataUpdate = False

    Set cmd = deIms.Commands("PORECUPDATE_SP")


    'Check for valid data.
    If Not ValidatePORECData() Then
        'Raise the ClassError event to detect invalid data.
        MsgBox Err.Description: Err.Clear
        Exit Function
    End If

    DoEvents
    'Set the parameter values for the command to be executed.
    cmd.Parameters("@User") = CurrentUser
    cmd.Parameters("@porc_ponumb") = Ponumb
    cmd.Parameters("@porc_npecode") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_npecode")
    cmd.Parameters("@porc_recpnumb") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_recpnumb")
    cmd.Parameters("@porc_rec") = GetPKValue(rsrecepList, rsrecepList.Bookmark, "porc_rec")

    'Execute the command.
    cmd.Execute

    DoEvents
    PutPORECDataUpdate = True
    
    Exit Function

errPutPORECDataUpdate:
    If rsrecepList.EOF Or rsrecepList.BOF Then Err.Clear
    If Err = -2147217900 Then Err.Clear: PutPORECDataUpdate = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Sub rsrecepList_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

On Error GoTo errWillChangeRecord

    If (adReason <> adRsnFirstChange) And (adReason <> adRsnUndoAddNew) And Not AddingRecord Then
    
    
       ' If adReason <> adRsnDelete Then adStatus = adStatusCancel

        If IsEmpty(rsrecepList(0).OriginalValue) Then

            If Not PutPORECDataInsert Then
                adStatus = adStatusCancel
            End If

        Else

            Select Case adReason
                Case adRsnUpdate
                    If Not DeleteInProgress Then

                        If Not PutPORECDataUpdate Then
                            adStatus = adStatusCancel
                        End If

                    End If
                Case adRsnAddNew

                    If (Not (PutPORECDataInsert) Or (Not (PutPORECDataUpdate))) Then
                        adStatus = adStatusCancel
                    End If

                Case adRsnDelete

                    If Not PutPORECDataDelete Then
                        adStatus = adStatusCancel
                    End If

                    DeleteInProgress = True
            End Select

        End If
    End If

    DoEvents
    If Err = -2147217900 Then Err.Clear
    
    Exit Sub

errWillChangeRecord:
If Err = -2147217900 Then Err.Clear
If rsrecepList.EOF Or rsrecepList.BOF Then Err.Clear
If Err Then MsgBox Err.Description: Err.Clear


End Sub


Private Sub PORECChange()
On Error Resume Next
Dim editmode(1) As Long, STR As String

    
    Set rsrecepList = deIms.rsPOREC
    editmode(0) = CLng(rsPO.editmode)
    Set NavBar1.Recordset = rsrecepList
    
    With rsrecepList
        editmode(1) = CLng(.editmode)
        
        If Err Then Err.Clear
        
        STR = CStr(!porc_ponumb)
        
        If Ponumb = "" Then STR = "@"
        If STR = Ponumb Then
        
            If Err Then Err.Clear
                
        Else
            
            If Err Then: Requery(0) = True: Err.Clear
            
            If (Not (Requery(0))) Then
                rsrecepList!POI_PONUMB = Ponumb
                Exit Sub
            End If
            
            If editmode(1) <> adEditNone Then .Update
            
            deIms.rsPOREC.Close
            Call deIms.POREC(Ponumb, FNamespace)
            
            'deIms.rsPOREC.Close
            'deIms.rsPOREC.CursorType = adOpenStatic
            'deIms.rsPOREC.LockType = adLockBatchOptimistic
            
            'deIms.rsPOREC.Open
            Set rsrecepList = deIms.rsPOREC
            Set NavBar1.Recordset = rsrecepList
            
            If Err Then Err.Clear
            
            'dgRecipientList.ReBind
            Set dgRecipientList.DataSource = Nothing
            Set dgRecipientList.DataSource = deIms
            
        End If
            
    End With
    
    Requery(0) = False
End Sub

Private Function ValidatePOClauseData() As Boolean

    Dim i As Long

    ValidatePOClauseData = False

    rsPOCLAUSE!poc_npecode = FNamespace
    rsPOCLAUSE!poc_ponumb = rsPO!po_ponumb
    'rsPOCLAUSE!poc_clau = txtClause.Text
    
    For i = 0 To rsPOCLAUSE.Fields.Count - 1
    
        Select Case LCase(rsPOCLAUSE.Fields(i).Name)
        
            Case "poc_ponumb", "poc_npecode", "poc_linenumb", "poc_clau"
                
                If IsEmpty(rsPOCLAUSE(i)) And Not rsPOCLAUSE(i).Type = adBoolean Then
                    MsgBox rsPOCLAUSE(i).Name & " error."
                    Exit Function
                End If
                
        End Select
        
    Next i

    DoEvents
    'Verify the field is not null.
    If IsNull(rsPOCLAUSE("poc_clau")) Then
        MsgBox "The field ' poc_clau ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rsPOCLAUSE("poc_clau")) Then
        If Len(Trim(rsPOCLAUSE("poc_clau"))) = 0 Then
            MsgBox "The field ' poc_clau ' does not contain valid text."
            Exit Function
        End If
    End If


    ValidatePOClauseData = True

End Function

Private Function PutPoCluaseInsert() As Boolean

    Dim cmd As Command

    On Error GoTo errPutPoCluaseInsert

    PutPoCluaseInsert = False

    Set cmd = deIms.Commands("POCLAUSEINSERT_SP")


    'Check for valid data.
    If Not ValidatePOClauseData() Then

        Exit Function
    End If

    DoEvents
    'Set the parameter values for the command to be executed.
    cmd.Parameters("@USER") = CurrentUser
    cmd.Parameters("@poc_ponumb") = Ponumb
    cmd.Parameters("@poc_npecode") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_npecode")
    cmd.Parameters("@poc_linenumb") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_linenumb")
    cmd.Parameters("@poc_clau") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_clau")

    'Execute the command.
    cmd.Execute

    DoEvents
    PutPoCluaseInsert = True
    
    Exit Function

errPutPoCluaseInsert:
If Err = -2147217900 Then Err.Clear: PutPoCluaseInsert = True
If Err Then MsgBox Err.Description: Err.Clear
End Function


Private Function PutPoCluaseUpdate() As Boolean

    Dim cmd As Command

    On Error GoTo errPutPoCluaseUpdate

    PutPoCluaseUpdate = False

    Set cmd = deIms.Commands("POCLAUSEUPDATE_SP")


    'Check for valid data.
    If Not ValidatePOClauseData() Then
        Exit Function
    End If

    'Set the parameter values for the command to be executed.
    cmd.Parameters("@USER") = CurrentUser
    cmd.Parameters("@poc_ponumb") = Ponumb
    cmd.Parameters("@poc_npecode") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_npecode")
    cmd.Parameters("@poc_linenumb") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_linenumb")
    cmd.Parameters("@poc_clau") = GetPKValue(rsPOCLAUSE, rsPOCLAUSE.Bookmark, "poc_clau")

    'Execute the command.
    cmd.Execute
    DoEvents

    PutPoCluaseUpdate = True
    
    Exit Function

errPutPoCluaseUpdate:
    If Err = -2147217900 Then Err.Clear: PutPoCluaseUpdate = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Sub rspoclause_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

On Error GoTo errWillChangeRecord

    If (adReason <> adRsnFirstChange) And (adReason <> adRsnUndoAddNew) And Not AddingRecord Then
        
        If IsEmpty(rsPOCLAUSE(0).OriginalValue) Then
        
            If Not PutPoCluaseInsert Then
                adStatus = adStatusCancel
            End If
            
        Else
        
            Select Case adReason
            
                Case adRsnUpdate
                    If Not DeleteInProgress Then
                        If Not PutPoCluaseUpdate Then
                            adStatus = adStatusCancel
                        End If
                    End If
                    
                Case adRsnAddNew
                    If Not PutPoCluaseInsert Or Not PutPoCluaseUpdate Then
                        adStatus = adStatusCancel
                    End If
                Case adRsnDelete
                
                    DeleteInProgress = True
            End Select
            
        End If
    End If
    
    DoEvents
    If Err = -2147217900 Then Err.Clear
    
    Exit Sub

errWillChangeRecord:
End Sub

Private Function ValidatePOREMData() As Boolean
Dim i As Long

    ValidatePOREMData = False

    rsPOREM!por_ponumb = rsPO!po_ponumb
    rsPOREM!por_npecode = rsPO!po_npecode
    rsPOREM!por_remk = txtRemarks.text
    
    For i = 0 To rsPOREM.Fields.Count - 1
        Select Case LCase(rsPOREM.Fields(i).Name)
            Case "por_ponumb", "por_npecode", "por_linenumb", "por_remk"
                If IsEmpty(rsPOREM(i)) And Not rsPOREM(i).Type = adBoolean Then
                    MsgBox rsPOREM(i).Name & " error."
                    Exit Function
                End If
        End Select
    Next i

    DoEvents
    'Verify the field is not null.
    If IsNull(rsPOREM("por_remk")) Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00073") 'J added
        MsgBox IIf(msg1 = "", "The field ' Remark ' cannot be null.", msg1) 'J modified
        '---------------------------------------------

        Exit Function
    Else
        If Len(rsPOREM("por_remk") & "") > rsPOREM("por_remk").DefinedSize Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00074") 'J added
            msg2 = translator.Trans("M00075") 'J added
            MsgBox IIf(msg1 = "", "Remark Number", msg1) + " " & rsPOREM("por_linenumb") & IIf(msg2 = "", "exceedes", msg2) + " " & rsPOREM("por_remk").DefinedSize 'J modified
            '---------------------------------------------
        
        End If
    End If


    ValidatePOREMData = True

End Function

Private Function PutPOREMDataInsert() As Boolean
Dim cmd As Command

    On Error GoTo errPutPOREMDataInsert

    PutPOREMDataInsert = False

    Set cmd = deIms.Commands("POREMINSERT_SP")


    'Check for valid data.
    If Not ValidatePOREMData() Then
        Exit Function
    End If

    'Set the parameter values for the command to be executed.
    
    cmd.Parameters("@por_ponumb") = Ponumb
    cmd.Parameters("@por_npecode") = deIms.NameSpace
    cmd.Parameters("@por_linenumb") = GetPKValue(rsPOREM, rsPOREM.Bookmark, "por_linenumb")
    cmd.Parameters("@por_remk") = GetPKValue(rsPOREM, rsPOREM.Bookmark, "por_remk")
    cmd.Parameters("@User") = CurrentUser
    'Execute the command.
    cmd.Execute

    PutPOREMDataInsert = True
    
    Exit Function

errPutPOREMDataInsert:
    If rsPOREM.EOF Or rsPOREM.BOF Then Err.Clear
    If Err = -2147217900 Then Err.Clear: PutPOREMDataInsert = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function


Private Function PutPOREMDataUpdate() As Boolean
Dim cmd As Command

    On Error GoTo errPutPOREMDataUpdate

    PutPOREMDataUpdate = False

    Set cmd = deIms.Commands("POREMUPDATE_SP")

    'Check for valid data.
    If Not ValidatePOREMData() Then
        Exit Function
    End If

    DoEvents
    'Set the parameter values for the command to be executed.
    cmd.Parameters("@User") = CurrentUser
    cmd.Parameters("@por_ponumb") = Ponumb
    cmd.Parameters("@por_npecode") = GetPKValue(rsPOREM, rsPOREM.Bookmark, "por_npecode")
    cmd.Parameters("@por_linenumb") = GetPKValue(rsPOREM, rsPOREM.Bookmark, "por_linenumb")
    cmd.Parameters("@por_remk") = GetPKValue(rsPOREM, rsPOREM.Bookmark, "por_remk")

    'Execute the command.
    cmd.Execute
    
    DoEvents
    PutPOREMDataUpdate = True
    
    Exit Function

errPutPOREMDataUpdate:
    If rsPOREM.EOF Or rsPOREM.BOF Then Err.Clear
    If Err = -2147217900 Then Err.Clear: PutPOREMDataUpdate = True
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Function FindPo(PONum As String) As Boolean
On Error Resume Next

Dim sCriteria As String, BK As Variant

'    If IndexOfDataCombo(cboPurchase, PoNum) > -1 Then
    With deIms.rsPO
    
        .CancelUpdate
        Call .CancelBatch(adAffectCurrent)
        BK = .Bookmark
        
        If .editmode = adEditAdd Then
            .CancelUpdate
            Call .CancelBatch(adAffectCurrent)
            .MoveLast
        End If
        
        sCriteria = "po_ponumb = " & "'" & PONum & "'"
        Call .Find(sCriteria, 0, adSearchForward, adBookmarkFirst)
        If Not .EOF Then FindPo = True: .Move 0: Exit Function
        
        .Bookmark = BK
    End With
'    End If
End Function


Private Sub ChangeMode(FMode As FormMode)
On Error Resume Next
Dim bl As Boolean

    LockWindowUpdate (HWND)
    
    If FMode = mdCreation Then
        lblStatus.ForeColor = vbRed
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00125") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '---------------------------------------------
        
    ElseIf FMode = mdModification Then
        lblStatus.ForeColor = vbBlue
                
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00126") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Modification", msg1) 'J modified
        '---------------------------------------------
  
     ElseIf FMode = mdVisualization Then
        lblStatus.ForeColor = vbGreen
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00092") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
    
    End If
    
       
    fm = FMode
    Call MakeReadOnly(fm = mdVisualization)
    Call ShowActiveRecords(False)
    
    GetUnits ("")
    ToggleNavButtons
    LockWindowUpdate (0)
End Sub

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
    fra_Purchase.Enabled = Value
    fra_FaxSelect.Enabled = Value
    dgRecipientList.Enabled = Value
    dcboDocumentType.Enabled = Value
    'RefreshAll
    
    On Error Resume Next
    
    If Value Then
        If LCase(SysUom) = "seco" Then
            dcboSecUnit.Enabled = Not rsPO!po_fromstckmast
        
        Else
            dcboUnit.Enabled = Not rsPO!po_fromstckmast
        End If
    Else
        dcboUnit.Enabled = False
        dcboSecUnit.Enabled = False
    End If
    
    Value = Not (Value)
    
    On Error Resume Next
    If Value Then deIms.cnIms.RollbackTrans
    
    If Err Then Err.Clear
End Sub

Private Function POExist() As Boolean
On Error Resume Next
Dim retval As Long
    
    POExist = deIms.POExist(rsPO!po_ponumb, FNamespace)
End Function

Private Function PutDataInsert() As Boolean

    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert = False

    Set cmd = deIms.Commands("POITEMUPDATE_SP")
    
    DoEvents

    'Set the parameter values for the command to be executed.
    cmd.Parameters(0) = Null
    cmd.Parameters("@USER") = CurrentUser
    cmd.Parameters("@poi_ponumb") = Ponumb
    cmd.Parameters("@poi_npecode") = deIms.NameSpace
    cmd.Parameters("@poi_liitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitnumb")
    cmd.Parameters("@poi_desc") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_desc")
    cmd.Parameters("@poi_primreqdqty") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primreqdqty")
    cmd.Parameters("@poi_primuom") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primuom")
    cmd.Parameters("@poi_secoreqdqty") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_secoreqdqty")
    cmd.Parameters("@poi_secouom") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_secouom")
    
    
    
    cmd.Parameters("@poi_totaprice") = CalculateTotalPrice
    
    cmd.Parameters("@poi_unitprice") = GetPrice
    cmd.Parameters("@poi_qtydlvd") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtydlvd")
    cmd.Parameters("@poi_qtyship") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtyship")
    cmd.Parameters("@poi_qtyinvt") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtyinvt")
    cmd.Parameters("@poi_comm") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_comm")
    cmd.Parameters("@poi_requnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_requnumb")
    cmd.Parameters("@poi_requliitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_requliitnumb")
    cmd.Parameters("@poi_quotnum") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_quotnum")
    cmd.Parameters("@poi_quotliitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_quotliitnumb")
    cmd.Parameters("@poi_locatax") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_locatax")
    cmd.Parameters("@poi_remk") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_remk")
    cmd.Parameters("@poi_serlnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_serlnumb")
    cmd.Parameters("@poi_manupartnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_manupartnumb")
    cmd.Parameters("@poi_liitreqddate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitreqddate")
    cmd.Parameters("@poi_liitrelsdate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitrelsdate")
    cmd.Parameters("@poi_starrendate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_starrendate")
    cmd.Parameters("@poi_endrentdate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_endrentdate")
    cmd.Parameters("@poi_stasliit") = "OH"
    cmd.Parameters("@poi_stasdlvy") = "NR"
    cmd.Parameters("@poi_stasship") = "NS"
    cmd.Parameters("@poi_stasinvt") = "NI"
    cmd.Parameters("@poi_currcode") = rsPO!po_currcode
    cmd.Parameters("@poi_afe") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_afe")
    cmd.Parameters("@poi_custcate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_custcate")
    cmd.Parameters("@poi_lastinvcnumb") = Null
    cmd.Parameters("@poi_qtytobedlvd") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primreqdqty")
    
    DoEvents
    'Execute the command.
    Call cmd.Execute(Options:=adExecuteNoRecords)

    DoEvents
    PutDataInsert = True

    Exit Function

errPutDataInsert:
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Function PutDataPOItemUpdate() As Boolean
On Error Resume Next

    Dim cmd As Command
    PutDataPOItemUpdate = False
    On Error GoTo errPutDataPOItemUpdate
    Set cmd = deIms.Commands("POITEMUPDATE_SP")

    DoEvents
    'Set the parameter values for the command to be executed.
    cmd.Parameters(0) = 0
    cmd.Parameters("@USER") = CurrentUser
    cmd.Parameters("@poi_ponumb") = Ponumb
    cmd.Parameters("@poi_npecode") = deIms.NameSpace
    cmd.Parameters("@poi_currcode") = rsPO!po_currcode
    cmd.Parameters("@poi_desc") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_desc")
    cmd.Parameters("@poi_comm") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_comm")
    cmd.Parameters("@poi_liitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitnumb")
    
    DoEvents
    'On Error Resume Next
    cmd.Parameters("@poi_primreqdqty") = CDbl(GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primreqdqty") & "")
    cmd.Parameters("@poi_primuom") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primuom")
    cmd.Parameters("@poi_secoreqdqty") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_secoreqdqty")
    cmd.Parameters("@poi_secouom") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_secouom")
    
  'Modified by Muzammil 08/07/00
 '   cmd.Parameters("@poi_totaprice") = CalculateTotalPrice                        'M
    cmd.Parameters("@poi_totaprice") = FormatNumber$(rsPOITEM!poi_totaprice, 2)        'M
    
    cmd.Parameters("@poi_unitprice") = GetPrice
    
    cmd.Parameters("@poi_qtydlvd") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtydlvd")
    cmd.Parameters("@poi_qtyship") = CDbl(GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtyship") & "")
    cmd.Parameters("@poi_qtyinvt") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_qtyinvt")
    cmd.Parameters("@poi_requnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_requnumb")
    cmd.Parameters("@poi_requliitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_requliitnumb")
    cmd.Parameters("@poi_quotnum") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_quotnum")
    cmd.Parameters("@poi_quotliitnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_quotliitnumb")
    cmd.Parameters("@poi_locatax") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_locatax")
    cmd.Parameters("@poi_remk") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_remk")
    cmd.Parameters("@poi_serlnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_serlnumb")
    cmd.Parameters("@poi_manupartnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_manupartnumb")
    cmd.Parameters("@poi_liitreqddate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitreqddate")
    cmd.Parameters("@poi_liitrelsdate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_liitrelsdate")
    cmd.Parameters("@poi_starrendate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_starrendate")
    cmd.Parameters("@poi_endrentdate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_endrentdate")
    
    cmd.Parameters("@poi_stasliit") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_stasliit")
    cmd.Parameters("@poi_stasdlvy") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_stasdlvy")
    cmd.Parameters("@poi_stasship") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_stasship")
    cmd.Parameters("@poi_stasinvt") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_stasinvt")

    cmd.Parameters("@poi_afe") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_afe")
    cmd.Parameters("@poi_custcate") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_custcate")
    cmd.Parameters("@poi_qtytobedlvd") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_primreqdqty")
    cmd.Parameters("@poi_lastinvcnumb") = GetPKValue(rsPOITEM, rsPOITEM.Bookmark, "poi_lastinvcnumb")

    DoEvents
    Call cmd.Execute(Options:=adExecuteNoRecords)

    DoEvents
    PutDataPOItemUpdate = True

    Exit Function

errPutDataPOItemUpdate:
    If Err Then MsgBox Err.Description: Err.Clear
End Function

Private Sub RefreshAll()
On Error Resume Next
Dim ct As Control

    For Each ct In Me.Controls
        ct.Refresh
        If Err Then Err.Clear
    Next ct
    
End Sub

Private Function IsRecipientInList(RecepientName As String, Optional ShowMessage As Boolean = True)
On Error Resume Next
Dim BK As Variant
    
    
    rsrecepList.MoveFirst
    If Not (rsrecepList.EOF Or rsrecepList.BOF) Then BK = rsrecepList.Bookmark
    
    rsrecepList.Filter = "porc_rec = '" & RecepientName & "'"
    
    If Not (rsrecepList.EOF) Then
    
        If ((Not (rsrecepList.RecordCount = 0))) Then
        
            #If DBUG Then
                If Err Then Stop
            #Else
                On Error Resume Next
            #End If
            
            If ShowMessage Then
                If opt_Email Then
                
                    'Modified by Juan (8/29/2000) for Multilingual
                    msg1 = translator.Trans("M00076") 'J added
                    MsgBox IIf(msg1 = "", "Email Address Already in list", msg1) 'J modified
                    '------------------------------------------
                    
                ElseIf opt_FaxNum Then
                
                    'Modified by Juan (8/29/2000) for Multilingual
                    msg1 = translator.Trans("M00077") 'J added
                    MsgBox IIf(msg1 = "", "Fax Number Already in list", msg1) 'J modified
                    '---------------------------------------------
                    
                End If
            End If
            IsRecipientInList = True
        End If
    End If
    
     rsrecepList.Filter = adFilterNone
     If IsRecipientInList Then Call rsrecepList.Find("porc_rec = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
     
     If rsrecepList.RecordCount = 0 Then IsRecipientInList = 0
     
     'If rsrecepList.EOF Then rsrecepList.MoveFirst
    If Err Then Err.Clear
    
End Function

Private Sub AddLIDef()
On Error Resume Next
Dim Rs As ADODB.Recordset



    Set Rs = New ADODB.Recordset
            
        Rs.ActiveConnection = deIms.cnIms
        Rs.CursorType = adOpenForwardOnly
        Rs.LockType = adLockReadOnly
        
        Rs.Open ("select max(poi_liitnumb) iNumb from POITEM where " & _
                "poi_ponumb = '" & deIms.rsPO!po_ponumb & "' and " & _
                "poi_npecode = '" & deIms.rsPO!po_npecode & "'")
                
        'Debug.Print rs.Source
        'Debug.Print rs!inumb
    
    With rsPOITEM
        txt_Price.Tag = ""
        
        !poi_qtyship = 0
        !poi_qtyinvt = 0
        !poi_qtydlvd = 0
        !poi_totaprice = 0
        !poi_comm = Null
        !poi_primuom = Null
        !poi_stasdlvy = "NR"
        !poi_stasship = "NS"
        !poi_stasinvt = "NI"
        !poi_stasliit = "OH"
        !poi_npecode = FNamespace
        !poi_afe = deIms.rsPO!po_chrgto
        !POI_PONUMB = deIms.rsPO!po_ponumb
        !poi_currcode = deIms.rsPO!po_currcode
        !poi_liitnumb = IIf(IsNull(Rs!iNumb), 1, Rs!iNumb + 1)
        !poi_liitnumb = IIf(!poi_liitnumb < .RecordCount, .RecordCount, !poi_liitnumb)
        !poi_liitreqddate = CDate(deIms.rsPO!po_reqddelvdate)
        
        dcboUnit.text = ""
        ssdcboCommoditty.text = ""
    End With
    
    Rs.Close
    Set Rs = Nothing

End Sub

Private Sub AddRemDef()
On Error Resume Next
Dim i As Long
Dim Rs As ADODB.Recordset


    Set Rs = New ADODB.Recordset
    
    Rs.ActiveConnection = deIms.cnIms
    Rs.CursorType = adOpenForwardOnly
    Rs.LockType = adLockReadOnly
    
    LogExec ("Adding Remarks")
    Rs.Open ("select max(por_linenumb) iNumb from POREM where " & _
            "por_ponumb = '" & Ponumb & "' and " & _
            "por_npecode = '" & FNamespace & "'")
            
            
    With rsPOREM
        !por_npecode = FNamespace
        !por_ponumb = deIms.rsPO!po_ponumb
        !por_linenumb = IIf(IsNull(Rs!iNumb), 1, Rs!iNumb + 1)
        !por_linenumb = IIf(!por_linenumb < .RecordCount, .RecordCount, !por_linenumb)
        Call LogExec("Remarks LIne Number " & !por_linenumb & " Addedd")
        
        'For i = 0 To 1000
        '    !por_remk = !por_remk & "1"
        'Next
    End With
    
        
End Sub

Private Sub AddClauseDef()
On Error Resume Next

    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    Rs.ActiveConnection = deIms.cnIms
    Rs.CursorType = adOpenForwardOnly
    Rs.LockType = adLockReadOnly
    
    Rs.Open ("select max(poc_linenumb) iNumb from POCLAUSE where " & _
            "poc_ponumb = '" & deIms.rsPO!po_ponumb & "' and " & _
            "poc_npecode = '" & deIms.rsPO!po_npecode & "'")
            
    With rsPOCLAUSE
        !poc_npecode = FNamespace
        !poc_ponumb = deIms.rsPO!po_ponumb
        !POC_LINENUMB = IIf(IsNull(Rs!iNumb), 1, Rs!iNumb + 1)
        !POC_LINENUMB = IIf(!POC_LINENUMB < .RecordCount, .RecordCount, !POC_LINENUMB)
    End With

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
        .ParameterFields(1) = "ponumb;" + rsPO!po_ponumb + ";TRUE"
    End With
    
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::BeforePrint", Err.Description, Err)
    End If
End Sub

Private Sub GetUnits(StockNumber As String)
On Error Resume Next
Dim STR As String
Dim cmd As ADODB.Command
    
    deIms.rsSECONDARYUNIT.Close
    deIms.rsGET_UNIT_OF_MEASURE.Close
    
    Set dcboUnit.RowSource = Nothing
    Set dcboUnit.DataSource = Nothing
    Set dcboSecUnit.RowSource = Nothing
    Set dcboSecUnit.DataSource = Nothing
    Set ssdcboManNumber.DataSource = Nothing
    Set ssdcboManNumber.DataSourceList = Nothing
    dcboUnit.RowMember = "GET_UNIT_OF_MEASURE"
    
    Call deIms.SecondaryUnit(deIms.NameSpace, StockNumber)
    Call deIms.GET_UNIT_OF_MEASURE(deIms.NameSpace, StockNumber)
    
    Set dcboUnit.RowSource = deIms
    Set dcboUnit.DataSource = deIms
    
   
    Set dcboSecUnit.RowSource = deIms
    Set dcboSecUnit.DataSource = deIms
    
    DoEvents
    If deIms.rsSECONDARYUNIT.RecordCount = 1 Then
        STR = Trim$(deIms.rsSECONDARYUNIT!uni_code & "")
        ComFactor = ComputingFactor(FNamespace, StockNumber, deIms.cnIms)
        
        dcboSecUnit.Enabled = False
        rsPOITEM!poi_secouom = STR
        dcboSecUnit.BoundText = rsPOITEM!poi_secouom
        dcboSecUnit.text = deIms.rsSECONDARYUNIT!uni_desc & ""
        
        dcboSecUnit.Refresh
        'deIms.rsPOITEM!poi_primuom = str
    End If
        
    If deIms.rsGET_UNIT_OF_MEASURE.RecordCount = 1 Then
        STR = Trim$(deIms.rsGET_UNIT_OF_MEASURE!uni_code & "")
        
        dcboUnit.Enabled = False
        rsPOITEM!poi_primuom = STR
        dcboUnit.BoundText = rsPOITEM!poi_primuom
        dcboUnit.text = deIms.rsGET_UNIT_OF_MEASURE!uni_desc & ""
        
        dcboSecUnit.Refresh
        'deIms.rsPOITEM!poi_secouom = str
    End If
        
    Set cmd = MakeCommand(deIms.cnIms, adCmdStoredProc)
    
    With cmd
        Dim Rs As ADODB.Recordset
        .CommandText = "GetStockManufacturer"
        .Parameters.Append .CreateParameter("np", adVarChar, adParamInput, 5, FNamespace)
        .Parameters.Append .CreateParameter("stock", adVarChar, adParamInput, 20, StockNumber)
        
        Set Rs = .Execute
        
        Rs.Filter = "stm_flag <> 0"
        Set ssdcboManNumber.DataSource = deIms
        Set ssdcboManNumber.DataSourceList = Rs.Clone
        
        Set Rs = Nothing
    End With
    
    DoEvents
    Set cmd = Nothing
    If Err Then Err.Clear
End Sub

Private Function GetDocumentType(All As Boolean) As Long
On Error Resume Next
Dim Rs As ADODB.Recordset

    With deIms
        GetDocumentType = 0
        
        If All Then
            Set Rs = .UserDocumentType("", GetDocumentType)
        Else
            Set Rs = .UserDocumentType(CurrentUser, GetDocumentType)
        End If
        
    End With
    
    If Err Then Err.Clear
    
    dcboDocumentType.RowMember = ""
    Set dcboDocumentType.RowSource = Nothing
    Set dcboDocumentType.DataSource = Nothing
    
    Set dcboDocumentType.RowSource = Rs
    Set dcboDocumentType.DataSource = deIms
    
    dcboDocumentType.ReFill

End Function

Private Sub txtRemarks_Validate(Cancel As Boolean)
Dim i As Long
On Error Resume Next

    i = rsPOREM.editmode
    
    If i <> adEditNone Then
        rsPOREM!por_remk = IIf(Len(Trim$(txtRemarks.text)), txtRemarks, Null)
        rsPOREM.Update
    End If
    If Err Then Err.Clear
End Sub

Private Function GetRequisitions(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
On Error Resume Next
    
    Dim cmd As ADODB.Command
    
    Set cmd = MakeCommand(cn, adCmdStoredProc)
    
    With cmd
        .CommandText = "GET_BRQ"
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, NameSpace)
        
        Set GetRequisitions = .Execute
    End With
    
    Set cmd = Nothing
    If Err Then Err.Clear
    
End Function

Private Sub GetDistributors(Gender As Variant)
Dim i As Long
Dim Values(1) As String
Dim Rs As ADODB.Recordset
On Error Resume Next

    'Set rs = GetPORecipients(FNamespace, dcboDocumentType.BoundText, "a", deIms.cnIms, i)
    'ORIGINAl
    Set Rs = GetPORecipients(FNamespace, dcboDocumentType.BoundText, deIms.cnIms, i)

    
    If i <= 0 Then Exit Sub
    
    opt_FaxNum = False
    
    Do Until Rs.EOF
        Values(0) = Rs(0) & ""
        Values(1) = FixFaxNumber(Rs(1) & "")

        
        If Len(Values(0)) Then Call AddRecepient(Values(0), False)
        If Len(Values(1)) > 4 Then Call AddRecepient(Values(1), False)
        

        If Err Then Err.Clear

        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    If Err Then Err.Clear

End Sub

Private Sub GetSupplierFax()
On Error Resume Next

  
    opt_FaxNum = False
    Call AddRecepient(FixFaxNumber(GetSupplierFaxNumber()), False)
    If Err Then Call LogErr(Name & "::GetSupplierFax", Err.Description, Err.number, True)

End Sub

Private Function FixFaxNumber(Faxnumber As String) As String
On Error Resume Next

    If Len(Faxnumber) < 7 Then Exit Function

    If Left$(Faxnumber, 1) = "+" Then
        Faxnumber = Right$(Faxnumber, Len(Faxnumber) - 1)
    End If
    
    If Mid$(Faxnumber, 1, 4) <> "FAX!" Then _
        FixFaxNumber = "FAX!" & Faxnumber

    'Modified by Juan (9/14/2000) for Multilingual
    msg1 = translator.Trans("M00078") 'J added
    If Err Then Err.Clear: MsgBox IIf(msg1 = "", "err occured", msg1) 'J modified
    '---------------------------------------------

End Function

'modified by muzammil 08/04/00 /// function did not have the " AS STRING" in the declaration before
Private Function GetSupplierFaxNumber() As String
On Error Resume Next
Dim STR As String

STR = Trim$(dcboSupplier.Value)


    GetSupplierFaxNumber = GetSupplierEmailForPO(FNamespace, STR, deIms.cnIms)
    If Err Then Call LogErr(Name & "::GetSupplierFax", Err.Description, Err.number, True)


End Function

Private Sub GetUnitOfMeasurement()
On Error GoTo CleanUp
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        
        .CommandText = " SELECT ? = psys_uom"
        .CommandText = .CommandText & " From PESYS "
        .CommandText = .CommandText & " WHERE psys_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND psys_usercode =  'PE' "
    

        .Parameters.Append .CreateParameter("", adVarChar, adParamOutput, 4)
        .Execute , , adExecuteNoRecords
        
        DoEvents
        SysUom = LCase$(.Parameters(0).Value & "")
    End With
    
    Exit Sub
    
CleanUp:
    Set cmd = Nothing
End Sub

Private Sub txtSecRequested_LostFocus()
On Error Resume Next
    Call txtSecRequested_Validate(False)
End Sub

Private Sub txtSecRequested_Validate(Cancel As Boolean)
Dim i As Double
On Error Resume Next


    If Not Editting Then Exit Sub
    If SysUom = "prim" Or SysUom = "both" Then Exit Sub
    
    If Len(Trim$(txtSecRequested)) = 0 Then Exit Sub
    

    If Not IsStringEqual(rsPOITEM!poi_primuom & "", rsPOITEM!poi_secouom & "") Then
    
    'Added the next four 5 lines and the closing  "end if " /Muzammil/07/25/00
    
        If SysUom = "seco" Or SysUom = "both" Then
             
             If SysUom = "seco" Then
             
                 txt_Requested.text = FormatNumber(CDbl(txtSecRequested) * ComFactor / 10000, 4)
                 'If it is new Record or poi_unitprice is not modified
                 If (rsPOITEM.Status <> adRecModified Or rsPOITEM("poi_unitprice").Value = rsPOITEM("poi_unitprice").OriginalValue) Then       'M
                 'Multiply SecPrice and Unitprice only if the original price of poi_unitprice is empty.This happens
                 'in a case when it is a new record and the user enters a SecQuantity and then
                 'the UnitPrice and then again modifies the Secqunatity.At that time the First
                 'condition Executes.
                    
                    If rsPOITEM("poi_unitprice").OriginalValue = Empty Then 'M
                    rsPOITEM!poi_totaprice = txtSecRequested * txt_Price 'M
                    Else 'M
                    rsPOITEM!poi_totaprice = txt_Requested * txt_Price 'M
                    End If 'M
                    
                 ElseIf rsPOITEM("poi_unitprice").Value <> rsPOITEM("poi_unitprice").OriginalValue And rsPOITEM!poi_secoreqdqty <> txtSecRequested Then 'M
                     
                     rsPOITEM!poi_totaprice = txtSecRequested * txt_Price 'M
                 
                 End If  'M
              
             ElseIf SysUom = "both" Then 'M
                 
                 
            rsPOITEM!poi_totaprice = txtSecRequested * txt_Price
            txt_Requested.text = FormatNumber(CDbl(txtSecRequested) * ComFactor / 10000, 4)
            End If
            
            
        End If
    Else
        txt_Requested = txtSecRequested
        'When the S Unit and P Unit is Same and the S Quantity is Modified after Modifying the
        'Unit Price, the total would not change,So added this Piece of Code.
        rsPOITEM!poi_totaprice = txtSecRequested * txt_Price 'M
    End If
        
    
'    If SysUom <> "seco" Then Exit Sub
'    If Len(Trim$(txt_Price)) > 0 Then
'
'        If ComFactor = 0 Then
'            i = txt_Price
'        Else
'
'            If txt_Price.Tag = "" Then
'                txt_Price = FormatNumber$(CDbl(txt_Price) * 10000 / ComFactor, 2)
'                txt_Price.Tag = txt_Price
'            End If
'        End If
'    End If
        
    If Err Then Err.Clear

End Sub


Private Sub GetLocations(CompanyName)
On Error Resume Next
Dim Rs As ADODB.Recordset

    Set Rs = deIms.rsINVENTORYLOCATION
    
    If Rs.State And adStateOpen Then Rs.Close
    
    Call deIms.INVENTORYLOCATION(FNamespace, CompanyName)
    Set dcboInvLocation.RowSource = deIms
    
    
End Sub

Private Sub ShowActiveRecords(Active As Boolean)
On Error Resume Next

    Set dcboShipto.RowSource = Nothing
    'Set dcboOriginator.RowSource = Nothing
    'Set dcboToBeUsedFor.RowSource = Nothing
    Set dcboSupplier.DataSourceList = Nothing
    
    If Active Then

        Call deIms.ActiveTbu(FNamespace)
        Call deIms.SuppLookup(FNamespace)
        Call deIms.ActiveShipTo(FNamespace)
        Call deIms.ActiveOriginator(FNamespace)
        
        If Err Then Err.Clear
        dcboShipto.RowMember = "ActiveShipTo"
        'dcboToBeUsedFor.RowMember = "ActiveTbu"
        'dcboOriginator.RowMember = "ActiveOriginator"
        dcboSupplier.DataMemberList = "SuppLookup"
    Else
    
        Call deIms.ShipTo(FNamespace)
        Call deIms.Supplier(FNamespace)
        Call deIms.Originator(FNamespace)
        Call deIms.TOBEUSEDFOR(FNamespace)
        
        If Err Then Err.Clear
    
        dcboShipto.RowMember = "ShipTo"
        'dcboOriginator.RowMember = "ORIGINATOR"
        'dcboToBeUsedFor.RowMember = "TOBEUSEDFOR"
        dcboSupplier.DataMemberList = "Supplier"
    End If
        
    If Err Then Err.Clear
    
    Set dcboShipto.RowSource = deIms
    'Set dcboOriginator.RowSource = deIms
    'Set dcboToBeUsedFor.RowSource = deIms
    Set dcboSupplier.DataSourceList = deIms

    If Err Then Call LogErr(Name & "::ShowActiveRecords", Err.Description, Err, True)
End Sub

Public Sub GetActiveCompanies(Active As Boolean)
Dim Rs As ADODB.Recordset
On Error Resume Next

    Set dcboCompany.RowSource = Nothing
    
    
    If deIms.rsCOMPANY.State And adStateOpen Then
        Set Rs = deIms.rsCOMPANY.Clone
    Else
        deIms.Company (FNamespace)
        Set Rs = deIms.rsCOMPANY.Clone
    End If
        
    
    If Active Then _
        Rs.Filter = "com_actvflag <> 0"
        
    Set dcboCompany.RowSource = Rs

End Sub


'SQL statement get active stock number

Public Sub GetActiveStockNumbers(Active As Boolean)
On Error Resume Next
Dim Source As String
Dim cmd As ADODB.Command
Dim Rs As ADODB.Recordset

    deIms.rsStockMasterLookup.Close
    
    If Err Then Err.Clear
    
    
    Set cmd = New ADODB.Command
    Source = deIms.rsStockMasterLookup.Source
    If Active Then Source = Source & " AND (stk_flag <> 0)"
    
    Source = Source & " Order By 1"
    
    With cmd
        .CommandText = Source
        .ActiveConnection = deIms.cnIms
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, FNamespace)
        Set Rs = .Execute
    End With
    
    
    ssdcboCommoditty.DataMemberList = ""
    Set ssdcboCommoditty.DataSourceList = Nothing
    Set ssdcboCommoditty.DataSourceList = Rs

End Sub

Public Function IsPoApprove(PoNumber As String) As Boolean
On Error Resume Next
    
    IsPoApprove = rsPO!po_stas = "OP"
    If Err Then Call LogErr(Name & "::IsPoApprove", Err.Description, Err, True)
End Function

Private Function Editting() As Boolean
On Error Resume Next
    Editting = ((fm = mdCreation) Or (fm = mdModification))
End Function

Public Sub ToggleNavButtons()
On Error Resume Next
Dim rc As Long
Dim bl As Boolean
Dim ed As Boolean

    ed = Editting
    bl = sst_PO.Tab = 0
    rc = rsrecepList.RecordCount
    
    With NavBar1
        .CloseEnabled = bl
        .CancelEnabled = ed
        .NewEnabled = Not ed
        .EditEnabled = Not ed
        .SaveEnabled = bl And ed
        .PrintEnabled = Not ed And bl
        .EMailEnabled = .PrintEnabled And rc
    End With
    
    If sst_PO.Tab = 0 Then
        NavBar1.NewEnabled = Not ed
        
    ElseIf sst_PO.Tab > 1 Then
        NavBar1.NewEnabled = ed
    End If
End Sub

Private Sub WriteStatus(Msg As String)
    Call MDI_IMS.WriteStatus(Msg, 1)
End Sub

Private Function GetPrice() As Double
On Error Resume Next
Dim db As Double
Dim STR As String
Dim val(1) As Double

    val(0) = CDbl(FormatNumber$(rsPOITEM!poi_unitprice, 4))
    val(1) = CDbl(FormatNumber$(rsPOITEM!poi_unitprice.OriginalValue, 4))
    
    If val(0) = val(1) Then GetPrice = val(0):  Exit Function

    STR = Trim$(rsPOITEM!poi_totaprice & "")
    db = FormatNumber$(Trim$(rsPOITEM!poi_unitprice & ""), 4)
    'OldValue = FormatNumber$(Trim$(rsPOITEM!poi_unitprice.OriginalValue & ""), 4)
    
    
    If SysUom = "seco" Then
'
'        If Len(str) Then
'
'            If rsPOITEM.Status <> adRecNew Then
'
'                If fm = mdModification Then
'
'                    If ((db = OldValue) Or (OldValue = 0)) Then
'
'                        If SysUom = "seco" Then
'                            db = CDbl(str) / rsPOITEM!poi_secoreqdqty
'                        Else
'                            db = CDbl(str) / rsPOITEM!poi_primreqdqty
'                        End If
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
        If ComFactor = 0 Then
            GetPrice = db
        Else
            GetPrice = FormatNumber$(db * 10000 / ComFactor, 4)
        End If
    Else
        GetPrice = db
    End If
    
    rsPOITEM!poi_unitprice = db
    GetPrice = CDbl(FormatNumber$(GetPrice, 4))
    
    If Err Then Call LogErr(Name & "::Price", Err.Description, Err)
    
    Err.Clear
End Function

Private Function CalculateTotalPrice() As Double
On Error GoTo 0
On Error Resume Next
Dim db As Double

Dim val(1) As Double

    val(0) = CDbl(FormatNumber$(rsPOITEM!poi_totaprice, 4))
    val(1) = CDbl(FormatNumber$(rsPOITEM!poi_totaprice.OriginalValue, 4))
    
    If val(0) = val(1) Then CalculateTotalPrice = val(0): Exit Function
    
    db = Trim$(rsPOITEM!poi_unitprice & "")
    
    If SysUom = "seco" Then
        db = db * rsPOITEM!poi_secoreqdqty
    Else
        db = db * rsPOITEM!poi_primreqdqty
    End If
    
    CalculateTotalPrice = FormatNumber$(db, 2)
    If Err Then Call LogErr(Name & "::Price", Err.Description, Err)
    
    Err.Clear
End Function

Private Function SaveLineItems() As Boolean
On Error Resume Next
Dim i As Long, x As Long




    SaveLineItems = True
    If Requery(1) = True Then Exit Function
    If rsPOITEM.State = adStateClosed Then Exit Function
    
    rsPOITEM.Update
    rsPOITEM.Filter = 0
    rsPOITEM.Filter = adFilterPendingRecords
    
    
    
    rsPOITEM.MoveFirst
    SaveLineItems = False
    i = rsPOITEM.RecordCount
    
    Err.Clear
    Do Until rsPOITEM.EOF
    
        x = rsPOITEM.AbsolutePosition
        Call WriteStatus("Saving Line item " & x & " of " & i)
        
        
        
        If rsPOITEM.Status = adRecNew Then
            If Not PutDataInsert Then Exit Function
        Else
            If Not PutDataPOItemUpdate Then Exit Function
        End If
        Call WriteStatus("Line item " & x & " of " & i & " saved successfully")
        
        DoEvents
        
        If Err Then
            If Err <> 3021 Then MsgBox Err.Description
            Call LogErr(Name & "::SaveLineItems", Err.Description, Err, True)
            Exit Function
        End If
        rsPOITEM.MoveNext
    Loop

    SaveLineItems = True
End Function

Public Function SaveRemarks() As Boolean

On Error Resume Next
Dim STR As String, i As Integer, x As Integer

    SaveRemarks = True
    If Requery(2) = True Then Exit Function
    If rsPOREM.State = adStateClosed Then Exit Function

    rsPOREM.Update
    rsPOREM.Filter = adFilterPendingRecords


    rsPOREM.MoveFirst
    SaveRemarks = False
    i = rsPOREM.RecordCount

    Do Until rsPOREM.EOF
    
        x = rsPOREM.AbsolutePosition
        
        
        STR = Trim$(rsPOREM!por_remk & "")
          
       'Modified by Muzammil 08/11/00
       'Reason - VBCRLFs before the text would block Email Generation.
          
          Do While InStr(1, STR, vbCrLf) = 1    'M
             STR = Mid(STR, 3, Len(STR))        'M
          Loop                                  'M
             STR = LTrim$(STR)
             txtRemarks.text = STR
        If Len(STR) > 0 Then
        
            Call WriteStatus("Saving Remarks " & x & " of " & i)
            
            If rsPOREM.Status = adRecNew Then
                If Not PutPOREMDataInsert Then Exit Function
            Else
                If Not PutPOREMDataUpdate Then Exit Function
            End If
            
            If Err = 0 Then
                Call WriteStatus("Remarks " & x & " of " & i & " saved successfully")
            Else
                Call LogErr(Name & "::SaveRemarks", Err.Description, Err, True)
                Exit Function
            End If
        
        End If
        
        DoEvents
        
        If Err Then
            If Err <> 3021 Then MsgBox Err.Description
            Call LogErr(Name & "::SaveRemarks", Err.Description, Err, True)
        End If
        
        rsPOREM.MoveNext
    Loop



    rsPOREM.Filter = 0
    SaveRemarks = True
    If Err Then Err.Clear
End Function

Public Function SaveClause() As Boolean

On Error Resume Next
Dim STR As String, i As Integer, x As Integer

    SaveClause = True
    If Requery(3) = True Then Exit Function
    If rsPOCLAUSE.State = adStateClosed Then Exit Function

    rsPOCLAUSE.Update
    rsPOCLAUSE.Filter = 0
    rsPOCLAUSE.Filter = adFilterPendingRecords


    rsPOCLAUSE.MoveFirst
    SaveClause = False
    i = rsPOCLAUSE.RecordCount

    Err.Clear
    Do Until rsPOCLAUSE.EOF
    
        x = rsPOCLAUSE.AbsolutePosition
        STR = Trim$(rsPOCLAUSE!poc_clau & "")


        If Err Then Exit Function
        
        If Len(STR) > 0 Then
        
            Call WriteStatus("Saving Clause " & x & " of " & i)
            
            If rsPOCLAUSE.Status = adRecNew Then
                If Not PutPoCluaseInsert Then Exit Function
            Else
                If Not PutPoCluaseUpdate Then Exit Function
            End If
            
            If Err = 0 Then
                Call WriteStatus("Clause " & x & " of " & i & " saved successfully")
            Else
                Call LogErr(Name & "::SaveClause", Err.Description, Err, True)
                
                Exit Function
            End If
        
        End If
        
        DoEvents
        
        If Err Then
            MsgBox Err.Description
            Call LogErr(Name & "::SaveClause", Err.Description, Err, True)
        End If
        rsPOCLAUSE.MoveNext
    Loop

    rsPOCLAUSE.Filter = 0
    SaveClause = True
End Function

Public Function SaveRecipients() As Boolean

On Error Resume Next
Dim STR As String, i As Integer, x As Integer

    SaveRecipients = True
    If Requery(0) = True Then Exit Function
    If rsrecepList.State = adStateClosed Then Exit Function

    rsrecepList.Update
    'rsrecepList.Filter = 0
    'rsrecepList.Filter = adFilterPendingRecords


    rsrecepList.MoveFirst
    SaveRecipients = False
    i = rsrecepList.RecordCount
    
    
    Do Until rsrecepList.EOF
    
        x = rsrecepList.AbsolutePosition
        STR = Trim$(rsrecepList!porc_rec & "")


        If Err Then Exit Function

        If Len(STR) > 0 Then
        
            Call WriteStatus("Saving Recipients " & x & " of " & i)
            
            If rsrecepList.Status = adRecNew Then
                If Not PutPORECDataInsert Then Exit Function
            Else
                If Not PutPORECDataUpdate Then Exit Function
            End If
            
            If Err = 0 Then
                Call WriteStatus("Recipients " & x & " of " & i & " saved successfully")
            Else
                Call LogErr(Name & "::SaveRecipients", Err.Description, Err, True)
                Exit Function
            End If
        
        End If
        
        DoEvents
        
        If Err Then
            MsgBox Err.Description
            Call LogErr(Name & "::SaveRecipients", Err.Description, Err, True)
            
            Exit Function
        End If
        rsrecepList.MoveNext
    Loop


    rsPOREM.Filter = 0
    SaveRecipients = True
End Function

Public Sub LoadData()
    POIChange
    PORemChange
    PORECChange
    POCLAUSEChange
End Sub
