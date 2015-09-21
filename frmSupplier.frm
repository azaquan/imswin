VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frm_IntSupe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier "
   ClientHeight    =   7695
   ClientLeft      =   180
   ClientTop       =   210
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11700
   Tag             =   "01010101"
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   3720
      TabIndex        =   51
      Top             =   6360
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   54853633
      CurrentDate     =   37316
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   1200
      TabIndex        =   52
      Top             =   6600
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   54853633
      CurrentDate     =   37316
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   240
      TabIndex        =   15
      Top             =   7080
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frmSupplier.frx":0000
      EditVisible     =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown ssdddContacts 
      Height          =   1335
      Left            =   8880
      TabIndex        =   14
      Top             =   4440
      Width           =   1215
      _Version        =   196617
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2143
      _ExtentY        =   2355
      _StockProps     =   77
   End
   Begin TabDlg.SSTab sstSup 
      Height          =   6855
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   512
      TabCaption(0)   =   "Supplier Record"
      TabPicture(0)   =   "frmSupplier.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl_remarks"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_InterSupp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_City"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Address2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Email"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_PhoneNum"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl_telNo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lbl_Fax"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SSOleDBCombo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "SSDBLine"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtRemarks"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtSuppCode"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_forSEARCH"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Remarks"
      TabPicture(1)   =   "frmSupplier.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Recipients"
      TabPicture(2)   =   "frmSupplier.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1(0)"
      Tab(2).Control(1)=   "cmd_Add"
      Tab(2).Control(2)=   "cmd_Remove"
      Tab(2).Control(3)=   "ssdbRecepientList"
      Tab(2).Control(4)=   "lbl_Recipients"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Contacts"
      TabPicture(3)   =   "frmSupplier.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ssdbgContacts"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Contract"
      TabPicture(4)   =   "frmSupplier.frx":008C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ssdbContract"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txt_forSEARCH 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   420
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3040
         Width           =   3530
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   780
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   5655
         Begin VB.TextBox Txt_contaname 
            DataField       =   "sup_contaname"
            DataMember      =   "INtSupplier"
            DataSource      =   "deIms"
            Height          =   288
            Left            =   2400
            MaxLength       =   35
            TabIndex        =   10
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txt_FaxNumber 
            DataField       =   "sup_faxnumb"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   9
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Lbl_Name 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Name"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lbl_FaxNum 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax #"
            DataMember      =   "SUPPLIER"
            Height          =   210
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.TextBox TxtSuppCode 
         DataField       =   "sup_code"
         DataMember      =   "INtSupplier"
         DataSource      =   "deIms"
         Enabled         =   0   'False
         Height          =   288
         Left            =   6000
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3135
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   3135
         ScaleWidth      =   11175
         TabIndex        =   20
         Top             =   2760
         Width           =   11175
         Begin ImsMailVB.Imsmail Imsmail1 
            Height          =   3135
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   5530
         End
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74760
         TabIndex        =   19
         Top             =   2115
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74775
         TabIndex        =   18
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox TxtRemarks 
         DataField       =   "sup_remk"
         DataMember      =   "INTSUPPLIER"
         Height          =   915
         Left            =   120
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   5760
         Width           =   11175
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   2115
         Left            =   -73245
         TabIndex        =   22
         Top             =   585
         Width           =   9090
         _Version        =   196617
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   8176
         Columns(0).Caption=   "Recipients"
         Columns(0).Name =   "Recp"
         Columns(0).DataField=   "Recipients"
         Columns(0).FieldLen=   256
         _ExtentX        =   16034
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBLine 
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   11115
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldSeparator  =   ";"
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   6218
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "sup_name"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3493
         Columns(1).Caption=   "Telephone"
         Columns(1).Name =   "Telephone"
         Columns(1).DataField=   "sup_phonnumb"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3043
         Columns(2).Caption=   "City"
         Columns(2).Name =   "City"
         Columns(2).DataField=   "sup_city"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3440
         Columns(3).Caption=   "Fax"
         Columns(3).Name =   "Fax"
         Columns(3).DataField=   "sup_faxnumb"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   5292
         Columns(4).Caption=   "Email"
         Columns(4).Name =   "Email"
         Columns(4).DataField=   "Sup_mail"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   19606
         _ExtentY        =   3625
         _StockProps     =   79
         BackColor       =   -2147483638
         DataMember      =   "INTSUPPLIER"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgContacts 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   11100
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
         BorderStyle     =   0
         FieldSeparator  =   ";"
         Col.Count       =   5
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
         stylesets(0).Picture=   "frmSupplier.frx":00A8
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
         stylesets(1).Picture=   "frmSupplier.frx":00C4
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowRowSizing  =   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   5292
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Contacts"
         Columns(0).Name =   "Contacts"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4419
         Columns(1).Caption=   "Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3175
         Columns(2).Caption=   "Tel"
         Columns(2).Name =   "Tel"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3175
         Columns(3).Caption=   "Fax"
         Columns(3).Name =   "Fax"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3519
         Columns(4).Caption=   "Email"
         Columns(4).Name =   "Email"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   19579
         _ExtentY        =   10610
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
         DataSource      =   "deIms"
         Height          =   330
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
         ListAutoValidate=   0   'False
         AutoRestore     =   0   'False
         _Version        =   196617
         Cols            =   6
         Columns(0).Width=   3200
         _ExtentX        =   6165
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   8160
         TabIndex        =   35
         Top             =   480
         Width           =   3135
         Begin VB.TextBox Txt_contaFax 
            DataField       =   "sup_contaFax"
            DataMember      =   "INtSupplier"
            DataSource      =   "deIms"
            Height          =   288
            Left            =   0
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1800
            Width           =   1830
         End
         Begin VB.TextBox Txt_contaPH 
            DataField       =   "sup_contaph"
            DataMember      =   "INtSupplier"
            DataSource      =   "deIms"
            Height          =   288
            Left            =   0
            MaxLength       =   25
            TabIndex        =   12
            Top             =   2160
            Width           =   1830
         End
         Begin VB.CheckBox chk_Active 
            Alignment       =   1  'Right Justify
            Caption         =   "Active?"
            DataField       =   "sup_actvflag"
            DataMember      =   "INTSUPPLIER"
            Height          =   192
            Left            =   1800
            TabIndex        =   36
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txt_PhoneNumber 
            DataField       =   "sup_phonnumb"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   0
            MaxLength       =   25
            TabIndex        =   8
            Top             =   1320
            Width           =   1830
         End
         Begin VB.TextBox txt_Email 
            DataField       =   "sup_mail"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   0
            MaxLength       =   59
            TabIndex        =   1
            Top             =   240
            Width           =   3030
         End
         Begin VB.TextBox txt_Address2 
            DataField       =   "sup_adr2"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   0
            MaxLength       =   25
            TabIndex        =   3
            Top             =   600
            Width           =   3024
         End
         Begin VB.TextBox txt_City 
            DataField       =   "sup_city"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   0
            MaxLength       =   25
            TabIndex        =   6
            Top             =   960
            Width           =   3024
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txt_Name 
            DataField       =   "sup_name"
            DataMember      =   "intsupplier"
            Height          =   285
            Left            =   2400
            MaxLength       =   35
            TabIndex        =   45
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox txt_Address1 
            DataField       =   "sup_adr1"
            DataMember      =   "intsupplier"
            Height          =   288
            Left            =   2400
            MaxLength       =   25
            TabIndex        =   2
            Top             =   480
            Width           =   3024
         End
         Begin VB.TextBox txt_State 
            DataField       =   "sup_stat"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   4
            Top             =   840
            Width           =   525
         End
         Begin VB.TextBox txt_Country 
            DataField       =   "sup_ctry"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   2400
            MaxLength       =   25
            TabIndex        =   7
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txt_Zipcode 
            DataField       =   "sup_zipc"
            DataMember      =   "INTSUPPLIER"
            Height          =   288
            Left            =   4260
            MaxLength       =   11
            TabIndex        =   5
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lbl_Sup_Name 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Width           =   2325
         End
         Begin VB.Label lbl_Address1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   285
            Left            =   0
            TabIndex        =   44
            Top             =   480
            Width           =   2325
         End
         Begin VB.Label lbl_State 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   285
            Left            =   0
            TabIndex        =   43
            Top             =   840
            Width           =   2325
         End
         Begin VB.Label lbl_Zip 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code"
            DataMember      =   "SUPPLIER"
            Height          =   285
            Left            =   3000
            TabIndex        =   42
            Top             =   885
            Width           =   1125
         End
         Begin VB.Label lbl_Country 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   285
            Left            =   0
            TabIndex        =   41
            Top             =   1200
            Width           =   2325
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbContract 
         Height          =   5295
         Left            =   -74760
         TabIndex        =   50
         Top             =   840
         Width           =   10875
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldSeparator  =   ";"
         Col.Count       =   3
         HeadFont3D      =   4
         DefColWidth     =   5292
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         MaxSelectedRows =   10
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   5292
         Columns(0).Caption=   "contractNum"
         Columns(0).Name =   "contractNum"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4339
         Columns(1).Caption=   "startDate"
         Columns(1).Name =   "startDate"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4313
         Columns(2).Caption=   "stopDate"
         Columns(2).Name =   "stopDate"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   19182
         _ExtentY        =   9340
         _StockProps     =   79
         BackColor       =   -2147483638
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Search Field"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   49
         Top             =   3060
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   48
         Top             =   3060
         Width           =   255
      End
      Begin VB.Label Lbl_Fax 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Fax #"
         Height          =   255
         Left            =   5880
         TabIndex        =   34
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Lbl_telNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Telephone #"
         Height          =   255
         Left            =   5280
         TabIndex        =   33
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label lbl_PhoneNum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone #"
         Height          =   285
         Left            =   6000
         TabIndex        =   32
         Top             =   1800
         Width           =   2085
      End
      Begin VB.Label lbl_Email 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID"
         Height          =   285
         Left            =   6000
         TabIndex        =   31
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label lbl_Address2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Left            =   6000
         TabIndex        =   30
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label lbl_City 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   285
         Left            =   6000
         TabIndex        =   29
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label lbl_InterSupp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "International Supplier"
         DataMember      =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   27
         Top             =   0
         Width           =   2835
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Lbl_remarks 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   5520
         Width           =   2415
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
      Left            =   6360
      TabIndex        =   16
      Top             =   6960
      Width           =   3300
   End
End
Attribute VB_Name = "frm_IntSupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim rsReceptList As ADODB.Recordset
Dim OrigEdit As Boolean
Dim OrigNew As Boolean
Dim FormMode As FormMode
Dim x As Integer
Dim xx As Integer
Dim rowguid, locked As Boolean
Dim dbtablename As String 'jawdat


Private Function SaveContract()
Dim x As Integer, y As Integer
Dim SupCode As String, np As String
Dim cmd As ADODB.Command

Err.Clear
On Error Resume Next

    If Not FindSupplier(rs!sup_code) Then SaveContract = False: Exit Function

    x = ssdbContract.Rows - 1
    If x < 0 Then Exit Function
 
    np = deIms.NameSpace
    ssdbContract.MoveFirst
    Set cmd = New ADODB.Command
    SupCode = rs!sup_code & ""
    SupCode = Trim$(SupCode)

    'If SupCode = "" Then Stop 'Hidden by Juan
 
    With cmd
        .Prepared = False
        .CommandType = adCmdText
       .ActiveConnection = deIms.cnIms
         Call BeginTransaction(deIms.cnIms)
        .CommandText = "Delete from SUPPLIERCONTRACT where scrt_supcode = ? and scrt_npecode = ?"
        Call .Execute(0, Array(SupCode, np), adExecuteNoRecords)
        .Prepared = True
        Call CommitTransaction(deIms.cnIms)

        Call BeginTransaction(deIms.cnIms)
        Dim sql As String
        
        .CommandText = "Insert into SUPPLIERCONTRACT(scrt_npecode, scrt_supcode, scrt_contractnum,scrt_startdate,scrt_stopdate)"
        .CommandText = .CommandText & "VALUES(?,?,?,?,?)"
        cmd.parameters.Refresh
        For y = 0 To x
            Call .Execute(0, Array(np, SupCode, ssdbContract.Columns(0).value, ssdbContract.Columns(1).value, ssdbContract.Columns(2).value), adExecuteNoRecords)
            ssdbContract.MoveNext
        Next
        Call CommitTransaction(deIms.cnIms)
    End With

    Set cmd = Nothing
    If Err Then Err.Clear
End Function

'add current cecipient to recipient list

Private Sub cmd_Add_Click()
    Imsmail1.AddCurrentRecipient
End Sub

'remove current cecipient from recipient list

Private Sub cmd_Remove_Click()
On Error Resume Next

    If rsReceptList.RecordCount Then rsReceptList.Delete
    If Err Then Err.Clear
End Sub

Private Sub MonthView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then MonthView1.Visible = False
    If KeyAscii = 13 Then MonthView1.Visible = False
End Sub

Private Sub MonthView1_LostFocus()
    MonthView1.Visible = False
End Sub

Private Sub MonthView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim dates As Date
 Dim s
s = MonthView1.HitTest(x, y, dates)

If s = 1 Or s = 2 Or s = 3 Then

    MonthView1.value = dates

    ssdbContract.Columns(1).Text = dates

    MonthView1.Visible = False

    ' Call GetDataForTheSelection

End If
End Sub


Private Sub MonthView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then MonthView2.Visible = False
    If KeyAscii = 13 Then MonthView2.Visible = False
End Sub

Private Sub MonthView2_LostFocus()
    MonthView2.Visible = False
End Sub

Private Sub MonthView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim dates As Date
 Dim s
s = MonthView1.HitTest(x, y, dates)

If s = 1 Or s = 2 Or s = 3 Then

    MonthView1.value = dates

    ssdbContract.Columns(2).Text = dates

    MonthView2.Visible = False

    'Call GetDataForTheSelection

End If
End Sub


Private Sub NavBar1_OnEditClick()

'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
 Dim imsLock As imsLock.Lock
 Dim ListOfPrimaryControls() As String
 Set imsLock = New imsLock.Lock

  ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)

  Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)   'lock should be here, added by jawdat, 2.1.02

If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
locked = True
End If

'jawdat, end copy

FormMode = ChangeModeOfForm(lblStatus, mdModification)
If FormMode = mdModification Then MakeReadOnly (True)
End Sub

Private Sub ssdbContract_AfterUpdate(RtnDispErrMsg As Integer)
    'If ContactExist Then
    '    RtnDispErrMsg = False

        'Modified by Juan (9/11/2000) for Multilingual
  '      msg1 = translator.Trans("M00257") 'J added
   '     MsgBox IIf(msg1 = "", "Contact is already in the list", msg1) 'J modified
        '---------------------------------------------

'        Call ssdbgContacts.RemoveItem(ssdbgContacts.row)
 '   End If
End Sub

Private Sub ssdbContract_Click()
    If ssdbContract.Col = 1 Then
       Call SetFocusOnDatesColumns(1)
    ElseIf ssdbContract.Col = 2 Then
       Call SetFocusOnDatesColumns(2)
    End If
End Sub


Private Sub ssdbContract_InitColumnProps()
'    LockWindowUpdate (HWND)
'    With ssdbgContacts.Columns(0)

        'Modified by Juan (9/12/2000) for Multilingual
'        msg1 = translator.Trans("L00165") 'J added
'        .Caption = IIf(msg1 = "", "Contacts", msg1) 'J modified
        '---------------------------------------------

'        .Name = "Contacts"
'        .CaptionAlignment = 2
'        .DataField = "sct_contcode"
'        .DataType = 8
'        .FieldLen = 10
'        .HeadStyleSet = "ColHeader"
'        .StyleSet = "RowFont"
'        .Width = 6400
'        .DropDownHwnd = ssdddContacts.HWND
'    End With

'    ssdbgContacts.Refresh
'    ssdbgContacts.MoveFirst
'    LockWindowUpdateOff
End Sub

Private Sub SSDBLine_KeyPress(KeyAscii As Integer)
''SSDBLine.ListAutoValidate
'If NavBar1.SaveEnabled = True Then
''Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
' Dim imsLock As imsLock.lock
' Dim ListOfPrimaryControls() As String
' Set imsLock = New imsLock.lock
'
'  ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'
'  Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid)   'lock should be here, added by jawdat, 2.1.02
'
'If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
'Else
'locked = True
'End If
'
'End If
End Sub

Private Sub SSOleDBCombo1_Click()
     deIms.rsINtSupplier.CancelUpdate
     deIms.rsINtSupplier.MoveFirst
If Not deIms.rsINtSupplier.State = 0 Then deIms.rsINtSupplier.Find "sup_code='" & (SSOleDBCombo1.Columns(0).value) & "'", adSearchForward
End Sub

Private Sub SSOleDBCombo1_DropDown()
'SSOleDBCombo1.DroppedDown = True
If deIms.rsINTSUPPLIERLOOKUP.State = 1 Then deIms.rsINTSUPPLIERLOOKUP.Close
Call deIms.INTSUPPLIERLOOKUP(deIms.NameSpace)

End Sub

Private Sub SSOleDBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
   If FormMode = mdvisualization Then
    If Not SSOleDBCombo1.DroppedDown Then SSOleDBCombo1.DroppedDown = True
   End If
End Sub

Private Sub SSOleDBCombo1_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
       SSOleDBCombo1.Text = SSOleDBCombo1.SelText
          deIms.rsINtSupplier.CancelUpdate
          deIms.rsINtSupplier.MoveFirst
          If Not deIms.rsINtSupplier.State = 0 Then deIms.rsINtSupplier.Find "sup_code='" & (SSOleDBCombo1.Columns(0).value) & "'", adSearchForward
    Else
      SSOleDBCombo1.MoveNext
  
 End If
'Dim data As String
'data = SSOleDBCombo1.text & Chr$(KeyAscii)
'SSOleDBCombo1_DropDown
'deIms.rsINtSupplier.MoveFirst
'deIms.rsINtSupplier.Find "sup_name like '" & Trim$(data) & "%'", adSearchForward
''''Call SSOleDBCombo1.Scroll(1, 15)
 
End Sub

Private Sub txt_forSEARCH_Change()
Dim n As Integer
    With deIms.rsINtSupplier
        If .RecordCount > 0 Then
            n = Len(txt_forSEARCH)
            If n > 0 And Not .EOF Then
                If UCase(Left(!sup_name, n)) <> UCase(txt_forSEARCH) Then
                    .MoveFirst
                    .Find "sup_name like '" + txt_forSEARCH + "%'"
                End If
            Else
                .MoveFirst
            End If
        End If
    End With
End Sub

Private Sub Txt_contaFax_GotFocus()
    Call HighlightBackground(Txt_contaFax)
End Sub

Private Sub Txt_contaFax_LostFocus()
    Call NormalBackground(Txt_contaFax)
End Sub


Private Sub Txt_contaname_GotFocus()
    Call HighlightBackground(Txt_contaname)
End Sub


Private Sub Txt_contaname_LostFocus()
    Call NormalBackground(Txt_contaname)
End Sub


Private Sub Txt_contaPH_GotFocus()
    Call HighlightBackground(Txt_contaPH)
End Sub


Private Sub Txt_contaPH_LostFocus()
    Call NormalBackground(Txt_contaPH)
End Sub

Private Sub txt_Name_GotFocus()
    Call HighlightBackground(txt_Name)
End Sub


Private Sub txt_Name_LostFocus()
    Call NormalBackground(txt_Name)
End Sub


Private Sub txtRemarks_GotFocus()
    If FormMode = mdvisualization Then NavBar1.SaveEnabled = False
End Sub


'get supplier codeDATAGRI
'Muzammil 12/18/00
'Reason - DcboSuppcode is now a Text box and will be Disabled all the time.
'''''Private Sub TxtSuppCode_Click(Area As Integer)
'''''Dim locked As Boolean
'''''On Error Resume Next
'''''
'''''    Dim STR As String
'''''
'''''  '  dcboSuppCode.locked = True
'''''
'''''    If deIms.rsIntSupplier.editmode = adEditAdd Then 'M
'''''    Area = 1  'M
'''''    End If    'M
'''''
'''''
'''''
'''''    If Area = 2 Then
'''''
'''''        'locked = dcboSuppCode.locked
'''''        'dcboSuppCode.locked = False
'''''        deIms.rsIntSupplier.CancelUpdate
'''''        If Err Then Err.Clear
'''''
'''''        STR = TxtSuppCode
'''''        Call FindSup(STR)
'''''        Call FindSup(STR)
'''''
'''''    End If
'''''
'''''End Sub

'set combo back ground color

Private Sub TxtSuppCode_GotFocus()
    Call HighlightBackground(TxtSuppCode)
End Sub

'Private Sub dcboSuppCode_KeyDown(KeyCode As Integer, Shift As Integer)
   ' Call dcboSuppCode_Click(2)
'End Sub

'Private Sub dcboSuppCode_KeyPress(KeyAscii As Integer)
'Dim rs As ADODB.Recordset

'    Set rs = deIms.rsIntSupplier
'    If rs.EditMode = adEditAdd Then Exit Sub
'
'    If KeyAscii = 13 Then KeyAscii = 0
'    If (((KeyAscii = 8) Or (KeyAscii > 31)) Or (KeyAscii = 0)) Then _
'        If GetNearestComboItem(dcboSuppCode, KeyAscii) Then FindSup (dcboSuppCode)
'End Sub


'Muzammil - 12/18/00.
'Reason- Since dcboSuppcode is now a text box and it Will Always be disabled
',Since we are generating Auto Code for it.

''Private Sub dcboSuppCode_KeyPress(KeyAscii As Integer)
''If deIms.rsIntSupplier.editmode <> adEditAdd Then
''   KeyAscii = 0
'' End If
''End Sub

'set combo back ground color

''Private Sub dcboSuppCode_LostFocus()
''    Call NormalBackground(dcboSuppCode)
''End Sub

'set combo back ground color

Private Sub chk_Active_GotFocus()
    Call HighlightBackground(chk_Active)
End Sub

'set combo back ground color

Private Sub chk_Active_LostFocus()
    Call NormalBackground(chk_Active)
End Sub
'Muzammil - 12/18/00.
'Reason- Since dcboSuppcode is now a text box and it Will Always be disabled
'lock supplier combo

''''Private Sub dcboSuppCode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''''    dcboSuppCode.locked = False
''''End Sub
''''
'''''lock supplier combo
''''
''''Private Sub dcboSuppCode_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
''''
''''  dcboSuppCode.locked = Not NavBar1.NewEnabled
''''End Sub

'Muzammil - 12/18/00.
'Reason- Since dcboSuppcode is now a text box and it Will Always be disabled
'Private Sub dcboSuppCode_Validate(Cancel As Boolean)
'Dim msg As String
'Dim code As String
'Dim str As String, OldNum As String
'
'
'    str = LCase(dcboSuppCode.Text)
'    OldNum = LCase(Trim$(rs("sup_code").OriginalValue & ""))
'
'
'    If Len(OldNum) Then
'
'        If OldNum <> str Then
'            dcboSuppCode.Text = OldNum
'            MsgBox "Supplier code cannot be changed once saved"
'        End If
'
'    Else
'
'    If Len(str) Then
'         If Len(Trim$(dcboSuppCode)) <> 0 Then
'                code = Trim(dcboSuppCode)
'                If CheckSupplierCode(code, False) Then
'                        Cancel = True
'                        MsgBox "Supplier code  ' " & str & " ' already exist, Please make new one."
'
'                End If
'        End If
'    End If
'    End If
'            If deIms.StockNumberExist(str, False) Then
'                Cancel = True
'                MsgBox "Stock number " & str & " already exist"
'            End If
'
'        End If
'
'    If Len(Trim$(dcboSuppCode)) <> 0 Then
'         msg = LCase(Trim$(rs("sup_code").OriginalValue & ""))
'
'         If Len(msg) Then
'            If (LCase(Trim$(dcboSuppCode)) <> (msg)) Then
''            If (LCase(Trim$(rs("sup_code"))) <> (msg)) Then
''
'                rs("sup_code") = msg
'                MsgBox "Supplier code cannot be changed once it is saved"
'            End If
'
'        End If
'    Else
'        If Len(Trim$(dcboSuppCode)) <> 0 Then
'                code = Trim(dcboSuppCode)
'                If CheckSupplierCode(code) Then
'
'                        MsgBox "Supplier code exist, Please make new one."
'                        Exit Sub
'                End If
'        End If
''    End If


'Muzammil - 12/18/00.
'Reason- Since dcboSuppcode is now a text box and it Will Always be disabled

'Added by Muzammil. Ticket no - 54.11/06/00
''''If FindSupplier(Trim$(dcboSuppCode.text)) Then
''''    'msg1 = translator.Trans("
''''    msg1 = translator.Trans("M00254") 'M added
''''    MsgBox IIf(msg1 = "", "Supplier code exist, Please use a different code.", msg1) 'M modified
''''    Cancel = True
''''    dcboSuppCode.SetFocus
'''' End If

'End Sub

'Private Sub dcboSuppCode_Validate(Cancel As Boolean)
'    If Len(Trim$(dcboSuppCode)) <> 0 Then
'        If CheckSupplierCode(dcboSuppCode) Then
'
'                MsgBox "You can not change Supplier Code"
'        End If
'    End If
'
'End Sub

'unload supplier form, and free memory

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

    Hide
    rs.CancelUpdate
   ' Imsmail1.Connected = False 'M

    rs.Update
    rs.UpdateBatch
    Set rs = Nothing
    deIms.rsINtSupplier.Close
    Set frm_IntSupe = Nothing
     If open_forms <= 5 Then ShowNavigator

    If Err Then Err.Clear
    

   Dim imsLock As imsLock.Lock
   Set imsLock = New imsLock.Lock
   Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


End Sub

'get recepient list and populate recepient data grid

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

'cancel record set update

Private Sub NavBar1_OnCancelClick()
  Dim i As EditModeEnum
 'Muzammil 12/19/00


 Dim imsLock As imsLock.Lock
  Set imsLock = New imsLock.Lock
  Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat

 
 
  FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
 If FormMode = mdvisualization Then MakeReadOnly (False)
 
    'Added by Juan Gonzalez 2007-7-11
    If sstSup.Tab = 4 Then
        ssdbContract.CancelUpdate
    Else '------------------------------
        If sstSup.Tab = 3 Then
            ssdbgContacts.CancelUpdate
        ElseIf sstSup.Tab = 0 Then
            i = deIms.rsINtSupplier.EditMode
            rs.CancelUpdate
            deIms.rsINtSupplier.CancelUpdate
            
            If Not i = adEditAdd Then GetOriginalValues
            
            'Call deIms.rsIntSupplier.CancelBatch(adAffectCurrent)
            'dcboSuppCode.locked = True
        End If
    End If 'JG 2007-7-11

End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Function chk_LI() As String
'    Dim chk_LI As String
'    Dim ls_code As String
'    Dim Focus_Flag As Boolean
'
'    chk_LI = ""
'    Focus_Flag = False
'    If Trim(txt_SupName.Text) = "" Then
'        chk_LI = chk_LI & "Name, "
'        If Focus_Flag = False Then
'            txt_SupName.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(txt_Address1.Text) = "" Then
'        chk_LI = chk_LI & "Address(1), "
'        If Focus_Flag = False Then
'            txt_Address1.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(txt_City.Text) = "" Then
'        chk_LI = chk_LI & "City, "
'        If Focus_Flag = False Then
'            txt_City.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(txt_State.Text) = "" Then
'        chk_LI = chk_LI & "State, "
'        If Focus_Flag = False Then
'            txt_State.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(txt_Country.Text) = "" Then
'        chk_LI = chk_LI & "Country, "
'        If Focus_Flag = False Then
'            txt_Country.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(txt_PhoneNumber.Text) = "" Then
'        chk_LI = chk_LI & "Phone Number, "
'        If Focus_Flag = False Then
'            txt_PhoneNumber.SetFocus
'            Focus_Flag = True
'        End If
'    End If
'    If Trim(dcboSuppCode.Text) = "" Then
'        chk_LI = chk_LI & "Code, "
'        If Focus_Flag = False Then
'            dcboSuppCode.SetFocus
'            Focus_Flag = True
'        End If
'    End If
End Function

'get crystal report parameter and application path

Private Sub PrintCurrent()
Dim Path As String
On Error GoTo ErrHandler

    Path = FixDir(App.Path) + "CRreports\"

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = Path & "Supplier.rpt"
        .ParameterFields(2) = "IntLoc;" & "INT" & ";TRUE"
        .ParameterFields(1) = "suppcode;" & TxtSuppCode & ";TRUE"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"

        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .WindowTitle = IIf(msg1 = "", "Supplier", msg1) 'J modified
        Call translator.Translate_Reports("Supplier.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------

    End With

    Exit Sub

ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'get crystal report parameter and application path

Private Sub PrintAll()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Supplier.rpt"
        .ParameterFields(2) = "IntLoc;" & "INT" & ";TRUE"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "suppcode;ALL;TRUE"

        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("L00128") 'J added
        .WindowTitle = IIf(msg1 = "", "Supplier", msg1) 'J modified
        Call translator.Translate_Reports("Supplier.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------

    End With
End Sub

'delete a record from data grid

Private Sub NavBar1_OnDeleteClick()

    'Added by Juan 2007-7-11
    If sstSup.Tab = 4 Then
        Call ssdbContract.RemoveItem(ssdbContract.row)
    Else '------
        If sstSup.Tab = 3 Then
            Call ssdbgContacts.RemoveItem(ssdbgContacts.row)
            NavBar1.SaveEnabled = True 'JCG 2008/1/13
        End If
    End If 'JG 2007-7-1
End Sub


'get email parameter to email report,set memory free

Private Sub NavBar1_OnEMailClick()

Dim ParamsForRPTI(2) As String

Dim rptinf As RPTIFileInfo

Dim ParamsForCrystalReports(2) As String

Dim subject As String

Dim FieldName As String

Dim Message As String

Dim attention As String

On Error Resume Next

If rsReceptList Is Nothing Then Exit Sub




    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"

    ParamsForCrystalReports(1) = "suppcode;" + TxtSuppCode + ";TRUE"

    ParamsForCrystalReports(2) = "Intloc;" + "INT" + ";TRUE"

    ParamsForRPTI(0) = "namespace=" & deIms.NameSpace

    ParamsForRPTI(1) = "suppcode=" & TxtSuppCode
    
    ParamsForRPTI(2) = "Intloc=" & "INT"

    FieldName = "Recipients"
    
    subject = "Supplier"
    
    If ConnInfo.EmailClient = Outlook Then

        'Call sendOutlookEmailandFax("Supplier.rpt", "Supplier", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, attention)  MM 030209 EFCR11
        Call sendOutlookEmailandFax(Report_EmailFax_Supplier_name, "Supplier", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, attention)

    ElseIf ConnInfo.EmailClient = ATT Then

        Call SendAttFaxAndEmail("Supplier.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, Message, FieldName)

    ElseIf ConnInfo.EmailClient = Unknown Then

        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

    End If

    Set rsReceptList = Nothing

    Set ssdbRecepientList.DataSource = Nothing


'''''Call SelectGatewayAndSendOutMails

End Sub

'move record set to first position

Private Sub NavBar1_OnFirstClick()
Dim i As EditModeEnum

'If locked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else

On Error Resume Next

'Added by Juan Gonzalez 2007-7-11
If sstSup.Tab = 4 Then
    ssdbContract.MoveFirst
Else '----------
    If sstSup.Tab = 3 Then
        ssdbgContacts.MoveFirst
    Else

        With deIms.rsINtSupplier

            i = .EditMode

            If (ValidateData) Then
                If ((i = adEditAdd)) Then NavBar1_OnSaveClick

                .MoveFirst
            End If

        End With

    End If
End If 'JG 2007-7-11
End Sub

'move record set to last position

Private Sub NavBar1_OnLastClick()
Dim i As EditModeEnum


'If locked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else

On Error Resume Next



'Added by Juan Gonzalez 2007-7-11
If sstSup.Tab = 4 Then
    ssdbContract.MoveLast
Else '----------------
    If sstSup.Tab = 3 Then
        ssdbgContacts.MoveLast
    Else

        With deIms.rsINtSupplier

            i = .EditMode

            If (ValidateData) Then
                If ((i = adEditAdd)) Then NavBar1_OnSaveClick

                .MoveLast
            End If

        End With

    End If
End If 'JG 2007-7-11
End Sub

'add new click set suplier modity user to current user and
'create user to current user, name space to current name space

Private Sub NavBar1_OnNewClick()
Dim i As EditModeEnum

'If locked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else

On Error Resume Next

FormMode = ChangeModeOfForm(lblStatus, mdCreation)
If FormMode = mdCreation Then MakeReadOnly (True)
    
    If sstSup.Tab = 3 Then
        FormMode = ChangeModeOfForm(lblStatus, mdModification)
        If FormMode = mdCreation Or FormMode = mdModification Then
             NavBar1.NewEnabled = True
         End If
        ssdbgContacts.Update
        Call ssdbgContacts.AddItem("")
        ssdbgContacts.row = ssdbgContacts.Rows - 1
    Else
        'Added by Juan 2007/7/7
        If sstSup.Tab = 4 Then
            FormMode = ChangeModeOfForm(lblStatus, mdModification)
            If FormMode = mdCreation Or FormMode = mdModification Then
                 NavBar1.NewEnabled = True
            End If
             
            ssdbContract.Update
            Call ssdbContract.AddItem("")
            ssdbContract.row = ssdbContract.Rows - 1
        '-------------------------
        Else
            With deIms.rsINtSupplier
    
               i = .EditMode
    
               'If (ValidateData) Then  'M
               '    If ((i = adEditAdd)) Then NavBar1_OnSaveClick  'M
    
                    .AddNew
                  'cOMMENTED Out By Muzammil.20/01/01
                    
    '''                !sup_creauser = CurrentUser
    '''                !sup_modiuser = CurrentUser
                    !SUP_FLAG = True
                    !sup_npecode = deIms.NameSpace
               ' End If
            End With
        End If
'    End If
  If Err.number > 0 Then Err.Clear
    TxtSuppCode.locked = False
End If
End Sub

'move record set to next position

Private Sub NavBar1_OnNextClick()
On Error Resume Next

'If locked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else

Dim i As EditModeEnum


'Added by Juan Gonzalez 2007-7-11
If sstSup.Tab = 4 Then
    ssdbContract.MoveNext
Else '---------
    If sstSup.Tab = 3 Then
        ssdbContract.MoveNext
    Else

        With deIms.rsINtSupplier
            i = .EditMode

            If (ValidateData) Then
                If ((i = adEditAdd)) Then NavBar1_OnSaveClick

                .MoveNext
                If ((.EOF) And (.RecordCount > 0)) Then .MoveLast
            End If
        End With
    End If
End If 'JG 2007-7-11
End Sub

'move record set to previous position

Private Sub NavBar1_OnPreviousClick()
On Error Resume Next

'If locked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'Else


Dim i As EditModeEnum


'Added by Juan Gonzalez 2007-7-11
If sstSup.Tab = 4 Then
    ssdbContract.MovePrevious
Else '------
    If sstSup.Tab = 3 Then
        ssdbgContacts.MovePrevious
    Else

        With deIms.rsINtSupplier
            i = .EditMode

            If (ValidateData) Then
                If ((i = adEditAdd)) Then NavBar1_OnSaveClick

                .MovePrevious
                If ((.BOF) And (.RecordCount > 0)) Then .MoveFirst
            End If
        End With

    End If
End If 'JG 2007-7-11
End Sub

'call function to print report

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handled

Dim retval As PrintOpts

    Load frmPrintDialog
    'frmPrintDialog.optprintSel.Visible = False
    With frmPrintDialog

        .Show 1
        retval = .Result

        DoEvents: DoEvents
        If retval = poPrintCurrent Then

            PrintCurrent

        ElseIf retval = poPrintAll Then
            PrintAll

        Else
            Exit Sub

        End If

    End With


'    MDI_IMS.CrystalReport1 = "Supplier"
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset

    Unload frmPrintDialog
    Set frmPrintDialog = Nothing

Handled:
    If Err Then MsgBox Err.Description
End Sub

'load form populate data and set navbar button

Private Sub Form_Load()


    
    
Dim ctl As Control
Dim rst As ADODB.Recordset



On Error Resume Next

sstSup.TabVisible(1) = False
    'Added by Juan (9/12/2000) for Multilingual
    Call translator.Translate_Forms("frm_IntSupe")
    '------------------------------------------
   'SSDBLine.IsItemInList
    Me.BackColor = frm_Color.txt_WBackground.BackColor

    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl

    sstSup.Tab = 0

    If deIms.rsPHONEDIR.State And adStateOpen Then
        Set rst = deIms.rsPHONEDIR.Clone
    Else
        deIms.PHONEDIR (deIms.NameSpace)
        Set rst = deIms.rsPHONEDIR.Clone
        deIms.rsPHONEDIR.Close
    End If

    DoEvents
    'Call deIms.Supplier(deIms.NameSpace)
    Call deIms.INtSupplier(deIms.NameSpace)
    'Call deIms.INTSUPPLIERLOOKUP(deIms.NameSpace)

'deIms.rsINtSupplier.CancelUpdate
'Commented out by Muz
'     deIms.rsINtSupplier.CancelUpdate

  'by Muz
    'SSOleDBCombo1.DataMemberList = "INtsupplierLookUp"
    'Set SSOleDBCombo1.DataSourceList = deIms
    
    
    DoEvents
    Call BindAll(Me, deIms)
    Set rs = deIms.rsINtSupplier
    Set TxtSuppCode.DataSource = deIms
    Set ssdddContacts.DataSource = rst

    deIms.rsINtSupplier.MoveFirst
    Call DisableButtons(Me, NavBar1)
    
    'Added By Muzammil - 12/18/00
    'Saving the oroginal state of the Navbar
    OrigEdit = NavBar1.EditEnabled
    OrigNew = NavBar1.NewEnabled
    
    'ssdbgContacts.DataMode = ssDataModeBound
    ssdbgContacts.FieldSeparator = Chr(1)

    Imsmail1.NameSpace = deIms.NameSpace
    'Imsmail1.Connected = True 'M
    Imsmail1.SetActiveConnection deIms.cnIms 'M
    Imsmail1.Language = Language 'M

    frm_IntSupe.Caption = frm_IntSupe.Caption + " - " + frm_IntSupe.Tag
'    PrepareImsMail (frm_IntSupe.Imsmail1)
    TxtSuppCode.locked = True   'M
    
    
    
    
    'Added By Muzammil - 12/18/00
    'Reason - To make The Form Operate in Modes
    
    FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
    If FormMode = mdvisualization Then MakeReadOnly (False)
    
    txt_forSEARCH.Enabled = True
txt_forSEARCH.BackColor = "&H00C0E0FF&"



    With frm_IntSupe
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With

End Sub

Private Sub lst_terms_DblClick()
'    Load frm_termedit
End Sub

'before save validate data format and check supplier code exist or not
'if supplier code already exist, show message

Private Sub NavBar1_OnSaveClick()
Screen.MousePointer = 11



On Error Resume Next
Dim Code As String
Dim list As Integer


    'kin add function to check supplier code exit or not
    list = TxtSuppCode
'Reason -  when there was no value in dcbosuppcode the list should be 0 but it's
'value becomes >0,due to which auto numbering is not activated.This statement forces it's Value to 0 in such a case

  NavBar1.SaveEnabled = False

    If list <> 0 And deIms.rsINtSupplier.EditMode = adEditAdd And Len(TxtSuppCode) = 0 Then 'M
          list = 0  'M
     End If         'M
'Added by Juan Gonzalez 2007-7-11
If sstSup.Tab = 4 Then
    ssdbContract.Update
    Call SaveContract
    
     msg1 = translator.Trans("M00255") 'J added
     If Err.number = 0 Then
          MsgBox "Contracts Updated Successfully"
         FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
         NavBar1.EditEnabled = True
         NavBar1.SaveEnabled = False
     End If
Else '----------------------
    If sstSup.Tab = 3 Then
        ssdbgContacts.Update
        Call SaveContacts
        
         msg1 = translator.Trans("M00255") 'J added
         If Err.number = 0 Then
              MsgBox "Contacts Updated Successfully"
             FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
             NavBar1.EditEnabled = True
             NavBar1.SaveEnabled = False
             NavBar1.NewEnabled = True
         End If
         
    Else

       If list = 0 Then
         If ValidateData = True Then

           If Len(Trim$(TxtSuppCode)) <> 0 Then
                Code = Trim(TxtSuppCode)
                
                'If CheckSupplierCode(Code) Then  'M 12/19/00

                    'Modified by Juan (9/11/2000) for Multilingual
                  'COMMECTED out by Muzammil.Since no need to Display MEssage for repetitive Code 12/19/00
                  
                 '   msg1 = translator.Trans("M00254") 'J added
                 '   MsgBox IIf(msg1 = "", "Supplier code exist, Please use a different code.", msg1) 'J modified
                    '---------------------------------------------

                    'Exit Sub 'M 12/19/00
                    
                    'Do While Not CheckSupplierCode(Code)
                    
                    
                'End If 'M 12/19/00

            Else
                    If TxtSuppCode = "" Then GetAutoNumber

            End If   'M
            
            
                       
                      ' deIms.rsINtSupplier!sup_name = Trim$(SSOleDBCombo1.text)
                        
                        deIms.rsINtSupplier!sup_name = Trim$(txt_Name)
                        
                        deIms.rsINtSupplier.Update

                       'Added by juan 2007/9/14
                        Dim sql As String
                        sql = "update SUPPLIER set "
                        If FormMode = mdCreation Then
                            sql = sql + "sup_creauser='" + CurrentUser + "', "
                             deIms.rsINtSupplier!sup_creauser = CurrentUser
                        End If
                        sql = sql + "sup_modiuser='" + CurrentUser + "' "
                        sql = sql + "where sup_code='" + deIms.rsINtSupplier!sup_code + "'"
                        deIms.rsINtSupplier!sup_modiuser = CurrentUser
                        deIms.cnIms.Execute sql + Err.Description
                         '---------------------------------------------

                        'Call deIms.rsINtSupplier.UpdateBatch(adAffectCurrent)
                          deIms.rsINtSupplier.Update
                          'Added by Muz  .2/22/01
                          'rs.CancelBatch
                          'deIms.rsINtSupplier.CancelBatch
                          
                         ' SSDBLine.Refresh
'                          SSDBLine.DataSource = Nothing
'                          SSDBLine.DataMember = "deIms.INtSupplier"
'                          SSDBLine.DataSource = deIms.cnIms
                          
                          
                          
                        Call SaveContacts
                        Call deIms.rsINtSupplier.Move(0)

                        'Modified by Juan (9/11/2000) for Multilingual
                        msg1 = translator.Trans("M00255") 'J added
                        MsgBox IIf(msg1 = "", "Insert into Supplier was completed", msg1) 'J modified
                        '---------------------------------------------
                        
                        ''Set SSOleDBCombo1.DataSourceList = Nothing
                        
                        
                        ''deIms.rsINTSUPPLIERLOOKUP.Close
                        ''Call deIms.INTSUPPLIERLOOKUP(deIms.NameSpace)
                        ''Set SSOleDBCombo1.DataSourceList = deIms
                        
                        
                        'SSOleDBCombo1.Refresh
                         
                         FormMode = ChangeModeOfForm(lblStatus, mdvisualization)
                         If FormMode = mdvisualization Then MakeReadOnly (False)
                         
         Else
             FormMode = ChangeModeOfForm(lblStatus, mdCreation)
             NavBar1.EditEnabled = True
             NavBar1.SaveEnabled = True
         End If

      End If
     End If
End If 'JG 2007-7-11

    If list <> 0 And sstSup.Tab = 0 Then
    
        If Len(Trim$(TxtSuppCode)) <> 0 Then
            Code = Trim(TxtSuppCode)
            If Trim(CheckSupplierCodeexit(Code)) <> Code Then

                'Modified by Juan (9/12/2000) for Multilingual
                msg1 = translator.Trans("M00256") 'J added
                MsgBox IIf(msg1 = "", "You can not change Supplier code, Please enter new code.", msg1) 'J modified
                '---------------------------------------------

                Exit Sub
            Else
                'Added by juan 2007/9/14
                 If FormMode = mdCreation Then
                      deIms.rsINtSupplier!sup_creauser = CurrentUser
                 End If
                deIms.rsINtSupplier!sup_modiuser = CurrentUser
                deIms.rsINtSupplier.Update
                
                Call deIms.rsINtSupplier.UpdateBatch(adAffectCurrent)


                Call SaveContacts
                Call deIms.rsINtSupplier.Move(0)

                'Modified by Juan (9/11/2000) for Multilingual
                msg1 = translator.Trans("M00255") 'J added
                MsgBox "Supplier was Updated Successfully" 'J modified
                '---------------------------------------------

            End If
        End If
        
        Unload Me
    End If
    
    Screen.MousePointer = 0
  'TxtSuppCode.Locked = True
 ' NavBar1.SaveEnabled = True
  


  Dim imsLock As imsLock.Lock
  Set imsLock = New imsLock.Lock
  Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat

End Sub


Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

''    If adReason > adRsnFirstChange Or adReason = adRsnMove Then
''
''        If sstSup.Tab = 0 Then
''
''            If Not ((Rs.EOF) Or (Rs.BOF)) Then
''                If Rs!SUP_FLAG <> 0 Then
''                    'optInternational = True
''                Else
''                    'optLocal = True
''                End If
''
''            End If
''
''        End If
''
''    End If

If NavBar1.SaveEnabled = True And FormMode = 3 Then
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
 Dim imsLock As imsLock.Lock
 Dim ListOfPrimaryControls() As String
 Set imsLock = New imsLock.Lock

  ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)

  Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)   'lock should be here, added by jawdat, 2.1.02

If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
locked = True
End If
End If

End Sub

'check contact exist or not if it exist show message
'and remove the item

Private Sub ssdbgContacts_AfterUpdate(RtnDispErrMsg As Integer)
    If ContactExist Then
        RtnDispErrMsg = False

        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("M00257") 'J added
        MsgBox IIf(msg1 = "", "Contact is already in the list", msg1) 'J modified
        '---------------------------------------------

        Call ssdbgContacts.RemoveItem(ssdbgContacts.row)
    End If

End Sub

'set window size and caption

Private Sub ssdbgContacts_InitColumnProps()

    LockWindowUpdate (HWND)
    'With ssdbgContacts.Columns(0)

        'Modified by Juan (9/12/2000) for Multilingual
    '    msg1 = translator.Trans("L00165") 'J added
    '    .Caption = IIf(msg1 = "", "Contacts", msg1) 'J modified
        '---------------------------------------------

    '    .Name = "Contacts"
    '    .CaptionAlignment = 2
    '    .DataField = "sct_contcode"
    '    .DataType = 8
    '    .FieldLen = 10
    '    .HeadStyleSet = "ColHeader"
    '    .StyleSet = "RowFont"
    '    .Width = 6400
    '    .DropDownHwnd = ssdddContacts.HWND
    'End With
    
    ssdbgContacts.Columns(0).DataField = "sct_contcode"
    ssdbgContacts.Columns(1).DataField = "sct_name"
    ssdbgContacts.Columns(1).FieldLen = 25
    ssdbgContacts.Columns(2).DataField = "sct_tel"
    ssdbgContacts.Columns(2).FieldLen = 25
    ssdbgContacts.Columns(3).DataField = "sct_fax"
    ssdbgContacts.Columns(3).FieldLen = 25
    ssdbgContacts.Columns(4).DataField = "sct_email"
    ssdbgContacts.Columns(4).FieldLen = 255

    ssdbgContacts.Refresh
    ssdbgContacts.MoveFirst
    LockWindowUpdateOff
End Sub

'set size and caption for combo data grid

Private Sub ssdddContacts_InitColumnProps()
    ssdddContacts.Columns.RemoveAll

    Call ssdddContacts.Columns.Add(0)
    Call ssdddContacts.Columns.Add(1)
    Call ssdddContacts.Columns.Add(2)
    Call ssdddContacts.Columns.Add(3)

    ssdddContacts.Columns(0).Width = 3200

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    ssdddContacts.Columns(0).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    '---------------------------------------------

    ssdddContacts.Columns(0).Name = "Name"
    ssdddContacts.Columns(0).DataField = "phd_name"

    ssdddContacts.Columns(0).FieldLen = 256
    ssdddContacts.Columns(1).Width = 1500

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00004") 'J added
    ssdddContacts.Columns(1).Caption = IIf(msg1 = "", "City", msg1) 'J modified
    '---------------------------------------------

    ssdddContacts.Columns(1).Name = "City"
    ssdddContacts.Columns(1).DataField = "phd_city"

    ssdddContacts.Columns(1).FieldLen = 256
    ssdddContacts.Columns(2).Width = 1500

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00130") 'J added
    ssdddContacts.Columns(2).Caption = "Phone Number" 'J modified
    '---------------------------------------------

    ssdddContacts.Columns(2).Name = "PhoneNumber"
    ssdddContacts.Columns(2).DataField = "phd_phonnumb"
    ssdddContacts.Columns(2).FieldLen = 256

    ssdddContacts.Columns(3).Width = 5292
    ssdddContacts.Columns(3).Visible = 0    'False

    'Modified by Juan (9/12/2000) for Multilingual
    msg1 = translator.Trans("L00028") 'J added
    ssdddContacts.Columns(3).Caption = IIf(msg1 = "", "Code", msg1) 'J modified
    '---------------------------------------------

    ssdddContacts.Columns(3).Name = "Code"
    ssdddContacts.Columns(3).DataField = "phd_code"
    ssdddContacts.Columns(3).FieldLen = 256

    ssdddContacts.ForeColorEven = 8388608
    ssdddContacts.BackColorEven = 16771818
    ssdddContacts.BackColorOdd = 16777215

    ssdddContacts.DataFieldList = "phd_code"
    ssdddContacts.DataFieldToDisplay = "phd_name"
    'ssdbgContacts.Columns(0).DropDownHwnd = ssdddContacts.HWND
End Sub

'depend tab status to set navbar button

Private Sub sstSup_Click(PreviousTab As Integer)


Dim rst As ADODB.Recordset
Dim iEditMode(1) As Integer
Dim x As Integer
    NavBar1.DeleteVisible = False
    NavBar1.DeleteEnabled = True
    
    Select Case PreviousTab
    
    Case 0
      If xx = 2 Then Exit Sub
      xx = 1
        If FormMode = mdModification Or FormMode = mdCreation Then
         
            MsgBox "Please save the Supplier before moving to any other tab."
            sstSup.Tab = 0
        End If
    Case 3
     If xx = 1 Then Exit Sub
        xx = 2
      If FormMode = mdModification Then
         
            MsgBox " Please save the contacts before moving to any other tab."
            sstSup.Tab = 3
      End If
    End Select
    

    Select Case sstSup.Tab

        Case 0, 1
            xx = 0
        Case 3
            xx = 0
            
          If Not FormMode = mdModification Then
            'If ssdbgContacts.Columns(0).TagVariant <> TxtSuppCode Then  'commented out by JCG 2008/1/13
                '------- added by JCG 2008/1/13
                NavBar1.DeleteVisible = True
                NavBar1.DeleteEnabled = True
                '--------
                
                'NavBar1.DeleteVisible = False  'commented out by JCG 2008/1/13
                'NavBar1.DeleteEnabled = False  'commented out by JCG 2008/1/13
                
                NavBar1.Width = 1
                Call AddContacts(deIms.SupplierContacts(TxtSuppCode))
            'End If   'commented out by JCG 2008/1/13
          End If
        Case 4 'Added by Juan Gonzalez 2007-7-8
            xx = 0
            If Not FormMode = mdModification Then
              If ssdbContract.Columns(0).TagVariant <> TxtSuppCode Then
                  NavBar1.DeleteVisible = False
                  NavBar1.DeleteEnabled = False
                  NavBar1.Width = 1
                  Call AddContracts
              End If
            End If
            '-----------------------
    End Select
    NavBar1.Enabled = True
End Sub

'set back ground color

Private Sub txt_Address1_GotFocus()
    Call HighlightBackground(txt_Address1)
End Sub

'set back ground color

Private Sub txt_Address1_LostFocus()
    Call NormalBackground(txt_Address1)
End Sub

'set back ground color

Private Sub txt_Address2_GotFocus()
    Call HighlightBackground(txt_Address2)
End Sub

'set back ground color

Private Sub txt_Address2_LostFocus()
    Call NormalBackground(txt_Address2)
End Sub

'set back ground color

Private Sub txt_City_GotFocus()
    Call HighlightBackground(txt_City)
End Sub

'set back ground color

Private Sub txt_City_LostFocus()
    Call NormalBackground(txt_City)
End Sub

'set back ground color

Private Sub txt_Country_GotFocus()
    Call HighlightBackground(txt_Country)
End Sub

'set back ground color

Private Sub txt_Country_LostFocus()
    Call NormalBackground(txt_Country)
End Sub

'set back ground color

Private Sub txt_Email_GotFocus()
    Call HighlightBackground(txt_Email)
End Sub

'set back ground color

Private Sub txt_Email_LostFocus()
    Call NormalBackground(txt_Email)
End Sub

'set back ground color

Private Sub txt_FaxNumber_GotFocus()
    Call HighlightBackground(txt_FaxNumber)
End Sub

'set back ground color

Private Sub txt_FaxNumber_LostFocus()
    Call NormalBackground(txt_FaxNumber)
End Sub

'set back ground color

Private Sub txt_PhoneNumber_GotFocus()
    Call HighlightBackground(txt_PhoneNumber)
End Sub

'set back ground color

Private Sub txt_PhoneNumber_LostFocus()
    Call NormalBackground(txt_PhoneNumber)
End Sub

'set back ground color

Private Sub txt_State_GotFocus()
    Call HighlightBackground(txt_State)
End Sub

'set back ground color

Private Sub txt_State_LostFocus()
    Call NormalBackground(txt_State)
End Sub

'set back ground color

'Private Sub txt_SupName_GotFocus()
'    Call HighlightBackground(txt_SupName)
'End Sub

'set back ground color

'Private Sub txt_SupName_LostFocus()
    'Call NormalBackground(txt_SupName)
'End Sub

'set back ground color

Private Sub txt_Zipcode_GotFocus()
    Call HighlightBackground(txt_Zipcode)
End Sub

'set back ground color

Private Sub txt_Zipcode_LostFocus()
    Call NormalBackground(txt_Zipcode)
End Sub

'check supplier code exist or not

Private Function FindSup(SupCode As String) As Boolean
On Error Resume Next
Dim sCriteria As String, BK As Variant

    With deIms.rsINtSupplier

        .CancelUpdate
        Call .CancelBatch(adAffectCurrent)
        If .EditMode = adEditAdd Then Exit Function

        .Bookmark = BK
        sCriteria = "sup_code = '" & SupCode & "'"
        Call .Find(sCriteria, 0, adSearchForward, adBookmarkFirst)

        Call .Resync(adAffectCurrent, adResyncAllValues)
        If Not .EOF Then FindSup = True: Exit Function

        .Bookmark = BK
    End With

End Function

'get crystal report parameter and application path
'to print report

'''Private Sub BeforePrint()
'''    With MDI_IMS.CrystalReport1
'''
'''        If Not (optInternational) Then
'''            .ReportFileName = FixDir(App.Path) + "CRreports\locsupp.rpt"
'''
'''            'Modified by Juan (10/5/2000) for Multilingual
'''            Call translator.Translate_Reports("locsupp.rpt") 'J added
'''            Call translator.Translate_SubReports 'J added
'''            '---------------------------------------------
'''
'''        Else
'''            .ReportFileName = FixDir(App.Path) + "CRreports\intsupp.rpt"
'''
'''            'Modified by Juan (10/5/2000) for Multilingual
'''            Call translator.Translate_Reports("intsupp.rpt") 'J added
'''            Call translator.Translate_SubReports 'J added
'''            '---------------------------------------------
'''
'''        End If
'''
'''        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'''    End With
'''End Sub

'load data to contact data grid

Private Sub AddContacts(rst As ADODB.Recordset)
On Error Resume Next

    ssdbgContacts.RemoveAll
    If rst Is Nothing Then Exit Sub
    If rst.EOF And rst.BOF Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub

    Do While Not rst.EOF
        'ssdbgContacts.AddItem ((rst!sct_contcode & "") & Chr(1) & (rst!sct_npecode & "") & Chr(1) & (rst!sct_supcode & ""))
        ssdbgContacts.AddItem ((rst!sct_contcode & "") & Chr(1) & (rst!sct_name & "") & Chr(1) & (rst!sct_tel & "") & Chr(1) & (rst!sct_fax & "") & Chr(1) & (rst!sct_email & "")) 'JCG 2008/01/13
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    ssdbgContacts.Columns(0).TagVariant = TxtSuppCode

End Sub

Private Sub AddContracts()
'On Error Resume Next
    ssdbContract.RemoveAll
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT * FROM SUPPLIERCONTRACT WHERE scrt_supcode='" + TxtSuppCode + "' AND scrt_npecode='" + deIms.NameSpace + "'"
    rs.Source = sql
    rs.ActiveConnection = deIms.cnIms
    rs.Open

    If rs Is Nothing Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    Do While Not rs.EOF
        ssdbContract.FieldSeparator = Chr(1)
        ssdbContract.AddItem ((rs!scrt_contractnum & "") & Chr(1) & (rs!scrt_startdate & "") & Chr(1) & (rs!scrt_stopdate & ""))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ssdbContract.Columns(0).TagVariant = TxtSuppCode
End Sub

Public Sub SetFocusOnDatesColumns(column As Integer)
If FormMode <> mdvisualization Then 'JCG 2007/01/12
    If column = 1 Then
        MonthView1.value = Now
       MonthView1.Top = 1200
       MonthView1.Left = ssdbContract.Columns(1).Left - 300
       
       MonthView1.Visible = True
       MonthView2.Visible = False
       MonthView1.SetFocus
    ElseIf column = 2 Then
        MonthView2.value = Now
       MonthView2.Top = 1200
       MonthView2.Left = ssdbContract.Columns(2).Left - 300
       
       MonthView1.Visible = False
       MonthView2.Visible = True
       MonthView2.SetFocus
    End If
End If

End Sub
'check contact exist or not
Private Function ContactExist() As Boolean
On Error Resume Next
Dim i As Integer, x As Integer
Dim Contact As String, y As Integer

    x = ssdbgContacts.Rows - 1
    If x < 0 Then Exit Function
    Contact = ssdbgContacts.Columns(0).value
    ssdbgContacts.MoveFirst

    For i = 0 To x
        If ssdbgContacts.Columns(0).value = Contact Then y = y + 1
        If y > 1 Then Exit For
        ssdbgContacts.MoveNext
    Next

    ContactExist = y > 1
End Function

'save contacts values
Private Function SaveContacts() As Boolean
Dim x As Integer, y As Integer
Dim SupCode As String, np As String
Dim cmd As ADODB.Command
On Error Resume Next

    If Not FindSupplier(rs!sup_code) Then _
        SaveContacts = False: Exit Function

    x = ssdbgContacts.Rows - 1
    If x < 0 Then Exit Function

    np = deIms.NameSpace
    ssdbgContacts.MoveFirst
    Set cmd = New ADODB.Command

    SupCode = rs!sup_code & ""

    SupCode = Trim$(SupCode)

    'If SupCode = "" Then Stop 'Hidden by Juan

    With cmd
        .Prepared = False
        .CommandType = adCmdText
       .ActiveConnection = deIms.cnIms
         Call BeginTransaction(deIms.cnIms)

        .CommandText = "Delete from suppliercontact where sct_supcode = ? and sct_npecode = ?"

        Call .Execute(0, Array(SupCode, np), adExecuteNoRecords)

        .Prepared = True
        Call CommitTransaction(deIms.cnIms)


        Call BeginTransaction(deIms.cnIms)

        '.CommandText = "Insert into SUPPLIERCONTACT(sct_npecode, sct_supcode, sct_contcode)"   Commented out by JCG 2008/01/13
        '.CommandText = .CommandText & "VALUES(?,?,?)"    Commented out by JCG 2008/01/13

        '----- JCG 2008/01/13
        .CommandText = "Insert into SUPPLIERCONTACT(sct_npecode, sct_supcode, sct_contcode, sct_name, sct_tel, sct_fax, sct_email)"
        .CommandText = .CommandText & "VALUES(?,?,?,?,?,?,?)"
        '---------

        cmd.parameters.Refresh

        For y = 0 To x
            'Call .Execute(0, Array(np, SupCode, ssdbgContacts.Columns(0).value), adExecuteNoRecords)  Commented out by JCG 2008/01/13
            Call .Execute(0, Array(np, SupCode, Format(y), ssdbgContacts.Columns(1).value, ssdbgContacts.Columns(2).value, ssdbgContacts.Columns(3).value, ssdbgContacts.Columns(4).value), adExecuteNoRecords)
            ssdbgContacts.MoveNext
        Next

        Call CommitTransaction(deIms.cnIms)
    End With

    Set cmd = Nothing
    If Err Then Err.Clear
End Function

'function get auto numbers

Private Sub GetAutoNumber()
Dim i As Integer, x As Integer
Dim str As String

    If Trim$(TxtSuppCode) <> "" Then Exit Sub


    str = VBA.Left$(ToAlphaChars(Trim$(rs!sup_name & "")), 4)
    str = str & VBA.Left$(ToAlphaChars(Trim$(rs!sup_city & "")), 3)

    i = 1

    Do While FindSupplier(str & IIf(i < 10, "0" + CStr(i), CStr(i)))

        i = i + 1
    Loop

    str = str & IIf(i < 10, "0" + CStr(i), CStr(i))

    rs!sup_code = str

End Sub

'search string values

Public Function ToAlphaChars(str As String) As String
Dim s As String
Dim i As Integer, x As Integer

    If Len(str) = 0 Then Exit Function
    str = Replace$(Replace$(Replace$(str, " ", ""), ",", ""), ".", "")

    x = Len(str)

    For i = 1 To x

        s = VBA.Right$(Mid$(str, 1, i), 1)

        If ((UCase(s) >= "A") And (UCase(s) <= "Z")) Then
            ToAlphaChars = ToAlphaChars & s
        Else
            i = i - 1
            str = Replace(str, s, "")
        End If

        If i >= Len(str) Then Exit For
    Next i

End Function

'call store procedure to find supplier name exist or not

Private Function FindSupplier(id As String) As Boolean
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdStoredProc)

    With cmd
        .Prepared = True
        .CommandText = "SupplierExist"
        .parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("ID", adVarChar, adParamInput, 10, Trim$(id))


        Call .Execute(0, , adExecuteNoRecords)
        FindSupplier = .parameters("RT").value
    End With

    Set cmd = Nothing
End Function

'check data format, if wrong data type entered, show message

Private Function ValidateData() As Boolean
On Error Resume Next
Dim msg As String

    ValidateData = False
    'If Len(Trim$(txt_SupName)) = 0 Then
    If Len(Trim$(txt_Name)) = 0 Then
        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00258") 'J added
        MsgBox IIf(msg1 = "", "Supplier name cannot be left empty", msg1) 'J modified
        '---------------------------------------------

        Exit Function

    ElseIf Len(Trim$(txt_City)) = 0 Then

        'Modified by Juan (9/12/20000) for Multilingual
        msg1 = translator.Trans("M00259") 'J added
        MsgBox IIf(msg1 = "", "City can not be left empty.", msg1) 'J modified
        '----------------------------------------------

        Exit Function

    ElseIf Len(Trim$(txt_PhoneNumber)) = 0 Then

        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00260") 'J added
        MsgBox IIf(msg1 = "", "Phone Number MUST have a value.", msg1) 'J modified
        '---------------------------------------------

        Exit Function

    ElseIf Len(Trim$(txt_PhoneNumber)) < 7 Then

        'Modified by Juan (9/12/2000) for Multilingual
        msg1 = translator.Trans("M00261") 'J added
        MsgBox IIf(msg1 = "", "Length of phone number field should be greater than 7.", msg1) 'J modified
        '---------------------------------------------

        Exit Function

    End If

     If Len(Trim$(TxtSuppCode)) <> 0 Then
         msg = LCase(Trim$(rs("sup_code").originalVALUE & ""))

         If Len(msg) Then
            If (LCase(Trim$(TxtSuppCode)) <> (msg)) Then
'            If (LCase(Trim$(rs("sup_code"))) <> (msg)) Then
'
                rs("sup_code") = msg

                'Modified by Juan (9/12/2000) for Multilingual
                msg1 = translator.Trans("M00262") 'J added
                MsgBox IIf(msg1 = "", "Supplier code cannot be changed once it is saved", msg1)
                '---------------------------------------------

                Exit Function
            End If

        End If
    End If

    ValidateData = True
    If Err Then Err.Clear
End Function

'check supplier code exist or not

Public Function CheckSupplierCode(Code As String, Optional Active As Boolean = False) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        .CommandText = "SELECT ? = count(*) "
        .CommandText = .CommandText & " From SUPPLIER "
        .CommandText = .CommandText & " Where sup_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sup_code = '" & Code & "'"

        .parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)

        Set rst = .Execute
        CheckSupplierCode = cmd.parameters("RT")
    End With


    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckSupplierCode", Err.Description, Err.number, True)
End Function


'check supplier code exist or not
'kin add new function check supplier code

Public Function CheckSupplierCodeexit(Code As String) As String
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        .CommandText = "SELECT sup_code "
        .CommandText = .CommandText & " From SUPPLIER "
        .CommandText = .CommandText & " Where sup_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sup_code = '" & Code & "'"

'        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)

        Set rst = .Execute
        CheckSupplierCodeexit = rst!sup_code
    End With


    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckSupplierCodeexit", Err.Description, Err.number, True)
End Function


Public Sub MakeReadOnly(value As Boolean)


Frame1.Enabled = value
Frame2.Enabled = value
Frame3.Enabled = value
txtRemarks.locked = Not value
''SSOleDBCombo1.AllowInput = Value
Imsmail1.Enabled = value
ssdbRecepientList.AllowUpdate = value
ssdbgContacts.AllowUpdate = value
ssdbContract.AllowUpdate = value

cmd_Add.Enabled = value
cmd_Remove.Enabled = value

Select Case FormMode

     Case mdCreation
             NavBar1.EditEnabled = False
             NavBar1.NewEnabled = False
             NavBar1.CancelEnabled = True
             NavBar1.SaveEnabled = True
             
     Case mdModification
            NavBar1.EditEnabled = False
             NavBar1.NewEnabled = False
             NavBar1.CancelEnabled = True
             NavBar1.SaveEnabled = True
             
     Case mdvisualization
             NavBar1.EditEnabled = OrigEdit
             NavBar1.NewEnabled = OrigNew
             NavBar1.CancelEnabled = False
             NavBar1.SaveEnabled = False
             
 End Select
End Sub

'Added by muzammil 12/18/00
'Reason - To Get the Original Values when The User clicks on CANCEL
Public Function GetOriginalValues()

''SSOleDBCombo1.text = deIms.rsINtSupplier("sup_name").OriginalValue

txt_Name = IIf(IsNull(deIms.rsINtSupplier("sup_name").originalVALUE), "", deIms.rsINtSupplier("sup_name").originalVALUE)

txt_Address1 = IIf(IsNull(deIms.rsINtSupplier("sup_adr1").originalVALUE), "", deIms.rsINtSupplier("sup_adr1").originalVALUE)
txt_Email = IIf(IsNull(deIms.rsINtSupplier("sup_mail").originalVALUE), "", deIms.rsINtSupplier("sup_mail").originalVALUE)
txt_State = IIf(IsNull(deIms.rsINtSupplier("sup_stat").originalVALUE), "", deIms.rsINtSupplier("sup_stat").originalVALUE)
txt_Zipcode = IIf(IsNull(deIms.rsINtSupplier("sup_zipc").originalVALUE), "", deIms.rsINtSupplier("sup_zipc").originalVALUE)
txt_Address2 = IIf(IsNull(deIms.rsINtSupplier("sup_adr2").originalVALUE), "", deIms.rsINtSupplier("sup_adr2").originalVALUE)
txt_Country = IIf(IsNull(deIms.rsINtSupplier("sup_ctry").originalVALUE), "", deIms.rsINtSupplier("sup_ctry").originalVALUE)
txt_City = IIf(IsNull(deIms.rsINtSupplier("sup_city").originalVALUE), "", deIms.rsINtSupplier("sup_city").originalVALUE)
txt_FaxNumber = IIf(IsNull(deIms.rsINtSupplier("sup_faxnumb").originalVALUE), "", deIms.rsINtSupplier("sup_faxnumb").originalVALUE)
txt_PhoneNumber = IIf(IsNull(deIms.rsINtSupplier("sup_phonnumb").originalVALUE), "", deIms.rsINtSupplier("sup_phonnumb").originalVALUE)
Txt_contaname = IIf(IsNull(deIms.rsINtSupplier("sup_contaname").originalVALUE), "", deIms.rsINtSupplier("sup_contaname").originalVALUE)
Txt_contaPH = IIf(IsNull(deIms.rsINtSupplier("sup_contaph").originalVALUE), "", deIms.rsINtSupplier("sup_contaph").originalVALUE)
Txt_contaFax = IIf(IsNull(deIms.rsINtSupplier("sup_contafax").originalVALUE), "", deIms.rsINtSupplier("sup_contafax").originalVALUE)
txtRemarks = IIf(IsNull(deIms.rsINtSupplier("sup_remk").originalVALUE), "", deIms.rsINtSupplier("sup_remk").originalVALUE)
End Function



'''''Private Sub SelectGatewayAndSendOutMails()
'''''
'''''If ConnInfo.EmailClient = Outlook Then
'''''
'''''    Call sendOutlookEmailandFax
'''''
'''''ElseIf ConnInfo.EmailClient = ATT Then
'''''
'''''    Call SendAttEmailandFax
'''''
'''''ElseIf ConnInfo.EmailClient = Outlook Then
'''''
'''''    MsgBox "Email is not set up properly. Please Configure the database for Emails.", vbInformation, "Imswin"
'''''
'''''End If
'''''
'''''End Sub

''''Public Sub SendAttEmailandFax()
''''Dim rpinf As RPTIFileInfo
''''Dim Params(1) As String
''''
''''    With rpinf
''''
''''        Params(1) = "suppcode=" & TxtSuppCode
''''        Params(0) = "namespace=" & deIms.NameSpace
''''        .ReportFileName = ReportPath & "Supplier.rpt"
''''
''''        'Added by Juan (10/5/2000) for Multilingual
''''        Call translator.Translate_Reports("Supplier.rpt") 'J added
''''        Call translator.Translate_SubReports 'J added
''''        '------------------------------------------
''''
''''        .Parameters = Params
''''    End With
''''
''''    Params(0) = ""
''''    Call WriteRPTIFile(rpinf, Params(0))
''''    'Modified by Muzammil 08/08/00
''''     PrintCurrent         'M
''''
''''
''''    Call SendEmailAndFax(rsReceptList, "Recipients", "Supplier", "", Params(0))
''''
''''    Set rsReceptList = Nothing
''''    Set ssdbRecepientList.DataSource = Nothing
''''
''''End Sub

'''''Public Function sendOutlookEmailandFax()
'''''Dim Params(1) As String
'''''Dim i As Integer
'''''Dim Attachments() As String
'''''Dim Subject As String
'''''Dim reports(0) As String
'''''Dim Recepients() As String
'''''Dim attention As String
'''''
'''''On Error GoTo errMESSAGE
'''''
'''''     'BeforePrint 'By M on 02/20. GenerateAttachemtn does almost the same thig.
'''''
'''''     If rsReceptList.RecordCount > 0 Then
'''''
'''''        Subject = "Supplier"
'''''        reports(0) = "supplier.rpt"
'''''
'''''        attention = "Attention Please "
'''''
'''''        'Send reports to it and creates the attachments and save them to a perticular FOLDER for AT&T
'''''
'''''        Attachments = generateattachments(reports)
'''''
'''''        Recepients = ToArrayFromRecordset(rsReceptList)
'''''        'Here we create the parameter FILE.
'''''        'Send the attachments ,the subject and the recepients to be written in the Parameter file.
'''''
'''''            Call WriteParameterFiles(Recepients, "", Attachments, Subject, attention)
'''''
'''''    Else
'''''
'''''         MsgBox "No Recipients to Send", , "Imswin"
'''''
'''''    End If
'''''
'''''errMESSAGE:
'''''
'''''    If Err.number <> 0 Then
'''''
'''''        MsgBox Err.Description
'''''
'''''    End If
'''''
'''''End Function

'''''Private Function generateattachments(reports() As String) As String()
'''''  Dim l
'''''  Dim Attachments(0) As String
'''''  Dim IFile As IMSFile
'''''  Dim FileName As String
'''''
'''''  Set IFile = New IMSFile
'''''  'l = UBound(reports)
'''''On Error GoTo errMESSAGE
'''''
''''''  For i = 0 To l
'''''
'''''
'''''    With MDI_IMS.CrystalReport1
'''''
'''''        .ReportFileName = ReportPath & "Supplier.rpt"
'''''
'''''        Call translator.Translate_Reports(reports(l))
'''''        Call translator.Translate_SubReports
'''''
'''''        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'''''        .ParameterFields(1) = "suppcode;" + TxtSuppCode + ";TRUE"
'''''        .ParameterFields(2) = "Intloc;" + "INT" + ";TRUE"
'''''
'''''    End With
'''''
'''''     Attachments(0) = "Report-" & "SUPPLIER" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
'''''
'''''     FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
'''''
'''''    If IFile.FileExists(FileName) Then IFile.DeleteFile (FileName)
'''''
'''''    If Not FileExists(FileName) Then MDI_IMS.SaveReport FileName, crptRTF
'''''
'''''     generateattachments = Attachments
'''''
'''''errMESSAGE:
'''''    If Err.number <> 0 Then
'''''        MsgBox Err.Description
'''''    End If
'''''
'''''End Function

''Private Sub BeforePrint()
''On Error Resume Next
''
''    With MDI_IMS.CrystalReport1
''        .ReportFileName = ReportPath & "supplier.rpt"
''
''        'Modified by Juan (8/28/2000) for Multilingual
''        Call translator.Translate_Reports("supplier.rpt") 'J added
''        Call translator.Translate_SubReports 'J added
''        '---------------------------------------------
''
''        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
''          .ParameterFields(1) = "suppcode;" + TxtSuppCode + ";TRUE"
''    End With
''
''    If Err Then
''        MsgBox Err.Description
''        Call LogErr(Name & "::BeforePrint", Err.Description, Err)
''    End If
''End Sub
''
