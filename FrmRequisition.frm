VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form FrmRequisition 
   BackColor       =   &H8000000A&
   Caption         =   "Requisition Managment"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10620
   Tag             =   "02020400"
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   960
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   16449537
      CurrentDate     =   37316
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDBDDStockNumber 
      Height          =   855
      Left            =   6840
      TabIndex        =   2
      Top             =   6960
      Width           =   3735
      DataFieldList   =   "stk_desc"
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "StockNumber"
      Columns(0).Name =   "StockNumber"
      Columns(0).DataField=   "stk_stcknumb"
      Columns(0).FieldLen=   256
      Columns(1).Width=   6535
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "stk_desc"
      Columns(1).FieldLen=   256
      _ExtentX        =   6588
      _ExtentY        =   1508
      _StockProps     =   77
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      DataFieldToDisplay=   "stk_stcknumb"
   End
   Begin TabDlg.SSTab SSTabRequisitions 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Requisition"
      TabPicture(0)   =   "FrmRequisition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCountPoitem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSDDBuyerDetails"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LROleDBNavBar1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSDDCompany"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSDDLocation"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSDDBuyer"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SSGridSelection"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SSTab1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MonthView2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Recepients"
      TabPicture(1)   =   "FrmRequisition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Recipients"
      Tab(1).Control(1)=   "Lbl_search"
      Tab(1).Control(2)=   "SSOLEDBFax"
      Tab(1).Control(3)=   "SSOLEDBEmail"
      Tab(1).Control(4)=   "dgRecipientList"
      Tab(1).Control(5)=   "fra_FaxSelect"
      Tab(1).Control(6)=   "cmd_Add"
      Tab(1).Control(7)=   "txt_Recipient"
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).Control(9)=   "Txt_search"
      Tab(1).Control(10)=   "OptFax"
      Tab(1).Control(11)=   "OptEmail"
      Tab(1).ControlCount=   12
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   360
         TabIndex        =   7
         Top             =   6120
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   16449537
         CurrentDate     =   37316
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4695
         Left            =   120
         TabIndex        =   9
         Top             =   2480
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Assign Req. To Buyer"
         TabPicture(0)   =   "FrmRequisition.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSOleFilter"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSGridHeaderDetails"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtsearch"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Req. Included in a PO"
         TabPicture(1)   =   "FrmRequisition.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtsearchDetl"
         Tab(1).Control(1)=   "SSGridPOITEMDETAILS"
         Tab(1).ControlCount=   2
         Begin VB.TextBox txtsearchDetl 
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   -74520
            TabIndex        =   27
            Text            =   "Hit Enter To see results"
            ToolTipText     =   "Hit Enter To see results"
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtsearch 
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   480
            TabIndex        =   26
            Text            =   "Hit Enter To see results"
            ToolTipText     =   "Hit Enter To see results"
            Top             =   480
            Width           =   1935
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGridHeaderDetails 
            Height          =   3615
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   10095
            _Version        =   196617
            DataMode        =   2
            Col.Count       =   9
            MultiLine       =   0   'False
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   106
            Columns.Count   =   9
            Columns(0).Width=   2593
            Columns(0).Caption=   "Requisition #"
            Columns(0).Name =   "Requisition #"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   1931
            Columns(1).Caption=   "Date Created"
            Columns(1).Name =   "Date Created"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   2196
            Columns(2).Caption=   "Date Approved"
            Columns(2).Name =   "Date Approved"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(3).Width=   2064
            Columns(3).Caption=   "Date Assigned"
            Columns(3).Name =   "Date Assigned"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3175
            Columns(4).Caption=   "Originator"
            Columns(4).Name =   "Originator"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(4).Locked=   -1  'True
            Columns(5).Width=   1588
            Columns(5).Caption=   "Day Open"
            Columns(5).Name =   "Day Open"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(5).Locked=   -1  'True
            Columns(6).Width=   3043
            Columns(6).Caption=   "Buyer"
            Columns(6).Name =   "Buyer"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "RowUpdated"
            Columns(7).Name =   "RowUpdated"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "BuyerCode"
            Columns(8).Name =   "BuyerCode"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            _ExtentX        =   17806
            _ExtentY        =   6376
            _StockProps     =   79
            ForeColor       =   -2147483630
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGridPOITEMDETAILS 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   11
            Top             =   960
            Width           =   10095
            _Version        =   196617
            DataMode        =   2
            Col.Count       =   7
            UseGroups       =   -1  'True
            BackColorOdd    =   16777215
            RowHeight       =   423
            ExtraHeight     =   106
            Groups.Count    =   3
            Groups(0).Width =   4286
            Groups(0).Caption=   "Requisition"
            Groups(0).Columns.Count=   2
            Groups(0).Columns(0).Width=   2937
            Groups(0).Columns(0).Caption=   "#"
            Groups(0).Columns(0).Name=   "No"
            Groups(0).Columns(0).DataField=   "Column 0"
            Groups(0).Columns(0).DataType=   8
            Groups(0).Columns(0).FieldLen=   256
            Groups(0).Columns(1).Width=   1349
            Groups(0).Columns(1).Caption=   "LI"
            Groups(0).Columns(1).Name=   "LI"
            Groups(0).Columns(1).DataField=   "Column 1"
            Groups(0).Columns(1).DataType=   8
            Groups(0).Columns(1).FieldLen=   256
            Groups(1).Width =   10927
            Groups(1).Caption=   "Purchase order"
            Groups(1).Columns.Count=   4
            Groups(1).Columns(0).Width=   2831
            Groups(1).Columns(0).Caption=   "#"
            Groups(1).Columns(0).Name=   "PONO"
            Groups(1).Columns(0).DataField=   "Column 2"
            Groups(1).Columns(0).DataType=   8
            Groups(1).Columns(0).FieldLen=   256
            Groups(1).Columns(1).Width=   2805
            Groups(1).Columns(1).Caption=   "LI"
            Groups(1).Columns(1).Name=   "POLI"
            Groups(1).Columns(1).DataField=   "Column 3"
            Groups(1).Columns(1).DataType=   8
            Groups(1).Columns(1).FieldLen=   256
            Groups(1).Columns(2).Width=   2963
            Groups(1).Columns(2).Caption=   "Total Value"
            Groups(1).Columns(2).Name=   "POTotalValue"
            Groups(1).Columns(2).DataField=   "Column 4"
            Groups(1).Columns(2).DataType=   8
            Groups(1).Columns(2).FieldLen=   256
            Groups(1).Columns(3).Width=   2328
            Groups(1).Columns(3).Caption=   "Creation Date"
            Groups(1).Columns(3).Name=   "Creation Date"
            Groups(1).Columns(3).DataField=   "Column 5"
            Groups(1).Columns(3).DataType=   8
            Groups(1).Columns(3).FieldLen=   256
            Groups(2).Width =   1482
            Groups(2).Caption=   "Days"
            Groups(2).CaptionAlignment=   0
            Groups(2).Columns(0).Width=   1482
            Groups(2).Columns(0).Caption=   "Elapsed"
            Groups(2).Columns(0).Name=   "Days Elapsed"
            Groups(2).Columns(0).DataField=   "Column 6"
            Groups(2).Columns(0).DataType=   8
            Groups(2).Columns(0).FieldLen=   256
            _ExtentX        =   17806
            _ExtentY        =   6376
            _StockProps     =   79
            ForeColor       =   0
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleFilter 
            Height          =   375
            Left            =   7080
            TabIndex        =   30
            Top             =   480
            Width           =   3135
            DataFieldList   =   "column 0"
            AllowInput      =   0   'False
            _Version        =   196617
            ColumnHeaders   =   0   'False
            ForeColorEven   =   0
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   93
            ForeColor       =   0
            BackColor       =   -2147483643
            DataFieldToDisplay=   "column 0"
         End
      End
      Begin VB.OptionButton OptEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -71805
         TabIndex        =   20
         Top             =   3030
         Width           =   735
      End
      Begin VB.OptionButton OptFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   -72645
         TabIndex        =   19
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox Txt_search 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   -72645
         MaxLength       =   60
         TabIndex        =   18
         Top             =   3750
         Width           =   3855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74445
         TabIndex        =   17
         Top             =   1230
         Width           =   1335
      End
      Begin VB.TextBox txt_Recipient 
         Height          =   288
         Left            =   -72645
         MaxLength       =   60
         TabIndex        =   16
         Top             =   3390
         Width           =   6150
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74445
         TabIndex        =   15
         Top             =   3390
         Width           =   1335
      End
      Begin VB.Frame fra_FaxSelect 
         Height          =   1170
         Left            =   -74505
         TabIndex        =   12
         Top             =   4035
         Width           =   1635
         Begin VB.OptionButton opt_FaxNum 
            Caption         =   "Fax Numbers"
            Height          =   330
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1275
         End
         Begin VB.OptionButton opt_Email 
            Caption         =   "Email"
            Height          =   288
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   795
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGridSelection 
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   10095
         ScrollBars      =   1
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   7
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   7
         Columns(0).Width=   2699
         Columns(0).Caption=   "Company"
         Columns(0).Name =   "Company"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2831
         Columns(1).Caption=   " Location"
         Columns(1).Name =   " Location"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2858
         Columns(2).Caption=   "Stock/Folio # Y/N"
         Columns(2).Name =   "Stock/Folio # Y/N"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Buyer"
         Columns(3).Name =   "Buyer"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1799
         Columns(4).Caption=   "Days Open"
         Columns(4).Name =   "Days Open"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1773
         Columns(5).Caption=   "From"
         Columns(5).Name =   "From"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1614
         Columns(6).Caption=   "To"
         Columns(6).Name =   "To"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         _ExtentX        =   17806
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Selection Criteria"
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDBuyer 
         Height          =   855
         Left            =   6600
         TabIndex        =   3
         Top             =   4680
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Buyer Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDLocation 
         Height          =   855
         Left            =   6840
         TabIndex        =   4
         Top             =   6840
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         GroupHeaders    =   0   'False
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Location Code"
         Columns(0).Name =   "Location Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDCompany 
         Height          =   855
         Left            =   5040
         TabIndex        =   5
         Top             =   4920
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "CompanyCode"
         Columns(0).Name =   "CompanyCode"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4260
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dgRecipientList 
         Height          =   2085
         Left            =   -72525
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   690
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
         stylesets(0).Picture=   "FrmRequisition.frx":0070
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
         stylesets(1).Picture=   "FrmRequisition.frx":008C
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
         Columns(0).Caption=   "Column 0"
         Columns(0).Name =   "Column 0"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10610
         _ExtentY        =   3678
         _StockProps     =   79
         ForeColor       =   -2147483630
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBEmail 
         Height          =   2055
         Left            =   -72600
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
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
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOLEDBFax 
         Height          =   2055
         Left            =   -72600
         TabIndex        =   25
         Top             =   4200
         Visible         =   0   'False
         Width           =   6195
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Name"
         Columns(0).Name =   "Name"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
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
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin LRNavigators.LROleDBNavBar LROleDBNavBar1 
         Height          =   375
         Left            =   3360
         TabIndex        =   28
         Top             =   7200
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDBuyerDetails 
         Height          =   855
         Left            =   6600
         TabIndex        =   29
         Top             =   5880
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "CompanyCode"
         Columns(0).Name =   "CompanyCode"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4260
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         DataFieldToDisplay=   "Column 1"
      End
      Begin VB.Label LblCountPoitem 
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
         Left            =   360
         TabIndex        =   31
         Top             =   2175
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Lbl_search 
         Caption         =   "Search by name"
         Height          =   255
         Left            =   -74325
         TabIndex        =   24
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74520
         TabIndex        =   23
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   8
         Top             =   2175
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SelectionCodes

    CompanyCode As String
    LocationCode As String
    StockNumber As String
    Buyer As String
    Fromdate As String
    Todate As String
    OpenFor As String

End Type

Private Type Filter

    AllReqs As String
    AssignedReqs As String
    UnAssignedreqs As String

End Type

Dim GGridFilledWithEmails As Boolean
Dim GGridFilledWithFax As Boolean
Dim GselectionCode As SelectionCodes
Dim GOldValue As String
Dim GValueChanged As Boolean
Dim GKeySearch As String
Dim rowguid, locked As Boolean       'jawdat
Dim mFilter As Filter
Dim GHeaderGridFilled As Boolean
Dim GDetailsGridFilled As Boolean
Dim GLocationPoDetails() As String
Private Sub Label4_Click()

End Sub



Private Sub Form_Load()

Dim RsDetails As ADODB.Recordset
Dim AggregateFont As StdFont

Set AggregateFont = New StdFont

AggregateFont.Size = 8.25
AggregateFont.Bold = True
AggregateFont.Name = "MS Sans Serif"


Me.Height = 8250
Me.Width = 10740

Load FrmShowApproving
Screen.MousePointer = 11
FrmShowApproving.Top = 4620
FrmShowApproving.Left = 3330
FrmShowApproving.Width = 3000
FrmShowApproving.Height = 1140

FrmShowApproving.Show
FrmShowApproving.Label2.Caption = " Loading Requisition ... "
DoEvents
DoEvents
Call PopulateBuyers
Call PopulateCompany
Call PopulateStockNumber
'Call PopulatebuyerForDetails

SSGridSelection.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9)

LROleDBNavBar1.NewEnabled = False
LROleDBNavBar1.CancelVisible = False
LROleDBNavBar1.DeleteVisible = False
LROleDBNavBar1.CancelLastSepVisible = False
LROleDBNavBar1.PreviousVisible = False
LROleDBNavBar1.NewVisible = False
LROleDBNavBar1.NextVisible = False
LROleDBNavBar1.LastVisible = False
LROleDBNavBar1.FirstVisible = False



SSGridHeaderDetails.StyleSets.Add ("RowBeingModified")
SSGridHeaderDetails.StyleSets("RowBeingModified").BackColor = vbYellow

SSGridHeaderDetails.StyleSets.Add ("CellBeingModified")
SSGridHeaderDetails.StyleSets("CellBeingModified").BackColor = &H80C0FF

SSGridHeaderDetails.StyleSets.Add ("RowModified")
SSGridHeaderDetails.StyleSets("RowModified").BackColor = &HFFFFC0

SSGridPOITEMDETAILS.StyleSets.Add ("BigFont")
Set SSGridPOITEMDETAILS.StyleSets("BigFont").Font = AggregateFont

'SSGridPOITEMDETAILS.Groups(0).HeadStyleSet = "BigFont"
'SSGridPOITEMDETAILS.Groups(1).HeadStyleSet = "BigFont"
'SSGridPOITEMDETAILS.Groups(2).HeadStyleSet = "BigFont"

SSGridPOITEMDETAILS.StyleSets.Add ("AggregateLine")
SSGridPOITEMDETAILS.StyleSets("AggregateLine").BackColor = &HC0FFC0   '&HE0E0E0 --> Gray Color
SSGridPOITEMDETAILS.StyleSets("AggregateLine").ForeColor = vbBlack
Set SSGridPOITEMDETAILS.StyleSets("AggregateLine").Font = AggregateFont

SSGridPOITEMDETAILS.StyleSets.Add ("RowBeingModified")
SSGridPOITEMDETAILS.StyleSets("RowBeingModified").BackColor = vbYellow
SSGridPOITEMDETAILS.StyleSets("RowBeingModified").ForeColor = vbBlack

SSGridPOITEMDETAILS.StyleSets.Add ("SumColumn")
SSGridPOITEMDETAILS.StyleSets("SumColumn").BackColor = &HC0E0FF 'Light Pink
SSGridPOITEMDETAILS.StyleSets("SumColumn").ForeColor = vbBlack
Set SSGridPOITEMDETAILS.StyleSets("SumColumn").Font = AggregateFont

SSGridHeaderDetails.ActiveRowStyleSet = "RowBeingModified"
SSGridHeaderDetails.ActiveCell.StyleSet = "CellBeingModified"

SSGridPOITEMDETAILS.ActiveRowStyleSet = "RowBeingModified"

Call DisableButtons(Me, LROleDBNavBar1)

SSGridSelection.AllowUpdate = True
txtsearch.Enabled = True
txtsearchDetl.Enabled = True

SSOleFilter.DataMode = ssDataModeAddItem
SSOleFilter.Text = "Show Un-Assigned Requisitions only"
SSOleFilter.AddItem "Show All Requisitions"
SSOleFilter.AddItem "Show Assigned Requisitions only"
SSOleFilter.AddItem "Show Un-Assigned Requisitions only"

mFilter.AllReqs = "Show All Requisitions"
mFilter.AssignedReqs = "Show Assigned Requisitions only"
mFilter.UnAssignedreqs = "Show Un-Assigned Requisitions only"


SSOleFilter.Columns.Add 0

SSOleFilter.Columns(0).Width = SSOleFilter.Width


'AllReqs As String  "Show All Requisitions"
'AssignedReqs AS String  "Show Assigned Requisitions only"
'UnAssignedreqs AS STRING  "Show Un-Assigned Requisitions only"


Dim I As Integer
For I = 0 To SSGridSelection.Cols - 1
            SSGridSelection.Columns(I).BackColor = &HFFFFC0
Next I

MonthView1.ToolTipText = "From Date"
MonthView2.ToolTipText = "To Date"

Unload FrmShowApproving



End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Dim X As SelectionCodes
     
         
     GGridFilledWithEmails = False
     GGridFilledWithFax = False
    GselectionCode = X
    GOldValue = ""
    GValueChanged = False
    GHeaderGridFilled = False
    GDetailsGridFilled = False
       Dim imsLock As imsLock.Lock
       Set imsLock = New imsLock.Lock
       Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
                    
    
     If open_forms <= 5 Then ShowNavigator
     

End Sub

Private Sub LROleDBNavBar1_BeforeSaveClick()
Call SaveRecord
End Sub

Private Sub LROleDBNavBar1_OnCloseClick()


'                    Dim imsLock As imsLock.Lock
'                    Set imsLock = New imsLock.Lock
'                    Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
                    
Unload Me
End Sub

Private Sub MonthView1_LostFocus()

MonthView1.Visible = False

End Sub

Private Sub MonthView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim dates As Date
 Dim s
s = MonthView1.HitTest(X, Y, dates)

If s = 1 Or s = 2 Or s = 3 Then

    MonthView1.value = dates

    SSGridSelection.Columns(5).Text = dates

    GselectionCode.Fromdate = CStr(dates)

    MonthView1.Visible = False

    Call GetDataForTheSelection

End If

End Sub

Private Sub MonthView2_LostFocus()

MonthView1.Visible = False

End Sub

Private Sub MonthView2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dates As Date
Dim s

s = MonthView2.HitTest(X, Y, dates)

If s = 1 Or s = 2 Or s = 3 Then

    MonthView2.value = dates

    SSGridSelection.Columns(6).Text = dates

    GselectionCode.Todate = CStr(dates)

    MonthView2.Visible = False

    Call GetDataForTheSelection

End If

End Sub

Private Sub SSDDBuyerDetails_DropDown()

If IsPoAlreadyInclude = True Then

    SSGridHeaderDetails.Columns(5).locked = True

    SSDDBuyerDetails.DroppedDown = False

    MsgBox "Can not modify the Requisition , it has already been assigned.", vbInformation, "Imswin"

 Else

    SSDDBuyerDetails.DroppedDown = True

End If

End Sub

Private Sub SSGridHeaderDetails_AfterUpdate(RtnDispErrMsg As Integer)
Dim I As Integer
Call SaveRecord

For I = 0 To SSGridHeaderDetails.Cols - 1
            SSGridHeaderDetails.Columns(I).CellStyleSet "RowModified", SSGridHeaderDetails.row
Next I

End Sub

Private Sub SSGridHeaderDetails_Change()
Dim I As Integer



 SSGridHeaderDetails.Columns(7).value = 1
 SSGridHeaderDetails.Columns(8).value = SSDDBuyerDetails.Columns(0).Text

LROleDBNavBar1.SaveEnabled = True

GOldValue = SSGridHeaderDetails.Columns(6).CellText(SSGridHeaderDetails.Bookmark)

GValueChanged = True

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

End Sub


Private Sub SSGridHeaderDetails_InitColumnProps()
SSGridHeaderDetails.Columns(6).DropDownHwnd = SSDDBuyerDetails.HWND
End Sub

Private Sub SSGridHeaderDetails_KeyPress(KeyAscii As Integer)

Dim Count As Integer

Select Case SSGridHeaderDetails.Col

Case 0

Case 5

    SSGridHeaderDetails.DroppedDown = True

End Select



End Sub

Private Sub SSGridPOITEMDETAILS_RowLoaded(ByVal Bookmark As Variant)
    Dim I As Integer
   
   If Trim(SSGridPOITEMDETAILS.Columns(0).Text) <> "'" And Len(Trim(SSGridPOITEMDETAILS.Columns(0).Text)) > 0 Then
       
        For I = 0 To SSGridPOITEMDETAILS.Cols - 1
        
            SSGridPOITEMDETAILS.Columns(I).CellStyleSet "AggregateLine"
            
        Next I
    
   End If

                    
    If Trim(SSGridPOITEMDETAILS.Columns(3).Text) = "Total Value" Then
       
        'For i = 0 To SSGridPOITEMDETAILS.Cols - 1
        
            SSGridPOITEMDETAILS.Columns(3).CellStyleSet "SumColumn"
            SSGridPOITEMDETAILS.Columns(4).CellStyleSet "SumColumn"
            
        'Next i
    
   End If
   
End Sub

Private Sub SSGridSelection_BeforeRowColChange(Cancel As Integer)

 Select Case SSGridSelection.Col

 Case 0

    If Len(Trim(SSGridSelection.Columns(0).value)) > 0 And SSGridSelection.IsItemInList = False Then

        MsgBox " Invalid Company. Please select a value from the drop downs.", vbInformation, "Imswin"

        Cancel = True

    End If

 Case 1

    If Len(Trim(SSGridSelection.Columns(1).value)) > 0 And SSGridSelection.IsItemInList = False Then

        MsgBox " Invalid Location. Please select a value from the drop downs.", vbInformation, "Imswin"

        Cancel = True

    End If

 Case 2

    If Len(Trim(SSGridSelection.Columns(2).value)) > 0 And SSGridSelection.IsItemInList = False Then

        MsgBox " Invalid Stocknumber. Please select a value from the drop downs.", vbInformation, "Imswin"

        Cancel = True

    End If

 Case 3

        If Len(Trim(SSGridSelection.Columns(3).value)) > 0 And SSGridSelection.IsItemInList = False Then

        MsgBox " Invalid Buyer. Please select a value from the drop downs.", vbInformation, "Imswin"

        Cancel = True

    End If

 Case 4

''       If IsNumeric(SSGridSelection.Columns(4).value) = False Then
''
''        MsgBox " Invalid Days Open. Please enter a valid value.", vbInformation, "Imswin"
''
''        Cancel = True
''
''       End If

 Case 5

       If IsDate(SSGridSelection.Columns(5).value) = False And Len(Trim(SSGridSelection.Columns(5).value)) > 0 Then

        MsgBox " Invalid From Date. Please enter a valid date.", vbInformation, "Imswin"

        Cancel = True

    End If

 Case 6

       If IsDate(SSGridSelection.Columns(6).value) = False And Len(Trim(SSGridSelection.Columns(6).value)) > 0 Then

        MsgBox " Invalid To date. Please enter a valid date.", vbInformation, "Imswin"

        Cancel = True

    End If


 End Select

End Sub

Private Sub SSGridSelection_Change()

Select Case SSGridSelection.Col


Case 0

    GselectionCode.CompanyCode = Trim(SSDDCompany.Columns(0).Text)

Case 1

    GselectionCode.LocationCode = Trim(SSDDLocation.Columns(0).Text)

Case 2

     GselectionCode.StockNumber = Trim(SSDBDDStockNumber.Columns(0).Text)

Case 3

   GselectionCode.Buyer = Trim(SSDDBuyer.Columns(0).Text)

Case 4


    GselectionCode.OpenFor = Trim(SSGridSelection.Columns(4).Text)


Case 5

Case 6

End Select

GHeaderGridFilled = False
GDetailsGridFilled = False

Call GetDataForTheSelection


End Sub

Private Sub SSGridSelection_Click()

Call LostFocusOnDatesColumns(1)

If SSGridSelection.Col = 1 Then

 Call PopulateLocation(Trim(GselectionCode.CompanyCode))

ElseIf SSGridSelection.Col = 5 Then

   Call SetFocusOnDatesColumns(5)

ElseIf SSGridSelection.Col = 6 Then

   Call SetFocusOnDatesColumns(6)

End If

End Sub

Private Sub SSGridSelection_InitColumnProps()

SSGridSelection.Columns(0).DropDownHwnd = SSDDCompany.HWND

SSGridSelection.Columns(1).DropDownHwnd = SSDDLocation.HWND

SSGridSelection.Columns(2).DropDownHwnd = SSDBDDStockNumber.HWND

SSGridSelection.Columns(3).DropDownHwnd = SSDDBuyer.HWND

End Sub

Public Function GetDataForTheSelection() 'As ADODB.Recordset

Screen.MousePointer = vbHourglass

If SSTab1.Tab = 0 Then

    GetDataForPOHeaderTab
    
    Label1.Visible = True
    LblCountPoitem.Visible = False
    
    
ElseIf SSTab1.Tab = 1 Then
   
   Label1.Visible = False
   LblCountPoitem.Visible = True
    
   
   GetDataForPODetailsTab

End If

Screen.MousePointer = vbArrow

End Function

Public Function PopulateBuyers()
Dim Rsbuyer As New ADODB.Recordset

Rsbuyer.source = "select buy_username,usr_username , buy_npecode from buyer,xuserprofile where buy_username = usr_userid and buy_npecode = usr_npecode and usr_npecode ='" & deIms.NameSpace & "' and usr_stas = 'A'"

Rsbuyer.ActiveConnection = deIms.cnIms

Rsbuyer.Open

SSDDBuyer.RemoveAll

SSDDBuyer.AddItem "ALL" & Chr(9) & "ALL"

  Do While Not Rsbuyer.EOF

    SSDDBuyer.AddItem Rsbuyer("buy_username") & Chr(9) & Rsbuyer("usr_username")
    
   SSDDBuyerDetails.AddItem Rsbuyer("buy_username") & Chr(9) & Rsbuyer("usr_username")
    
    Rsbuyer.MoveNext

   Loop

 Rsbuyer.Close

 Set Rsbuyer = Nothing


End Function

Public Function PopulateCompany()

Dim rsCOMPANY As New ADODB.Recordset

rsCOMPANY.source = "select com_compcode,  com_name  from company where  com_npecode  ='" & deIms.NameSpace & "'"

rsCOMPANY.ActiveConnection = deIms.cnIms

rsCOMPANY.Open

SSDDCompany.RemoveAll

SSDDCompany.AddItem "ALL" & Chr(9) & "ALL"

  Do While Not rsCOMPANY.EOF

   SSDDCompany.AddItem rsCOMPANY("com_compcode") & Chr(9) & rsCOMPANY("com_name")

    rsCOMPANY.MoveNext

   Loop

 rsCOMPANY.Close

 Set rsCOMPANY = Nothing


End Function

Public Function PopulateLocation(CompanyCode As String)

Dim RsLocation As New ADODB.Recordset

RsLocation.source = "select loc_locacode , loc_name    from location where loc_npecode='" & deIms.NameSpace & "' and loc_compcode  ='" & CompanyCode & "'"

RsLocation.ActiveConnection = deIms.cnIms

RsLocation.Open

SSDDLocation.RemoveAll

    SSDDLocation.AddItem "ALL" & Chr(9) & "ALL"

  Do While Not RsLocation.EOF

    SSDDLocation.AddItem RsLocation("loc_locacode") & Chr(9) & RsLocation("loc_name")

    RsLocation.MoveNext

   Loop

 RsLocation.Close

 Set RsLocation = Nothing

End Function

Public Function PopulateStockNumber()

Dim RsLocation As New ADODB.Recordset

RsLocation.source = "select stk_stcknumb , stk_desc  from stockmaster where stk_npecode='" & deIms.NameSpace & "' union select 'ALL','ALL'"

RsLocation.ActiveConnection = deIms.cnIms

RsLocation.Open

Set SSDBDDStockNumber.DataSource = RsLocation

End Function


Private Sub SSGridSelection_KeyPress(KeyAscii As Integer)

Dim X As Integer

Select Case SSGridSelection.Col

    Case 0

            SSGridSelection.DroppedDown = True

    Case 1
    
            SSGridSelection.DroppedDown = True
            
    Case 2
    
            SSGridSelection.DroppedDown = True
            
            
     Case 3
     
            SSGridSelection.DroppedDown = True
     

    Case 4

            If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 9) Then

                MsgBox "Invalid Days open, please enter a valid digit.", vbInformation, "Imswin"

                KeyAscii = 0

            End If

    Case 5

        SetFocusOnDatesColumns (5)

    Case 6

        SetFocusOnDatesColumns (6)

End Select

End Sub

Private Sub SSGridSelection_LostFocus()

''MonthView1.Visible = False
''
''MonthView2.Visible = False

End Sub


Public Sub SetFocusOnDatesColumns(column As Integer)

If column = 5 Then


   MonthView1.Top = 1400 'SSGridSelection.Columns(5).Top
   MonthView1.Left = SSGridSelection.Columns(5).Left - 1300
   
   MonthView1.Visible = True
   MonthView2.Visible = False

ElseIf column = 6 Then

   MonthView2.Top = 1200 'SSGridSelection.Columns(6).Top
   MonthView2.Left = SSGridSelection.Columns(6).Left - 1300 'Me.Left + Me.Width - MonthView2.Width - 150
   
   MonthView1.Visible = False
   MonthView2.Visible = True

End If


End Sub

Public Sub LostFocusOnDatesColumns(column As Integer)


   MonthView1.Visible = False
   MonthView2.Visible = False

End Sub

Private Sub SSoleDbDetails_InitColumnProps()
'SSoleDbDetails.ActiveCell
End Sub

Public Function GetDataForPOHeaderTab()

Dim rs As ADODB.Recordset
Dim str As String
Dim Location As String
Dim StockNumber As String
Dim Buyer As String
Dim DaysOpen As String
Dim Fromdate As String
Dim Todate As String
Dim CompanyCode As String

If GHeaderGridFilled = True Then Exit Function

CompanyCode = GselectionCode.CompanyCode

Location = GselectionCode.LocationCode

StockNumber = GselectionCode.StockNumber

Buyer = GselectionCode.Buyer

DaysOpen = GselectionCode.OpenFor

Fromdate = GselectionCode.Fromdate

Todate = GselectionCode.Todate

Set rs = New ADODB.Recordset

str = " SELECT distinct po_ponumb, po_date, po_apprby, po_datesent, po_orig, DATEDIFF(dd, po_date, GETDATE()) AS DatesOpen, po_AssignreqDate,"

str = str & " dbo.XUSERPROFILE.usr_username FROM dbo.PO "

str = str & " INNER JOIN dbo.POITEM ON  dbo.POITEM.poi_ponumb =po.po_ponumb  AND dbo.POITEM.poi_npecode =po.po_npecode "

str = str & " LEFT OUTER JOIN dbo.XUSERPROFILE ON po_buyr = dbo.XUSERPROFILE.usr_userid AND po_npecode = dbo.XUSERPROFILE.usr_npecode "

str = str & " WHERE (po_docutype = 'r') AND (po_npecode = '" & deIms.NameSpace & "') and po_ponumb not in "

str = str & " ( select distinct poi_requnumb from poitem where poi_npecode='" & deIms.NameSpace & "' and len(ltrim(rtrim(isnull(poi_requnumb,'')))) >0) and po_stas ='op' "

If Trim(Len(CompanyCode)) > 0 And CompanyCode <> "ALL" Then str = str & " and po.po_compcode = '" & Trim(CompanyCode) & "'"

If Trim(Len(Location)) > 0 And Location <> "ALL" Then str = str & " and po.po_invloca = '" & Trim(Location) & "'"

If Trim(Len(Buyer)) > 0 And Buyer <> "ALL" Then str = str & " and po.po_buyr = '" & Trim(Buyer) & "'"

If Trim(Len(DaysOpen)) > 0 Then str = str & " and datediff(dd,po.po_date, GETDATE()) <=" & CInt(DaysOpen)

If Trim(Len(Fromdate)) > 0 Then str = str & " and po_date >=  '" & CDate(Trim(Fromdate)) & "'"

If Trim(Len(Todate)) > 0 Then str = str & " and po_date <='" & CDate(Trim(Todate)) & "'"

If Trim(Len(StockNumber)) > 0 And StockNumber <> "ALL" Then str = str & " and poitem.poi_comm ='" & Trim(StockNumber) & "'"

If Trim(SSOleFilter) = mFilter.UnAssignedreqs Then

        str = str & " and len(ltrim(rtrim(isnull(po_buyr,'')))) = 0 "
        
        
ElseIf Trim(SSOleFilter) = mFilter.AssignedReqs Then

        str = str & " and len(ltrim(rtrim(isnull(po_buyr,'')))) <> 0 "
        
ElseIf Trim(SSOleFilter) = mFilter.AllReqs Then

        'str = str & " and len(ltrim(rtrim(isnull(po_buyr,'')))) = 0 "
        
End If

str = str & " ORDER BY po.po_ponumb"

rs.source = str

rs.ActiveConnection = deIms.cnIms

rs.Open , , adOpenKeyset, adLockOptimistic

Label1.Caption = rs.RecordCount & " records found."

SSGridHeaderDetails.RemoveAll

Do While Not rs.EOF

    SSGridHeaderDetails.AddItem rs("po_ponumb") & vbTab & FormatDateTime(rs("po_date"), vbShortDate) & vbTab & FormatDateTime(rs("po_datesent"), vbShortDate) & vbTab & FormatDateTime(rs("po_AssignreqDate"), vbShortDate) & vbTab & rs("po_orig") & vbTab & rs("DatesOpen") & vbTab & rs("usr_username")

    rs.MoveNext

Loop

GHeaderGridFilled = True

End Function

Public Function GetDataForPODetailsTab()

Dim Location As String
Dim StockNumber As String
Dim Buyer As String
Dim DaysOpen As String
Dim Fromdate As String
Dim Todate As String
Dim CompanyCode As String
Dim TotalPrice As String
Dim query As String
Dim RsPOITEMS As New ADODB.Recordset
Dim PreviousPO As String
Dim Ponumb As String
Dim TotalIncludedPrice As Double
Dim Count As Integer
Dim I As Integer
Dim CurrencyX As String
If GDetailsGridFilled = True Then

    Call MoveToPoinPoDetails(SSGridHeaderDetails.Columns(0).Text)
    
    Exit Function

End If

CompanyCode = GselectionCode.CompanyCode

Location = GselectionCode.LocationCode

StockNumber = GselectionCode.StockNumber

Buyer = GselectionCode.Buyer

DaysOpen = GselectionCode.OpenFor

Fromdate = GselectionCode.Fromdate

Todate = GselectionCode.Todate


query = " SELECT TOP 100 PERCENT RequisitionLI.poi_ponumb AS PONUMB, RequisitionLI.poi_liitnumb AS LineNumb,"

query = query & "(select sum(poi_totaprice)  from poitem where poi_requnumb = RequisitionLI.poi_ponumb AND  "

query = query & " poi_npecode = RequisitionLI.poi_npecode ) 'TotalIncludedPrice',"

query = query & " POITEMINCLUDED.poi_ponumb AS POsIncludedIn, POITEMINCLUDED.poi_liitnumb AS LineItemsIncludedIn, POITEMINCLUDED.poi_totaprice as POItemTotalPrice,"

query = query & " POIncluded.po_creadate POCreationDate, Requisition.po_AssignreqDate RequsitionAssignedDate, datediff(dd,Requisition.po_AssignreqDate,POIncluded.po_creadate ) AS DaysElapsed, POIncluded.po_currcode as 'Currency'"

query = query & " FROM dbo.PO Requisition INNER JOIN"

query = query & " dbo.POITEM RequisitionLI ON Requisition.po_ponumb = RequisitionLI.poi_ponumb AND"

query = query & " Requisition.po_npecode = RequisitionLI.poi_npecode LEFT OUTER JOIN"

query = query & " dbo.POITEM POITEMINCLUDED ON RequisitionLI.poi_ponumb = POITEMINCLUDED.poi_requnumb AND"

query = query & " RequisitionLI.poi_liitnumb = POITEMINCLUDED.poi_requliitnumb AND RequisitionLI.poi_npecode = POITEMINCLUDED.poi_npecode LEFT OUTER JOIN"

query = query & " dbo.PO POIncluded ON POIncluded.po_ponumb = POITEMINCLUDED.poi_ponumb AND"

query = query & " POIncluded.po_npecode = POITEMINCLUDED.poi_npecode"

query = query & " where Requisition.po_docutype='r' and Requisition.po_stas ='OP'"

If Trim(Len(CompanyCode)) > 0 And CompanyCode <> "ALL" Then query = query & " and Requisition.po_compcode = '" & Trim(CompanyCode) & "'"

If Trim(Len(Location)) > 0 And Location <> "ALL" Then query = query & " and Requisition.po_invloca = '" & Trim(Location) & "'"

If Trim(Len(Buyer)) > 0 And Buyer <> "ALL" Then query = query & " and Requisition.po_buyr = '" & Trim(Buyer) & "'"

If Trim(Len(DaysOpen)) > 0 Then query = query & " and datediff(dd,Requisition.po_date, GETDATE()) <=" & CInt(DaysOpen)

If Trim(Len(Fromdate)) > 0 Then query = query & " and Requisition.po_date >=  '" & CDate(Trim(Fromdate)) & "'"

If Trim(Len(Todate)) > 0 Then query = query & " and Requisition.po_date <='" & CDate(Trim(Todate)) & "'"

If Trim(Len(StockNumber)) > 0 And StockNumber <> "ALL" Then query = query & " and RequisitionLI.poi_comm ='" & Trim(StockNumber) & "'"

query = query & " ORDER BY RequisitionLI.poi_ponumb, cast(RequisitionLI.poi_liitnumb as int)"

RsPOITEMS.source = query

RsPOITEMS.ActiveConnection = deIms.cnIms

RsPOITEMS.Open

SSGridPOITEMDETAILS.RemoveAll

If RsPOITEMS.RecordCount = 0 Then ReDim Preserve GLocationPoDetails(1, Count)

If RsPOITEMS.RecordCount > 0 Then

Do While Not RsPOITEMS.EOF


     If PreviousPO <> Trim(RsPOITEMS("PONUMB")) Then

       If SSGridPOITEMDETAILS.Rows > 0 Then
        
            SSGridPOITEMDETAILS.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "Total Value" & vbTab & CurrencyX & " " & Format(TotalIncludedPrice, "0.00")
            
        End If

        TotalIncludedPrice = Format(IIf(Len(Trim(RsPOITEMS("TotalIncludedPrice") & "")) = 0, "0.00", Trim(RsPOITEMS("TotalIncludedPrice"))), "0.00")
        
        CurrencyX = RsPOITEMS("currency") & ""
        
        SSGridPOITEMDETAILS.AddItem RsPOITEMS("PONUMB") & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
        
        TotalPrice = ""
        
        ReDim Preserve GLocationPoDetails(1, Count)

        GLocationPoDetails(0, Count) = Trim(RsPOITEMS("PONUMB"))
        
        GLocationPoDetails(1, Count) = SSGridPOITEMDETAILS.Rows - 1 ' SSGridPOITEMDETAILS.AddItemRowIndex(SSGridPOITEMDETAILS.Bookmark)
    
        Count = Count + 1
        

         
     End If
    
    If Len(Trim(RsPOITEMS("POItemTotalPrice") & "")) > 0 Then

        TotalPrice = RsPOITEMS("Currency") & " " & (Format(RsPOITEMS("POItemTotalPrice"), "00.00"))
        
    Else
    
        TotalPrice = "0.00"
        
    End If

    
     SSGridPOITEMDETAILS.AddItem "    '" & vbTab & RsPOITEMS("LineNumb") & vbTab & RsPOITEMS("POsIncludedIn") & vbTab & RsPOITEMS("LineItemsIncludedIn") & vbTab & TotalPrice & vbTab & FormatDateTime(RsPOITEMS("POCreationDate"), vbShortDate) & vbTab & RsPOITEMS("DaysElapsed")
    
     PreviousPO = Trim(RsPOITEMS("PONUMB"))
    
     RsPOITEMS.MoveNext

Loop

    SSGridPOITEMDETAILS.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "Total Value" & vbTab & "US$ " & Format(TotalIncludedPrice, "0.00")

End If

Call MoveToPoinPoDetails(SSGridHeaderDetails.Columns(0).Text)

LblCountPoitem.Caption = RsPOITEMS.RecordCount & " records found."

GDetailsGridFilled = True

End Function

Private Sub SSOleFilter_Click()
    GHeaderGridFilled = False
    GDetailsGridFilled = False
    Call GetDataForPOHeaderTab

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Screen.MousePointer = vbHourglass

Call GetDataForTheSelection



Screen.MousePointer = vbArrow

End Sub


Public Sub PopulatebuyerForDetails()


Dim Rsbuyer As New ADODB.Recordset

Rsbuyer.source = "select buy_username,usr_username , buy_npecode from buyer,xuserprofile where buy_username = usr_userid and buy_npecode = usr_npecode and usr_npecode ='" & deIms.NameSpace & "'"

Rsbuyer.ActiveConnection = deIms.cnIms

Rsbuyer.Open

SSDDBuyerDetails.RemoveAll

  Do While Not Rsbuyer.EOF

    SSDDBuyerDetails.AddItem Rsbuyer("buy_username") & Chr(9) & Rsbuyer("usr_username")

    Rsbuyer.MoveNext

   Loop

 Rsbuyer.Close

 Set Rsbuyer = Nothing

End Sub

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

If SSTabRequisitions.Tab = 1 Then
opt_Email.value = True
LROleDBNavBar1.Visible = False
ElseIf SSTabRequisitions.Tab = 0 Then
LROleDBNavBar1.Visible = True
End If
End Sub

Private Sub Txt_search_Change()

Dim Grid As SSOleDBGrid

Dim X As Integer

Dim Count As Integer

Dim I As Integer

If SSOLEDBEmail.Visible = True Then Set Grid = SSOLEDBEmail

If SSOLEDBFax.Visible = True Then Set Grid = SSOLEDBFax

I = Len(Txt_search)

Count = 1

    Grid.MoveFirst

    Do While Not Grid.Rows = Count

        If UCase(Txt_search) = UCase(Mid(Grid.Columns(0).value, 1, I)) Then

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

rst.source = Sql

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

rst.source = Sql

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
Dim X As Integer

If Len(dgRecipientList.SelBookmarks(0)) = 0 Then

    MsgBox "Please make a selection first.", vbInformation, "Imswin"

    Exit Sub

 End If



    dgRecipientList.DeleteSelected



End Sub
Private Sub SSOLEDBFax_DblClick()
On Error Resume Next

    'dgRecipientList.AddItem SSOLEDBFax.Columns(1).Value

    AddRecepient SSOLEDBFax.Columns(1).value

    If Err Then Err.Clear
End Sub

Public Function IsPoAlreadyInclude() As Boolean

Dim rsPO As New ADODB.Recordset

On Error GoTo Errhandler

IsPoAlreadyInclude = True

rsPO.source = "select count(*) RecordCount from poitem where poi_requnumb='" & Trim(SSGridHeaderDetails.Columns(0).value) & "' and poi_npecode ='" & deIms.NameSpace & "'"

rsPO.ActiveConnection = deIms.cnIms

rsPO.Open

If rsPO("RecordCount") = 0 Then IsPoAlreadyInclude = False

Exit Function

Errhandler:

MsgBox "Some errors occurred while trying to verify if the Requisition has already been assigned.", vbCritical, "Imswin"

Err.Clear

End Function

Public Function SaveRecord()

Dim I As Integer

Dim rsPO As ADODB.Recordset
Dim imsLock As imsLock.Lock
'this valraible is used because the record is saved using two doors, the navbar and the other is from afterupdate event of the grid.
'the variable makes sure that the two events are not steppig on each other.

If GValueChanged = False Then Exit Function

If MsgBox("Are you sure you want to save the changes.", vbInformation + vbYesNo, "Imswin") = vbYes Then

On Error GoTo Errhandler


    Set rsPO = New ADODB.Recordset

    rsPO.source = "Update PO set po_buyr = '" & SSGridHeaderDetails.Columns(8).value & "' , po_AssignreqDate =getdate()  where po_ponumb ='" & Trim(SSGridHeaderDetails.Columns(0).value) & "' and po_npecode ='" & deIms.NameSpace & "'"

    rsPO.ActiveConnection = deIms.cnIms

    rsPO.Open
    
    'If Len(Trim(SSGridHeaderDetails.Columns(3).text)) = 0 Then
    SSGridHeaderDetails.Columns(3).Text = Date

GOldValue = ""

Else

SSGridHeaderDetails.Columns(6).Text = GOldValue

GOldValue = ""

End If

GValueChanged = False

    
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
                    

Exit Function

Errhandler:

MsgBox "Errors Occurred while trying to update the BUYER for the Transaction " & SSGridHeaderDetails.Columns(0).value & ". Could not save the record." & Err.Description, vbCritical, "Imswin"

Err.Clear

LROleDBNavBar1.SaveEnabled = True

GOldValue = ""

    
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
  
                    

End Function

Private Sub txtsearch_GotFocus()
If Trim(txtsearch.Text) = "Hit Enter To see results" Then txtsearch = ""
End Sub

Private Sub txtsearch_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then Call MoveToPO(txtsearch, SSGridHeaderDetails)

End Sub

Private Sub txtsearchDetl_GotFocus()
If Trim(txtsearchDetl.Text) = "Hit Enter To see results" Then txtsearchDetl = ""
End Sub

Private Sub SSGridHeaderDetails_RowLoaded(ByVal Bookmark As Variant)
    Dim I As Integer
                    
    If Len(SSGridHeaderDetails.Columns(7).Text) > 0 Then
       If SSGridHeaderDetails.Columns(7).Text = 1 Then
        For I = 0 To SSGridHeaderDetails.Cols - 1
            SSGridHeaderDetails.Columns(I).CellStyleSet "RowModified" ', SSGridHeaderDetails.row
        Next I
    End If
   End If
End Sub


Public Function MoveToPO(PoNumber As String, SSoleGrid As SSOleDBGrid)
Dim I As Integer
PoNumber = Trim(PoNumber)

SSoleGrid.MoveFirst

For I = 0 To SSoleGrid.Rows

If UCase(Mid(SSoleGrid.Columns(0).Text, 1, Len(Trim(PoNumber)))) = UCase(Trim(PoNumber)) Then

    Exit For

End If

SSoleGrid.MoveNext

Next I

End Function

Private Sub txtsearchDetl_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call MoveToPoinPoDetails(txtsearchDetl)
End Sub

Public Function MoveToPoinPoDetails(Ponumb As String)

Dim I As Integer

If IsArrayLoaded(GLocationPoDetails) = False Then Exit Function
If Len(Trim(Ponumb)) = 0 Then Exit Function
For I = 0 To UBound(GLocationPoDetails, 2)

    If UCase(Trim(Ponumb)) = GLocationPoDetails(0, I) Then
    
        SSGridPOITEMDETAILS.MoveFirst
        
        SSGridPOITEMDETAILS.MoveRecords GLocationPoDetails(1, I)
        
        Exit For
        
     End If
    

Next I

End Function
