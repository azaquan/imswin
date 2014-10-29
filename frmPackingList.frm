VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmPackingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List/Manifest Management"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   13905
   Tag             =   "02030200"
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh Data"
      Height          =   255
      Left            =   12240
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13685
      _ExtentX        =   24130
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Header"
      TabPicture(0)   =   "frmPackingList.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label(15)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label(17)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label(19)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label(20)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label(21)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label(22)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label(23)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label(24)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label(25)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label(26)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label(27)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label(28)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label(29)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Shape1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblPodNamevalue"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblpodname"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblpoddatevalue"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblpoddate"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ShipperList"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ShipToList"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "SoldToList"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "destinationList(4)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "destinationList(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "destinationList(3)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "destinationList(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "destinationList(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "PriorityList"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "PackingListList"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtClause"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "DTPicker1"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cell(0)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "cell(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cell(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cell(3)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cell(4)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cell(5)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cell(6)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "cell(7)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cell(8)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cell(9)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cell(10)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cell(11)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cell(12)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cell(13)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cell(14)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cell(15)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cell(16)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cell(17)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cell(18)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "cell(19)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "cell(20)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "cell(21)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "cell(22)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cell(23)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cell(24)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "cell(25)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cell(26)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "cell(27)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "cell(28)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "cell(29)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "remarkBUTTON"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).ControlCount=   78
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "frmPackingList.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Recipients"
      Tab(1).Control(1)=   "Imsmail1"
      Tab(1).Control(2)=   "cmd_Add"
      Tab(1).Control(3)=   "cmd_Remove"
      Tab(1).Control(4)=   "RecipientList"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Line Item List"
      TabPicture(2)   =   "frmPackingList.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "packinglistLABEL"
      Tab(2).Control(1)=   "shiptoLABEL"
      Tab(2).Control(2)=   "POlist"
      Tab(2).Control(3)=   "Command1"
      Tab(2).Control(4)=   "TextLINE"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton remarkBUTTON 
         Caption         =   "Remark"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   29
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   80
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   28
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   78
         Top             =   4320
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   27
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   70
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   26
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   68
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   66
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   64
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   23
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   62
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   9360
         TabIndex        =   60
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   9360
         TabIndex        =   58
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   9360
         TabIndex        =   56
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   54
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   5400
         MaxLength       =   40
         TabIndex        =   52
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   50
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   48
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   5400
         TabIndex        =   46
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   5400
         TabIndex        =   44
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   42
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   5400
         TabIndex        =   40
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   5400
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   36
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   34
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   32
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   30
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1920
         MaxLength       =   24
         TabIndex        =   28
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   26
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   24
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   20
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6000
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         CustomFormat    =   "MMMM/dd/yyyy"
         Format          =   60424195
         CurrentDate     =   36867
      End
      Begin VB.TextBox TextLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -68640
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Show Only Selection"
         Height          =   375
         Left            =   -63550
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid RecipientList 
         Height          =   2535
         Left            =   -72960
         TabIndex        =   9
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4471
         _Version        =   393216
         RowHeightMin    =   240
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74520
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74520
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin ImsMailVB.Imsmail Imsmail1 
         Height          =   3375
         Left            =   -74640
         TabIndex        =   5
         Top             =   3480
         Width           =   13135
         _ExtentX        =   23178
         _ExtentY        =   5953
      End
      Begin VB.TextBox txtClause 
         Height          =   1635
         Left            =   240
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   5640
         Width           =   13320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid POlist 
         Height          =   6120
         Left            =   -74880
         TabIndex        =   4
         Top             =   1080
         Width           =   13420
         _ExtentX        =   23680
         _ExtentY        =   10795
         _Version        =   393216
         Cols            =   17
         RowHeightMin    =   285
         BackColorSel    =   16761024
         ForeColorSel    =   0
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   17
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid PackingListList 
         Height          =   975
         Left            =   1920
         TabIndex        =   11
         Top             =   990
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid PriorityList 
         Height          =   975
         Left            =   1920
         TabIndex        =   72
         Top             =   2430
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid destinationList 
         Height          =   975
         Index           =   1
         Left            =   5400
         TabIndex        =   74
         Top             =   1710
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         DataMember      =   "12"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid destinationList 
         Height          =   975
         Index           =   0
         Left            =   5400
         TabIndex        =   73
         Top             =   1350
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         DataMember      =   "11"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid destinationList 
         Height          =   975
         Index           =   3
         Left            =   5400
         TabIndex        =   76
         Top             =   3150
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         DataMember      =   "15"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid destinationList 
         Height          =   975
         Index           =   2
         Left            =   5400
         TabIndex        =   75
         Top             =   2790
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         DataMember      =   "14"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid destinationList 
         Height          =   975
         Index           =   4
         Left            =   1920
         TabIndex        =   77
         Top             =   4230
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         DataMember      =   "7"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SoldToList 
         Height          =   975
         Left            =   9015
         TabIndex        =   84
         Top             =   1710
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ShipToList 
         Height          =   975
         Left            =   9015
         TabIndex        =   83
         Top             =   1350
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ShipperList 
         Height          =   975
         Left            =   9000
         TabIndex        =   82
         Top             =   960
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblpoddate 
         Caption         =   "POD Date"
         Height          =   270
         Left            =   4440
         TabIndex        =   90
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblpoddatevalue 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   89
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label lblpodname 
         Caption         =   "POD Name"
         Height          =   255
         Left            =   960
         TabIndex        =   88
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblPodNamevalue 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   275
         Left            =   1920
         TabIndex        =   87
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         Height          =   4335
         Left            =   240
         Top             =   720
         Width           =   15
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Marks 4"
         Height          =   255
         Index           =   29
         Left            =   7920
         TabIndex        =   81
         Top             =   4725
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Marks 3"
         Height          =   255
         Index           =   28
         Left            =   7920
         TabIndex        =   79
         Top             =   4365
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Marks 2"
         Height          =   255
         Index           =   27
         Left            =   7920
         TabIndex        =   71
         Top             =   4005
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Marks 1"
         Height          =   255
         Index           =   26
         Left            =   7920
         TabIndex        =   69
         Top             =   3645
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Custom's Ref."
         Height          =   255
         Index           =   25
         Left            =   7920
         TabIndex        =   67
         Top             =   2925
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Forwarder's Ref."
         Height          =   255
         Index           =   24
         Left            =   7920
         TabIndex        =   65
         Top             =   2565
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipper's Ref."
         Height          =   255
         Index           =   23
         Left            =   7920
         TabIndex        =   63
         Top             =   2205
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Sold to"
         Height          =   255
         Index           =   22
         Left            =   7920
         TabIndex        =   61
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship to"
         Height          =   255
         Index           =   21
         Left            =   7920
         TabIndex        =   59
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipper"
         Height          =   255
         Index           =   20
         Left            =   7920
         TabIndex        =   57
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "HAWB / TBL"
         Height          =   255
         Index           =   19
         Left            =   3840
         TabIndex        =   55
         Top             =   4725
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume CBM"
         Height          =   255
         Index           =   18
         Left            =   3840
         TabIndex        =   53
         Top             =   4365
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross Weight Kgs"
         Height          =   255
         Index           =   17
         Left            =   3840
         TabIndex        =   51
         Top             =   4005
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Pieces"
         Height          =   255
         Index           =   16
         Left            =   3840
         TabIndex        =   49
         Top             =   3645
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Index           =   15
         Left            =   3840
         TabIndex        =   47
         Top             =   2925
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   255
         Index           =   14
         Left            =   3840
         TabIndex        =   45
         Top             =   2565
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Flight Voyage 2"
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   43
         Top             =   2205
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   41
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   39
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Flight Voyage 1"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   37
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Estimated Arrive"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   35
         Top             =   4725
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Estimated Departure"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   33
         Top             =   4365
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Destination"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   4005
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipping Term"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   3645
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Via Carrier"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   2925
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "AWB / BL"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   2565
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipping Mode"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   2205
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Shipping Date"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Document Date"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Packing/Manifest"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   765
         Width           =   1575
      End
      Begin VB.Label shiptoLABEL 
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
         Left            =   -70920
         TabIndex        =   14
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label packinglistLABEL 
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
         Left            =   -74760
         TabIndex        =   13
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74520
         TabIndex        =   8
         Top             =   720
         Width           =   1260
      End
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelEnabled   =   0   'False
      EMailVisible    =   -1  'True
      FirstVisible    =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      Wrappable       =   -1  'True
      Mode            =   3
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin VB.Label lblStatu 
      Alignment       =   1  'Right Justify
      Caption         =   "Visualization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   7560
      Width           =   4335
   End
End
Attribute VB_Name = "frmPackingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Form As FormMode
Dim readyFORsave As Boolean
Dim rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim lastCELL, focusHERE, nextCELL As Integer
Dim SaveEnabled As Boolean
Dim selectionSTART As Integer
Dim multiMARKED As Boolean
Dim WithEvents st As frm_ShipTerms
Attribute st.VB_VarHelpID = -1
Dim rowguid, recLocked As Boolean, dbtablename As String, grid1 As Boolean, grid2 As Boolean
Dim POValue 'jawdat
Dim OPENEDFORM As Boolean
Dim settingUP As Boolean
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    If cellACTIVE <> "" Then
        With gridACTIVE
            If .Rows < 1 Then Exit Sub
            If settingUP = False Then
                If Not .Visible Then .Visible = True
            End If
            If IsNumeric(.Tag) Then
                .row = val(.Tag)
                .Col = column
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
            End If
            .Col = column
            .Tag = ""
            found = False
            For i = 0 To .Rows - 1
                word = Trim(UCase(.TextMatrix(i, column)))
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    .row = i
                    .CellBackColor = &H800000 'Blue
                    .CellForeColor = &HFFFFFF 'White
                    .Tag = .row
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                .row = 0
                .Tag = ""
            End If
            If IsNumeric(.Tag) Then .topROW = val(.Tag)
        End With
    End If
End Sub

Sub arrowKEYS(direction As String, Index As Integer)
Dim Grid As MSHFlexGrid
    With cell(Index)
        Select Case Index
            Case 0
                Set Grid = PackingListList
            Case 3
                Set Grid = PriorityList
            Case 7
                Set Grid = destinationList(4)
            Case 11
                Set Grid = destinationList(0)
            Case 12
                Set Grid = destinationList(1)
            Case 14
                Set Grid = destinationList(2)
            Case 15
                Set Grid = destinationList(3)
            Case 20
                Set Grid = ShipperList
            Case 21
                Set Grid = ShipToList
            Case 22
                Set Grid = SoldToList
        End Select
        
        Select Case Index
            Case 0, 3, 7, 11, 12, 14, 15, 20, 21, 22
                If IsNumeric(Grid.Tag) Then
                    Grid.row = val(Grid.Tag)
                    Grid.CellBackColor = &HFFFF00   'Cyan
                    Grid.CellForeColor = &H80000008 'Default Window Text
                End If
                Select Case direction
                Case "down"
                    If Grid.row < (Grid.Rows - 1) Then
                        If Grid.row = 0 And .Text = "" Then
                            .Text = Grid.Text
                        Else
                            Grid.row = Grid.row + 1
                        End If
                    Else
                        Grid.row = Grid.Rows - 1
                    End If
                Case "up"
                    If Grid.row > 0 Then
                        Grid.row = Grid.row - 1
                    Else
                        Grid.row = 1
                    End If
            End Select
            If Not Grid.Visible Then
                Grid.Visible = True
            End If
            Grid.ZOrder
            Grid.topROW = Grid.row
            Grid.SetFocus
        End Select
    End With
End Sub

Sub BeforePrint()
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\packinglist.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "manifestnumb;" + cell(0) + ";true"
        Call translator.Translate_Reports("packinglist.rpt")
    End With
End Sub

Sub begining()
Dim i
On Error Resume Next
    With PriorityList
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .Visible = False
        .ColWidth(0) = 1600
        .ColWidth(1) = 0
    End With
    Call getPRIORITYlist
    
    With destinationList
        For i = 0 To 4
            With .Item(i)
                .TextMatrix(0, 0) = ""
                .TextMatrix(0, 1) = ""
                .Visible = False
                .ColWidth(0) = 2000
                .ColWidth(1) = 0
            End With
        Next
    End With
    Call getDestinationList
    
    With ShipperList
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .Visible = False
        .ColWidth(0) = 2715
        .ColWidth(1) = 0
    End With
    Call getSHIPPERlist
    
    With ShipToList
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .Visible = False
        .ColWidth(0) = 2715
        .ColWidth(1) = 0
    End With
    Call getSHIPTOlist
    
    With SoldToList
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .Visible = False
        .ColWidth(0) = 2715
        .ColWidth(1) = 0
    End With
    Call getSOLDTOlist
End Sub


Sub Coloring(dye) 'JCG 2008/6/21 inserting new col 3
Dim currentCOL As Integer
Dim i As Integer
    With POlist
        currentCOL = .Col
        'For i = 1 To 12
        For i = 1 To 13
            .Col = i
            .CellBackColor = dye
        Next
        .Col = currentCOL
    End With
End Sub

Sub markROW() 'JCG 2008/6/21 inserting new col 3
Dim i, t
    With POlist
        If .TextMatrix(.row, 1) <> "" And .TextMatrix(.row, 1) <> Chr(34) Then
            .Col = 0
            .CellFontName = "Wingdings 3"
            .CellFontSize = 10
            If .Text = "" Then
            
                'lock jawdat
                'jawdat, start copy
                Dim currentformname, currentformname1
                currentformname = Me.Name
                currentformname1 = Me.Name
                Dim imsLock As imsLock.Lock
                Dim ListOfPrimaryControls() As String
                Set imsLock = New imsLock.Lock
                
                ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
                
                Call imsLock.Check_Lock(recLocked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)   'lock should be here, added by jawdat, 2.1.02
    
                If recLocked = True Then
                    t = .TextMatrix(.row, 1)
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 1) = t And .TextMatrix(i, 0) = "Æ" Then
                            recLocked = False
                            Exit For
                        End If
                    Next
                End If
                If recLocked = True Then                                        'sets locked = true because another user has this record open in edit mode
                    .Text = ""
                    Exit Sub
                Else
                    If Form = mdModification Then
                        'If .TextMatrix(.row, 15) = "" Then
                        If .TextMatrix(.row, 16) = "" Then
                            .Text = "Æ"
                            '.TextMatrix(.row, 9) = .TextMatrix(.row, 8)
                            .TextMatrix(.row, 10) = .TextMatrix(.row, 9)
                        Else
                            .Text = "}"
                        End If
                    Else
                        .Text = "Æ"
                        '.TextMatrix(.row, 9) = .TextMatrix(.row, 8)
                        .TextMatrix(.row, 10) = .TextMatrix(.row, 9)
                    End If
                End If
                recLocked = False
                'jawdat, end copy
                
            Else
                If Form = mdModification Then
                    'If .TextMatrix(.row, 15) = "" Then
                    If .TextMatrix(.row, 16) = "" Then
                        .Text = ""
                        '.TextMatrix(.row, 9) = "0.00"
                        .TextMatrix(.row, 10) = "0.00"
                    End If
                Else
                    .Text = ""
                    '.TextMatrix(.row, 9) = "0.00"
                    .TextMatrix(.row, 10) = "0.00"
                End If
            End If
            '.TextMatrix(.row, 10) = ""
            .TextMatrix(.row, 11) = ""
            '.TextMatrix(.row, 12) = FormatNumber(CDbl(.TextMatrix(.row, 9)) * CDbl(.TextMatrix(.row, 11)), 2)
            .TextMatrix(.row, 13) = FormatNumber(CDbl(.TextMatrix(.row, 10)) * CDbl(.TextMatrix(.row, 12)), 2)
        End If
    End With
End Sub

Sub clearDOCUMENT()
Dim i As Integer
    readyFORsave = False
    For i = 1 To 29
        cell(i) = ""
        cell(i).BackColor = txtClause.BackColor
    Next
    PackingListList.Visible = False
    PriorityList.Visible = False
    For i = 0 To 4
        destinationList(i).Visible = False
    Next
    ShipperList.Visible = False
    ShipToList.Visible = False
    SoldToList.Visible = False
    DTPicker1.Visible = False
    txtClause = ""
    cell(0).SetFocus
    Command1.Caption = "&Show Only Selection"
    
    'POD
'    lblpoddatevalue.Caption = "" 'MM 111609
'    lblPodNamevalue.Caption = ""  'MM 111609
    
End Sub

Function controlOBJECT(controlNAME As String) As Control
Dim c As Control
    For Each c In Me.Controls
        If c.Name = controlNAME Then
            Exit For
        End If
        Set c = Nothing
    Next
    Set controlOBJECT = c
End Function

Sub datePICKER(controlNAME As String)
Dim h, i As Integer
Dim c As Control

    With DTPicker1
        .Tag = ""
        For Each c In Me.Controls
            If c.Name = controlNAME Then
                Exit For
            End If
            Set c = Nothing
        Next
        If c Is Nothing Then Exit Sub
        .Tag = controlNAME
    
        .Left = c.Left + c.ColWidth(0)
        .Height = c.RowHeight(i)
        If c.row = 0 Then
            .Top = c.Top
            .Height = .Height - 80
        Else
            h = 20
            For i = 0 To c.row - 1
                h = h + c.RowHeight(i)
            Next
            .Top = h + c.Top - 30
            .Height = .Height + 10
        End If
        .Visible = True
        .value = IIf(IsDate(c.Text), c.Text, Now)
        .SetFocus
        Call DTPicker1_DropDown
    End With
End Sub

Sub getDestinationList()
Dim Sql As String
Dim i As Integer
Dim dataPAKING As New ADODB.Recordset
On Error Resume Next
    For i = 0 To 4
        destinationList(i).Rows = 0
        If Err.number = -2147417848 Then Err.Clear
    Next
    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT des_destcode, des_destname FROM Destination WHERE des_npecode = '" + deIms.NameSpace + "' " _
        & "ORDER BY des_destname"
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                For i = 0 To 4
                    destinationList(i).AddItem !des_destname + vbTab + !des_destcode
                Next
                .MoveNext
            Loop
        End If
        For i = 0 To 4
            destinationList(i).row = 0
        Next
    End With
End Sub

Sub getPACKINGLIST() 'JCG 2008/6/21 inserting new col 3
'On Error Resume Next
Dim dataPAKING  As New ADODB.Recordset
Dim dataRECIP  As New ADODB.Recordset
Dim Sql, Text As String
        
    Screen.MousePointer = 11
    Call clearDOCUMENT
    Sql = "SELECT * from Packing_List_Full_Header WHERE NameSpace = '" + deIms.NameSpace + "' " _
        & "AND PackingListNumber = '" + Trim(cell(0).Text) + "' "
    Set dataPAKING = New ADODB.Recordset
    dataPAKING.Open Sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    With dataPAKING
        If .RecordCount > 0 Then
            'POlist.Cols = 13
            POlist.Cols = 14
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
            NavBar1.EditEnabled = True
            cell(1) = IIf(IsNull(!DocumentDate), "", Format(!DocumentDate, "MMMM/dd/yyyy"))
            cell(2) = IIf(IsNull(!ShipDate), "", Format(!ShipDate, "MMMM/dd/yyyy"))
            
            cell(3) = IIf(IsNull(!Priority), "", !Priority)
            PriorityList.TextMatrix(0, 0) = cell(3)
            PriorityList.TextMatrix(0, 1) = !PriorityCode
            cell(4) = IIf(IsNull(!AWB), "", !AWB)
            cell(5) = IIf(IsNull(!ViaCarrier), "", !ViaCarrier)
            
            cell(6) = IIf(IsNull(!shipterm), "", !shipterm)
            cell(7) = IIf(IsNull(!Destination), "", !Destination)
            destinationList(4).TextMatrix(0, 0) = cell(7)
            destinationList(4).TextMatrix(0, 1) = IIf(IsNull(!DestinationCode), "", !DestinationCode)
            cell(8) = IIf(IsNull(!Etd), "", Format(!Etd, "MMMM/dd/yyyy"))
            cell(9) = IIf(IsNull(!ETA), "", Format(!ETA, "MMMM/dd/yyyy"))
            
            cell(10) = IIf(IsNull(!flight1), "", !flight1)
            cell(11) = IIf(IsNull(!From1), "", !From1)
            destinationList(0).TextMatrix(0, 0) = cell(11)
            destinationList(0).TextMatrix(0, 1) = IIf(IsNull(!DestinationCode), "", !DestinationCode)
            cell(12) = IIf(IsNull(!to1), "", !to1)
            destinationList(1).TextMatrix(0, 0) = cell(12)
            destinationList(1).TextMatrix(0, 1) = IIf(IsNull(!DestinationCode), "", !DestinationCode)
            
            cell(13) = IIf(IsNull(!flight2), "", !flight2)
            destinationList(2).TextMatrix(0, 0) = cell(13)
            destinationList(2).TextMatrix(0, 1) = IIf(IsNull(!DestinationCode), "", !DestinationCode)
            cell(14) = IIf(IsNull(!from2), "", !from2)
            destinationList(3).TextMatrix(0, 0) = cell(14)
            destinationList(3).TextMatrix(0, 1) = IIf(IsNull(!DestinationCode), "", !DestinationCode)
            cell(15) = IIf(IsNull(!to2), "", !to2)
            
            cell(16) = IIf(IsNull(!NumberOfPieces), "", !NumberOfPieces)
            cell(17) = IIf(IsNull(!GrossWeight), "", !GrossWeight)
            cell(18) = IIf(IsNull(!Volume), "", !Volume)
            cell(19) = IIf(IsNull(!HAWB), "", !HAWB)
            
            cell(20) = IIf(IsNull(!Shipper), "", !Shipper)
            
            If ShipperList.Rows > 0 Then
                ShipperList.TextMatrix(0, 0) = cell(20)
                ShipperList.TextMatrix(0, 1) = !ShipperCode
            End If
            cell(21) = IIf(IsNull(!shipto), "", !shipto)
            If ShipToList.Rows > 0 Then
                ShipToList.TextMatrix(0, 0) = cell(21)
                ShipToList.TextMatrix(0, 1) = IIf(IsNull(!ShipToCode), "", !ShipToCode)
            End If
            cell(22) = IIf(IsNull(!SoldTo), "", !SoldTo)
            If SoldToList.Rows > 0 Then
                SoldToList.TextMatrix(0, 0) = cell(22)
                SoldToList.TextMatrix(0, 1) = IIf(IsNull(!SoldToCode), "", !SoldToCode)
            End If
            
            cell(23) = IIf(IsNull(!ShippersReference), "", !ShippersReference)
            cell(24) = IIf(IsNull(!ForwardsReference), "", !ForwardsReference)
            cell(25) = IIf(IsNull(!CustomsReference), "", !CustomsReference)
            
            cell(26) = IIf(IsNull(!mark1), "", !mark1)
            cell(27) = IIf(IsNull(!mark2), "", !mark2)
            cell(28) = IIf(IsNull(!mark3), "", !mark3)
            cell(29) = IIf(IsNull(!mark4), "", !mark4)
            
            txtClause = IIf(IsNull(!remark), "", !remark)
            
            'POD
'            lblpoddatevalue.Caption = (!pl_pod_datetime & "") 'MM 111609
'            lblPodNamevalue.Caption = (!pl_pod_name & "") 'MM 111609
            
            
            'Recipients
            Set dataRECIP = New ADODB.Recordset
            Sql = "SELECT plrc_rec FROM PACKINGREC " _
                & "WHERE plrc_npecode = '" + deIms.NameSpace + "' AND plrc_manfnumb = '" + Trim(cell(0).Text) + "' "
            dataRECIP.Open Sql, deIms.cnIms, adOpenForwardOnly
            If Err.number = 0 Then
                Set rsReceptList = New ADODB.Recordset
                Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
                rsReceptList.Open
                If dataRECIP.RecordCount > 0 Then
                    Do While Not dataRECIP.EOF
                        rsReceptList.AddNew
                        rsReceptList(0) = dataRECIP(0)
                        dataRECIP.MoveNext
                    Loop
                End If
                RecipientList.Rows = 2
                RecipientList.TextMatrix(1, 1) = ""
                Call getRECIPIENTSlist
                RecipientList.Refresh
            Else
                Err.Clear
            End If
            
            'Details
            Call getLINEitems(cell(0))
            cell(1).SelStart = 0
            cell(0).SelLength = Len(cell(0))
            cell(0).SetFocus
            
            
        Else
            NavBar1.PrintEnabled = False
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00088")
            MsgBox IIf(msg1 = "", "Does not exist yet", msg1)
            cell(0) = ""
        End If
    End With
    If PackingListList.Rows > 0 Then
        PackingListList.Visible = True
    End If
    Screen.MousePointer = 0
End Sub

Sub getPackingListList()
On Error Resume Next
Dim Sql As String
Dim dataPAKING As New ADODB.Recordset
On Error GoTo errorTRACK

    With PackingListList
        .Visible = False
        .ColWidth(0) = 1600
    End With
    
    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT pl_manfnumb FROM PACKINGLIST WHERE pl_npecode = '" + deIms.NameSpace + "' ORDER BY pl_manfnumb"
    PackingListList.Clear
    PackingListList.Rows = 0
    
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                PackingListList.AddItem " " + Trim(!pl_manfnumb)
                .MoveNext
            Loop
        End If
        If PackingListList.Rows > 0 Then PackingListList.row = 0
    End With
    Exit Sub
    
errorTRACK:
    If Err.number = -2147417848 Then
        Err.Clear
        Resume Next
    Else
        MsgBox Err.Description
        Exit Sub
    End If
End Sub

Sub getLINEitems(packinglist As String) 'JCG 2008/6/21 inserting new col 3
Dim dataPACKING  As New ADODB.Recordset
Dim Sql, rowTEXT, current As String
Dim i As Integer
    Err.Clear
    'On Error Resume Next
    makeDETAILgrid
    If packinglist = "*" Then
        Sql = "SELECT * from POs_for_Packing_List WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND poi_stasliit = 'OP' AND poi_stasdlvy IN ('RP', 'RC') AND poi_stasship IN ('SP', 'NS') " _
            & "AND shipped < delivered " _
            & "ORDER BY PO, CONVERT(integer, LineItem)"
    Else
        packinglist = Trim(packinglist)
        Sql = "SELECT * from Packing_List_Details WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PackingList = '" + packinglist + "' ORDER BY PO, CONVERT(integer, LineItem)"
    End If
    
    Set dataPACKING = New ADODB.Recordset
    dataPACKING.Open Sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then Exit Sub
    With dataPACKING
        If .RecordCount > 0 Then
            current = ""
            Do While Not .EOF
                If current <> !PO Then
                    If POlist.Rows > 2 Then
                        POlist.AddItem ""
                    End If
                    current = !PO
                End If
                rowTEXT = "" + vbTab
                rowTEXT = rowTEXT + IIf(IsNull(!PO), "", !PO) + vbTab 'PO Number
                rowTEXT = rowTEXT + IIf(IsNull(!lineITEM), "", !lineITEM) + vbTab 'PO Line Item
                
                rowTEXT = rowTEXT + IIf(IsNull(!Commodity), "", !Commodity) + vbTab 'PO Commodity
                
                
                rowTEXT = rowTEXT + IIf(IsNull(!Description), "", !Description) + vbTab 'PO Description
                rowTEXT = rowTEXT + FormatNumber(!Quantity, 2) + vbTab 'Quantity Requested
                rowTEXT = rowTEXT + IIf(IsNull(!Unit), "", Trim(!Unit)) + vbTab 'Unit
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!delivered), 0, !delivered), 2) + vbTab 'Quantity Already Delivered
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!Shipped), 0, !Shipped), 2) + vbTab 'Quantity Already Shipped
                If !delivered <= !Shipped Then
                    rowTEXT = rowTEXT + "0.00" + vbTab 'Quantity To Ship
                Else
                    If IsNull(!delivered) Then
                        rowTEXT = rowTEXT + "0.00" + vbTab 'Quantity To Ship
                    Else
                        If !delivered > 0 Then
                            rowTEXT = rowTEXT + FormatNumber(!delivered - IIf(IsNull(!Shipped), 0, !Shipped), 2) + vbTab 'Quantity To Ship
                        End If
                    End If
                End If
                
                If packinglist = "*" Then
                    rowTEXT = rowTEXT + "0.00" + vbTab 'Quantity Being Shipped
                    rowTEXT = rowTEXT + "" + vbTab 'Box Number
                Else
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!BeingShipped), 0, !BeingShipped), 2) + vbTab 'Quantity Being Shipped
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!box) Or Not IsNumeric(!box), "0", !box), 0) + vbTab  'Box Number
                End If
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice), 0, !UnitPrice), 2) + vbTab 'Unit Price
                
                If packinglist = "*" Then
                    rowTEXT = rowTEXT + "0.00" 'Total Amount
                Else
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!TotalPrice), 0, !TotalPrice), 2) 'Total Amount
                End If
      'me.POlist.TextMatrix(11,13)    jawdat 2.5.02
                rowTEXT = rowTEXT & vbTab
                
                    If packinglist = "*" Then
                        rowTEXT = rowTEXT + "0.00" + vbTab 'Quantity Being Shipped
                        rowTEXT = rowTEXT + "" + vbTab 'Box Number
                    Else
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!BeingShipped), 0, !BeingShipped), 2) + vbTab 'Quantity Being Shipped
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!box) Or Not IsNumeric(!box), "0", !box), 0) + vbTab  'Box Number
                    End If
                    
       'me.POlist.TextMatrix(11,13)    jawdat 2.5.02
       
                POlist.AddItem rowTEXT
                If packinglist <> "*" Then
                    If Form = mdvisualization And (!Quantity <> !Quantity2) Then
                        POlist.row = POlist.Rows - 1
                        POlist.Col = 0
                        POlist = "u"
                        POlist.CellFontName = "Wingdings"
                        POlist.CellFontSize = 8
                        'POlist.AddItem "w" + vbTab + Chr(34) + vbTab + Chr(34) + vbTab + Chr(34) + vbTab + FormatNumber(IIf(IsNull(!Quantity2), 0, !Quantity2), 2) + vbTab + !unit2
                        POlist.AddItem "w" + vbTab + Chr(34) + vbTab + Chr(34) + vbTab + Chr(34) + vbTab + Chr(34) + vbTab + FormatNumber(IIf(IsNull(!Quantity2), 0, !Quantity2), 2) + vbTab + !unit2
                        POlist.row = POlist.Rows - 1
                        POlist.Col = 0
                        POlist.CellFontName = "Wingdings"
                        'For i = 1 To 13
                        For i = 1 To 14
                            'If i < 5 Then POlist.CellAlignment = 4
                            If i < 6 Then POlist.CellAlignment = 4
                            POlist.Col = i
                            POlist.BandExpandable(0) = True
                            POlist.CellBackColor = &HFFFFC0
                        Next
                    End If
                End If
                .MoveNext
            Loop
            POlist.RemoveItem (1)
        End If
    End With
End Sub


Sub getPRIORITYlist()
'On Error Resume Next
Dim Sql As String
Dim dataPAKING As New ADODB.Recordset

    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT pri_code, pri_desc FROM PRIORITY WHERE pri_actvflag=1 AND pri_npecode = '" + deIms.NameSpace + "' ORDER BY pri_desc"
    PriorityList.Rows = 0
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                PriorityList.AddItem !pri_desc + vbTab + !pri_code
                .MoveNext
            Loop
        End If
        PriorityList.row = 0
    End With
End Sub


Sub getRECIPIENTSlist()
    If Not IsNothing(rsReceptList) Then
        With rsReceptList
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    RecipientList.AddItem "" + vbTab + .Fields(0)
                    .MoveNext
                Loop
            End If
        End With
        If RecipientList.Rows > 2 And RecipientList.TextMatrix(1, 1) = "" Then RecipientList.RemoveItem 1
    End If
End Sub

Sub getSHIPPERlist()
On Error Resume Next
Dim Sql As String
Dim dataPAKING As New ADODB.Recordset
On Error GoTo errorTRACK

    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT shi_code, shi_name FROM SHIPPER WHERE shi_npecode = '" + deIms.NameSpace + "' " _
        & "AND shi_actvflag = 1 ORDER BY shi_name"
    ShipperList.Rows = 0
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                ShipperList.AddItem !shi_name + vbTab + !shi_code
                .MoveNext
            Loop
        End If
        ShipperList.row = 0
    End With
    Exit Sub
    
errorTRACK:
    If Err.number = -2147417848 Then
        Err.Clear
        Resume Next
    Else
        MsgBox Err.Description
        Exit Sub
    End If
End Sub

Sub getSHIPTOlist()
On Error Resume Next
Dim Sql As String
Dim dataPAKING As New ADODB.Recordset

    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT sht_code, sht_name FROM SHIPTO WHERE sht_npecode = '" + deIms.NameSpace + "' " _
        & "AND sht_actvflag = 1 ORDER BY sht_name"
    ShipToList.ColAlignment(0) = 0
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                ShipToList.AddItem !sht_name + vbTab + !sht_code
                .MoveNext
            Loop
        End If
        If ShipToList.Rows > 1 Then ShipToList.RemoveItem 0
        ShipToList.row = 0
    End With
End Sub

Sub getSOLDTOlist()
On Error Resume Next
Dim Sql As String
Dim dataPAKING As New ADODB.Recordset

    Set dataPAKING = New ADODB.Recordset
    Sql = "SELECT slt_code, slt_name FROM SOLDTO WHERE slt_npecode = '" + deIms.NameSpace + "' " _
        & "ORDER BY slt_name"
    SoldToList.Rows = 0
    With dataPAKING
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            Do While Not .EOF
                SoldToList.AddItem !slt_name + vbTab + !slt_code
                .MoveNext
            Loop
        End If
        SoldToList.row = 0
    End With
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

Sub gridONfocus(ByRef Grid As MSHFlexGrid)
Dim i, x As Integer
    With Grid
        If .Rows > 0 Then
            x = .Col
            For i = 0 To .Cols - 1
                .Col = i
                .CellBackColor = &H800000   'Blue
                .CellForeColor = &HFFFFFF   'White
            Next
            .Col = x
            .Tag = .row
        End If
    End With
End Sub

Sub lockDOCUMENT(locked As Boolean)
Dim i As Integer
    For i = 1 To 29
        If locked Then
            cell(i).locked = True
        Else
            cell(i).locked = False
        End If
    Next
    If locked Then
        txtClause.locked = True
        remarkBUTTON.Enabled = False
        Imsmail1.Enabled = False
        cmd_Add.Enabled = False
        cmd_Remove.Enabled = False
    Else
        txtClause.locked = False
        remarkBUTTON.Enabled = True
        Imsmail1.Enabled = True
        cmd_Add.Enabled = True
        cmd_Remove.Enabled = True
    End If
End Sub

Sub makeDETAILgrid() 'JCG 2008/6/21 inserting new col 3
Dim i
    With POlist
        .Clear
        .Rows = 2
        'For i = 0 To 12
        For i = 0 To 13 'new
            .ColWidth(i) = 780
            .ColAlignment(i) = 6
            .ColAlignmentFixed(i) = 4
        Next
        .ColAlignment(0) = 4
        .ColAlignment(1) = 0
        .ColAlignment(3) = 0 'new col
        '.ColAlignment(3) = 0
        .ColAlignment(4) = 0
        '.ColAlignment(5) = 0
        .ColAlignment(6) = 0
        .ColWidth(0) = 285
        .ColWidth(1) = 1100
        .ColWidth(2) = 400
        .ColWidth(3) = 1550 'new col
        '.ColWidth(3) = 3000 - 285
        .ColWidth(4) = 3000 - 285
        '.ColWidth(5) = 690
        .ColWidth(6) = 650
        '.ColWidth(10) = 400
        .ColWidth(11) = 400
        '.ColWidth(11) = 1040
        .ColWidth(12) = 1040
        '.ColWidth(12) = 1040
        .ColWidth(13) = 1040
        .RowHeight(0) = 500
        .RowHeightMin = 240
        .WordWrap = True
        .row = 0
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .TextMatrix(0, 0) = "®"
        .TextMatrix(0, 1) = "Transaction"
        .TextMatrix(0, 2) = "Line #"
        .TextMatrix(0, 3) = "Commodity #" 'new col
        '.TextMatrix(0, 3) = "Description"
        .TextMatrix(0, 4) = "Description"
        '.TextMatrix(0, 4) = "Quantity"
        .TextMatrix(0, 5) = "Quantity"
        '.TextMatrix(0, 5) = "Unit"
        .TextMatrix(0, 6) = "Unit"
        '.TextMatrix(0, 6) = "Already Delivered"
        .TextMatrix(0, 7) = "Already Delivered"
        '.TextMatrix(0, 7) = "Already Shipped"
        .TextMatrix(0, 8) = "Already Shipped"
        '.TextMatrix(0, 8) = "To Ship"
        .TextMatrix(0, 9) = "To Ship"
        '.TextMatrix(0, 9) = "Being Shipped"
        .TextMatrix(0, 10) = "Being Shipped"
        '.TextMatrix(0, 10) = "Box #"
        .TextMatrix(0, 11) = "Box #"
        '.TextMatrix(0, 11) = "Unit Price"
        .TextMatrix(0, 12) = "Unit Price"
        '.TextMatrix(0, 12) = "Total Amount"
        .TextMatrix(0, 13) = "Total Amount"
        
        If Form <> mdvisualization Then
            '.Cols = 15
            .Cols = 16
            'Invisible columns
            '.ColWidth(13) = 0
            .ColWidth(14) = 0
            '.TextMatrix(0, 13) = "Real Height"
            .TextMatrix(0, 14) = "Real Height"
            '.ColWidth(14) = 0
            .ColWidth(15) = 0
            '.TextMatrix(0, 14) = "Old value"
            .TextMatrix(0, 15) = "Old value"
        End If
        .row = 1
        .Col = 1
    End With
End Sub

Function PLexists() As Boolean
Dim Sql, packinglist As String
Dim dataPAKING  As New ADODB.Recordset
    On Error Resume Next
    PLexists = True
    packinglist = Trim(cell(0))
    Sql = "SELECT pl_manfnumb from PACKINGLIST WHERE pl_npecode = '" + deIms.NameSpace + "' " _
        & "AND pl_manfnumb = '" + packinglist + "'"
    Set dataPAKING = New ADODB.Recordset
    dataPAKING.Open Sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then Exit Function
    If dataPAKING.RecordCount < 1 Then
        PLexists = False
    End If
End Function

Sub showDTPicker1(cellNUMBER As Integer)
    With cell(cellNUMBER)
        DTPicker1.Tag = cellNUMBER
        DTPicker1.Top = .Top
        DTPicker1.Height = .Height
        DTPicker1.Left = .Left
        DTPicker1.Width = .Width
        DTPicker1.ZOrder
        DTPicker1.Visible = True
        DTPicker1.SetFocus
    End With
End Sub

Sub showLIST(ByRef Grid As MSHFlexGrid)
    With Grid
        .ZOrder
        .Visible = True
    End With
End Sub

Sub showTEXTline(column As Integer) 'JCG 2008/6/21 inserting new col 3
Dim positionX, positionY, i As Integer
    With POlist
        If .TextMatrix(.row, 0) <> "" Then
            .Col = column
            positionX = .Left + 30 + .ColPos(.Col)
            positionY = .Top + .RowPos(.row) + 30
            TextLINE.Text = .Text
            TextLINE.Left = positionX
            TextLINE.Width = .ColWidth(.Col) - 10
            TextLINE.Top = positionY
            TextLINE.Height = .RowHeight(.row) - 10
            TextLINE.Tag = .row
            TextLINE.SelStart = 0
            TextLINE.SelLength = Len(TextLINE.Text)
            TextLINE.Visible = True
            TextLINE.SetFocus
        End If
    End With
End Sub

Sub textBOX(ByVal mainCONTROL As MSHFlexGrid, standard As Boolean)
Dim h, i As Integer
Dim box As textBOX

    With mainCONTROL
        box.Height = .RowHeight(i)
        box.Height = box.Height + 10
        If .row = 0 And .FixedRows > 0 Then
            box.Top = .Top
            box.Height = box.Height - 80
        Else
            If standard Then
                box.Left = .Left + .ColWidth(0)
                h = 20
                For i = 0 To .row - 1
                    h = h + .RowHeight(i)
                Next
                box.Top = h + .Top - 30
                box.Width = .ColWidth(1)
            Else
                box.Left = .Left
                box.Top = .Top - box.Height
                box.Width = .ColWidth(0)
            End If
        End If
        box.Visible = True
        box.Text = .Text
        If standard Then
            box.SetFocus
        End If
    End With
End Sub



Private Sub cell_Change(Index As Integer)
    If Me.ActiveControl.Name = "cell" Then
        With cell(Index)
            Select Case Index
                Case 0
                    If Form = mdvisualization Then
                        If cell(Index) <> "" Then Call alphaSEARCH(cell(Index), PackingListList, 0)
                    End If
                Case 3
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), PriorityList, 0)
                Case 7
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), destinationList(4), 0)
                Case 11
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), destinationList(0), 0)
                Case 12
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), destinationList(1), 0)
                Case 14
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), destinationList(2), 0)
                Case 15
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), destinationList(3), 0)
                Case 20
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), ShipperList, 0)
                Case 21
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), ShipToList, 0)
                Case 22
                    If Form <> mdvisualization Then Call alphaSEARCH(cell(Index), SoldToList, 0)
            End Select
        End With
    End If
End Sub

Private Sub cell_Click(Index As Integer)
    focusHERE = Index
    cell(Index).SetFocus
    cell(Index).Refresh
    Select Case Index
        Case 0
            If Form = mdvisualization Then
                If PackingListList.Rows > 0 Then Call showLIST(PackingListList)
            Else
                PackingListList.Visible = False
            End If
    End Select
End Sub

Private Sub cell_GotFocus(Index As Integer)
    With cell(Index)
        If Not .locked Then
            If focusHERE = -1 And nextCELL > -1 Then
                cell(nextCELL).SetFocus
                nextCELL = -1
            End If
            .BackColor = vbYellow
            .Appearance = 1
            .Refresh
            .Tag = .Text
            Select Case Index
                Case 0
                    If Form = mdvisualization Then
                        If PackingListList.Visible Then
                            PackingListList.Visible = False
                        Else
                            If PackingListList.Rows > 0 Then
                                PackingListList.Visible = True
                                PackingListList.ZOrder
                            End If
                        End If
                    End If
                Case 1, 2, 8, 9
                    If IsDate(cell(Index)) Then DTPicker1.value = CDate(cell(Index))
                    If Form <> mdvisualization Then
                        If DTPicker1.Enabled Then Call showDTPicker1(Index)
                    End If
                Case 3
                    Call showLIST(PriorityList)
                Case 4, 5, 6
                Case 7
                    Call showLIST(destinationList(4))
                Case 10, 13
                Case 11
                    Call showLIST(destinationList(0))
                Case 12
                    Call showLIST(destinationList(1))
                Case 14
                    Call showLIST(destinationList(2))
                Case 15
                    Call showLIST(destinationList(3))
                Case 16 To 19
                    
                Case 20
                    Call showLIST(ShipperList)
                Case 21
                    Call showLIST(ShipToList)
                Case 22
                    Call showLIST(SoldToList)
                Case 23 To 29
            End Select
            focusHERE = -1
        End If
    End With
End Sub

Private Sub cell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    With cell(Index)
        If Not .locked Then
            activeARROWS = False
            If Index = 0 And Form = mdvisualization Then activeARROWS = True
            If Index > 0 And Form <> mdvisualization Then activeARROWS = True
            If activeARROWS Then
                Select Case KeyCode
                    Case 40
                        Call arrowKEYS("down", Index)
                    Case 38
                        Call arrowKEYS("up", Index)
                End Select
            End If
        End If
    End With
End Sub
Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
    With cell(Index)
        If Not .locked Then
            Select Case KeyAscii
                Case 13
                    KeyAscii = 0
                    If cell(Index) <> "" Then
                        Select Case Index
                            Case 0
                                Select Case Form
                                    Case mdvisualization
                                        If PackingListList.Visible Then
                                            If cell(0) = Left(Trim(PackingListList), Len(cell(0))) Then
                                                cell(0) = PackingListList
                                            End If
                                            PackingListList.Visible = False
                                        End If
                                        Call getPACKINGLIST
                                    Case mdCreation
                                        If PLexists Then
                                            msg1 = translator.Trans("M00282")
                                            MsgBox IIf(msg1 = "", "Packing List Entered Number is already exist", msg1)
                                            Exit Sub
                                        Else
                                            cell(1).SetFocus
                                        End If
                                End Select
                                PackingListList.Visible = False
                            Case 3
                                cell(Index) = PriorityList.Text
                                PriorityList.Visible = False
                            Case 7
                                cell(Index) = destinationList(4).Text
                                destinationList(4).Visible = False
                            Case 11
                                cell(Index) = destinationList(0).Text
                                destinationList(0).Visible = False
                            Case 12
                                cell(Index) = destinationList(1).Text
                                destinationList(1).Visible = False
                            Case 14
                                cell(Index) = destinationList(2).Text
                                destinationList(2).Visible = False
                            Case 15
                                cell(Index) = destinationList(3).Text
                                destinationList(3).Visible = False
                            Case 16, 17
                                Call cell_Validate(Index, True)
                            Case 20
                                cell(Index) = ShipperList.Text
                                ShipperList.Visible = False
                            Case 21
                                cell(Index) = ShipToList.Text
                                ShipToList.Visible = False
                            Case 22
                                cell(Index) = SoldToList.Text
                                SoldToList.Visible = False
                        End Select
                    End If
                    If Index < 29 Then
                        If Index > 0 Then cell(Index + 1).SetFocus
                    Else
                        txtClause.SetFocus
                    End If
                Case 27
                    .Text = cell(Index).Tag
                    Select Case Index
                        Case 0
                            PackingListList.Visible = False
                        Case 3
                            PriorityList.Visible = False
                        Case 7
                            destinationList(4).Visible = False
                        Case 11
                            destinationList(0).Visible = False
                        Case 12
                            destinationList(1).Visible = False
                        Case 14
                            destinationList(2).Visible = False
                        Case 15
                            destinationList(3).Visible = False
                        Case 20
                            ShipperList.Visible = False
                        Case 21
                            ShipToList.Visible = False
                        Case 22
                            SoldToList.Visible = False
                    End Select
            End Select
        End If
    End With
End Sub

Private Sub cell_LostFocus(Index As Integer)
On Error Resume Next
    lastCELL = Index - 1
    With cell(Index)
        If Not .locked Then
            .BackColor = txtClause.BackColor
            Select Case Index
                Case 0
                    Select Case Form
                        Case mdvisualization
                            'PackingListList.Visible = False
                            If cell(0) = Right(packinglistLABEL, Len(cell(0))) Then Exit Sub
                            If Not PackingListList.Visible Then
                                Exit Sub
                                'Call getPACKINGLIST
                            End If
                        Case mdCreation
                            If PLexists Then
                                msg1 = translator.Trans("M00282")
                                MsgBox IIf(msg1 = "", "Packing List Entered Number is already exist", msg1)
                                cell(0) = ""
                                cell(0).SetFocus
                            End If
                        Exit Sub
                    End Select
                Case 1, 2, 8, 9
                    .Text = .Tag
                    If Me.ActiveControl.Name <> "DTPicker1" Then
                        DTPicker1.Visible = False
                    End If
                Case 3
                    If Me.ActiveControl.Name <> "PriorityList" Then
                        PriorityList.Visible = False
                    End If
                Case 7
                    If Me.ActiveControl.Name <> "destinationList" Then
                       destinationList(4).Visible = False
                    End If
                Case 11
                    If Me.ActiveControl.Name <> "destinationList" Then
                        destinationList(0).Visible = False
                    End If
                Case 12
                    If Me.ActiveControl.Name <> "destinationList" Then
                        destinationList(1).Visible = False
                    End If
                Case 14
                    If Me.ActiveControl.Name <> "destinationList" Then
                        destinationList(2).Visible = False
                    End If
                Case 15
                    If Me.ActiveControl.Name <> "destinationList" Then
                        destinationList(3).Visible = False
                    End If
                Case 20
                    If Me.ActiveControl.Name <> "ShipperList" Then
                        ShipperList.Visible = False
                    End If
                Case 21
                    If Me.ActiveControl.Name <> "ShipToList" Then
                        ShipToList.Visible = False
                    End If
                Case 22
                    If Me.ActiveControl.Name <> "SoldToList" Then
                        SoldToList.Visible = False
                    End If
                Case 29
                    txtClause.SetFocus
            End Select
        End If
    End With
End Sub



Public Sub cell_Validate(Index As Integer, Cancel As Boolean)
  If Form <> mdvisualization Then
'      If Form = mdCreation Then
    
        With cell(Index)
            If Not .locked Then
                If .Text <> "" Then
                    If Form = mdCreation Then
                        Select Case Index
                            Case 0
                            Case 1, 2, 8, 9
                                If Not IsDate(.Text) Then
                                    .Text = ""
                                End If
                            Case 3
                                If .Text <> PriorityList Then
                                    .Text = ""
                                End If
                            Case 4
                            Case 5
                            Case 6
                            Case 7
                                If .Text <> destinationList(4) Then
                                    .Text = ""
                                End If
                            Case 10
                            Case 11
                                If .Text <> destinationList(0) Then
                                    .Text = ""
                                End If
                            Case 12
                                If .Text <> destinationList(1) Then
                                    .Text = ""
                                End If
                            Case 13
                            Case 14
                                If .Text <> destinationList(2) Then
                                    .Text = ""
                                End If
                            Case 15
                                If .Text <> destinationList(3) Then
                                    .Text = ""
                                End If
                            Case 16, 17, 18
                                If IsNumeric(.Text) Then
                                    If CDbl(.Text) > 0 Then
                                        .Text = FormatNumber(CDbl(.Text), 2)
                                    Else
                                        .Text = ""
                                    End If
                                Else
                                    .Text = ""
                                End If
                            Case 19
                            Case 20
                                If .Text <> ShipperList Then
                                    .Text = ""
                                End If
                            Case 21
                                If .Text <> ShipToList Then
                                    .Text = ""
                                End If
                            Case 22
                                If .Text <> SoldToList Then
                                    .Text = ""
                                End If
                            Case 23
                            Case 24
                            Case 25
                            Case 26
                            Case 27
                            Case 28
                            Case 29
                        End Select
                    End If
                End If
            End If
        End With
    End If






'    End If
End Sub

Private Sub cmd_Add_Click()
    Imsmail1.AddCurrentRecipient
End Sub

Private Sub cmd_Remove_Click()
'On Error Resume Next
'    If RecipientList.row > 0 Then
'        If RecipientList.TextMatrix(RecipientList.row, 1) <> "" Then
'            rsReceptList.MoveFirst
'            rsReceptList.Find "Recipients = '" + RecipientList.TextMatrix(RecipientList.row, 1)
'            If Not rsReceptList.EOF Then
'                rsReceptList.Delete
'                rsReceptList.Update
'            End If
'        End If
        If RecipientList.row > 0 Then
            If RecipientList.Rows > 2 Then
                RecipientList.RemoveItem (RecipientList.row)
            Else
                RecipientList.TextMatrix(1, 1) = ""
            End If
        End If
        'Call getRECIPIENTSlist
    'End If
    If Err Then Err.Clear
End Sub


Private Sub Command1_Click() 'JCG 2008/6/21 inserting new col 3
Dim showAll As Boolean
Dim i As Integer
    If Command1.Caption = "&Show Only Selection" Then
        Command1.Caption = "&Show All Records"
        showAll = False
    Else
        Command1.Caption = "&Show Only Selection"
        showAll = True
    End If
    
    With POlist
        .Col = 0
        If showAll Then
            .RowHeightMin = 240
        Else
            For i = 1 To .Rows - 1
                If .RowHeight(i) > 240 Then
                    .TextMatrix(i, 13) = .RowHeight(i)
                End If
            Next
            .RowHeightMin = 0
            .RowHeight(-1) = 0
            .RowHeight(0) = 500
            For i = .Rows - 1 To 1 Step -1
                If Not showAll Then
                    If .TextMatrix(i, 0) <> "" Then
                        .RowHeight(i) = 240
                    End If
                End If
            Next
        End If
        For i = 1 To .Rows - 1
            'If IsNumeric(.TextMatrix(i, 13)) Then
            If IsNumeric(.TextMatrix(i, 14)) Then
                'If val(.TextMatrix(i, 13)) > 240 Then .RowHeight(i) = val(.TextMatrix(i, 13))
                If val(.TextMatrix(i, 14)) > 240 Then .RowHeight(i) = val(.TextMatrix(i, 14))
            End If
        Next
    End With
End Sub

Private Sub lblpodname_Click()

End Sub

Private Sub NavBar1_OnEditClick() 'JCG 2008/6/21 inserting new col 3
Dim packinglist, Sql, i, inPOINT, t
Dim datax As ADODB.Recordset
Dim commandx As ADODB.Command
Dim gotIT As Boolean
Dim gotPO As Boolean
Dim addTHIS As Boolean
Dim gotITEM As Boolean
Dim rowTEXT
Dim qty As Double
    Call ChangeMode(mdModification)
    Call begining
    cell(0).SetFocus
    settingUP = True
    cell(11) = cell(11) + " "
    cell(11) = Trim(cell(11))
    destinationList(0).Visible = False
    cell(12) = cell(12) + " "
    cell(12) = Trim(cell(12))
    destinationList(1).Visible = False
    cell(14) = cell(14) + " "
    cell(14) = Trim(cell(14))
    destinationList(2).Visible = False
    cell(15) = cell(15) + " "
    cell(15) = Trim(cell(15))
    destinationList(3).Visible = False
    settingUP = False
    
    Call lockDOCUMENT(False)
    With NavBar1
        .NewEnabled = False
        .CancelEnabled = True
        .EditEnabled = False
        .EMailEnabled = False
        .SaveEnabled = True
        SaveEnabled = True
    End With
    
    Call getLINEitems("*")
    
    packinglist = Trim(cell(0))
    Sql = "SELECT * from Packing_List_Details WHERE NameSpace = '" + deIms.NameSpace + "' " _
        & "AND PackingList = '" + packinglist + "' ORDER BY PO, CONVERT(integer, LineItem)"
    Set commandx = New ADODB.Command
    commandx.ActiveConnection = deIms.cnIms
    commandx.CommandType = adCmdStoredProc
    commandx.CommandText = "POs_for_Editing_Packing_List"
    commandx.parameters("@NameSpace").value = deIms.NameSpace
    commandx.parameters("@Packing").value = packinglist
    Set datax = commandx.Execute
    
    With POlist
        '.Cols = 17
        .Cols = 18
        '.ColWidth(15) = 0
        .ColWidth(16) = 0
        '.ColWidth(16) = 0
        .ColWidth(17) = 0
        If datax.RecordCount > 0 Then
            Do While Not datax.EOF
                gotPO = True
                gotITEM = False
                addTHIS = False
                inPOINT = 0
                i = 1
                Do While True
                    If .TextMatrix(i, 1) <> "" Then
                        If .TextMatrix(i, 1) = datax!PO Then
                            gotPO = True
                            Select Case .TextMatrix(i, 2)
                                Case datax!lineITEM
                                    gotITEM = True
                                    inPOINT = i
                                    Exit Do
                                Case Is > datax!lineITEM
                                    inPOINT = i
                                Case Is < datax!lineITEM
                                    inPOINT = i + 1
                            End Select
                        Else
                            If .TextMatrix(i, 1) > datax!PO Then
                                If inPOINT = 0 Then
                                    inPOINT = i
                                End If
                                Exit Do
                            End If
                        End If
                    End If
                    i = i + 1
                    If i >= .Rows Then Exit Do
                Loop
                If inPOINT = 0 Then inPOINT = .Rows - 1
                If gotPO Then
                    If Not gotITEM Then
                        addTHIS = True
                    End If
                    i = inPOINT
                Else
                    addTHIS = True
                    i = -1
                End If
                
                If gotPO And gotITEM = False Then
                    rowTEXT = "" + vbTab
                    rowTEXT = rowTEXT + IIf(IsNull(datax!PO), "", datax!PO) + vbTab 'PO Number
                    rowTEXT = rowTEXT + IIf(IsNull(datax!lineITEM), "", datax!lineITEM) + vbTab 'PO Line Item
                    
                   'PO Commodity NEW col
                    
                    rowTEXT = rowTEXT + IIf(IsNull(datax!Description), "", datax!Description) + vbTab 'PO Description
                    rowTEXT = rowTEXT + FormatNumber(datax!Quantity, 2) + vbTab 'Quantity Requested
                    rowTEXT = rowTEXT + IIf(IsNull(datax!Unit), "", Trim(datax!Unit)) + vbTab 'Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!delivered), 0, datax!delivered), 2) + vbTab 'Quantity Already Delivered
                    qty = IIf(IsNull(datax!Shipped), 0, datax!Shipped) - IIf(IsNull(datax!qty), 0, datax!qty)
                    If qty < 0 Then qty = 0
                    rowTEXT = rowTEXT + FormatNumber(qty, 2) + vbTab 'Quantity Already Shipped
                    
                    If datax!delivered > 0 Then
                        qty = datax!delivered - IIf(IsNull(datax!Shipped), 0, datax!Shipped) + datax!qty
                    Else
                        qty = datax!qty
                    End If
                    rowTEXT = rowTEXT + FormatNumber(qty, 2) + vbTab 'Quantity To Ship
                    
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!qty), 0, datax!qty), 2) + vbTab 'Quantity Being Shipped
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!box) Or Not IsNumeric(datax!box), "0", datax!box), 0) + vbTab  'Box Number
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!UnitPrice), 0, datax!UnitPrice), 2) + vbTab 'Unit Price
                    
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!TotalPrice), 0, datax!TotalPrice), 2) 'Total Amount
                    rowTEXT = rowTEXT & vbTab
                    
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!qty), 0, datax!qty), 2) + vbTab 'Quantity Being Shipped
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!box) Or Not IsNumeric(datax!box), "0", Format(datax!box, "0")), 0) + vbTab  'Box Number
                    rowTEXT = rowTEXT + "original" + vbTab 'Flag for the orignal transaction
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(datax!Line), 0, datax!Line)) 'Original line #
                
                    .AddItem rowTEXT, inPOINT
                Else
                    '.TextMatrix(i, 7) = Format(CDbl(.TextMatrix(i, 7)) - datax!qty, "0.00") 'Already Shipped
                    .TextMatrix(i, 8) = Format(CDbl(.TextMatrix(i, 8)) - datax!qty, "0.00") 'Already Shipped
                    '.TextMatrix(i, 8) = Format(CDbl(.TextMatrix(i, 8)) + datax!qty, "0.00") 'To Ship
                    .TextMatrix(i, 9) = Format(CDbl(.TextMatrix(i, 9)) + datax!qty, "0.00") 'To Ship
                    '.TextMatrix(i, 9) = Format(datax!qty, "0.00")  'Being Shipped
                    .TextMatrix(i, 10) = Format(datax!qty, "0.00")  'Being Shipped
                    '.TextMatrix(i, 10) = FormatNumber(IIf(IsNull(datax!box) Or Not IsNumeric(datax!box), "0", Format(datax!box, "0")), 0)  'Box Number
                    .TextMatrix(i, 11) = FormatNumber(IIf(IsNull(datax!box) Or Not IsNumeric(datax!box), "0", Format(datax!box, "0")), 0)  'Box Number
                
                    '.TextMatrix(i, 12) = FormatNumber(CDbl(.TextMatrix(i, 9)) * CDbl(.TextMatrix(i, 11)), 2)
                    .TextMatrix(i, 13) = FormatNumber(CDbl(.TextMatrix(i, 10)) * CDbl(.TextMatrix(i, 12)), 2)
                    '.TextMatrix(i, 13) = .TextMatrix(i, 9) 'Originally Being Shipped
                    .TextMatrix(i, 14) = .TextMatrix(i, 10) 'Originally Being Shipped
                    '.TextMatrix(i, 14) = .TextMatrix(i, 10) 'Originally Box
                    .TextMatrix(i, 15) = .TextMatrix(i, 11) 'Originally Box
                    '.TextMatrix(i, 15) = "original" 'Flag for the orignal transaction
                    .TextMatrix(i, 16) = "original" 'Flag for the orignal transaction
                    '.TextMatrix(i, 16) = FormatNumber(IIf(IsNull(datax!Line), 0, datax!Line)) 'Original line #
                    .TextMatrix(i, 17) = FormatNumber(IIf(IsNull(datax!Line), 0, datax!Line)) 'Original line #
                End If
                
                If i > 0 Then
                    .row = i
                Else
                    .row = .Rows - 1
                End If
                Call markROW
                If .TextMatrix(i + 1, 1) <> .TextMatrix(i, 1) Then
                    If .TextMatrix(i + 1, 1) <> "" Then
                        .AddItem "", i + 1
                    End If
                End If
                datax.MoveNext
            Loop
        End If
        .row = 1
        .Col = 1
    End With
    
    Command1.Enabled = True
    Command3.Enabled = False
End Sub

Private Sub remarkBUTTON_Click()
Screen.MousePointer = 11
    If st Is Nothing Then Set st = New frm_ShipTerms
    st.Show
    st.txt_Description.SetFocus
Screen.MousePointer = 0
End Sub

Private Sub st_Completed(Cancelled As Boolean, Terms As String)
On Error Resume Next

    If Not Cancelled Then
        txtClause.Text = txtClause.Text & Terms
        txtClause.SelStart = Len(txtClause)
    End If
    
    Set st = Nothing
End Sub

Private Sub Command3_Click()
Dim answer
    Call begining
    answer = MsgBox("If you want to refresh your lineitems, you are going to lose your current selection.  Do you want to continue?", vbYesNo + vbDefaultButton2)
    If answer = vbYes Then
        Call getLINEitems("*")
    End If
End Sub

Private Sub destinationList_Click(Index As Integer)
    'Datamember used only as a string field to put the attached cell
    cell(val(destinationList(Index).DataMember) + 1).SetFocus
End Sub

Private Sub destinationList_EnterCell(Index As Integer)
    With destinationList(Index)
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then
            If Me.ActiveControl.Index = .Index Then
                cell(val(destinationList(Index).DataMember)) = .Text
            End If
        End If
    End With
End Sub


Private Sub destinationList_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            nextCELL = 12
        Case 1
            nextCELL = 13
        Case 2
            nextCELL = 15
        Case 3
            nextCELL = 16
        Case 4
            nextCELL = 8
    End Select
    Call gridONfocus(destinationList(Index))
End Sub

Private Sub destinationList_KeyPress(Index As Integer, KeyAscii As Integer)
    With destinationList(Index)
        Select Case KeyAscii
            Case 13
                cell(val(.DataMember) + 1).SetFocus
            Case 27
                .Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(val(.DataMember)) = Chr(KeyAscii)
                Call alphaSEARCH(cell(val(.DataMember)), destinationList(Index), 0)
                .Tag = ""
                cell(val(.DataMember)).SetFocus
                cell(val(.DataMember)).SelStart = Len(cell(val(.DataMember)))
                cell(val(.DataMember)).SelLength = 0
        End Select
    End With
End Sub

Private Sub destinationList_LeaveCell(Index As Integer)
    With destinationList(Index)
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub destinationList_LostFocus(Index As Integer)
    With destinationList(Index)
        If IsNumeric(.Tag) Then cell(val(destinationList(Index).DataMember)) = .Text
        .Visible = False
    End With
End Sub

Public Sub DTPicker1_DropDown()
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    With DTPicker1
        Select Case KeyCode
            Case 16
                If Shift = 1 Then
                    nextCELL = val(DTPicker1.Tag) - 1
                End If
            Case 13
                cell(val(.Tag)).Text = Format(.value, "MMMM/dd/yyyy")
                cell(val(.Tag) + 1).SetFocus
        End Select
    End With
End Sub

Private Sub DTPicker1_LostFocus()
Dim indexCELL As Integer
    With DTPicker1
        .Visible = False
        If IsNumeric(.Tag) Then
            cell(val(.Tag)).Text = Format(.value, "MMMM/dd/yyyy")
            If indexCELL = val(.Tag) Then
                indexCELL = val(.Tag) + 1
            Else
                indexCELL = val(.Tag)
            End If
            .Visible = False
            If focusHERE > 0 Then
                cell(focusHERE).SetFocus
                focusHERE = -1
            Else
                If nextCELL > -1 Then
                    cell(nextCELL).SetFocus
                Else
                    If lastCELL = indexCELL - 1 Then
                        If (lastCELL + 2) > 0 Then
                            cell(lastCELL + 2).SetFocus
                        Else
                            cell(indexCELL).SetFocus
                        End If
                    Else
                        cell(indexCELL).SetFocus
                        lastCELL = indexCELL
                    End If
                End If
            End If
        End If
        .value = Now
        lastCELL = val(DTPicker1.Tag)
    End With
End Sub


Private Sub Form_Activate()
Dim rights, Sql
Dim datax As New ADODB.Recordset
    
    If OPENEDFORM Then Exit Sub
    OPENEDFORM = True
    Screen.MousePointer = 11
    Me.Refresh
    With RecipientList
        .ColWidth(0) = 300
        .ColWidth(1) = 9095
        .Rows = 2
        .Clear
        .TextMatrix(0, 1) = "Recipient List"
    End With

    DoEvents
    Set datax = New ADODB.Recordset
    Sql = "SELECT dis_mail FROM DISTRIBUTION WHERE dis_id = 'F' AND dis_npecode = '" + deIms.NameSpace + "'"
    datax.Open Sql, deIms.cnIms, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        Do While Not datax.EOF
            RecipientList.AddItem "" + vbTab + datax!dis_mail
            datax.MoveNext
        Loop
        If RecipientList.Rows > 2 And RecipientList.TextMatrix(1, 1) = "" Then RecipientList.RemoveItem (1)
    End If
    
    With NavBar1
        If Form = mdvisualization Then
            .EditVisible = True
            .Width = 2835
            .SaveEnabled = False
            .CancelEnabled = False
            If PLexists Then
                .PrintEnabled = True
                .EMailEnabled = True
            End If
        End If
    End With
    cell(0).SetFocus
    Screen.MousePointer = 0
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim rights, Sql
Dim datax As New ADODB.Recordset
    OPENEDFORM = False
    Call translator.Translate_Forms("frmPackingList")
    Form = mdvisualization
    Screen.MousePointer = 11
    SSTab1.Tab = 0
    Call lockDOCUMENT(True)
    Call getPackingListList
    Imsmail1.NameSpace = deIms.NameSpace
    Imsmail1.SetActiveConnection deIms.cnIms
    Imsmail1.Language = Language
    rights = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    SaveEnabled = rights
    
    NavBar1.NewEnabled = SaveEnabled
    cell(0).Enabled = True
    frmPackingList.Caption = frmPackingList.Caption + " - " + frmPackingList.Tag
    frmPackingList.Left = Int((MDI_IMS.Width - frmPackingList.Width) / 2)
    frmPackingList.Top = Int((MDI_IMS.Height - frmPackingList.Height) / 2) - 500
    lastCELL = -1
    focusHERE = -1
    nextCELL = -1

    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.number, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim closing
    If Form <> mdvisualization Then
        closing = MsgBox("Do you really want to close and lose your last record?", vbYesNo)
        If closing = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(recLocked, deIms.cnIms, CurrentUser, rowguid, grid1, dbtablename, , grid2) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
    Set imsLock = Nothing
End Sub

Private Sub IMSMail1_OnAddClick(ByVal address As String)
On Error Resume Next

    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        rsReceptList.Open
    End If
    
    If (InStr(1, address, "@") > 0) = 0 Then
        address = UCase(address)
    End If
    
    If Not IsInList(address, "Recipients", rsReceptList) Then
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))
    End If
    If InStr(UCase(address), "INTERNET") > 0 Then address = Mid(address, InStr(UCase(address), "INTERNET") + 8)
    If InStrRev(address, "!") > 0 Then address = Mid(address, InStrRev(address, "!") + 1)
    RecipientList.AddItem "" + vbTab + address
    If RecipientList.Rows > 2 And RecipientList.TextMatrix(1, 1) = "" Then RecipientList.RemoveItem (1)
    'Call getRECIPIENTSlist
    
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
'Dim grid2 As Boolean
grid2 = True
grid1 = False
Call imsLock.Unlock_Row(recLocked, deIms.cnIms, CurrentUser, rowguid, grid1, dbtablename, POValue, grid2) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
 Set imsLock = Nothing
    
    
    
End Sub

Private Sub NavBar1_BeforeSaveClick() 'JCG 2008/6/21 inserting new col 3
Dim wrong As Boolean
Dim i, n
Dim position As Integer
Dim wasSAVED As Boolean

Screen.MousePointer = 11
    
    'Revision for Header
    NavBar1.SaveEnabled = True
    wrong = False
    For i = 0 To 12
        If cell(i) = "" Then
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00016")
            MsgBox IIf(msg1 = "", "Cannot be left empty", msg1)
            cell(i).SetFocus
            If Me.ActiveControl = "DTPicker1" Then
                DTPicker1.Enabled = False
            End If
            cell(i).SetFocus
            DTPicker1.Enabled = True
            Exit Sub
        End If
    Next
    For i = 16 To 17
        position = i
        If IsNumeric(cell(i)) Then
            If CDbl(cell(i)) <= 0 Then
                wrong = True
                Exit For
            End If
        Else
            wrong = True
            Exit For
        End If
    Next
    If cell(18) = "" Then
        Screen.MousePointer = 0
        msg1 = translator.Trans("M00016")
        MsgBox IIf(msg1 = "", "Cannot be left empty", msg1)
        cell(18).SetFocus
        Exit Sub
    End If

    If cell(13) + cell(14) + cell(15) <> "" Then
        For i = 13 To 15
            If cell(i) = "" Then
                position = i
            End If
        Next
    End If
    
    If wrong Then
        Screen.MousePointer = 0
        msg1 = translator.Trans("M00122")
        MsgBox IIf(msg1 = "", "Invalid Value", msg1)
        cell(position).SetFocus
        Exit Sub
    End If
    
    'Revision for Recipients
    With RecipientList
        For i = 1 To .Rows - 1
            If Len(.TextMatrix(i, 1)) > 60 Then
                Screen.MousePointer = 0
                msg1 = translator.Trans("M00342")
                MsgBox IIf(msg1 = "", "Value is too large", msg1)
                SSTab1.Tab = 1
                .row = i
                .SetFocus
                Exit Sub
            End If
        Next
    End With
    
    'Revision for Details
    wrong = True
    position = 0
    For i = 1 To POlist.Rows - 1
    
        If POlist.TextMatrix(i, 0) <> "" Then
            'If IsNumeric(POlist.TextMatrix(i, 9)) Then
            If IsNumeric(POlist.TextMatrix(i, 10)) Then
                If Form = mdModification Then
                    n = -0.01
                Else
                    n = 0
                End If
                'If CDbl(POlist.TextMatrix(i, 9)) > n Then
                If CDbl(POlist.TextMatrix(i, 10)) > n Then
                    'If CDbl(POlist.TextMatrix(i, 8)) < CDbl(POlist.TextMatrix(i, 9)) Then
                    If CDbl(POlist.TextMatrix(i, 9)) < CDbl(POlist.TextMatrix(i, 10)) Then
                        wrong = True
                        position = i
                        readyFORsave = False
                        Exit For
                    Else
                        wrong = False
                        readyFORsave = True
                    End If
                Else
                    readyFORsave = False
                    wrong = True
                    position = i
                    Exit For
                End If
            End If
        End If
    Next
    If wrong Then
        SSTab1.Tab = 2
        If position > 0 Then
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00122")
            MsgBox IIf(msg1 = "", "Invalid Value", msg1)
            POlist.row = position
            'POlist.Col = 9
            POlist.Col = 10
            POlist.SetFocus
        Else
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00707")
            MsgBox IIf(msg1 = "", "You have to select at least one line item.", msg1)
        End If
    Else
        Call SAVE
        Call ChangeMode(mdvisualization)
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        NavBar1.NewEnabled = True
        wasSAVED = True
    End If
        
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(recLocked, deIms.cnIms, CurrentUser, rowguid, grid1, dbtablename, , grid2) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
       
    Set imsLock = Nothing
    If wasSAVED Then
        Call getPACKINGLIST
    End If
    Screen.MousePointer = 0
End Sub

Private Sub NavBar1_GotFocus()
    If nextCELL > -1 Then
        cell(nextCELL).SetFocus
    End If
    If Me.ActiveControl.Name = "cell" Then
        NavBar1.Tag = ""
    End If
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As String
    msg1 = translator.Trans("M00706")
    msg2 = translator.Trans("L00441")
    response = MsgBox(IIf(msg1 = "", "Are you sure you want to cancel changes?", msg1), vbYesNo, IIf(msg2 = "", "Cancel", msg2))
    If response = vbYes Then
        With NavBar1
            cell(0) = ""
            Call ChangeMode(mdvisualization)
            If SSTab1.Tab > 0 Then SSTab1.Tab = 0
            Call lockDOCUMENT(True)
            Call clearDOCUMENT
            .NewEnabled = SaveEnabled
            .CancelEnabled = False
            .SaveEnabled = False
            .PrintEnabled = False
            .EditEnabled = False
            .CancelEnabled = False
            .EMailEnabled = True
        End With
    End If
    
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    'Dim grid2 As Boolean
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(recLocked, deIms.cnIms, CurrentUser, rowguid, grid1, dbtablename, , grid2) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
    Set imsLock = Nothing
End Sub

Private Sub NavBar1_OnCloseClick()
Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(recLocked, deIms.cnIms, CurrentUser, rowguid, grid1, dbtablename, POValue, grid2)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
    Set imsLock = Nothing
    Unload Me
End Sub


Private Sub NavBar1_OnEMailClick()

Dim ParamsForRPTI(1) As String

Dim rptinf As RPTIFileInfo

Dim ParamsForCrystalReports(1) As String

Dim subject As String

Dim FieldName As String

Dim Message As String

Dim attention As String

On Error Resume Next

If rsReceptList Is Nothing Then Exit Sub
                
    ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    
    ParamsForCrystalReports(1) = "manifestnumb;" + cell(0) + ";true"
    
    ParamsForRPTI(0) = "namespace=" + deIms.NameSpace
    
    ParamsForRPTI(1) = "manifestnumb=" + cell(0)
    
    FieldName = "Recipients"
    
    subject = "Packing list #" + cell(0)
    
        attention = "Attention Please "
    
    If ConnInfo.EmailClient = Outlook Then
    
        'Call sendOutlookEmailandFax("packinglist.rpt", "Packing List", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, attention)  MM 030209  EFCR11
        Call sendOutlookEmailandFax(Report_EmailFax_PackingManifest_name, "Packing List", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, attention)
    
    ElseIf ConnInfo.EmailClient = ATT Then
    
        Call SendAttFaxAndEmail("packinglist.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rsReceptList, subject, Message, FieldName)

    ElseIf ConnInfo.EmailClient = Unknown Then
    
        MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

    End If




'''''Dim Params(1) As String
'''''Dim i As Integer
'''''Dim Attachments(0) As String
'''''Dim subject As String
'''''Dim reports(0) As String
'''''Dim Recepients() As String
'''''Dim attention As String
'''''Dim rptinfo As RPTIFileInfo
'''''Dim FileName As String
'''''Dim IFile As IMSFile
'''''Screen.MousePointer = 11
'''''
'''''On Error GoTo errMESSAGE
'''''    Set IFile = New IMSFile
'''''     BeforePrint
'''''     MDI_IMS.CrystalReport1.PrintFileType = crptRTF
'''''
'''''    If rsReceptList.RecordCount > 0 Then
'''''
'''''        subject = "Packing list #" + cell(0)
'''''        reports(0) = "packinglist.rpt"
'''''
'''''        attention = "Attention Please "
'''''
'''''        With rptinfo
'''''            Params(0) = "namespace=" + deIms.NameSpace
'''''            Params(1) = "manifestnumb=" + cell(0)
'''''            .ReportFileName = ReportPath & "packinglist.rpt"
'''''            Call translator.Translate_Reports("packinglist.rpt")
'''''            .Parameters = Params
'''''        End With
'''''
'''''        Call WriteRPTIFile(rptinfo, Left(MDI_IMS.CrystalReport1.ReportFileName, Len(MDI_IMS.CrystalReport1.ReportFileName) - 3) + "rtf")
'''''
'''''        Recepients = ToArrayFromRecordset(rsReceptList)
'''''
'''''        Attachments(0) = "Report-" & "Packinglist" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
'''''        FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
'''''        If IFile.FileExists(FileName) Then IFile.DeleteFile (FileName)
'''''        If Not FileExists(FileName) Then MDI_IMS.SaveReport FileName, crptRTF
'''''
'''''        Call WriteParameterFiles(Recepients, "", Attachments, subject, attention)
'''''    Else
'''''        MsgBox "No Recipients to Send", , "Imswin"
'''''    End If
'''''    Screen.MousePointer = 0
'''''
'''''
'''''errMESSAGE:
'''''    If Err.number <> 0 Then
'''''        MsgBox Err.Description
'''''    End If
End Sub

Private Sub NavBar1_OnNewClick()
Dim response As String
    Screen.MousePointer = 11
    With NavBar1
        Call clearDOCUMENT
        Call ChangeMode(mdCreation)
        Call begining
        Call lockDOCUMENT(True)
        Screen.MousePointer = 0
        .EditEnabled = False
        .NewEnabled = False
        .CancelEnabled = True
        .SaveEnabled = SaveEnabled
        .PrintEnabled = False
        .EMailEnabled = False
        cell(0) = ""
        cell(1) = Format(Now, "MMMM/dd/yyyy")
        cell(2) = Format(Now, "MMMM/dd/yyyy")
        cell(8) = Format(Now, "MMMM/dd/yyyy")
        cell(9) = Format(Now, "MMMM/dd/yyyy")
        
        Screen.MousePointer = 11
        Call getLINEitems("*")
        Call lockDOCUMENT(False)
        If RecipientList.TextMatrix(1, 1) <> "" Then
            msg1 = translator.Trans("M00716")
            msg2 = translator.Trans("L00241")
            response = MsgBox(IIf(msg1 = "", "Do you want to use the current Recipient List?", msg1), vbYesNo, IIf(msg2 = "", "Recipient List", msg2))
            If response = vbNo Then
                RecipientList.Rows = 2
                RecipientList.TextMatrix(1, 1) = ""
            End If
        End If
    End With
    Screen.MousePointer = 0
    cell(0).SetFocus
    
    'jawdat
    NavBar1.EditEnabled = False
End Sub

Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next
    Select Case FMode
        Case mdCreation
            lblStatu.ForeColor = vbRed
            msg1 = translator.Trans("L00125")
            lblStatu.Caption = IIf(msg1 = "", "Creation", msg1)
            lblStatu.Tag = "Creation"
        Case mdvisualization
            lblStatu.ForeColor = vbGreen
            msg1 = translator.Trans("L00092") 'J added
            lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
            lblStatu.Tag = "Visualization"
         
         'jawdat
         Case mdModification
            lblStatu.ForeColor = vbBlue
            msg1 = translator.Trans("L00126") 'J added
            lblStatu.Caption = IIf(msg1 = "", "Modification", msg1) 'J modified
            lblStatu.Tag = "Modification"
                                    
    End Select
    ChangeMode = True
    Form = FMode
End Function

Private Sub NavBar1_OnPrintClick()
On Error Resume Next
Screen.MousePointer = 11
    With MDI_IMS.CrystalReport1
        Call BeforePrint
        msg1 = translator.Trans("L00213")
        .WindowTitle = IIf(msg1 = "", "Packing List", msg1)
        .Action = 1
    End With
Screen.MousePointer = 0
End Sub

Sub SAVE() 'JCG 2008/6/21 inserting new col 3
Dim header As New ADODB.Recordset
Dim details As New ADODB.Recordset
Dim details2 As New ADODB.Recordset
Dim Recipients As New ADODB.Recordset
Dim addREC As Boolean

Dim PoItem As New ADODB.Recordset
Dim POread As New ADODB.Recordset
Dim POdata As New ADODB.Recordset

Dim i As Integer
Dim lineN As Integer
Dim Sql, shippingSTATUS As String
Dim Q As Double
Dim baseQ As Double
Dim newDETAIL As Boolean
Dim touchDETAIL As Boolean
Dim itemSTATUS
On Error GoTo errTRACK
    Screen.MousePointer = 11
    
    FrmShowApproving.Label2 = "Saving Packing List"
    FrmShowApproving.Show
    FrmShowApproving.Refresh
    
    If readyFORsave Then
        'Header routine
        msg1 = translator.Trans("M00708")
        MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Header", msg1)
        deIms.cnIms.BeginTrans
        Set header = New ADODB.Recordset
        With header
            Select Case Form
                Case mdCreation
                    Sql = "SELECT * FROM PACKINGLIST WHERE pl_npecode = '" + deIms.NameSpace + "' AND pl_manfnumb = ''"
                    header.Open Sql, deIms.cnIms, adOpenDynamic, adLockOptimistic
                    .AddNew
                Case mdModification
                    Sql = "SELECT * FROM PACKINGLIST WHERE pl_npecode = '" + deIms.NameSpace + "' AND pl_manfnumb = '" + Trim(cell(0)) + "'"
                    header.Open Sql, deIms.cnIms, adOpenDynamic, adLockOptimistic
                    If header.RecordCount = 0 Then
                        Screen.MousePointer = 0
                        MsgBox "Transaction failed"
                        Exit Sub
                    End If
            End Select
            !pl_creauser = CurrentUser
            !pl_npecode = deIms.NameSpace
            
            !pl_manfnumb = Trim(cell(0))
            !pl_docudate = CDate(cell(1))
            !pl_shipdate = CDate(cell(2))
            !pl_priocode = PriorityList.TextMatrix(PriorityList.row, 1)
            !pl_awbnumb = cell(4)
            !pl_viacarr = cell(5)
            !pl_shipterm = cell(6)
            !pl_dest = destinationList(4).TextMatrix(destinationList(4).row, 1)
            !pl_etd = CDate(cell(8))
            !pl_eta = CDate(cell(9))
            !pl_fig1 = cell(10)
            !pl_from1 = destinationList(0).TextMatrix(destinationList(0).row, 1)
            !pl_to1 = destinationList(1).TextMatrix(destinationList(1).row, 1)
            
            If cell(13) <> "" Then
                !pl_fig2 = cell(13)
                If cell(14) = "" Then
                Else
                    !pl_from2 = destinationList(2).TextMatrix(destinationList(2).row, 1)
                End If
                If cell(15) = "" Then
                Else
                    !pl_to2 = destinationList(3).TextMatrix(destinationList(3).row, 1)
                End If
            End If
            !pl_numbpiec = CInt(cell(16))
            !pl_grosweig = CDbl(cell(17))
            !pl_totavolu = cell(18)
            !pl_hawbnumb = IIf(cell(19) = "", Null, cell(19))
            If ShipperList.Rows > 0 Then
                !pl_shipcode = IIf(cell(20) = "", "", ShipperList.TextMatrix(ShipperList.row, 1))
            End If
            If ShipToList.Rows > 0 Then
                !pl_shtocode = IIf(cell(21) = "", Null, ShipToList.TextMatrix(ShipToList.row, 1))
            End If
            If SoldToList.Rows > 0 Then
                !pl_sltcode = IIf(cell(22) = "", Null, SoldToList.TextMatrix(SoldToList.row, 1))
            End If
            !pl_shiprefe = IIf(cell(23) = "", Null, cell(23))
            !pl_forwrefe = IIf(cell(24) = "", Null, cell(24))
            !pl_custrefe = IIf(cell(25) = "", Null, cell(25))
            !pl_mark1 = IIf(cell(26) = "", Null, cell(26))
            !pl_mark2 = IIf(cell(27) = "", Null, cell(27))
            !pl_mark3 = IIf(cell(28) = "", Null, cell(28))
            !pl_mark4 = IIf(cell(29) = "", Null, cell(29))
            !pl_remk = Format(txtClause)
            .Update
        End With
        
        'Recipients routine
        msg1 = translator.Trans("M00709")
        MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Recipients", msg1)
        Set Recipients = New ADODB.Recordset
        Sql = "SELECT * FROM PACKINGREC WHERE plrc_npecode = '" + deIms.NameSpace + "' AND plrc_manfnumb = ''"
        Recipients.Open Sql, deIms.cnIms, adOpenKeyset, adLockOptimistic
        
        For i = 1 To RecipientList.Rows - 1
            With Recipients
                If RecipientList.TextMatrix(i, 1) <> "" Then
                    Select Case Form
                        Case mdCreation
                            addREC = True
                        Case mdModification
                            Set Recipients = New ADODB.Recordset
                            Sql = "SELECT * FROM PACKINGREC WHERE plrc_npecode = '" + deIms.NameSpace _
                                & "' AND plrc_manfnumb = '" + Trim(cell(0)) + "' AND plrc_rec = '" + RecipientList.TextMatrix(i, 1) + "'"
                            Recipients.Open Sql, deIms.cnIms, adOpenKeyset, adLockOptimistic
                            If Recipients.RecordCount = 0 Then
                                addREC = True
                            End If
                    End Select
                    
                    If addREC Then
                        Sql = "INSERT INTO PACKINGREC (plrc_manfnumb, plrc_npecode, plrc_recpmumb, plrc_rec,plrc_creauser) VALUES ('" _
                            & Trim(cell(0)) + "', '" + deIms.NameSpace + "', " + Format(i) + ", '" + RecipientList.TextMatrix(i, 1) + "', '" + CurrentUser + "')"
                        deIms.cnIms.Execute Sql
'                        .AddNew
'                        !plrc_manfnumb = Trim(cell(0))
'                        !plrc_npecode = deIms.NameSpace
'                        !plrc_recpmumb = i
'                        !plrc_rec = RecipientList.TextMatrix(i, 1)
'                        !plrc_creauser = CurrentUser
'                        .Update
                    End If
                    addREC = False
               End If
            End With
        Next
        
        'Details routine
        msg1 = translator.Trans("M00710")
        MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Details", msg1)
        Set details = New ADODB.Recordset
        With details
            Select Case Form
                Case mdCreation
                    Sql = "SELECT * FROM PACKINGDETL WHERE pld_npecode = '" + deIms.NameSpace + "' AND pld_ponum = ''"
                    details.Open Sql, deIms.cnIms, adOpenKeyset, adLockOptimistic
                Case mdModification
                    Sql = "SELECT * FROM PACKINGDETL WHERE pld_npecode = '" + deIms.NameSpace + "' AND pld_manfnumb = '" + Trim(cell(0)) + "'"
                    details.Open Sql, deIms.cnIms, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        lineN = .RecordCount + 1
                    Else
                        lineN = 1
                    End If
            End Select
            For i = 1 To POlist.Rows - 1
                newDETAIL = False
                touchDETAIL = False
                If POlist.TextMatrix(i, 0) <> "" Then
                    If POlist.TextMatrix(i, 1) <> "" Then
                        Select Case Form
                            Case mdCreation
                                newDETAIL = POlist.TextMatrix(i, 0) <> "" And POlist.TextMatrix(i, 0) <> "w"
                                touchDETAIL = newDETAIL
                            Case mdModification
                                Set details2 = New ADODB.Recordset
                                Sql = "SELECT * FROM PACKINGDETL WHERE pld_npecode = '" + deIms.NameSpace _
                                    & "' AND pld_manfnumb = '" + Trim(cell(0)) + "' AND pld_manfsrl = " + IIf(POlist.TextMatrix(i, 17) = "", "0", POlist.TextMatrix(i, 17))
                                    '& "' AND pld_manfnumb = '" + Trim(cell(0)) + "' AND pld_manfsrl = " + IIf(POlist.TextMatrix(i, 16) = "", "0", POlist.TextMatrix(i, 16))
                                details2.Open Sql, deIms.cnIms, adOpenKeyset, adLockPessimistic
                                If details2.RecordCount = 0 Then
                                    Screen.MousePointer = 0
                                    touchDETAIL = False
                                End If
                                newDETAIL = POlist.TextMatrix(i, 0) <> "" And POlist.TextMatrix(i, 0) = "Æ"
                                touchDETAIL = POlist.TextMatrix(i, 0) <> "" Or POlist.TextMatrix(i, 0) <> "w"
                        End Select
                    End If
                End If
                
                If touchDETAIL Then
                    If newDETAIL Then
                        .AddNew
                        !pld_modiuser = CurrentUser
                        !pld_manfnumb = Trim(cell(0))
                        !pld_npecode = deIms.NameSpace
                        !pld_manfsrl = lineN
                        lineN = lineN + 1
                        !pld_ponum = POlist.TextMatrix(i, 1)
                        !pld_liitnumb = POlist.TextMatrix(i, 2)
                        '!pld_boxnumb = IIf(POlist.TextMatrix(i, 10) = "", Null, POlist.TextMatrix(i, 10))
                        !pld_boxnumb = IIf(POlist.TextMatrix(i, 11) = "", Null, POlist.TextMatrix(i, 11))
                        '!pld_reqdqty = CDbl(POlist.TextMatrix(i, 9))
                        !pld_reqdqty = CDbl(POlist.TextMatrix(i, 10))
                        '!pld_unitpric = CDbl(POlist.TextMatrix(i, 11))
                        !pld_unitpric = CDbl(POlist.TextMatrix(i, 12))
                        '!pld_totaprice = CDbl(POlist.TextMatrix(i, 12))
                        !pld_totaprice = CDbl(POlist.TextMatrix(i, 13))
                    Else
                        'details2!pld_boxnumb = IIf(POlist.TextMatrix(i, 10) = "", Null, POlist.TextMatrix(i, 10))
                        details2!pld_boxnumb = IIf(POlist.TextMatrix(i, 11) = "", Null, POlist.TextMatrix(i, 11))
                        'details2!pld_reqdqty = CDbl(POlist.TextMatrix(i, 9))
                        details2!pld_reqdqty = CDbl(POlist.TextMatrix(i, 10))
                        'details2!pld_unitpric = CDbl(POlist.TextMatrix(i, 11))
                        details2!pld_unitpric = CDbl(POlist.TextMatrix(i, 12))
                        'details2!pld_totaprice = CDbl(POlist.TextMatrix(i, 12))
                        details2!pld_totaprice = CDbl(POlist.TextMatrix(i, 13))
                        details2.Update
                    End If
                                       
                    'Update PO Shipping Status
                    If touchDETAIL Then
                        msg1 = translator.Trans("M00711")
                        MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Updating PO Line Item", msg1)
                        Set PoItem = New ADODB.Recordset
                        Sql = "SELECT poi_primreqdqty, poi_qtyship, poi_qtydlvd, poi_stasship,poi_tbs from POITEM " _
                            & "WHERE poi_npecode = '" + deIms.NameSpace + "' " _
                            & "AND poi_ponumb = '" + Trim(POlist.TextMatrix(i, 1)) + "' " _
                            & "AND poi_liitnumb = '" + Trim(POlist.TextMatrix(i, 2)) + "'"
                        PoItem.Open Sql, deIms.cnIms, adOpenForwardOnly, adLockReadOnly
                        If PoItem.RecordCount > 0 Then
                            If newDETAIL Then
                                'Q = IIf(IsNumeric(PoItem!poi_qtyship), PoItem!poi_qtyship, 0) + CDbl(POlist.TextMatrix(i, 9))
                                Q = IIf(IsNumeric(PoItem!poi_qtyship), PoItem!poi_qtyship, 0) + CDbl(POlist.TextMatrix(i, 10))
                            Else
                                'Q = IIf(IsNumeric(PoItem!poi_qtyship), PoItem!poi_qtyship, 0) - CDbl(POlist.TextMatrix(i, 13)) + CDbl(POlist.TextMatrix(i, 9))
                                Q = IIf(IsNumeric(PoItem!poi_qtyship), PoItem!poi_qtyship, 0) - CDbl(POlist.TextMatrix(i, 14)) + CDbl(POlist.TextMatrix(i, 10))
                            End If
                                 
                            baseQ = PoItem!poi_primreqdqty
                            'If baseQ < PoItem!poi_qtydlvd Then baseQ = PoItem!poi_qtydlvd
                            Select Case Q
                                Case Is >= baseQ
                                    itemSTATUS = "SC"
                                Case Is < baseQ
                                    itemSTATUS = "SP"
                            End Select
                            
                            Sql = "UPDATE POITEM " _
                                & "SET poi_stasship = '" + itemSTATUS + "' ," _
                                & "poi_tbs = 1, poi_qtyship = " + Format(Q) _
                                & "WHERE poi_npecode = '" + deIms.NameSpace + "' " _
                                & "AND poi_ponumb = '" + Trim(POlist.TextMatrix(i, 1)) + "' " _
                                & "AND poi_liitnumb = '" + Trim(POlist.TextMatrix(i, 2)) + "'"
                            deIms.cnIms.Execute Sql
                    
                            'Check PO General Shipping Status
                            msg1 = translator.Trans("M00712")
                            MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Checking Shipping Status", msg1)
                            Set POread = New ADODB.Recordset
                            Sql = "DECLARE @PO AS VARCHAR(15), @nameSPACE AS VARCHAR(5) " _
                                & "SET @PO = '" + Trim(POlist.TextMatrix(i, 1)) + "' " _
                                & "SET @nameSPACE = '" + deIms.NameSpace + "' " _
                                & "SELECT poi_ponumb, " _
                                & "(SELECT COUNT(*) FROM POITEM WHERE poi_npecode = @nameSPACE and " _
                                    & "poi_stasship = 'SC' AND POI_PONUMB = @PO ) AS SC, " _
                                & "(SELECT COUNT(*) FROM POITEM WHERE poi_npecode = @nameSPACE and " _
                                    & "poi_stasship = 'SP' AND POI_PONUMB = @PO) AS SP, " _
                                & "(SELECT COUNT(*) FROM POITEM WHERE poi_npecode = @nameSPACE and " _
                                    & "poi_stasship = 'NS'AND POI_PONUMB = @PO) AS NS " _
                                & "From POITEM WHERE poi_npecode = @nameSPACE and POI_PONUMB = @PO GROUP BY poi_ponumb"
                            POread.Open Sql, deIms.cnIms, adOpenStatic, adLockReadOnly
                            If POread.RecordCount > 0 Then
                                If POread!NS = 0 Then
                                    If POread!SP = 0 Then
                                        shippingSTATUS = "SC"
                                    Else
                                        shippingSTATUS = "SP"
                                    End If
                                Else
                                    If POread!SP = 0 Then
                                        If POread!SC = 0 Then
                                            shippingSTATUS = "NS"
                                        Else
                                            shippingSTATUS = "SP"
                                        End If
                                    Else
                                        shippingSTATUS = "SP"
                                    End If
                                End If
                                
                                'Update PO General Shipping Status
                                msg1 = translator.Trans("M00713")
                                MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Updating PO Header", msg1)
                                Set POdata = New ADODB.Recordset
                                Sql = "SELECT * from PO WHERE po_npecode = '" + deIms.NameSpace + "' " _
                                    & "AND po_ponumb = '" + Trim(POlist.TextMatrix(i, 1)) + "'"
                                POdata.Open Sql, deIms.cnIms, adOpenDynamic, adLockOptimistic
                                If POdata.RecordCount > 0 Then
                                    POdata!po_stasship = shippingSTATUS
                                    POdata!PO_tbs = 1
                                    POdata.Update
                                Else
                                    msg1 = translator.Trans("M00078")
                                    MsgBox IIf(msg1 = "", "Error occurred", msg1)
                                    deIms.cnIms.RollbackTrans
                                    Exit Sub
                                End If
                            Else
                                msg1 = translator.Trans("M00078")
                                MsgBox IIf(msg1 = "", "Error occurred", msg1)
                                deIms.cnIms.RollbackTrans
                                Exit Sub
                            End If
                        Else
                            msg1 = translator.Trans("M00078")
                            MsgBox IIf(msg1 = "", "Error occurred", msg1)
                            deIms.cnIms.RollbackTrans
                            Exit Sub
                        End If
                    End If
                End If
            Next
            msg1 = translator.Trans("M00714")
            .UpdateBatch
        End With
    End If
    msg1 = translator.Trans("M00715")
    MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Commiting Transaction", msg1)
    deIms.cnIms.CommitTrans
    MDI_IMS.StatusBar1.Panels(1).Text = ""
    Screen.MousePointer = 0
    msg1 = translator.Trans("M00306")
    MsgBox IIf(msg1 = "", "Insert into Packing List is completed successfully", msg1)
    Screen.MousePointer = 11
    Call lockDOCUMENT(True)
    Call clearDOCUMENT
    Call getPackingListList
    Unload FrmShowApproving
    Screen.MousePointer = 0
    Exit Sub
    
errTRACK:
    If Err.number = -2147217873 Then
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub PackingListList_Click()
    Screen.MousePointer = 11
    Select Case Form
        Case mdvisualization
            PackingListList.Tag = PackingListList.row
            cell(0) = Trim(PackingListList)
            Call getPACKINGLIST
        Case mdCreation
            cell(1).SetFocus
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub PackingListList_KeyPress(KeyAscii As Integer)
    With PackingListList
        Select Case KeyAscii
            Case 13
                Select Case Form
                    Case mdvisualization
                        cell(0) = .Text
                        Call getPACKINGLIST
                    Case mdCreation
                        cell(1).SetFocus
                End Select
            Case 27
                PackingListList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(0) = Chr(KeyAscii)
                Call alphaSEARCH(cell(0), PackingListList, 0)
                .Tag = ""
                cell(0).SetFocus
                cell(0).SelStart = Len(cell(0))
                cell(0).SelLength = 0
        End Select
    End With
End Sub

Private Sub POlist_Click() 'JCG 2008/6/21 inserting new col 3
Dim i, currentCOL As Integer
  If Form <> mdvisualization Then
        With POlist
            If .TextMatrix(.row, 1) <> "" Then
                selectionSTART = .MouseRow
                POlist.Tag = .row
                currentCOL = .MouseCol
                If .MouseCol = 0 Then
                    Select Case currentCOL
                        Case 0, 1, 2, 3 'new col 3 added
                            If multiMARKED Then
                                multiMARKED = False
                            Else
                                Call markROW
                            End If
                        'Case 9, 10
                        Case 10, 11
                    End Select
                    '.ColSel = 12
                    .ColSel = 13
                Else
                    Select Case currentCOL
                        Case 0, 1, 2, 3 'new col 3 added
                        'Case 9, 10
                        Case 10, 11
                            Call showTEXTline(currentCOL)
                    End Select
                End If
            End If
        End With
    End If
End Sub

Private Sub PackingListList_EnterCell()
    With PackingListList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
    End With
End Sub

Private Sub PackingListList_GotFocus()
    Call gridONfocus(PackingListList)
End Sub

Private Sub PackingListList_LeaveCell()
    With PackingListList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub


Private Sub PackingListList_LostFocus()
    With PackingListList
        cell(0).Text = Trim(.Text)
    End With
End Sub

Public Sub PackingListList_Validate(Cancel As Boolean)
    cell(0) = Trim(PackingListList)
End Sub

Private Sub POlist_DblClick() 'JCG 2008/6/21 inserting new col 3
Dim w, i, Col As Integer
    w = 0
    With POlist
        If .MouseRow = 0 Then
            Col = .MouseCol
            If Col <> 0 Then
                'If (Col <> 3) Then
                If (Col <> 4) Then
                    For i = .topROW To .Rows - 1
                        If w < TextWidth(.TextMatrix(i, Col)) Then
                            w = TextWidth(.TextMatrix(i, Col))
                        End If
                        If Not .RowIsVisible(i) Then Exit For
                    Next
                    If w > TextWidth(.TextMatrix(0, Col)) Then
                        .ColWidth(Col) = w + 120
                    Else
                        .ColWidth(Col) = TextWidth(.TextMatrix(0, Col)) + 120
                    End If
                End If
            End If
        Else
            If Form <> mdvisualization Then
                If .TextMatrix(.row, 1) <> "" Then
                    Call markROW
                End If
            End If
        End If
        '.ColSel = 12
        .ColSel = 13
    End With
End Sub

Private Sub POlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'JCG 2008/6/21 inserting new col 3
On Error Resume Next
Dim i
    With POlist
        If Shift = 1 Then
            multiMARKED = True
            If selectionSTART > 0 Then
                For i = selectionSTART To .MouseRow
                    If .TextMatrix(i, 1) <> "" Then
                        .row = i
                        .Col = 0
                        .CellFontName = "Wingdings 3"
                        .CellFontSize = 10
                        .Text = "Æ"
                        '.TextMatrix(.row, 9) = .TextMatrix(.row, 8)
                        .TextMatrix(.row, 10) = .TextMatrix(.row, 9)
                        '.TextMatrix(.row, 10) = ""
                        .TextMatrix(.row, 11) = ""
                        '.TextMatrix(.row, 12) = FormatNumber(CDbl(.TextMatrix(.row, 9)) * CDbl(.TextMatrix(.row, 11)), 2)
                        .TextMatrix(.row, 13) = FormatNumber(CDbl(.TextMatrix(.row, 10)) * CDbl(.TextMatrix(.row, 12)), 2)
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub POlist_Scroll()
   If Form <> mdvisualization Then TextLINE.Visible = False
End Sub

Private Sub POlist_SelChange()
    With POlist
       If Form <> mdvisualization Then
            If .TextMatrix(.row, 1) <> "" Then
                If .RowHeight(POlist.row) > 240 Then
                    .TextMatrix(POlist.row, 13) = .RowHeight(POlist.row)
                End If
            End If
        End If
    End With
End Sub

Private Sub PriorityList_Click()
    cell(4).SetFocus
End Sub

Private Sub PriorityList_EnterCell()
    With PriorityList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(3) = .Text
    End With
End Sub


Private Sub PriorityList_GotFocus()
    nextCELL = 4
    Call gridONfocus(PriorityList)
End Sub

Private Sub PriorityList_KeyPress(KeyAscii As Integer)
    With PriorityList
        Select Case KeyAscii
            Case 13
                cell(4).SetFocus
            Case 27
                PriorityList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(3) = Chr(KeyAscii)
                Call alphaSEARCH(cell(3), PriorityList, 0)
                .Tag = ""
                cell(3).SetFocus
                cell(3).SelStart = Len(cell(3))
                cell(3).SelLength = 0
        End Select
    End With
End Sub

Private Sub PriorityList_LeaveCell()
    With PriorityList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub PriorityList_LostFocus()
    With PriorityList
        cell(3).Text = .Text
        .Visible = False
    End With
End Sub

Private Sub PriorityList_Validate(Cancel As Boolean)
    cell(3) = PriorityList
    PriorityList.Visible = False
End Sub

Private Sub txtClause_LostFocus()
    'If Me.ActiveControl.Name <> "cell" Then cell(0).SetFocus
End Sub

Private Sub ShipperList_Click()
    cell(21).SetFocus
End Sub

Private Sub ShipperList_EnterCell()
    With ShipperList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(20) = .Text
    End With
End Sub


Private Sub ShipperList_GotFocus()
    nextCELL = 21
    Call gridONfocus(ShipperList)
End Sub

Private Sub ShipperList_KeyPress(KeyAscii As Integer)
    With ShipperList
        Select Case KeyAscii
            Case 13
                cell(21).SetFocus
            Case 27
                ShipperList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(20) = Chr(KeyAscii)
                Call alphaSEARCH(cell(20), ShipperList, 0)
                .Tag = ""
                cell(20).SetFocus
                cell(20).SelStart = Len(cell(20))
                cell(20).SelLength = 0
        End Select
    End With
End Sub

Private Sub ShipperList_LeaveCell()
    With ShipperList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub ShipperList_LostFocus()
    With ShipperList
        cell(20).Text = .Text
        .Visible = False
    End With
End Sub

Private Sub ShipperList_Validate(Cancel As Boolean)
    cell(20) = ShipperList.Text
    ShipperList.Visible = False
End Sub

Private Sub ShipToList_Click()
    cell(22).SetFocus
End Sub


Private Sub ShipToList_EnterCell()
    With ShipToList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(21) = .Text
    End With
End Sub


Private Sub ShipToList_GotFocus()
    nextCELL = 22
    Call gridONfocus(ShipToList)
End Sub

Private Sub ShipToList_KeyPress(KeyAscii As Integer)
    With ShipToList
        Select Case KeyAscii
            Case 13
                cell(22).SetFocus
            Case 27
                ShipToList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(21) = Chr(KeyAscii)
                Call alphaSEARCH(cell(21), ShipToList, 0)
                .Tag = ""
                cell(21).SetFocus
                cell(21).SelStart = Len(cell(21))
                cell(21).SelLength = 0
        End Select
    End With
End Sub

Private Sub ShipToList_LeaveCell()
    With ShipToList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub ShipToList_LostFocus()
    With ShipToList
        cell(21).Text = .Text
        .Visible = False
    End With
End Sub

Private Sub ShipToList_Validate(Cancel As Boolean)
    cell(21) = ShipToList.Text
    ShipToList.Visible = False
End Sub

Private Sub SoldToList_Click()
    cell(23).SetFocus
End Sub


Private Sub SoldToList_EnterCell()
    With SoldToList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(22) = .Text
    End With
End Sub


Private Sub SoldToList_GotFocus()
    nextCELL = 23
    Call gridONfocus(SoldToList)
End Sub

Private Sub SoldToList_KeyPress(KeyAscii As Integer)
    With SoldToList
        Select Case KeyAscii
            Case 13
                cell(23).SetFocus
            Case 27
                SoldToList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(22) = Chr(KeyAscii)
                Call alphaSEARCH(cell(22), SoldToList, 0)
                .Tag = ""
                cell(22).SetFocus
                cell(22).SelStart = Len(cell(22))
                cell(22).SelLength = 0
        End Select
    End With
End Sub

Private Sub SoldToList_LeaveCell()
    With SoldToList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub SoldToList_LostFocus()
    With SoldToList
        cell(22).Text = .Text
        .Visible = False
    End With
End Sub

Private Sub SoldToList_Validate(Cancel As Boolean)
    cell(22) = SoldToList.Text
    SoldToList.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case PreviousTab
        Case 0
            If cell(0) = "" Then
                SSTab1.Tab = 0
            Else
                packinglistLABEL = "Packing List # " + cell(0)
                Select Case Form
                    Case mdModification, mdCreation
                        Command1.Enabled = True
                    Case Else
                        Command1.Enabled = False
                End Select
            End If
    End Select
    With NavBar1
        If SSTab1.Tab = 0 Then
            If Form = mdvisualization Then
                .NewEnabled = SaveEnabled
            Else
                .SaveEnabled = SaveEnabled
            End If
        Else
            .NewEnabled = False
            .SaveEnabled = False
        End If
    End With
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
            Call TextLINE_Validate(False)
        End If
    End With
End Sub

Public Sub TextLINE_Validate(Cancel As Boolean) 'JCG 2008/6/21 inserting new col 3
    If Form = mdCreation Or mdModification Then
        With TextLINE
            'If POlist.Col > 8 Then
            If POlist.Col > 9 Then
                If IsNumeric(.Text) Then
                    If CDbl(.Text) = 0 Then
                        'If POlist.TextMatrix(val(.Tag), 15) <> "" Then POlist.TextMatrix(val(.Tag), 9) = "0.00"
                        If POlist.TextMatrix(val(.Tag), 16) <> "" Then POlist.TextMatrix(val(.Tag), 10) = "0.00"
                    Else
                        'If CDbl(POlist.TextMatrix(val(.Tag), 8)) >= CDbl(.Text) Then
                        If CDbl(POlist.TextMatrix(val(.Tag), 9)) >= CDbl(.Text) Then
                            POlist.TextMatrix(val(.Tag), POlist.Col) = FormatNumber(.Text, 2)
                            'POlist.TextMatrix(POlist.row, 12) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 9)) * CDbl(POlist.TextMatrix(POlist.row, 11)), 2)
                            POlist.TextMatrix(POlist.row, 13) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 10)) * CDbl(POlist.TextMatrix(POlist.row, 12)), 2)
                            .Tag = ""
                            .Text = ""
                            .Visible = False
                            Exit Sub
                        End If
                    End If
                End If
                'If POlist.TextMatrix(val(.Tag), 15) = "" Then
                If POlist.TextMatrix(val(.Tag), 16) = "" Then
                    If .Text <> "" Then
                        msg1 = translator.Trans("M00122")
                        MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                        TextLINE = ""
                    End If
                End If
            Else
                'If POlist.Col > 8 Then
                If POlist.Col > 9 Then
                    POlist.TextMatrix(val(.Tag), POlist.Col) = IIf(Len(.Text) > 2, Left(.Text, 2), .Text)
                    .Tag = ""
                    .Text = ""
                    .Visible = False
                End If
            End If
        End With
    Else         '    jawdat 2.5.02
        Call Update_PackingList
    End If
End Sub
 Public Function Update_PackingList() 'JCG 2008/6/21 inserting new col 3

 
    With TextLINE
        'If POlist.Col = 9 Then
        If POlist.Col = 10 Then
            If IsNumeric(.Text) Then
                If CDbl(.Text) > 0 Then
   'POlist.TextMatrix(POlist.row, 9) = .Text
   POlist.TextMatrix(POlist.row, 10) = .Text


'1. already shipped = already shipped - being shipped

'POlist.TextMatrix(POlist.row, 7) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 6) - POlist.TextMatrix(POlist.row, 9)))
POlist.TextMatrix(POlist.row, 8) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 7) - POlist.TextMatrix(POlist.row, 10)))

'2. to ship = already delivered -  already shipped

'POlist.TextMatrix(POlist.row, 8) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 6) - POlist.TextMatrix(POlist.row, 7)))
POlist.TextMatrix(POlist.row, 9) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 7) - POlist.TextMatrix(POlist.row, 8)))

              
POlist.TextMatrix(val(.Tag), POlist.Col) = FormatNumber(.Text, 2)
'POlist.TextMatrix(POlist.row, 12) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 9)) * CDbl(POlist.TextMatrix(POlist.row, 11)), 2)
POlist.TextMatrix(POlist.row, 13) = FormatNumber(CDbl(POlist.TextMatrix(POlist.row, 10)) * CDbl(POlist.TextMatrix(POlist.row, 12)), 2)
                       
                     
.Tag = ""
.Text = ""
.Visible = False
              
                End If
        Else
            POlist.TextMatrix(val(.Tag), POlist.Col) = IIf(Len(.Text) > 2, Left(.Text, 2), .Text)
            .Tag = ""
            .Text = ""
            .Visible = False
    End If
    End If
    End With
    
NavBar1.SaveEnabled = True
 
End Function

