VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmSapAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAP Adjustment"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   8415
   Tag             =   "02050900"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "SapAdjustment.frx":0000
      CancelVisible   =   0   'False
      PreviousVisible =   0   'False
      NewVisible      =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      EMailVisible    =   -1  'True
      CloseEnabled    =   0   'False
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab sstbSapAdjustment 
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Sap Adjustment"
      TabPicture(0)   =   "SapAdjustment.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDesc(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEntyNumb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDesc(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCurrSap"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDesc(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDesc(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDesc(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDesc(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDesc(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblType"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDesc(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDesc(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblUser"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDesc(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDesc(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ssdcboCompany"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ssdcboCondition"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ssdcboStockNumb"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ssdcboWarehouse"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNewSap"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtRemarks"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cbo_Transaction"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "&Recipients"
      TabPicture(1)   =   "SapAdjustment.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Recipients"
      Tab(1).Control(1)=   "ssdbRecepientList"
      Tab(1).Control(2)=   "Picture1"
      Tab(1).Control(3)=   "cmd_Remove"
      Tab(1).Control(4)=   "cmd_Add"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox cbo_Transaction 
         Height          =   315
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "3"
         Top             =   480
         Width           =   2400
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74835
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74835
         TabIndex        =   15
         Top             =   1635
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   -74880
         ScaleHeight     =   2895
         ScaleWidth      =   7455
         TabIndex        =   29
         Top             =   2040
         Width           =   7455
         Begin ImsMailVB.Imsmail Imsmail 
            Height          =   2775
            Left            =   70800
            TabIndex        =   31
            Top             =   -4320
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4895
         End
      End
      Begin VB.TextBox txtRemarks 
         Height          =   2175
         Left            =   180
         MaxLength       =   7000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2640
         Width           =   7815
      End
      Begin VB.TextBox txtNewSap 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         TabIndex        =   4
         Tag             =   "5"
         Top             =   2160
         Width           =   1425
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboWarehouse 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Tag             =   "1"
         Top             =   825
         Width           =   2385
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
         stylesets(0).Picture=   "SapAdjustment.frx":0054
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
         stylesets(1).Picture=   "SapAdjustment.frx":0070
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboStockNumb 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1470
         Width           =   2400
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
         stylesets(0).Picture=   "SapAdjustment.frx":008C
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
         stylesets(1).Picture=   "SapAdjustment.frx":00A8
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         ExtraHeight     =   370
         Columns.Count   =   2
         Columns(0).Width=   3519
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4233
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCondition 
         Height          =   315
         Left            =   5580
         TabIndex        =   3
         Tag             =   "4"
         Top             =   1470
         Width           =   2385
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
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
         stylesets(0).Picture=   "SapAdjustment.frx":00C4
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
         stylesets(1).Picture=   "SapAdjustment.frx":00E0
         stylesets(1).AlignmentText=   1
         HeadFont3D      =   4
         DefColWidth     =   5292
         ForeColorEven   =   8388608
         BackColorEven   =   16771818
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4339
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1799
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbRecepientList 
         Height          =   1455
         Left            =   -73260
         TabIndex        =   16
         Top             =   480
         Width           =   6150
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
         Columns(0).Width=   9710
         Columns(0).Caption=   "Recp"
         Columns(0).Name =   "Recp"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   10848
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Recipient List"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboCompany 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Tag             =   "0"
         Top             =   480
         Width           =   2400
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
         BackColorOdd    =   16771818
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4445
         Columns(0).Caption=   "Description"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4233
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74835
         TabIndex        =   30
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblDesc 
         Caption         =   "Condition"
         Height          =   315
         Index           =   7
         Left            =   4080
         TabIndex        =   27
         Top             =   1470
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Caption         =   "Stock #"
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   26
         Top             =   1470
         Width           =   1300
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5580
         TabIndex        =   10
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         Caption         =   "User"
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   25
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Caption         =   "Company"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   480
         Width           =   1300
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SAP ADJUSTMENT"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   1140
         Width           =   1560
      End
      Begin VB.Label lblDesc 
         Caption         =   "Type"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   23
         Top             =   1140
         Width           =   1300
      End
      Begin VB.Label lblDate 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5580
         TabIndex        =   12
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         Caption         =   "Date"
         Height          =   315
         Index           =   5
         Left            =   4080
         TabIndex        =   22
         Top             =   1140
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Caption         =   "Transac #"
         Height          =   315
         Index           =   2
         Left            =   4080
         TabIndex        =   21
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label lblDesc 
         Caption         =   "Warehouse"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   840
         Width           =   1300
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Current Sap"
         Height          =   195
         Index           =   11
         Left            =   4080
         TabIndex        =   19
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lblCurrSap 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5580
         TabIndex        =   13
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label lblDesc 
         Caption         =   "New Sap"
         Height          =   315
         Index           =   12
         Left            =   4080
         TabIndex        =   18
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblEntyNumb 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   2385
      End
      Begin VB.Label lblDesc 
         Caption         =   "Entry #"
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   17
         Top             =   1800
         Width           =   1300
      End
   End
End
Attribute VB_Name = "frmSapAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Transnumb As String
Dim CompCode As String
Dim rsReceptList As ADODB.Recordset

Dim SaveEnabled As Boolean
'populate warehouse data grids
Dim rowguid, locked As Boolean, dbtablename As String, j1 As Integer, rowguid1 As String      'jawdat

Private Sub AddWhareHouses(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.State And adStateOpen = adStateClosed Then Exit Sub
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    ssdcboWarehouse.RemoveAll
    
    Do While ((Not rst.EOF))

        ssdcboWarehouse.AddItem rst!loc_name & "" & ";" & rst!loc_locacode & ""
        rst.MoveNext
    Loop

CleanUp:
    rst.Close
    Set rst = Nothing
End Sub



Private Sub cbo_Transaction_Change()
'jawdat, start copy
'Dim currentformname, currentformname1
'currentformname = Forms(3).Name
'currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
'Dim ListOfPrimaryControls() As String
'Set imsLock = New imsLock.lock
'
'ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
'
'Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
'
' If NavBar1.SaveEnabled = False Then Call ClearScreen
'
'If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
'
''Dim imsLock As imsLock.lock
'Set imsLock = New imsLock.lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'
'
'
'                                              'Exit Edit sub because theres nothing the user can do
'Else
'locked = True
'End If
'
''jawdat, end copy
'
'
 
 
 
End Sub

'set transaction data combo bottum

Private Sub cbo_Transaction_Click()


Dim bl As Boolean

    Transnumb = Trim$(cbo_Transaction)
    
    bl = CBool(Len(Transnumb))
    
    Imsmail.Enabled = bl
    cmd_Add.Enabled = bl
    cmd_Remove.Enabled = bl
    NavBar1.PrintEnabled = bl
    'NavBar1.SaveEnabled = False
    ssdbRecepientList.Enabled = bl
'    NavBar1.EMailEnabled = ssdbRecepientList.Rows
    NavBar1.EMailEnabled = True
    


  
    
End Sub

'call function get issue number and populate data grid

Private Sub cbo_Transaction_GotFocus()

On Error Resume Next

Dim rst As ADODB.Recordset

    Set rst = deIms.rsIssueNumber
    If rst.State And adStateOpen = adStateOpen Then rst.Close
    Call deIms.IssueNumber(deIms.NameSpace, CompCode, "SI")
    
    Call PopuLateFromRecordSet(cbo_Transaction, rst, rst.Fields(0).Name, False)
    
    rst.Close
    Set rst = Nothing
    If Err Then Err.Clear
    Call HighlightBackground(cbo_Transaction)

End Sub

'do not allow enter data to transaction combo

Private Sub cbo_Transaction_KeyPress(KeyAscii As Integer)
    If NavBar1.NewEnabled = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo_Transaction_LostFocus()
    cbo_Transaction.Clear
    Call NormalBackground(cbo_Transaction)
    ssdcboCompany.SetFocus
End Sub

'call function add current reciptient to recipien list

Private Sub cmd_Add_Click()
    Imsmail.AddCurrentRecipient
End Sub

'delete recipient form recipient list

Private Sub cmd_Remove_Click()
On Error Resume Next

    rsReceptList.Delete
    rsReceptList.Update
    
    If Err Then Err.Clear
End Sub

'call function get data grids datas and set navbar button

Private Sub Form_Load()
Dim fComp As String


Dim np As String, cn As ADODB.Connection
    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    NavBar1.NewEnabled = SaveEnabled
    NavBar1.SaveEnabled = SaveEnabled

    'Added by Juan (9/27/2000) for Multilingual
    Call translator.Translate_Forms("frmSapAdjustment")
    '------------------------------------------

    np = deIms.NameSpace
    Set cn = deIms.cnIms
    fComp = GetCompany(np, "PE", cn)
    ssdcboStockNumb.FieldSeparator = Chr(1)
    'CompCode = GetCompanyCode(np, fComp, cn)
    
    AddCompanies
    Imsmail.NameSpace = deIms.NameSpace
    
    'IMSMail.Connected = True 'M
    Imsmail.SetActiveConnection deIms.cnIms   'M
    Imsmail.Language = Language 'M
    Call DisableButtons(Me, NavBar1)
    NavBar1.CloseEnabled = True
    
    frmSapAdjustment.Caption = frmSapAdjustment.Caption + " - " + frmSapAdjustment.Tag
    
    sstbSapAdjustment.TabVisible(1) = False
    NavBar1.EMailVisible = False
    
    cbo_Transaction.locked = False
    cbo_Transaction.Enabled = True
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'populate condition data grid

Public Sub AddConditions(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    If rst.EOF And rst.BOF Then Exit Sub
    
    rst.MoveFirst
    ssdcboCondition.RemoveAll
    ssdcboCondition.FieldSeparator = Chr(1)
    
    Do While (Not rst.EOF)
    
        ssdcboCondition.AddItem rst.Fields(0) & "" & Chr(1) & rst.Fields(1) & ""
        
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
End Sub

'populate stock master numbers data grid

Public Sub AddStockNumbers(rst As ADODB.Recordset)
    
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then Exit Sub
    If rst.EOF And rst.BOF Then Exit Sub
    
    rst.MoveFirst
    ssdcboStockNumb.RemoveAll
    
    Do While (Not rst.EOF)
    
        ssdcboStockNumb.AddItem rst.Fields(0) & "" & Chr(1) & rst.Fields(1) & ""
        
        rst.MoveNext
        
    Loop
    
    rst.Close
    Set rst = Nothing
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim closing
'    If txtNewSap <> "" Or ssdcboStockNumb <> "" Then
'        closing = MsgBox("Do you really want to close and lose your last record?", vbYesNo)
'        If closing = vbNo Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If


    Imsmail.Enabled = False
    If open_forms <= 5 Then ShowNavigator
    
 Dim grid2
grid2 = True
'If locked = True Then
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
'End If
  
   
    

End Sub

'call function add current recipient to recipient list

Private Sub IMSMail_OnAddClick(ByVal address As String)
On Error Resume Next

    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        
        rsReceptList.Open
    End If
    
       'Modified by Muzammil 08/14/00
'Reason - To Add "INTERNET!" before email.
If (InStr(1, address, "@") > 0) And InStr(1, UCase(address), "INTERNET!") = 0 Then address = "INTERNET!" & UCase(address)
    
    
    If Not IsInList(address, "Recipients", rsReceptList) Then _
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))

    Set ssdbRecepientList.DataSource = rsReceptList
    ssdbRecepientList.Columns(0).DataField = "Recipients"
    
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'call function to print crystal report

Private Sub NavBar1_OnPrintClick()
On Error Resume Next
    BeforePrint
    MDI_IMS.CrystalReport1.Action = 1
    MDI_IMS.CrystalReport1.Reset
    If Err Then MsgBox Err.Description: Err.Clear
End Sub

'get parmeters for crystal report

Public Sub BeforePrint()
On Error GoTo errHandler

     With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath + "wareSISE.rpt"
        .ParameterFields(0) = "transnumb;" & cbo_Transaction & ";TRUE"
        .ParameterFields(1) = "namespace;" & deIms.NameSpace & ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00161") 'J Added
        .WindowTitle = IIf(msg1 = "", "SAP Adjustment", msg1) 'J Modified
        Call translator.Translate_Reports("wareSISE.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
    End With
    Exit Sub
    
errHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'validate data format before save

Private Sub NavBar1_OnSaveClick()

NavBar1.SaveEnabled = SaveEnabled
MDI_IMS.StatusBar1.Panels(1).Text = "Saving"
Screen.MousePointer = 11



On Error Resume Next

Dim Cancel As Boolean

    If Len(txtNewSap) Then
        MDI_IMS.StatusBar1.Panels(1).Text = "Validation"
        Call txtNewSap_Validate(Cancel)
    Else
        Cancel = True
        Screen.MousePointer = 0
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00393") 'J added
        MsgBox IIf(msg1 = "", "New SAP cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        Screen.MousePointer = 11
    End If
    
    If Cancel Then
        Screen.MousePointer = 0
        MDI_IMS.StatusBar1.Panels(1).Text = ""
        Exit Sub
    End If
    
    
    'Modified by Muzammil 08/11/00
       'Reason - VBCRLFs before the text would block Email Generation.
'          MDI_IMS.StatusBar1.Panels(1).text = "Saving Remarks"
'          Do While InStr(1, txtRemarks, vbCrLf) = 1                   'M
'             txtRemarks = Mid(txtRemarks, 3, Len(txtRemarks))         'M
'          Loop                                                        'M
'             txtRemarks = LTrim$(txtRemarks)                          'M
        
    
    If Len(txtRemarks) = 0 Then
        Screen.MousePointer = 0
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00394") 'J added
        MsgBox IIf(msg1 = "", "Remarks cannot be empty", msg1) 'J modified
        '---------------------------------------------
        MDI_IMS.StatusBar1.Panels(1).Text = ""
        NavBar1.SaveEnabled = SaveEnabled
        Exit Sub
    End If
        Screen.MousePointer = 11
        'NavBar1.SaveEnabled = False
        
    
    If SapAdjustmentNew Then
        
        'doevents
        MDI_IMS.StatusBar1.Panels(1).Text = "Saving Items"
        Call cbo_Transaction.AddItem(Transnumb, cbo_Transaction.ListCount)
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, Transnumb)
        
'        BeforePrint
'        Call SendWareHouseMessage(deIms.NameSpace, "Automatic Distribution", _
'                                  lblType, deIms.cnIms, CreateRpti)
    End If
        
         'NavBar1.SaveEnabled = False
    
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err)
    End If
    MDI_IMS.StatusBar1.Panels(1).Text = ""
    NavBar1.SaveEnabled = False
    Screen.MousePointer = 0
End Sub

'call function get company recordset and populate data grid

Private Sub ssdcboCompany_Click()
    cbo_Transaction.ListIndex = CB_ERR
    CompCode = ssdcboCompany.Columns("Code").Text
    
    ssdcboWarehouse = ""
    ssdcboWarehouse.RemoveAll
    Call AddWhareHouses(GetLocation(deIms.NameSpace, "OTHER", CompCode, deIms.cnIms, False))

    NavBar1.SaveEnabled = SaveEnabled
    
    ssdcboCompany.SelLength = 0
    ssdcboCompany.SelStart = 0
    
End Sub

Private Sub ssdcboCompany_GotFocus()
Call HighlightBackground(ssdcboCompany)
End Sub

Private Sub ssdcboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboCompany.DroppedDown = True
    End If
End Sub


Private Sub ssdcboCompany_LostFocus()
Call NormalBackground(ssdcboCompany)
End Sub

'call function sap stock values and format tofour decimal digit

Private Sub ssdcboCondition_Click()
Dim db As Double
Dim stckNumb As String, condCond As String

    stckNumb = ssdcboStockNumb.Text
    cbo_Transaction.ListIndex = CB_ERR
    condCond = ssdcboCondition.Columns("Code").Text
    Call Get_Sap_Stock_Values(stckNumb, condCond, db)
    
    lblCurrSap = FormatNumber((db), 4)
    txtNewSap.Enabled = True
    
    ssdcboCondition.SelLength = 0
    ssdcboCondition.SelStart = 0

End Sub

Private Sub ssdcboCondition_GotFocus()
Call HighlightBackground(ssdcboCondition)
End Sub

Private Sub ssdcboCondition_LostFocus()
Call NormalBackground(ssdcboCondition)
End Sub

'call function to get condition  values and fill data grid

Private Sub ssdcboStockNumb_Click()
Dim str As String
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

'jawdat, start copy

Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.Lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock

ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)

Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02

j1 = j1 + 1
If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
'If j1 > 1 Then

NavBar1.SaveEnabled = False
cbo_Transaction.Enabled = False
ssdcboCondition.Enabled = False
txtNewSap.Enabled = False
txtRemarks.Enabled = False
lblEntyNumb.Enabled = False
lblType.Enabled = False
lblUser.Enabled = False
lblDate.Enabled = False
lblCurrSap.Enabled = False


'Dim imsLock As imsLock.lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid1, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

                                             'Exit Edit sub because theres nothing the user can do
Else

cbo_Transaction.Enabled = True
ssdcboCondition.Enabled = True
txtNewSap.Enabled = True
txtRemarks.Enabled = True
lblEntyNumb.Enabled = True
lblType.Enabled = True
lblUser.Enabled = True
lblDate.Enabled = True
lblCurrSap.Enabled = True
NavBar1.SaveEnabled = True

j1 = 1

rowguid1 = rowguid
locked = True
End If

    ClearFields
    
    Call AddConditions(Get_Sap_Stock_Values(ssdcboStockNumb.Columns(0).Text))
If locked = True Then
    If ssdcboCondition.Rows Then
        str = ssdcboStockNumb.Columns(1).Text
        Call FindInGrid(ssdcboCondition, str, True, 1)
        ssdcboCondition.Text = ssdcboCondition.Columns(0).Text
        Call ssdcboCondition_Click: ssdcboCondition.Enabled = True
End If
    
    ssdcboStockNumb.SelLength = 0
    ssdcboStockNumb.SelStart = 0
 End If
    
End Sub

Private Sub ssdcboStockNumb_GotFocus()
Call HighlightBackground(ssdcboStockNumb)
End Sub

Private Sub ssdcboStockNumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboStockNumb.DroppedDown Then ssdcboStockNumb.DroppedDown = True

End Sub

Private Sub ssdcboStockNumb_LostFocus()
Call NormalBackground(ssdcboStockNumb)
End Sub

Private Sub ssdcboStockNumb_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboStockNumb)) > 0 Then
         If Not ssdcboStockNumb.IsItemInList Then
                ssdcboStockNumb.Text = ""
            End If
            If Len(Trim$(ssdcboStockNumb)) = 0 Then
           ssdcboStockNumb.SetFocus
            Cancel = True
            End If
            End If
End Sub

'assign values to lables and get stock number

Private Sub ssdcboWarehouse_Click()
    ClearFields
    lblUser = CurrentUser
    cbo_Transaction.ListIndex = CB_ERR
    lblDate = Format$(Date, "mm/dd/yyyy")
    Call AddStockNumbers(Get_Sap_Stock_Values)
    ssdcboStockNumb.Enabled = ssdcboStockNumb.Rows
    
    If Len(ssdcboWarehouse) <> 0 Then
        NavBar1.SaveEnabled = SaveEnabled
    End If
    
    ssdcboWarehouse.SelStart = 0
    ssdcboWarehouse.SelLength = 0
End Sub

'call class function to get sap stock values

Private Function Get_Sap_Stock_Values(Optional StockNumb = Null, Optional Cond = Null, Optional value As Double) As ADODB.Recordset
Dim WH As String
Dim cl As imsspInt

    
    Set cl = New imsspInt
    WH = ssdcboWarehouse.Columns("Code").Text
    
    With deIms
        Set Get_Sap_Stock_Values = cl.Get_Sap_Stock_Values(.NameSpace, CompCode, WH, .cnIms, StockNumb, Cond, value)
    End With
    
    Set cl = Nothing
    
End Function

Private Sub ssdcboWarehouse_GotFocus()
Call HighlightBackground(ssdcboWarehouse)
End Sub

Private Sub ssdcboWarehouse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        ssdcboWarehouse.DroppedDown = True
    End If
End Sub

Private Sub ssdcboWarehouse_LostFocus()
Call NormalBackground(ssdcboWarehouse)
End Sub

'set navbar buttoms

Private Sub sstbSapAdjustment_Click(PreviousTab As Integer)
Dim blFlag As Boolean


    blFlag = 1
    
    If sstbSapAdjustment.Tab = 0 Then 'This line was modifed by Juan (9/27/2000) for Multilingual
        NavBar1.SaveEnabled = SaveEnabled
        If Len(cbo_Transaction) <> 0 Then
            NavBar1.PrintEnabled = True
            NavBar1.SaveEnabled = SaveEnabled
        End If
        
        With NavBar1
    '        .SaveEnabled = 0
            .CloseEnabled = True
            .PrintEnabled = .SaveEnabled And cbo_Transaction.ListIndex <> CB_ERR
            .EMailEnabled = ((ssdbRecepientList.Rows) And (.PrintEnabled))
        End With
    Else
        NavBar1.SaveEnabled = False
        NavBar1.PrintEnabled = False
        NavBar1.CloseEnabled = False
    End If

End Sub

Private Sub txtNewSap_GotFocus()

Call HighlightBackground(txtNewSap)

End Sub

Private Sub txtNewSap_LostFocus()
Call NormalBackground(txtNewSap)
End Sub

'validate new SAP value and format, set to four digit

Private Sub txtNewSap_Validate(Cancel As Boolean)
    
 
    If Len(txtNewSap) Then
    
        If Not IsNumeric(txtNewSap) Then
        
            'Modified by Juan (9/27/2000) for Multilingual
            msg1 = translator.Trans("M00122") 'J added
            MsgBox IIf(msg1 = "", "Invalid Value", msg1) 'J modified
            '---------------------------------------------
            
        ElseIf txtNewSap <= 0 Then
                
            'Modified by Juan (9/27/2000) for Multilingual
            msg1 = translator.Trans("M00394") 'J added
            MsgBox IIf(msg1 = "", "SAP can not be less than or equal to zero", msg1) 'J modified
            '---------------------------------------------
                
                
        ElseIf txtNewSap = lblCurrSap Then
        
            'Modified by Juan (9/27/2000) for Multilingual
            msg1 = translator.Trans("M00395") 'J added
            MsgBox IIf(msg1 = "", "SAP can not be equal to current sap", msg1) 'J modified
            '---------------------------------------------
        
        Else
            txtNewSap = FormatNumber((txtNewSap), 4)
            Exit Sub
            
        End If
        
        Cancel = True
    Else
    
'        txtNewSap = FormatNumber((txtNewSap), 4)
        Cancel = False: Exit Sub
    End If
        
    
    
    Cancel = True
End Sub

'set parameters and call function send email and fax

Private Sub NavBar1_OnEMailClick()
Dim Filename As String
BeforePrint
    Call WriteRPTIFile(CreateRpti, Filename)
    Call SendEmailAndFax(rsReceptList, "Recipients", "Sap Adjustment", "", Filename)

    Set rsReceptList = Nothing
    Set ssdbRecepientList.DataSource = Nothing

End Sub

Private Function SapAdjustment() As Boolean
'On Error Resume Next

Dim db As Double
Dim cmd As New ADODB.Command
Dim Params As ADODB.parameters

' Set command properties
    Set Params = cmd.parameters
    cmd.CommandText = "SAPADJUSTMENT"
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = deIms.cnIms
    
    SapAdjustment = False
  
    ' Define stored procedure parameters
    ' and append to command.
    
    db = 0
    db = CDbl(Trim$(txtNewSap))
    
    If Err Then Err.Clear
    
    If db = 0 Then
    
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00396") 'J added
        MsgBox IIf(msg1 = "", "New SAP contains an invalid value", msg1) 'J modified
        '---------------------------------------------
        
        Exit Function
    End If
    
'    If Len(Trim$(txtRemarks)) > 1000 Then
'
'        'Modified by Juan (9/27/2000) for Multilingual
'        msg1 = translator.Trans("M00397") 'J added
'        MsgBox IIf(msg1 = "", "Remarks is too large", msg1) 'J modified
'        '---------------------------------------------
'
'        Exit Function
'    End If
            
    If Params.Count = 0 Then
        'Params.Refresh
        Params.Append cmd.CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5)
        Params.Append cmd.CreateParameter("@USER", adVarChar, adParamInput, 20)
        Params.Append cmd.CreateParameter("@COMPCODE", adVarChar, adParamInput, 10)
        Params.Append cmd.CreateParameter("@REMARKS", adVarChar, adParamInput, 7000)
        Params.Append cmd.CreateParameter("@LOCATION", adVarChar, adParamInput, 10)
        Params.Append cmd.CreateParameter("@NEWSAP", adDouble, adParamInput, 20)
        Params.Append cmd.CreateParameter("@STOCKNUMBER", adVarChar, adParamInput, 20)
        Params.Append cmd.CreateParameter("@CONDITION", adVarChar, adParamInput, 2)
        
        Params.Append cmd.CreateParameter("@SETRANSNUMB", adVarChar, adParamOutput, 15)
        Params.Append cmd.CreateParameter("@SITRANSNUMB", adVarChar, adParamOutput, 15)
    End If
    
    Params("@USER") = CurrentUser
    Params("@NAMESPACE") = deIms.NameSpace

    Params("@NEWSAP") = db
    Params("@COMPCODE") = Trim$(CompCode)
    Params("@REMARKS") = Trim$(txtRemarks)
    Params("@LOCATION") = Trim$(ssdcboWarehouse.Columns(1).Text)
    Params("@CONDITION") = Trim$(ssdcboCondition.Columns(1).Text)
    Params("@STOCKNUMBER") = Trim$(ssdcboStockNumb.Columns(0).Text)
    
    
    
    'doevents
    ' Execute the command
    Call cmd.Execute(Options:=adExecuteNoRecords)
    Transnumb = Trim$(Params("@SITRANSNUMB") & "")
    
    If Len(Transnumb) > 0 Then
        DisableControls
        SapAdjustment = True
        'doevents: 'doevents: 'doevents
        lblEntyNumb = cmd.parameters("@SETRANSNUMB") & ""
        Call cbo_Transaction.AddItem(Transnumb)
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, Transnumb)
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00018") + " " 'J added
        msg2 = " " + translator.Trans("M00398") + " " 'J added
        MsgBox IIf(msg1 = "", "Please note that your transaction number is ", msg1) & Transnumb & IIf(msg2 = "", " and ", msg2) & lblEntyNumb
        '---------------------------------------------
        
    End If

    Set cmd = Nothing
    Set Params = Nothing
    
  
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::SapAdjustment", Err.Description, Err)
    End If
End Function



Private Function SapAdjustmentNew() As Boolean
Dim answer As Boolean
On Error GoTo RollBack
Dim cn As ADODB.Connection
Dim cmd As ADODB.Command
Dim datax As ADODB.Recordset
Dim sql
Dim oldSap As Double
Dim currentSap As Double
Dim newSap As Double
Dim qty1 As Double
Dim qty2 As Double
Dim NameSpace, Company, Location, condition, StockNumber, User, remarks, subLocation, logical, stockDescription, serialNumber As String
Dim SI As String
Dim SE As String
Dim RecordsAffected As Long
Dim tranSerial As Integer

answer = False
Set cn = deIms.cnIms
NameSpace = deIms.NameSpace
Company = Trim$(CompCode)
Location = Trim$(ssdcboWarehouse.Columns(1).Text)
condition = Trim$(ssdcboCondition.Columns(1).Text)
StockNumber = Trim$(ssdcboStockNumb.Columns(0).Text)
User = CurrentUser
remarks = Trim$(txtRemarks)
newSap = CDbl(Trim$(txtNewSap))
    'PRELIMINARY
    sql = "SELECT qs1_desc FROM QTYST1 WHERE qs1_compcode = '" + Company + "' AND " _
        + "qs1_npecode = '" + NameSpace + "' AND qs1_ware = '" + Location + "' AND " _
        + "qs1_stcknumb = '" + StockNumber + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, deIms.cnIms, adOpenForwardOnly
    If datax.RecordCount = 0 Then
        MsgBox "Transaction can't be saved because could not get stock number description"
    Else
        stockDescription = datax!qs1_desc
    End If
    
    sql = "select sap_value from sap " _
    + "where sap_compcode='" + Company + "' and sap_npecode='" + NameSpace + "' and " _
    + "sap_loca='" + Location + "' and sap_stcknumb = '" + StockNumber + "' and sap_cond='" + condition + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        currentSap = 0
        oldSap = 0
    Else
        currentSap = datax!sap_value
        oldSap = datax!sap_value
    End If
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    Call BeginTransaction(cn)
    sql = "UPDATE SAP SET sap_value = " + Format(newSap) + " , " _
        + "sap_modiuser = '" + User + "' WHERE sap_loca = '" + Location + "' AND " _
        + "sap_npecode = '" + NameSpace + "' AND sap_compcode = '" + Company + "' AND " _
        + "sap_cond = '" + condition + "' AND sap_stcknumb = '" + StockNumber + "' "
    cmd.CommandText = sql
    cmd.Execute RecordsAffected, , adExecuteNoRecords
    If Err.number > 0 Or RecordsAffected = 0 Then
        Call RollbackTransaction(cn)
        MsgBox "The transaction could not be saved when updating SAP"
        Err.Clear
        Exit Function
    End If
    If oldSap <> newSap Then
        sql = "insert into saphistory (saph_compcode,saph_npecode,saph_loca,saph_stcknumb,saph_cond,saph_date,saph_value) " _
            + "values('" + Company + "', '" + NameSpace + "' ,  '" + Location + "' , '" + StockNumber + "' , " _
            + "'" + condition + "', GETDATE(), " + Format(oldSap) + ") "
        cmd.CommandText = sql
        cmd.Execute RecordsAffected, , adExecuteNoRecords
        If Err.number > 0 Or RecordsAffected = 0 Then
            Call RollbackTransaction(cn)
            MsgBox "The transaction could not be saved when inserting into SAP history"
            Err.Clear
            Exit Function
        End If
    End If
    'TRANSACTION NUMBERS
    Dim num As Integer
    Dim TransactionNumber As Long
    TransactionNumber = GetTransNumb(deIms.NameSpace, cn)
    SI = "SI-" + Format(TransactionNumber)
    TransactionNumber = GetTransNumb(deIms.NameSpace, cn)
    SE = "SE-" + Format(TransactionNumber)

    'QTYST5
    sql = "SELECT SUM (qs5_primqty) as qty1, SUM(qs5_secoqty) as qty2, qs5_logiware , qs5_subloca " _
        + "FROM QTYST5 WHERE ((qs5_compcode = '" + Company + "') AND (qs5_cond = '" + condition + "' ) AND " _
        + "(qs5_ware = '" + Location + "') AND (qs5_stcknumb = '" + StockNumber + "') AND (qs5_npecode = '" + NameSpace + "')) " _
        + "GROUP BY qs5_stcknumb, qs5_compcode, qs5_ware, qs5_subloca, qs5_logiware, qs5_cond " _
        + "HAVING SUM(qs5_primqty) > 0 "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount > 0 Then
        tranSerial = 1
        Do While Not datax.EOF
            subLocation = datax!qs5_subloca
            logical = datax!qs5_logiware
            qty1 = datax!qty1
            qty2 = datax!qty2
        
            If tranSerial = 1 Then
                'INSERTING TRANSACTIONS
                sql = "INSERT INTO INVTISSUE (ii_compcode, ii_npecode, ii_ware, ii_trannumb,ii_user, " _
                    + "ii_trandate,ii_trantype,ii_issuto,ii_requnumb,ii_valion,ii_valuby,ii_stcknumb, " _
                    + "ii_cond,ii_sap,ii_newsap,ii_entynumb, ii_creauser, ii_modiuser) " _
                    + "VALUES('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SI + "', '" + User + "', GETDATE(), " _
                    + "'SI','" + Location + "', null, GETDATE(), '" + User + "', '" + StockNumber + "','" + condition + "', " _
                    + Format(currentSap) + "," + Format(newSap) + ", '" + SE + "', '" + User + "', '" + User + "') "
                cmd.CommandText = sql
                cmd.Execute RecordsAffected, , adExecuteNoRecords
                If Err.number > 0 Or RecordsAffected = 0 Then
                    Call RollbackTransaction(cn)
                    MsgBox "The transaction could not be saved when inserting into INVTISSUE"
                    Err.Clear
                    Exit Function
                End If
                
                sql = "INSERT INTO INVTISSUEREM (iir_compcode, iir_npecode, iir_ware, iir_trannumb, " _
                    + "iir_linenumb, iir_remk, iir_creauser, iir_modiuser) " _
                    + "VALUES('" + Company + "','" + NameSpace + "','" + Location + "','" + SI + "', 1," _
                    + "'" + remarks + "', '" + User + "', '" + User + "') "
                cmd.CommandText = sql
                cmd.Execute RecordsAffected, , adExecuteNoRecords
                If Err.number > 0 Or RecordsAffected = 0 Then
                    Call RollbackTransaction(cn)
                    MsgBox "The transaction could not be saved when inserting into INVTISSUEREM"
                    Err.Clear
                    Exit Function
                End If
                
                sql = "INSERT INTO INVTRECEIPT (ir_compcode, ir_npecode, ir_ware, ir_trannumb, " _
                    + "ir_user, ir_trandate, ir_trantype, ir_tranfrom, ir_valion, ir_valuby, " _
                    + "ir_stcknumb, ir_cond, ir_sap, ir_newsap, ir_entynumb, ir_creauser, ir_modiuser) " _
                    + "VALUES('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SE + "', '" + User + "', GETDATE(), " _
                    + "'SE', '" + Location + "', GETDATE(), '" + User + "', '" + StockNumber + "', '" + condition + "', " _
                    + Format(newSap) + ", " + Format(newSap) + ", '" + SI + "', '" + User + "' , '" + User + "') "
                cmd.CommandText = sql
                cmd.Execute RecordsAffected, , adExecuteNoRecords
                If Err.number > 0 Or RecordsAffected = 0 Then
                    Call RollbackTransaction(cn)
                    MsgBox "The transaction could not be saved when inserting into INVTRECEIPT"
                    Err.Clear
                    Exit Function
                End If
                
                sql = "INSERT INTO INVTRECEIPTREM (irr_compcode, irr_npecode, irr_ware, irr_trannumb, " _
                    + "irr_linenumb, irr_remk, irr_creauser, irr_modiuser) " _
                    + "VALUES('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SE + "', 1, " _
                    + "'" + remarks + "', '" + User + "', '" + User + "') "
                cmd.CommandText = sql
                cmd.Execute RecordsAffected, , adExecuteNoRecords
                If Err.number > 0 Or RecordsAffected = 0 Then
                    Call RollbackTransaction(cn)
                    MsgBox "The transaction could not be saved when inserting into INVTRECEIPTREM"
                    Err.Clear
                    Exit Function
                End If
            End If
            
            'QTYST5 RECEIPT SIDE
            sql = "INSERT INTO QTYST5(qs5_npecode , qs5_compcode , qs5_ware , qs5_stcknumb , " _
                + "qs5_cond , qs5_subloca , qs5_logiware , qs5_secoqty , qs5_trantype , " _
                + "qs5_fromto , qs5_trancompcode , qs5_tranware , qs5_trantrannumb , qs5_transerl , " _
                + "qs5_primqty , qs5_tranlinenumb , qs5_trannumb,  qs5_creauser, qs5_modiuser) " _
                + "VALUES('" + NameSpace + "', '" + Company + "', '" + Location + "', '" + StockNumber + "', " _
                + "'" + condition + "', '" + subLocation + "', '" + logical + "', " + Format(qty2) + ", 'SE' , " _
                + "'" + Location + "', '" + Company + "', '" + Location + "', '" + SE + "', " + Format(tranSerial) + ", " _
                + "" + Format(qty1) + ", " + Format(tranSerial) + ", '" + SE + "', '" + User + "', '" + User + "')  "
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting RECEIPT into QTYST5"
                Err.Clear
                Exit Function
            End If
            
            'QTYST5 ISSUE SIDE
            sql = "INSERT INTO QTYST5(qs5_npecode ,qs5_compcode ,qs5_ware ,qs5_stcknumb ,qs5_cond , " _
                + "qs5_subloca ,qs5_logiware ,qs5_secoqty ,qs5_trantype ,qs5_fromto ,qs5_trancompcode , " _
                + "qs5_tranware ,qs5_trantrannumb ,qs5_transerl ,qs5_primqty ,qs5_tranlinenumb , " _
                + "qs5_trannumb, qs5_creauser, qs5_modiuser) " _
                + "VALUES('" + NameSpace + "', '" + Company + "', '" + Location + "', '" + StockNumber + "', '" + condition + "', " _
                + "'" + subLocation + "', '" + logical + "', " + Format(qty2 * -1) + ", 'SI', '" + Location + "', " _
                + "'" + Company + "', '" + Location + "', '" + SI + "', " + Format(tranSerial) + ", " + Format(qty1 * -1) + ",  " _
                + "" + Format(tranSerial) + ", '" + SI + "', '" + User + "', '" + User + "')  "
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting ISSUE into QTYST5"
                Err.Clear
                Exit Function
            End If
            
            'INVTISSUEDETL
            sql = "INSERT INTO INVTISSUEDETL (iid_compcode, iid_npecode, iid_ware, iid_trannumb, " _
                + "iid_transerl, iid_ponumb, iid_liitnumb, iid_stcknumb, iid_ps, iid_serl, iid_newcond, " _
                + "iid_stcktype, iid_ctry, iid_tosubloca, iid_tologiware, iid_owle, iid_leasecomp,  " _
                + "iid_primqty, iid_secoqty, iid_unitpric, iid_curr, iid_currvalu, iid_stckdesc,  " _
                + "iid_fromlogiware, iid_fromsubloca, iid_origcond, iid_creauser, iid_modiuser)  " _
                + "VALUES ('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SI + "', " + Format(tranSerial) + ", NULL, NULL,  " _
                + "'" + StockNumber + "', 1, NULL, '" + condition + "', NULL, 'USA', '" + subLocation + "', '" + logical + "', NULL, NULL,  " _
                + "" + Format(qty1) + ", " + Format(qty2) + ", " + Format(currentSap) + ", 'USD', 1, '" + stockDescription + "',  " _
                + "'" + logical + "', '" + subLocation + "', '" + condition + "', '" + User + "', '" + User + "')"
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting into INVTISSUEDETL"
                Err.Clear
                Exit Function
            End If
            
            'INVTRECEIPTDETL
            sql = "INSERT INTO INVTRECEIPTDETL(ird_compcode, ird_npecode, ird_ware, ird_trannumb, " _
                + "ird_transerl, ird_ponumb, ird_liitnumb, ird_stcknumb, ird_ps, ird_serl, " _
                + "ird_newcond, ird_stcktype, ird_ctry, ird_tosubloca, ird_tologiware, ird_owle, " _
                + "ird_leasecomp, ird_primqty, ird_secoqty, ird_unitpric, ird_curr, ird_currvalu,  " _
                + "ird_stckdesc, ird_fromlogiware, ird_fromsubloca, ird_origcond, ird_reprcost,  " _
                + "ird_newstcknumb, ird_newdesc, ird_creauser, ird_modiuser)  " _
                + "VALUES ('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SE + "', " + Format(tranSerial) + ", NULL, NULL,  " _
                + "'" + StockNumber + "', 1, NULL, '" + condition + "', NULL, 'USA', '" + subLocation + "',  " _
                + "'" + logical + "', NULL, NULL, " + Format(qty1) + ", " + Format(qty2) + ", " + Format(newSap) + ", " _
                + "'USD', 1, '" + stockDescription + "', '" + logical + "', '" + subLocation + "', '" + condition + "', " _
                + "NULL,NULL,NULL, '" + User + "', '" + User + "')"
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting into INVTRECEIPTDETL"
                Err.Clear
                Exit Function
            End If
            tranSerial = tranSerial + 1
            datax.MoveNext
        Loop
    End If
    
    sql = "SELECT qs6_logiware , qs6_subloca ,  qs6_serl , qs6_primqty, qs6_secoqty FROM QTYST6 " _
        + "WHERE ((qs6_compcode = '" + Company + "') AND (qs6_cond = '" + condition + "') AND " _
        + "(qs6_ware = '" + Location + "') AND (qs6_stcknumb = '" + StockNumber + "') AND " _
        + "(qs6_npecode = '" + NameSpace + "'))  AND qs6_primqty >0  " _
        + "GROUP BY qs6_stcknumb, qs6_compcode, qs6_ware, qs6_subloca, qs6_logiware, " _
        + "qs6_cond , qs6_serl , qs6_primqty , qs6_secoqty "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount > 0 Then
        tranSerial = 1
        Do While Not datax.EOF
            subLocation = datax!qs6_subloca
            logical = datax!qs6_logiware
            qty1 = datax!qs6_primqty
            qty2 = datax!qs6_secoqty
            serialNumber = datax!qs6_serl
            
            'INSERT RECEIPT INTO QTYST7
            sql = "INSERT INTO QTYST7(qs7_npecode, qs7_compcode, qs7_ware, qs7_stcknumb, " _
                + "qs7_cond, qs7_serl, qs7_subloca, qs7_logiware, qs7_secoqty, qs7_trantype, " _
                + "qs7_fromto, qs7_trancompcode, qs7_tranware, qs7_trantrannumb, qs7_transerl, " _
                + "qs7_primqty, qs7_tranlinenumb, qs7_trannumb,  qs7_creauser, qs7_modiuser) " _
                + "VALUES('" + NameSpace + "', '" + Company + "', '" + Location + "', '" + StockNumber + "', " _
                + "'" + condition + "', '" + serialNumber + "', '" + subLocation + "', '" + logical + "', " _
                + "" + Format(qty2) + ", 'SE', '" + Location + "', '" + Company + "', '" + Location + "', " _
                + "'" + SE + "', 1, " + Format(qty1) + ", " + Format(tranSerial) + ", '" + SE + "', " _
                + "'" + User + "', '" + User + "' )"
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting RECEIPT into QTYST7"
                Err.Clear
                Exit Function
            End If
            
            'INSERT ISSUE INTO QTYST7
            sql = "INSERT INTO QTYST7(qs7_npecode, qs7_compcode, qs7_ware, qs7_stcknumb, " _
                + "qs7_cond, qs7_serl, qs7_subloca, qs7_logiware, qs7_secoqty, qs7_trantype , " _
                + "qs7_fromto, qs7_trancompcode, qs7_tranware, qs7_trantrannumb, qs7_transerl, " _
                + "qs7_primqty, qs7_tranlinenumb, qs7_trannumb,  qs7_creauser, qs7_modiuser) " _
                + "VALUES('" + NameSpace + "', '" + Company + "', '" + Location + "', '" + StockNumber + "', " _
                + "'" + condition + "', '" + serialNumber + "', '" + subLocation + "', '" + logical + "', " _
                + "" + Format(qty2 * -1) + ", 'SI', '" + Location + "', '" + Company + "', '" + Location + "', " _
                + "'" + SI + "', " + Format(tranSerial) + ", " + Format(qty1 * -1) + ", " + Format(tranSerial) + ", " _
                + "'" + SI + "', '" + User + "', '" + User + "' )"
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting ISSUE into QTYST7"
                Err.Clear
                Exit Function
            End If
    
            'INVTISSUEDETL
            sql = "INSERT INTO INVTISSUEDETL (iid_compcode, iid_npecode, iid_ware, iid_trannumb, " _
                + "iid_transerl, iid_ponumb, iid_liitnumb, iid_stcknumb, iid_ps, iid_serl, " _
                + "iid_newcond, iid_stcktype, iid_ctry, iid_tosubloca, iid_tologiware, iid_owle, " _
                + "iid_leasecomp, iid_primqty, iid_secoqty, iid_unitpric, iid_curr, iid_currvalu," _
                + "iid_stckdesc, iid_fromlogiware, iid_fromsubloca, iid_origcond, iid_creauser, iid_modiuser) " _
                + "VALUES ('" + Company + "', '" + NameSpace + "', '" + NameSpace + "', '" + SI + "', " + Format(tranSerial) + ", " _
                + "NULL, NULL, '" + StockNumber + "', 0, '" + serialNumber + "', '" + condition + "', NULL, 'USA', " _
                + "'" + subLocation + "', '" + logical + "', NULL, NULL, " + Format(qty1) + ", " + Format(qty2) + ", " _
                + Format(currentSap) + ", 'USD', 1, '" + stockDescription + "', '" + logical + "', " _
                + "'" + subLocation + "', '" + condition + "', '" + User + "', '" + User + "' ) "
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting into INVTISSUEDETL"
                Err.Clear
                Exit Function
            End If
            
            'INVTRECEIPTDETL
            sql = "INSERT INTO INVTRECEIPTDETL (ird_compcode, ird_npecode, ird_ware, ird_trannumb, " _
                + "ird_transerl, ird_ponumb, ird_liitnumb, ird_stcknumb, ird_ps, ird_serl, " _
                + "ird_newcond, ird_stcktype, ird_ctry, ird_tosubloca, ird_tologiware, ird_owle, " _
                + "ird_leasecomp, ird_primqty, ird_secoqty, ird_unitpric, ird_curr, ird_currvalu, " _
                + "ird_stckdesc, ird_fromlogiware, ird_fromsubloca, ird_origcond, ird_reprcost, " _
                + "ird_newstcknumb, ird_newdesc, ird_creauser, ird_modiuser) " _
                + "VALUES ('" + Company + "', '" + NameSpace + "', '" + Location + "', '" + SE + "', " _
                + Format(tranSerial) + ", NULL, NULL, '" + StockNumber + "',0, '" + serialNumber + "', " _
                + "'" + condition + "', NULL, 'USA', '" + subLocation + "', '" + logical + "', NULL, NULL, " _
                + Format(qty1) + ", " + Format(qty2) + ", " + Format(newSap) + ", 'USD', 1, '" + stockDescription + "', " _
                + "'" + logical + "', '" + subLocation + "', '" + condition + "', NULL,NULL,NULL, '" + User + "', '" + User + "' )"
            cmd.CommandText = sql
            cmd.Execute RecordsAffected, , adExecuteNoRecords
            If Err.number > 0 Or RecordsAffected = 0 Then
                Call RollbackTransaction(cn)
                MsgBox "The transaction could not be saved when inserting into INVTRECEIPTDETL"
                Err.Clear
                Exit Function
            End If
            tranSerial = tranSerial + 1
            datax.MoveNext
        Loop
    End If
    Call CommitTransaction(cn)
    If Len(SI) > 0 And Len(SE) > 0 Then
        DisableControls
        answer = True
        lblEntyNumb = SI
        Call cbo_Transaction.AddItem(SI)
        cbo_Transaction.ListIndex = IndexOf(cbo_Transaction, SI)
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00018") + " " 'J added
        msg2 = " " + translator.Trans("M00398") + " " 'J added
        MsgBox IIf(msg1 = "", "Please note that your transaction number is ", msg1) & Transnumb & IIf(msg2 = "", " and ", msg2) & lblEntyNumb
        '---------------------------------------------
    End If
    
    SapAdjustmentNew = answer
    Set cmd = Nothing
    Set datax = Nothing
    Exit Function
    
RollBack:
    If Err.number <> 0 Then
        MsgBox "The transaction could not be saved due this error: " + Err.Description
        Call RollbackTransaction(cn)
        Err.Clear
    End If
End Function
Function CommitTransaction(cn As ADODB.Connection)
On Error Resume Next
    With MakeCommand(cn, adCmdText)
        .CommandText = "COMMIT TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    If Err Then Err.Clear
End Function
Function RollbackTransaction(cn As ADODB.Connection)
On Error Resume Next
    With MakeCommand(cn, adCmdText)
        .CommandText = "ROLLBACK TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    If Err Then Err.Clear
End Function
Function BeginTransaction(cn As ADODB.Connection)
    With MakeCommand(cn, adCmdText)
        .CommandText = "BEGIN TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
End Function
Function getDATA(Access, parameters) As ADODB.Recordset
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = deIms.cnIms
        .CommandType = adCmdStoredProc
        .CommandText = Access
        Set getDATA = .Execute(, parameters)
    End With
End Function

'clear data fields

Private Sub ClearFields()
    txtNewSap = ""
    txtRemarks = ""
    lblEntyNumb = ""
    ssdcboCondition.Text = ""
    'ssdcboStockNumb.Text = ""
    cbo_Transaction.ListIndex = CB_ERR
End Sub

'disable controls

Private Sub DisableControls()
    txtNewSap.Enabled = False
    txtRemarks.Enabled = False
    ssdcboCondition.Enabled = False
    ssdcboStockNumb.Enabled = False
End Sub

'get company data recordset and populate data grid

Private Sub AddCompanies()
On Error Resume Next
Dim rs As ADODB.Recordset

    If deIms.rsCOMPANY.State Then
        Set rs = deIms.rsCOMPANY.Clone
    Else
        deIms.Company (deIms.NameSpace)
        
        Set rs = deIms.rsCOMPANY.Clone
        deIms.rsCOMPANY.Close
    End If
    
    rs.Filter = "com_actvflag <> 0"
    ssdcboCompany.FieldSeparator = Chr$(1)
    
    rs.MoveFirst
    If rs.RecordCount = 0 Then Exit Sub
    
    Do Until rs.EOF
        ssdcboCompany.AddItem rs("com_name") & Chr$(1) & rs("com_compcode")
        
        rs.MoveNext
    Loop
    
End Sub

'get crystal report parameters

Private Function CreateRpti() As RPTIFileInfo

    With CreateRpti
        ReDim .parameters(1)
        .ReportFileName = reportPath & "wareSISE.rpt"
        .parameters(0) = "transnumb=" & cbo_Transaction
        .parameters(1) = "namespace=" & deIms.NameSpace
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("wareSISE.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
    
    End With

End Function


Private Sub ClearScreen()
        
        ssdcboCompany = ""
        ssdcboWarehouse = ""
        lblType = ""
        ssdcboStockNumb = ""
        lblEntyNumb = ""
        ssdcboCondition = ""
        lblCurrSap = ""
        txtNewSap = ""
        txtRemarks = ""
        
End Sub

Private Sub txtRemarks_GotFocus()
Call HighlightBackground(txtRemarks)
End Sub

Private Sub txtRemarks_LostFocus()
Call NormalBackground(txtRemarks)

End Sub
