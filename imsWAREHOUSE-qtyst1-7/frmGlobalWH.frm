VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGlobalWH 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Tag             =   "02050801"
   Begin VB.TextBox searchFIELD 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1410
   End
   Begin VB.TextBox searchFIELD 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2880
      Width           =   6210
   End
   Begin VB.CommandButton searchButton 
      Caption         =   "Search"
      Height          =   255
      Left            =   2145
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2910
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show &Remarks"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox quantityBOX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   5400
      MousePointer    =   1  'Arrow
      TabIndex        =   29
      Text            =   "quantityBOX"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox savingLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5520
      ScaleHeight     =   945
      ScaleWidth      =   3105
      TabIndex        =   27
      Top             =   3360
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
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   12255
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton saveBUTTON 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10560
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton newBUTTON 
      Caption         =   "&New Transaction"
      Height          =   375
      Left            =   8880
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   4485
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   4485
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   4485
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
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
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   4485
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4485
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   4485
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   2
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   5
      Left            =   6240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
      Index           =   6
      Left            =   6240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
      Index           =   4
      Left            =   6240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid STOCKlist 
      Height          =   3540
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   6244
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   9
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   0
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid matrix 
      Height          =   735
      Left            =   8520
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox remarks 
      Height          =   3255
      Left            =   120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   3480
      Width           =   14055
   End
   Begin VB.CheckBox checkAll 
      Caption         =   "Transfer all items"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Label remarksLabel 
      Caption         =   "Remarks:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "To select rows individualy, uncheck ""Transfer all items"" and right click on them"
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label label 
      Caption         =   "To Namespace"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   20
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "To Warehouse"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   19
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "To Company"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   18
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "From Warehouse"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "From Company"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "From Namespace"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "Transaction"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmGlobalWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim thisFORM As FormMode
Dim usingARROWS As Boolean
Dim doChanges As Boolean
Dim inProgress As Boolean
Dim isReset As Boolean
Dim ratioList() As Double

Public stocknumb As String
Public stockDESC As String
Public FromWH As String
Public ToWH As String
Public WH As String
Public fromlogic As String
Public fromSubLoca As String
Public toLOGIC As String
Public toSUBLOCA As String
Public condition As String
Public unitPRICE As Double
Public serial As String
Public Item, item2
Public qty1 As Double
Public qty2 As Double
Public col As Integer

Public NP As String
Public CompCode As String
Public ToCompCode As String
Public fromCompCode As String
Public TranType As String
Public datay As New ADODB.Recordset
Public goAhead As Boolean

Function allSelected() As Boolean
allSelected = True
Dim i As Integer
    For i = 2 To 6
        If cell(i) = "" Then
            allSelected = False
            Exit For
        End If
    Next
End Function

Function doReceiptSide(Item) As Boolean
Dim retval As Boolean
retval = False
    retval = PutReceiptDetail(Item)
    If retval = False Then
        doReceiptSide = retval
        Exit Function
    End If
    retval = Update_Sap(NP, CompCode, stocknumb, ToWH, qty1, CDbl(1), unitPRICE, condition, CurrentUser, cn)
    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, stockDESC, CurrentUser, cn)
    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, CurrentUser, cn)
    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, toSUBLOCA, CurrentUser, cn)
    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, toSUBLOCA, condition, CurrentUser, cn)
    If serial = "" Or UCase(serial) = "POOL" Then
        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, toSUBLOCA, condition, Format(Transnumb), CDbl(Item), ToWH, "AE", CompCode, FromWH, Format(Transnumb), CompCode, CDbl(Item), CurrentUser, cn)
    Else
        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, toSUBLOCA, condition, serial, CurrentUser, cn)
        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, ToWH, qty1, qty2, toLOGIC, toSUBLOCA, condition, Format(Transnumb), FromWH, Val(serial), ToWH, "AE", CompCode, Format(Transnumb), CompCode, CDbl(Item), serial, CurrentUser, cn)
    End If
    doReceiptSide = retval
End Function

Sub generalCheck()
Dim sql As String
Dim datax As ADODB.Recordset
On Error Resume Next
    'Issue side
    'logical warehouse to
    sql = "select 1 from logwar " _
        + "where lw_code = 'GENERAL' and lw_npecode = '" + cell(1).tag + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into logwar " _
            + "(lw_code,lw_npecode,lw_desc,lw_actvflag,lw_type) " _
            + "values ('GENERAL', '" + cell(1).tag + "', 'GENERAL', 0, 'ACTUAL')"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'sublocation to
    sql = "select 1 from sublocation " _
        + "where sb_code = 'GENERAL' and sb_npecode = '" + cell(1).tag + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into sublocation " _
            + "(sb_code,sb_npecode,sb_desc,sb_actvflag) " _
            + "values ('GENERAL', '" + cell(1).tag + "', 'GENERAL', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'location to
    sql = "select 1 from location " _
        + "where loc_npecode = '" + cell(1).tag + "' and loc_compcode= '" + cell(2).tag + "' and loc_locacode ='IN-TRANSIT'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into location " _
            + "(loc_locacode,loc_npecode,loc_name,loc_compcode,loc_gender,loc_actvflag) " _
            + "values ('IN-TRANSIT', '" + cell(1).tag + "', 'In Transit','" + cell(2).tag + "','BASE', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'company to
    sql = "select 1 from company " _
        + "where com_npecode = '" + cell(1).tag + "' and com_compcode='IN-TRANSIT'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into company " _
            + "(com_compcode,com_npecode,com_name,com_adr1,com_city,com_ctry,com_actvflag) " _
            + "values ('IN-TRANSIT', '" + cell(1).tag + "', 'On Transit', 'N/A', 'N/A','N/A', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'Receipt side
    'logical warehouse to
    sql = "select 1 from logwar " _
        + "where lw_code = 'GENERAL' and lw_npecode = '" + cell(4).tag + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into logwar " _
            + "(lw_code,lw_npecode,lw_desc,lw_actvflag,lw_type) " _
            + "values ('GENERAL', '" + cell(4).tag + "', 'GENERAL', 0, 'ACTUAL')"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'sublocation to
    sql = "select 1 from sublocation " _
        + "where sb_code = 'GENERAL' and sb_npecode = '" + cell(4).tag + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into sublocation " _
            + "(sb_code,sb_npecode,sb_desc,sb_actvflag) " _
            + "values ('GENERAL', '" + cell(4).tag + "', 'GENERAL', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'location to
    sql = "select 1 from location " _
        + "where loc_npecode = '" + cell(4).tag + "' and loc_compcode= '" + cell(5).tag + "' and loc_locacode ='IN-TRANSIT'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into location " _
            + "(loc_locacode,loc_npecode,loc_name,loc_compcode,loc_gender,loc_actvflag) " _
            + "values ('IN-TRANSIT', '" + cell(4).tag + "', 'In Transit','" + cell(5).tag + "','BASE', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
    
    'company to
    sql = "select 1 from company " _
        + "where com_npecode = '" + cell(4).tag + "' and com_compcode='IN-TRANSIT'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into company " _
            + "(com_compcode,com_npecode,com_name,com_adr1,com_city,com_ctry,com_actvflag) " _
            + "values ('IN-TRANSIT', '" + cell(4).tag + "', 'On Transit', 'N/A', 'N/A','N/A', 0)"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
End Sub

Sub enableCells(Value As Boolean)
Dim i As Integer
    For i = 1 To 6
        cell(i).Enabled = Value
    Next
End Sub

Sub hideREMARKS()
    Command5.Caption = "Show Remarks"
    remarks.Visible = False
    STOCKlist.Visible = True
End Sub

Function inList(list() As String, Value) As Boolean
Dim i As Integer
Dim thing
inList = False
    For thing = LBound(list) To UBound(list)
        If Value = list(thing) Then
            inList = True
            Exit Function
        End If
    Next
End Function

Sub makeLists()
    With STOCKlist
        .cols = 10
        .TextMatrix(0, 0) = "#"
        .ColWidth(0) = 400
        .TextMatrix(0, 1) = "Commodity"
        .ColWidth(1) = 1400
        .TextMatrix(0, 2) = "Description"
        .ColWidth(2) = 3200
        
        
        
        
        
        
        .TextMatrix(0, 3) = "Unit Price"
        .ColWidth(3) = 1400
        .ColAlignmentFixed(3) = 7
        .TextMatrix(0, 4) = "Logical WH"
        .ColWidth(4) = 1000
        .ColAlignmentFixed(4) = 7
        .TextMatrix(0, 5) = "Sub Location"
        .ColWidth(5) = 1000

        .TextMatrix(0, 6) = "Prim. Qty"
        .ColWidth(6) = 1000
        .ColAlignmentFixed(6) = 7
        .TextMatrix(0, 7) = "Prim. Unit"
        .ColWidth(7) = 800
        .ColAlignmentFixed(7) = 4
        .TextMatrix(0, 8) = "Sec. Qty"
        .ColWidth(8) = 1000
        .ColAlignmentFixed(8) = 7
        .TextMatrix(0, 9) = "Sec. Unit"
        .ColWidth(9) = 800
        .ColAlignmentFixed(9) = 4
    End With
End Sub


Function saveEntry(oldItem) As Boolean
    stocknumb = .TextMatrix(oldItem, 1)
    stockDESC = .TextMatrix(oldItem, 2)
    unitPRICE = .TextMatrix(oldItem, 3)

    saveEntry = False
    NP = cell(4).tag
    CompCode = cell(5).tag
    ToWH = cell(6).tag
    toLOGIC = "GENERAL"
    toSUBLOCA = "GENERAL"
    fromCompCode = "GENERAL"
    FromWH = "in-transit"
    fromlogic = "GENERAL"
    fromSubLoca = "GENERAL"
    Call stockNumberCheck(stocknumb)
    saveEntry = doReceiptSide(oldItem)
    If saveEntry = False Then
        Call RollbackTransaction(cn)
        MsgBox "Error in Transaction - Entry side"
        Screen.MousePointer = 0
        savingLABEL.Visible = False
        Me.Enabled = True
        Exit Function
    End If
End Function

Sub showREMARKS()
    Command5.Caption = "Hide Remarks"
    quantityBOX.Visible = False
    STOCKlist.Visible = False
    remarks.Visible = True
    remarks.SetFocus
End Sub

Sub stockNumberCheck(StockNumber As String)
Dim datax As ADODB.Recordset
Dim sql As String
    sql = "select 1 from stockmaster " _
        + "where stk_stcknumb = '" + StockNumber + "' and stk_npecode = '" + cell(4).tag + "' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        sql = "insert into stockmaster (stk_stcknumb,stk_npecode,stk_desc,stk_descflag, " _
            + "stk_primuon,stk_secouom,stk_stcktype,stk_catecode,stk_poolspec,stk_compfctr, " _
            + "stk_hazmatclau,stk_mini,stk_maxi,stk_characctcode,stk_stdrcost,stk_estmprice, " _
            + "stk_grpe,stk_imge,stk_techspec,stk_flag,stk_eccnid,stk_eccnlicsreq,stk_eccnsourceid, " _
            + "stk_ratio1,stk_ratio2) " _
            + "select stk_stcknumb,'" + cell(4).tag + "', stk_desc,stk_descflag,stk_primuon, " _
            + "stk_secouom,stk_stcktype,stk_catecode,stk_poolspec,stk_compfctr,stk_hazmatclau, " _
            + "stk_mini,stk_maxi,stk_characctcode,stk_stdrcost,stk_estmprice,stk_grpe, " _
            + "stk_imge,stk_techspec,stk_flag,stk_eccnid,stk_eccnlicsreq,stk_eccnsourceid, " _
            + "stk_ratio1,stk_ratio2 from stockmaster Where " _
            + "stk_stcknumb = '" + StockNumber + "' and stk_npecode = '" + cell(1).tag + "' "
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenStatic
    End If
End Sub

Private Sub cell_Change(Index As Integer)
Dim n As Integer
    If saveBUTTON.Enabled = False Then Exit Sub
    If Not directCLICK Then
        If Index = 0 Then
            n = 0
        Else
            n = 1
        End If
        Call alphaSEARCH(cell(Index), combo(Index), n)
    Else
        directCLICK = False
    End If
End Sub

Private Sub cell_Click(Index As Integer)
Dim datax As New ADODB.Recordset
Dim sql As String
Dim i
If Index = 1 Then Exit Sub
Screen.MousePointer = 11
    With cell(Index)
        If saveBUTTON.Enabled Or Index = 0 Then
            If Index > 2 And Index <> 4 Then
                If Index = 3 Then
                    Call cleanSTOCKlist
                End If
                If combo(Index - 1) = "" And (Index - 1) > 1 Then
                    MsgBox "Please select " + label(Index - 1) + " first"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                If Index = 2 Then
                    cell(3) = ""
                    Call cleanSTOCKlist
                Else
                    Select Case Index
                        Case 4
                            cell(5) = ""
                            cell(6) = ""
                        Case 5
                            cell(6) = ""
                    End Select
                End If
            End If
            .SelStart = 0
            .SelLength = Len(.text)
            If Not (saveBUTTON.Enabled And Index = 0) Then
                Call showCOMBO(combo(Index), Index)
            End If
        End If
        Screen.MousePointer = 0
    End With
Screen.MousePointer = 0
End Sub

Sub hideCombos()
Dim i As Integer
    For i = 2 To 6
        combo(i).Visible = False
    Next
End Sub


Sub showCOMBO(ByRef grid As MSHFlexGrid, Index)
    If Index = 1 Then Exit Sub
    With grid
        Call fillCOMBO(grid, Index)
        If .Rows > 0 Then
            .Visible = True
            .ZOrder
            If Index < 5 Then .Top = cell(Index).Top + cell(Index).Height + 20
        End If
        .MousePointer = 0
    End With
End Sub
Sub fillCOMBO(ByRef grid As MSHFlexGrid, Index)
On Error Resume Next
Dim sql
Dim i
Dim datax As New ADODB.Recordset
Dim addCOMBO As Boolean
Dim namespaceVal, companyVal As String
    Err.Clear
    With combo(Index)
        .Rows = 2
        .RowHeightMin = 240
        .TextMatrix(1, 0) = ""
        If Index = 0 Then
            .cols = 1
            .TextMatrix(0, 0) = "Transaction ·"
            .ColWidth(0) = cell(0).width
            .ColAlignment(0) = 0
        Else
            .cols = 2
            .TextMatrix(0, 0) = "Description"
            .TextMatrix(0, 1) = "Code"
            .ColWidth(0) = 2800
            .ColAlignment(0) = 0
            .ColWidth(1) = 1400
            .ColAlignment(1) = 0
        End If
    End With
    
    Err.Clear
    If Index < 4 Then
        namespaceVal = cell(1).tag
        companyVal = cell(2).tag
    Else
        namespaceVal = cell(4).tag
        companyVal = cell(5).tag
    End If
    Select Case Index
        Case 0
            sql = "SELECT transac FROM issuesAndReceipts " _
                + "WHERE npecode = '" + namespaceVal + "' AND transac like 'GT-%'" _
                + "ORDER BY transac desc "
        Case 4
            sql = "select npce_name as namespaceName, npce_code as namespace from namespace " _
                + "order by namespaceName "
        Case 2, 5
            sql = "select com_name as companyName, com_compcode as company from company " _
                + "where com_npecode = '" + namespaceVal + "' and com_actvflag = 1 " _
                + "order by companyName "
        Case 3, 6
            sql = "select loc_name as locationName, loc_locacode as location from location " _
                + "where loc_npecode = '" + namespaceVal + "' and loc_compcode = '" + companyVal + "' " _
                + "and loc_gender='BASE' and loc_actvflag=1 " _
                + "order by locationName "
    End Select
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount < 1 Then Exit Sub
    Call doCOMBO(Index, datax)
    Set datax = New ADODB.Recordset
End Sub
Sub doCOMBO(Index, datax As ADODB.Recordset)
Dim i, extraW
Dim t As String
    Err.Clear
    With frmGlobalWH.combo(Index)
        Do While Not datax.EOF
            If Index = 0 Then
                .AddItem Trim(datax.Fields(0))
            Else
                .AddItem Trim(datax.Fields(0)) + vbTab + Trim(datax.Fields(1))
            End If
            datax.MoveNext
        Loop
        If .TextMatrix(1, 0) = "" Then .RemoveItem (1)
        .row = 1
        If .Rows < 9 Then
            extraW = 0
            .Height = (350 * .Rows)
            .ScrollBars = flexScrollBarNone
        Else
            extraW = 280
            .Height = 2340
            .ScrollBars = flexScrollBarVertical
        End If
        If frmGlobalWH.cell(Index).width > (.width + extraW) Then
            .width = frmGlobalWH.cell(Index).width
            .ColWidth(0) = .ColWidth(0) + (.width - .width) - extraW
        Else
            .width = .width + extraW
        End If
        If (frmGlobalWH.cell(Index).Left + .width) > frmGlobalWH.width Then
            .Left = frmGlobalWH.width - .width - 300
        Else
            .Left = frmGlobalWH.cell(Index).Left - 100
        End If
    End With
End Sub
Private Sub cell_GotFocus(Index As Integer)
    If Index = 1 Then Exit Sub
    If saveBUTTON.Enabled Or Index = 0 Then
        If Not (saveBUTTON.Enabled And Index = 0) Then
            With cell(Index)
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
        If Not .locked Then
                Select Case KeyCode
                    Case 27
                        combo(Index).Visible = False
                    Case 40
                        Call arrowKEYS("down", Index)
                    Case 38
                        Call arrowKEYS("up", Index)
                    Case Else
                    Dim col
                End Select
        End If
    End With
End Sub

Sub arrowKEYS(direction As String, Index As Integer)
Dim grid As MSHFlexGrid
    With cell(Index)
        Set grid = combo(Index)
            grid.Visible = True
            Call gridCOLORnormal(grid, Val(grid.tag))
            Select Case direction
                Case "down"
                    If grid.row < (grid.Rows - 1) Then
                        If grid.row = 0 And .text = "" Then
                            .text = grid.text
                        Else
                            grid.row = grid.row + 1
                        End If
                    Else
                        grid.row = grid.Rows - 1
                    End If
                Case "up"
                    If grid.row > 0 Then
                        grid.row = grid.row - 1
                    Else
                        grid.row = 1
                    End If
            End Select
            
            grid.tag = grid.row
            If Not grid.Visible Then
                grid.Visible = True
            End If
            grid.ZOrder
            grid.TopRow = IIf(grid.row = 0, 1, grid.row)
            usingARROWS = True
            Call gridCOLORdark(grid, grid.row)
            grid.SetFocus
    End With
End Sub
Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i, t, n
Dim gotIT As Boolean
    With cell(Index)
        Select Case KeyAscii
            Case 13
                KeyAscii = 0
                If Not .locked Then
                    justCLICK = False
                    gotIT = False
                    If Index = 4 Or Index = 0 Then
                        n = 0
                    Else
                        n = 1
                    End If
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
    If Index = 1 Then Exit Sub
    If usingARROWS Then
        usingARROWS = False
    Else
        If saveBUTTON.Enabled Or Index = 0 Then
            If Not (saveBUTTON.Enabled And Index = 0) Then
                If activeCELL <> 1 Then combo(activeCELL).Visible = False
            End If
        End If
    End If
    If saveBUTTON.Enabled Or Index = 0 Then
        With cell(Index)
            .backcolor = vbWhite
        End With
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 1 Then
        If findSTUFF(cell(Index), combo(Index), 0) = 0 Then cell(Index) = ""
    End If
End Sub

Private Sub checkAll_Click()
    If checkAll.Value Then
        STOCKlist.Enabled = False
    Else
        STOCKlist.Enabled = True
    End If
End Sub

Private Sub combo_Click(Index As Integer)
Dim i, sql, t
Dim datax As New ADODB.Recordset
Dim currentformname, currentformname1
Dim MSGBOXReply As VbMsgBoxResult
Dim labelname As String
Dim ratio As Integer
Dim showIt As Boolean
    combo(Index).Visible = False
    DoEvents
    Screen.MousePointer = 11
    DoEvents
    directCLICK = True
    Set datax = New ADODB.Recordset
    DoEvents
    showIt = False
    With combo(Index)
        If Index = 0 Then
            cell(0) = Trim(.TextMatrix(.row, 0))
        Else
            cell(Index) = Trim(.TextMatrix(.row, 0))
            cell(Index).tag = Trim(.TextMatrix(.row, 1))
        End If
        Select Case Index
            Case 0
                sql = "select * from issuesAndReceiptsDetail " _
                    + "where transac = '" + cell(0) + "' order by item "
            Case 3
                sql = "select * from stockinfoExtended where " _
                    + "NameSpace = '" + cell(1).tag + "' " _
                    + "and Company = '" + cell(2).tag + "' " _
                    + "and Location = '" + cell(3).tag + "' " _
                    + "order by stocknumber, Logic, SubLoca "
        End Select
        If sql = "" Then
        Else
            Set datax = New ADODB.Recordset
            datax.Open sql, cn, adOpenStatic
            If datax.RecordCount > 0 Then
                STOCKlist.Enabled = True
                Call cleanSTOCKlist
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
                Call fillSTOCKlist(datax)
                If Index = 0 And saveBUTTON.Enabled = False Then
                    sql = "select ii_npecode, npce_name, ii_compcode, com_name, ii_ware, loc_name from " _
                        + "invtissue , NameSpace, Company, Location " _
                        + "where ii_npecode = npce_code and " _
                        + "ii_npecode = com_npecode and ii_compcode = com_compcode and " _
                        + "ii_npecode = loc_npecode And ii_ware = loc_locacode and " _
                        + "ii_trannumb = '" + cell(0) + "' "
                    Set datax = New ADODB.Recordset
                    datax.Open sql, cn, adOpenForwardOnly
                    If datax.RecordCount > 0 Then
                        cell(1) = datax!npce_name
                        cell(1).tag = datax!ii_npecode
                        cell(2) = datax!com_name
                        cell(3) = datax!loc_name
                    End If
                    sql = "select ir_npecode, npce_name, ir_compcode, com_name, ir_ware, loc_name from " _
                        + "invtreceipt , NameSpace, Company, Location " _
                        + "where ir_npecode = npce_code and " _
                        + "ir_npecode = com_npecode and ir_compcode = com_compcode and " _
                        + "ir_npecode = loc_npecode And ir_ware = loc_locacode and " _
                        + "ir_trannumb = '" + cell(0) + "' "
                    Set datax = New ADODB.Recordset
                    datax.Open sql, cn, adOpenForwardOnly
                    If datax.RecordCount > 0 Then
                        cell(4) = datax!npce_name
                        cell(4).tag = datax!ir_npecode
                        cell(5) = datax!com_name
                        cell(6) = datax!loc_name
                    End If
                    sql = "select * from INVTISSUEREM " _
                        + "where iir_trannumb = '" + cell(0) + "' "
                    Set datax = New ADODB.Recordset
                    datax.Open sql, cn, adOpenForwardOnly
                    If datax.RecordCount > 0 Then
                        remarks = datax!iir_remk
                    End If
                End If
                If savingLABEL.Visible Then
                    Label3 = "SAVING..."
                    savingLABEL.Visible = False
                End If
            End If
        End If
        .Visible = False
        Dim nextVal As Integer
        nextVal = Index + 1
        If nextVal > 6 Then checkAll.SetFocus
        cell(Index).SetFocus
    End With
    Screen.MousePointer = 0
End Sub

Sub fillSTOCKlist(datax As ADODB.Recordset)
On Error GoTo errorHandler
Dim n, rec, i, qty2Value, lineNumber
Dim firstTime As Boolean
firstTime = True
lineNumber = 0
ReDim ratioList(datax.RecordCount) As Double
onDetailListInProcess = True
    With datax
        n = 0
        STOCKlist.Rows = 2
        STOCKlist.row = 1
        STOCKlist.col = 0
        STOCKlist.CellFontName = "MS Sans Serif"
        mainItemRow = 0
        Do While Not .EOF
            n = n + 1
            rec = ""
            STOCKlist.ColAlignment(0) = 7
            STOCKlist.ColAlignment(1) = 0
            STOCKlist.ColAlignment(4) = 0
            STOCKlist.ColAlignment(5) = 0
            STOCKlist.ColAlignment(6) = 7
            STOCKlist.ColAlignment(7) = 4
            STOCKlist.ColAlignment(8) = 7
            STOCKlist.ColAlignment(9) = 4
            rec = rec + Format(n) + vbTab
            rec = rec + Trim(!StockNumber) + vbTab
            rec = rec + Trim(!description) + vbTab
            rec = rec + Format(!unitPRICE, "#,###,##0.00") + vbTab
            rec = rec + Trim(!logic) + vbTab
            rec = rec + Trim(!subloca) + vbTab
            rec = rec + Format(!qty, "0.00") + vbTab
            rec = rec + IIf(IsNull(!unit), "", !unit) + vbTab
            rec = rec + Format(!qty2, "0.00") + vbTab
            rec = rec + IIf(IsNull(!unit2), "", !unit2) + vbTab
            STOCKlist.AddItem rec
            ratioList(n) = IIf(!qty = 0, 1, !qty2 / !qty)
            If n = 20 Then
                DoEvents
                STOCKlist.Refresh
            End If
            .MoveNext
        Loop

        If STOCKlist.Rows > 2 Then STOCKlist.RemoveItem (1)
        STOCKlist.RowHeightMin = 240
        STOCKlist.row = 0
    End With
    
errorHandler:
If Err.Number > 0 Then
    'MsgBox "fillSTOCKlist " + Err.description
    Err.Clear
    Resume Next
End If
End Sub
Private Sub combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    justCLICK = False
    With cell(Index)
        If Not .locked Then
            Select Case KeyCode
                Case 27
                    combo(Index).Visible = False
                Case 40
                    Call arrowKEYS("down", Index)
                Case 38
                    Call arrowKEYS("up", Index)
                Case Else
                Dim col
            End Select
        End If
    End With
End Sub


Private Sub combo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call combo_Click(Index)
        Case 27
            combo(Index).Visible = False
            Exit Sub
    End Select
    combo(Index).Visible = False
    If Index > 0 Then
        cell(Index + 1).SetFocus
        Call cell_Click(Index + 1)
    End If
End Sub


Private Sub combo_LostFocus(Index As Integer)
    combo(Index).Visible = False
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim reportPATH, cnSTRING, text
Screen.MousePointer = 11
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
    With CrystalReport1
        .Reset
        .LogOnServer "pdsodbc.dll", dsnF, dsnDSQ, dsnUID, dsnPWD
        reportPATH = repoPATH + "\"
        .ReportFileName = reportPATH & "wareGlobalTransferIssue.rpt"
        .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
        .ParameterFields(1) = "NAMESPACE;" & cell(1).tag & ";TRUE"
        Set thisrepo = CrystalReport1
        mainREPORT = True
        Call Translate_Reports(CrystalReport1.ReportFileName)
        Call Translate_SubReports
        .Action = 1
        
        .Reset
        .LogOnServer "pdsodbc.dll", dsnF, dsnDSQ, dsnUID, dsnPWD
        .ReportFileName = reportPATH & "wareGlobalTransferEntry.rpt"
        .ParameterFields(0) = "transnumb;" & cell(0) & ";TRUE"
        .ParameterFields(1) = "NAMESPACE;" & cell(4).tag & ";TRUE"
        Set thisrepo = CrystalReport1
        mainREPORT = True
        Call Translate_Reports(CrystalReport1.ReportFileName)
        Call Translate_SubReports
        .Action = 1
        
        .Reset
    End With
Screen.MousePointer = 0
End Sub

Private Sub Command5_Click()
    If STOCKlist.Visible Then
        Call showREMARKS
    Else
        Call hideREMARKS
    End If
End Sub

Private Sub Form_Activate()
Dim rights As Boolean
    inProgress = False
    Screen.MousePointer = 0
    rights = Getmenuuser(nameSP, CurrentUser, Me.tag, cn)
    newBUTTON.Enabled = rights
    Me.Visible = True
    If newBUTTON.Enabled Then newBUTTON.SetFocus
    Me.Refresh
    Call makeLists
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim datax As New ADODB.Recordset
    sql = "select npce_name from NAMESPACE where npce_code = '" + nameSP + "'"
    datax.source = sql
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        nameSPname = datax!npce_name
        datax.Close
    End If
    cell(1).text = nameSPname
    cell(1).tag = nameSP
    makeLists
    With frmGlobalWH
        .Left = Round((Screen.width - .width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub newBUTTON_Click()
Dim i
    isReset = True
    Call cleanSTOCKlist
    remarks = ""
    saveBUTTON.Enabled = True
    newBUTTON.Enabled = False
    Call enableCells(True)
    cell(0).backcolor = &HFFFFC0
    cell(0) = ""
    cell(2).SetFocus
    Call cell_Click(2)
End Sub


Sub cleanSTOCKlist()
Dim i
    With STOCKlist
        .Rows = 2
        For i = 0 To .cols - 1
            .TextMatrix(1, i) = ""
        Next
        .RowHeightMin = 0
        .RowHeight(1) = 0
    End With
End Sub



Private Sub quantityBOX_Click()
    With quantityBOX
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub


Private Sub quantityBOX_GotFocus()
    With quantityBOX
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub quantityBOX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call quantityBOX_Validate(False)
    End If
End Sub


Private Sub quantityBOX_LostFocus()
    quantityBOX.Visible = False
End Sub

Private Sub quantityBOX_Validate(Cancel As Boolean)
Dim qty As Double
Dim qty2 As Double
Dim newQty As Double
Dim newQty2 As Double
Dim ratioValue As Double
'On Error Resume Next
    If col = 6 Or col = 8 Then
        If IsNumeric(STOCKlist.TextMatrix(STOCKlist.row, 6)) Then
            qty = CDbl(STOCKlist.TextMatrix(STOCKlist.row, 6))
        Else
            qty = 0
        End If
        If IsNumeric(STOCKlist.TextMatrix(STOCKlist.row, 8)) Then
            qty2 = CDbl(STOCKlist.TextMatrix(STOCKlist.row, 8))
        Else
            qty2 = 0
        End If
        With quantityBOX
            If IsNumeric(.text) Then
                If CDbl(.text) > 0 Then
                    .text = Format(.text, "0.00")
                Else
                    .text = "0.00"
                End If
            Else
                .text = "0.00"
            End If
            If qty1 <> qty2 Then
                If IsNumeric(.text) Then
                    ratioValue = ratioList(STOCKlist.row)
                    If col = 6 Then
                        newQty = CDbl(.text)
                        newQty2 = newQty * ratioValue
                    Else
                        newQty2 = CDbl(.text)
                        newQty = newQty2 / ratioValue
                    End If
                End If
            Else
                newQty = CDbl(.text)
                newQty2 = newQty
            End If
            .SelStart = Len(.text)
            STOCKlist.TextMatrix(STOCKlist.row, 6) = Format(newQty, "###,####,###0.00")
            STOCKlist.TextMatrix(STOCKlist.row, 8) = Format(newQty2, "###,####,###0.00")
        End With
    End If
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


Private Sub saveBUTTON_Click()
Dim i, ii
Dim oldItem As String
Dim retval As Boolean
Dim refStockNumber, refCompany As String
Dim sql As String
Dim fullQty1, fullQty2 As Double
Dim datax As New ADODB.Recordset
    Screen.MousePointer = 11
    If remarks = "" Then
        Call showREMARKS
        Screen.MousePointer = 0
        MsgBox "Please include the remarks for this transaction"
        remarks.backcolor = &HC0FFFF
        remarks.SetFocus
        Exit Sub
    End If
    Call hideREMARKS
    
    If Not allSelected Then
        MsgBox "Please select all fields"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    savingLABEL.Visible = True
    savingLABEL.ZOrder
    Me.Enabled = False
    Me.Refresh
    Call BeginTransaction(cn)
    If Not retval Then Call RollbackTransaction(cn)
    TranType = "GT"
    'ISSUE side
    Call generalCheck
    retval = PutIssue("GT")
    Call PutIssueRemarks
    If retval = False Then
        Call RollbackTransaction(cn)
        MsgBox "Error in Transaction - Issue header"
        Screen.MousePointer = 0
        savingLABEL.Visible = False
        Me.Enabled = True
        Exit Sub
    End If
    retval = putReceipt("GT")
    If retval = False Then
        Call RollbackTransaction(cn)
        MsgBox "Error in Transaction - Entry header"
        Screen.MousePointer = 0
        savingLABEL.Visible = False
        Me.Enabled = True
        Exit Sub
    End If
    sql = "select * from location " _
        + "where loc_npecode = '" + cell(1).tag + "' " _
        + "and loc_compcode = '" + cell(2).tag + "' " _
        + "and loc_locacode = 'in-transit' "
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount > 0 Then
    Else
        Call RollbackTransaction(cn)
        MsgBox "It is required to have an 'in-transit' location"
        Screen.MousePointer = 0
        savingLABEL.Visible = False
        Me.Enabled = True
        Exit Sub
    End If
    goAhead = True
    refStockNumber = ""
    With STOCKlist
        Item = 0
        For i = 1 To .Rows - 1
            If checkAll.Value = False Then
                .row = i
                .col = 1
                If .CellBackColor = vbWhite Then
                    goAhead = False
                Else
                    goAhead = True
                End If
            End If
            If goAhead Then
                NP = cell(1).tag
                CompCode = cell(2).tag
                WH = cell(3).tag
                stocknumb = .TextMatrix(i, 1)
                stockDESC = .TextMatrix(i, 2)
                unitPRICE = .TextMatrix(i, 3)
                fromlogic = .TextMatrix(i, 4)
                fromSubLoca = .TextMatrix(i, 5)
                FromWH = cell(3).tag
                ToCompCode = cell(2).tag
                sql = "select * from stockinfoExtended " _
                    + "where NameSpace = '" + NP + "' " _
                    + "and Company = '" + CompCode + "' " _
                    + "and Location = '" + WH + "' " _
                    + "and logic = '" + fromlogic + "' " _
                    + "and subloca = '" + fromSubLoca + "' " _
                    + "and StockNumber = '" + stocknumb + "' "
                Set datax = New ADODB.Recordset
                datax.Open sql, cn, adOpenStatic
                If datax.RecordCount > 0 Then
                    'ISSUE side
                    Item = Item + 1
                    ToWH = "in-transit"
                    toLOGIC = "GENERAL"
                    toSUBLOCA = "GENERAL"
                    condition = datax!condition
                    qty1 = CDbl(.TextMatrix(i, 6))
                    qty2 = CDbl(.TextMatrix(i, 8))
                    
                    serial = datax!serialNumber

                    If qty1 = 0 Or qty2 = 0 Then
                        MsgBox "No zero values are allowed on either primary or secondary quantity"
                        Screen.MousePointer = 0
                        savingLABEL.Visible = False
                        Me.Enabled = True
                        Exit Sub
                    End If
                    condition = datax!condition
                    retval = PutIssueDetail(i)
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Screen.MousePointer = 0
                        savingLABEL.Visible = False
                        Me.Enabled = True
                        Exit Sub
                    End If
                    qty2 = qty2 * -1
                    qty1 = qty1 * -1
                    retval = retval And Quantity_In_stock1_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, stockDESC, CurrentUser, cn)
                    retval = retval And Quantity_In_stock2_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, CurrentUser, cn)
                    retval = retval And Quantity_In_stock3_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, fromSubLoca, CurrentUser, cn)
                    retval = retval And Quantity_In_stock4_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, fromSubLoca, condition, CurrentUser, cn)
                    If serial = "" Or UCase(serial) = "POOL" Then
                        retval = retval And Quantity_In_stock5_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, fromSubLoca, condition, Format(Transnumb), CDbl(i), ToWH, "GT", CompCode, ToWH, Format(Transnumb), CompCode, CDbl(i), CurrentUser, cn)
                    Else
                        retval = retval And Quantity_In_stock6_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, fromSubLoca, condition, serial, CurrentUser, cn)
                        retval = retval And Quantity_In_stock7_Insert(NP, CompCode, stocknumb, FromWH, qty1, qty2, fromlogic, fromSubLoca, condition, Format(Transnumb), ToWH, CDbl(i), ToWH, "GT", CompCode, Format(Transnumb), CompCode, CDbl(i), serial, CurrentUser, cn)
                    End If
                    If retval = False Then
                        Call RollbackTransaction(cn)
                        MsgBox "Error in Transaction"
                        Screen.MousePointer = 0
                        savingLABEL.Visible = False
                        Me.Enabled = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        'Entry side
        If retval Then
            goAhead = False
            qty1 = 0
            qty2 = 0
            Dim stockNumberList() As String
            ReDim stockNumberList(STOCKlist.Rows)
            Dim rItem As Integer
            rItem = 0
            For i = 1 To .Rows - 1
                If checkAll.Value Then
                    goAhead = True
                Else
                    .row = i
                    .col = 1
                    If .CellBackColor = vbWhite Then
                        goAhead = False
                    Else
                        goAhead = True
                    End If
                End If
                If goAhead Then
                    stocknumb = .TextMatrix(i, 1)
                    If inList(stockNumberList, stocknumb) Then
                        qty1 = qty1 + CDbl(.TextMatrix(i, 6))
                        qty2 = qty2 + CDbl(.TextMatrix(i, 8))
                    Else
                        'if this is true it means that an item has to be saved
                        If qty1 > 0 Then
                            retval = saveEntry(oldItem)
                            qty1 = 0
                        End If
                        rItem = rItem + 1 'stocknumber for receipt transaction side
                        stockNumberList(rItem) = stocknumb 'adding stocknumber to list
                        qty1 = CDbl(.TextMatrix(i, 6))
                        qty2 = CDbl(.TextMatrix(i, 8))
                    End If
                    oldItem = i
                End If
            Next
            If qty1 > 0 Then
                retval = saveEntry(oldItem)
            End If
        End If
    End With
    
    If retval Then
        Call CommitTransaction(cn)
        cell(0) = Transnumb
        cell(0).tag = cell(0)
        cell(0).Visible = True
        combo(0).Visible = False
        combo(0).TextMatrix(1, 0) = Transnumb
    End If
    
    Call enableCells(False)
    If Err Then Err.Clear
    newBUTTON.Enabled = True
    saveBUTTON.Enabled = False
    savingLABEL.Visible = False
    savingLABEL.Visible = False
    Me.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
RollBack:
    Call RollbackTransaction(cn)
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Function PutIssueRemarks() As Boolean
Dim cmd As ADODB.Command
    Set cmd = getCOMMAND("InvtIssuetRem_Insert")
    cmd.parameters("@LineNumb") = 1
    cmd.parameters("@REMARKS") = remarks
    cmd.parameters("@TranNumb") = Transnumb
    cmd.parameters("@CompanyCode") = cell(2).tag
    cmd.parameters("@NAMESPACE") = cell(1).tag
    cmd.parameters("@WhareHouse") = cell(3).tag
    cmd.parameters("@USER") = CurrentUser
    Call cmd.Execute(0, , adExecuteNoRecords)
    PutIssueRemarks = cmd.parameters(0).Value = 0
End Function
Function putReceipt(prefix) As Integer
Dim v As Variant
    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceipt_Insert"
        .parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, cell(4).tag)
        .parameters.Append .CreateParameter("@COMPANYCODE", adChar, adParamInput, 10, cell(5).tag)
        .parameters.Append .CreateParameter("@WHAREHOUSE", adChar, adParamInput, 10, cell(6).tag)
        .parameters.Append .CreateParameter("@TRANS", adVarChar, adParamInput, 15, Transnumb)
        .parameters.Append .CreateParameter("@TRANTYPE", adChar, adParamInput, 2, prefix)
        .parameters.Append .CreateParameter("@TRANFROM", adVarChar, adParamInput, 10, cell(6).tag)
        .parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, Null)
        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, Null)
        .parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, CurrentUser)
        Call .Execute(Options:=adExecuteNoRecords)
        putReceipt = .parameters("RV") = 0
    End With
    If putReceipt Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function

Private Function PutIssue(prefix) As Boolean
Dim NP As String
Dim cmd As Command
On Error GoTo errPutIssue

    PutIssue = False
    Set cmd = getCOMMAND("InvtIssue_Insert")
    NP = cell(1).tag
    Transnumb = prefix + "-" & GetGlobalTransactionNumber
    cmd.parameters("@NAMESPACE") = NP
    cmd.parameters("@TRANTYPE") = prefix
    cmd.parameters("@COMPANYCODE") = cell(2).tag
    cmd.parameters("@TRANSNUMB") = Transnumb
    cmd.parameters("@ISSUTO") = cell(3).tag
    cmd.parameters("@SUPPLIERCODE") = Null
    cmd.parameters("@WHAREHOUSE") = cell(3).tag
    cmd.parameters("@STCKNUMB") = Null
    cmd.parameters("@COND") = Null
    cmd.parameters("@SAP") = Null
    cmd.parameters("@NEWSAP") = Null
    cmd.parameters("@ENTYNUMB") = Null
    cmd.parameters("@USER") = CurrentUser
    cmd.Execute
    PutIssue = cmd.parameters(0).Value = 0
    Exit Function

errPutIssue:
    MsgBox Err.description
    Err.Clear
End Function


Public Function GetGlobalTransactionNumber() As Long
    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "GetGlobalTransactionNumber"
        .parameters.Append .CreateParameter("@numb", adInteger, adParamOutput, 4, Null)
        Call .Execute(Options:=adExecuteNoRecords)
        GetGlobalTransactionNumber = .parameters("@numb").Value
    End With
    If GetGlobalTransactionNumber Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function

Function PutIssueDetail(row) As Boolean
Dim cmd As Command
On Error Resume Next
    PutIssueDetail = False
    Set cmd = getCOMMAND("InvtIssueDetl_INSERT")


        'Set the parameter values for the command to be executed.
        cmd.parameters("@iid_trannumb") = Transnumb
        cmd.parameters("@iid_compcode") = cell(2).tag
        cmd.parameters("@iid_npecode") = cell(1).tag
        cmd.parameters("@iid_ware") = cell(3).tag
        cmd.parameters("@iid_transerl") = Item
        cmd.parameters("@iid_stcknumb") = stocknumb
        cmd.parameters("@iid_ps") = IIf(serial = "", 1, 0)
        cmd.parameters("@iid_serl") = IIf(serial = "", Null, serial)
        cmd.parameters("@iid_newcond") = condition
        cmd.parameters("@iid_stcktype") = "I"
        cmd.parameters("@iid_ctry") = "US"
        cmd.parameters("@iid_tosubloca") = toSUBLOCA
        cmd.parameters("@iid_tologiware") = toLOGIC
        cmd.parameters("@iid_owle") = 1
        cmd.parameters("@iid_leasecomp") = Null
        cmd.parameters("@iid_primqty") = qty1
        cmd.parameters("@iid_secoqty") = qty2
        cmd.parameters("@iid_unitpric") = unitPRICE
        cmd.parameters("@iid_curr") = "USD"
        cmd.parameters("@iid_currvalu") = 1
        cmd.parameters("@iid_stckdesc") = stockDESC
        cmd.parameters("@iid_fromlogiware") = fromlogic
        cmd.parameters("@iid_fromsubloca") = fromSubLoca
        cmd.parameters("@iid_origcond") = condition
        cmd.parameters("@user") = CurrentUser

    'Execute the command.
    Call cmd.Execute(Options:=adExecuteNoRecords)
    PutIssueDetail = True
End Function


Private Function PutReceiptDetail(Item) As Boolean
    Dim psVALUE, serial
    Dim cmd As Command
    On Error GoTo errPutDataInsert
    PutReceiptDetail = False
    Set cmd = getCOMMAND("INVTRECEIPTDETL_INSERT")

    'Set the parameter values for the command to be executed.
    cmd.parameters("@ird_curr") = "USD"
    cmd.parameters("@ird_currvalu") = 1
    cmd.parameters("@ird_ponumb") = Null
    cmd.parameters("@ird_lirtnumb") = Null
    cmd.parameters("@ird_compcode") = cell(5).tag
    cmd.parameters("@ird_trannumb") = Transnumb
    cmd.parameters("@ird_npecode") = cell(4).tag
    cmd.parameters("@ird_ware") = cell(6).tag
    cmd.parameters("@ird_transerl") = Item
    cmd.parameters("@ird_stcknumb") = stocknumb
    cmd.parameters("@ird_ps") = IIf(serial = "", 1, 0)
    cmd.parameters("@ird_serl") = IIf(serial = "", Null, serial)
    cmd.parameters("@ird_newcond") = condition
    cmd.parameters("@ird_stcktype") = ""
    cmd.parameters("@ird_ctry") = "US"
    cmd.parameters("@ird_tosubloca") = toSUBLOCA
    cmd.parameters("@ird_tologiware") = toLOGIC
    cmd.parameters("@ird_owle") = 1
    cmd.parameters("@ird_leasecomp") = Null
    cmd.parameters("@ird_primqty") = qty1
    cmd.parameters("@ird_secoqty") = qty2
    cmd.parameters("@ird_unitpric") = unitPRICE
    cmd.parameters("@ird_stckdesc") = stockDESC
    cmd.parameters("@ird_fromlogiware") = fromlogic
    cmd.parameters("@ird_fromsubloca") = fromSubLoca
    cmd.parameters("@ird_origcond") = condition
    cmd.parameters("@user") = CurrentUser
    
    
    'Execute the command.
    cmd.Execute
    PutReceiptDetail = True
    Exit Function

errPutDataInsert:
    MsgBox Err.description: Err.Clear
End Function



Private Sub searchButton_Click()
    Call searchStockNumber(0)
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


Public Sub searchFIELD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call searchStockNumber(Index)
    End If
End Sub

Public Sub searchStockNumber(Index As Integer)
Dim datax As New ADODB.Recordset
Dim sql, list, i, ii, t
Screen.MousePointer = 11
      
            If Index = 0 Then
                If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                    sql = "SELECT stk_stcknumb, stk_desc, uni_desc " _
                        & "FROM STOCKMASTER LEFT OUTER JOIN UNIT ON " _
                        & "stk_npecode = uni_npecode AND " _
                        & "stk_primuon = uni_code WHERE " _
                        & "(stk_npecode = '" + nameSP + "') AND " _
                        & "(stk_stcknumb like '" + searchFIELD(Index).text + "%')"
                    datax.Open sql, cn, adOpenStatic
                    With STOCKlist
                        .Rows = 2
                        .TextMatrix(1, 0) = ""
                        .TextMatrix(1, 1) = ""
                        .TextMatrix(1, 2) = ""
                        .TextMatrix(1, 3) = ""
                        If datax.RecordCount > 0 Then
                            STOCKlist.Rows = datax.RecordCount + 1
                            Dim r As Integer
                            r = 1
                            Do While Not datax.EOF
                                If findSTUFF(datax!stk_stcknumb, STOCKlist, 1) = 0 Then
                                    .TextMatrix(r, 1) = datax!stk_stcknumb
                                    .TextMatrix(r, 2) = datax!stk_desc
                                    .TextMatrix(r, 3) = datax!uni_desc
                                    r = r + 1
                                    'STOCKlist.addITEM "" + vbTab + datax!stk_stcknumb + vbTab + datax!stk_desc + vbTab + datax!uni_desc & "", 1
                                End If
                                datax.MoveNext
                                Loop
                                STOCKlist.RowHeight(1) = 240
                            If STOCKlist.Rows > 2 And STOCKlist.TextMatrix(1, 1) = "" Then STOCKlist.RemoveItem 1
                            Call reNUMBER(STOCKlist)
                        End If
                    End With
                End If

        
    STOCKlist.TopRow = 1
    Screen.MousePointer = 0
End Sub

Private Sub searchFIELD_LostFocus(Index As Integer)
    searchFIELD(Index).backcolor = &HC0E0FF
End Sub



Private Sub STOCKlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim row, i
    If checkAll.Value = False Then
        With STOCKlist
            If y > 240 Then
                If Button = 2 Then
                    row = Round((y - 60) / .RowHeight(1))
                    .row = row
                    If .TopRow > 1 Then
                        .row = .row + .TopRow - 1
                    End If
                    For i = 1 To .cols - 1
                        .col = i
                        If .CellBackColor = vbWhite Then
                            .CellBackColor = &H80C0FF    'light orange
                        Else
                            .CellBackColor = vbWhite
                        End If
                    Next
                Else
                    If .CellBackColor <> vbWhite Then
                        col = .MouseCol
                        If col = 6 Or col = 8 Then
                            quantityBOX = .TextMatrix(.row, col)
                            quantityBOX.Left = .ColPos(col) + .Left + 40
                            quantityBOX.width = .ColWidth(col) - 40
                            quantityBOX.Top = .CellTop + .Top - 20
                            quantityBOX.Height = .RowHeight(.row) - 20
                            quantityBOX.Visible = True
                            quantityBOX.ZOrder
                            quantityBOX.SetFocus
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub


