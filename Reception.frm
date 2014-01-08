VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmReception 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Freight Forwarder Receipt"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   10590
   Tag             =   "02030100"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   3960
      TabIndex        =   27
      Top             =   6240
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "Reception.frx":0000
      CancelVisible   =   0   'False
      PreviousVisible =   0   'False
      NewVisible      =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      EMailVisible    =   -1  'True
      PrintEnabled    =   0   'False
      EmailEnabled    =   -1  'True
      SaveEnabled     =   0   'False
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   10663
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Reception"
      TabPicture(0)   =   "Reception.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "POlist"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboPoNumb"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboRecepTion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ReceptionDate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "EditButton"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextLINE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "searchButton"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "searchForm"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "Reception.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "RecipientList"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "cmd_Add"
      Tab(1).Control(3)=   "cmd_Remove"
      Tab(1).Control(4)=   "lbl_Recipients"
      Tab(1).ControlCount=   5
      Begin VB.PictureBox searchForm 
         Height          =   5415
         Left            =   120
         ScaleHeight     =   5355
         ScaleWidth      =   10035
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   10095
         Begin VB.CommandButton print 
            Caption         =   "Print Freight Receipt"
            Height          =   375
            Left            =   360
            TabIndex        =   49
            Top             =   3960
            Width           =   2415
         End
         Begin VB.TextBox cell 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   360
            MaxLength       =   15
            TabIndex        =   46
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CommandButton closeForm 
            Caption         =   "Close Searching Option"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   4440
            Width           =   2415
         End
         Begin VB.CommandButton goSearch 
            Caption         =   "Go Search"
            Height          =   375
            Left            =   360
            TabIndex        =   41
            Top             =   3480
            Width           =   2415
         End
         Begin VB.CheckBox allPOs 
            Caption         =   "Show all PO's no  matter status"
            Height          =   255
            Left            =   360
            TabIndex        =   40
            Top             =   1680
            Width           =   2775
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid resultsList 
            Height          =   4920
            Left            =   3360
            TabIndex        =   43
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   8678
            _Version        =   393216
            Rows            =   4
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   285
            BackColorSel    =   16761024
            ForeColorSel    =   0
            GridColorFixed  =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.TextBox cell 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   360
            MaxLength       =   15
            TabIndex        =   44
            Top             =   1320
            Width           =   2415
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
            Height          =   1455
            Index           =   0
            Left            =   360
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            Index           =   1
            Left            =   360
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
         Begin VB.Label receiptLabel 
            Caption         =   "Label5"
            Height          =   375
            Left            =   480
            TabIndex        =   48
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "By Freight Receipt:"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label3 
            Caption         =   "By PO:"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Searching Option Form"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   360
            TabIndex        =   37
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.CommandButton searchButton 
         Caption         =   "Search..."
         Height          =   375
         Left            =   8040
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TextLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6240
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton EditButton 
         Caption         =   "Edit"
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   580
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid RecipientList 
         Height          =   1815
         Left            =   -73320
         TabIndex        =   30
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3201
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker ReceptionDate 
         Bindings        =   "Reception.frx":0054
         DataField       =   "rec_date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataMember      =   "RECEPTION"
         DataSource      =   "deIms"
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   900
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   36848
      End
      Begin VB.ComboBox cboRecepTion 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   540
         Width           =   1800
      End
      Begin VB.ComboBox cboPoNumb 
         Height          =   315
         Left            =   5400
         TabIndex        =   19
         Top             =   540
         Width           =   1740
      End
      Begin VB.Frame Frame1 
         Caption         =   "Supplier Information"
         Enabled         =   0   'False
         Height          =   1815
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   9210
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   17
            Top             =   1020
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_name"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   14
            Left            =   1800
            TabIndex        =   16
            Top             =   300
            Width           =   7260
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   14
            Top             =   1380
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address1"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr1"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   15
            Left            =   1800
            TabIndex        =   12
            Top             =   660
            Width           =   2640
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_city"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   16
            Left            =   1800
            TabIndex        =   11
            Top             =   1020
            Width           =   2640
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_ctry"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   17
            Left            =   1800
            TabIndex        =   10
            Top             =   1380
            Width           =   2640
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_adr2"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   18
            Left            =   6240
            TabIndex        =   9
            Top             =   660
            Width           =   2820
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_stat"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   19
            Left            =   6240
            TabIndex        =   7
            Top             =   1020
            Width           =   2820
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "sup_zipc"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   20
            Left            =   6240
            TabIndex        =   5
            Top             =   1380
            Width           =   2820
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   315
            Index           =   12
            Left            =   4560
            TabIndex        =   15
            Top             =   1020
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address2"
            Height          =   315
            Index           =   11
            Left            =   4560
            TabIndex        =   8
            Top             =   660
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
            Height          =   315
            Index           =   13
            Left            =   4560
            TabIndex        =   6
            Top             =   1380
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   -74880
         ScaleHeight     =   3255
         ScaleWidth      =   9255
         TabIndex        =   3
         Top             =   2520
         Width           =   9255
         Begin ImsMailVB.Imsmail Imsmail 
            Height          =   3375
            Left            =   -120
            TabIndex        =   28
            Top             =   0
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   5953
         End
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74760
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74760
         TabIndex        =   1
         Top             =   2115
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid POlist 
         Height          =   2520
         Left            =   120
         TabIndex        =   33
         Top             =   3360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4445
         _Version        =   393216
         Cols            =   8
         RowHeightMin    =   285
         BackColorSel    =   16761024
         ForeColorSel    =   0
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To look for older records"
         Height          =   315
         Index           =   3
         Left            =   8040
         TabIndex        =   35
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "po_buyr"
         DataMember      =   "GETPONUMBERSFORRECEPTION"
         DataSource      =   "deIms"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1260
         TabIndex        =   26
         Top             =   900
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reception Date"
         Height          =   315
         Index           =   2
         Left            =   3540
         TabIndex        =   24
         Top             =   900
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reception #"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order #"
         Height          =   315
         Index           =   0
         Left            =   3540
         TabIndex        =   22
         Top             =   540
         Width           =   1800
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74715
         TabIndex        =   21
         Top             =   540
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim sRecNum As String
Dim vPKValues As Variant
Dim FNamespace As String
Dim Rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Dim Recpdelt As imsReceptionDetail
Dim Reception As imsReception
Dim RsPOITEMS As ADODB.Recordset
Dim notREADY As Boolean
Dim SaveEnabled As Boolean
'Dim RecepCol As RecepCol
'Dim WithEvents Recp As RecepCol
Dim beginning As Boolean
Dim currentPO, currentRECEPTION, dbtablename As String
Dim rowguid, locked As Boolean, idleStateEngagedFlag As Boolean, bl1   'jawdat

Dim selectionSTART As Integer
Dim lastCELL, focusHERE, nextCELL As Integer

Dim directCLICK As Boolean
Dim activeCELL As Integer
Dim usingARROWS As Boolean
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    If cellACTIVE <> "" Then
        With gridACTIVE
            If Not .Visible Then .Visible = True
            If .Rows < val(.Tag) Then .Tag = 1
            If IsNumeric(.Tag) Then
                .Col = column
                Call gridCOLORnormal(gridACTIVE, val(.Tag))
            End If
            If .Cols <= column Then Exit Sub
            .Col = column
            .Tag = ""
            found = False
            
            For i = 1 To .Rows - 1
                word = Trim(UCase(.TextMatrix(i, column)))
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    Call gridCOLORdark(gridACTIVE, i)
                    .Tag = .row
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                .row = 0
                .Tag = ""
            End If
            If IsNumeric(.Tag) Then
                If .Tag = "0" Then
                    .topROW = 1
                Else
                    .topROW = val(.Tag)
                End If
            End If
        End With
    End If
End Sub

Sub cleanGrid()
    resultsList.Rows = 2
    resultsList.Rows = 3
    resultsList.RemoveItem (1)
    resultsList.FixedRows = 1
End Sub

Sub fillUpReceipt(receipt As String, poNumber As String)
On Error GoTo errorOptions
Dim rst As ADODB.Recordset
    If receipt = "" Then Exit Sub
    currentRECEPTION = receipt
    Screen.MousePointer = 11
    Set rst = deIms.rsGet_Reception_Info_From_PONumb
    Call deIms.Get_Reception_Info_From_PONumb(poNumber, FNamespace)
    rst.Filter = 0
    rst.Filter = "recd_recpnumb = '" & receipt & "'"
    If rst.RecordCount = 0 Then Exit Sub
    If ((Not IsNull(rst!rec_date)) Or IsEmpty(rst!rec_date)) Then
        ReceptionDate = Format(rst!rec_date, "mm/dd/yyyy")
    End If
    rst.MoveFirst
    Dim row
    row = 1
    POlist.Rows = 2
    Do While True
        With POlist
            .row = row
            .TextMatrix(row, 0) = rst!poi_liitnumb
            .TextMatrix(row, 1) = rst!poi_comm
            .TextMatrix(row, 2) = rst!poi_desc
            .TextMatrix(row, 3) = Format(rst!poi_unitprice, "###,##0.00")
            .TextMatrix(row, 4) = rst!poi_primreqdqty
            .TextMatrix(row, 5) = rst!poi_qtydlvd
            .TextMatrix(row, 6) = rst!poi_qtytobedlvd
            .TextMatrix(row, 7) = Format(rst!poi_totaprice, "###,##0.00")
        End With
        rst.MoveNext
        row = row + 1
        If rst.EOF Then Exit Do
        POlist.AddItem ""
    Loop
    rst.Close
    cboPoNumb = poNumber
    receiptLabel = receipt
    Call MakeGridReadonly(Not SaveEnabled)
    POlist.Enabled = False
    If Len(Trim$(receipt)) > 0 Then
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
        NavBar1.SaveEnabled = False
    End If
    currentRECEPTION = receipt
    If Err Then Err.Clear
    Screen.MousePointer = 0
    Me.Refresh
    POlist.TextMatrix(0, 5) = "Qty. Received in Transaction"
    EditButton.Visible = False
    POlist.TextMatrix(0, 5) = "Qty. Received to Date"
Screen.MousePointer = 0
Exit Sub

errorOptions:
MsgBox Err.Description
Screen.MousePointer = 0
Resume Next
End Sub

Sub gridCOLORdark(Grid As MSHFlexGrid, row)
    With Grid
        .row = row
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
    End With
End Sub

Public Function CommitTransaction(cn As ADODB.Connection)
On Error Resume Next
    With MakeCommand(cn, adCmdText)
        .CommandText = "COMMIT TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    If Err Then Err.Clear
End Function
Sub gridCOLORnormal(Grid As MSHFlexGrid, row)
    With Grid
        .row = row
        .CellBackColor = &HFFFFC0      'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub
Sub doGrid(datax As ADODB.Recordset, list() As String)
Dim rec, i, ii
Dim t As String
    Err.Clear
    With resultsList
        resultsList.Rows = datax.RecordCount + 1
        i = 1
        Do While Not datax.EOF
            rec = ""
            For ii = 0 To 5
                If IsNull(datax(list(ii))) Then
                Else
                    If VarType(datax(list(ii))) = vbDate Then
                        .TextMatrix(i, ii) = Format(datax(list(ii)), "M/d/yyyy")
                    Else
                        .TextMatrix(i, ii) = datax(list(ii))
                    End If
                End If
            Next
            i = i + 1
            datax.MoveNext
        Loop
        If .TextMatrix(1, 0) = "" Then .RemoveItem (1)
        .row = 1
    End With
End Sub

Sub doCOMBO(Index, datax As ADODB.Recordset, list, totalwidth)
Dim rec, i, extraW
Dim t As String
    Err.Clear
    With combo(Index)
        combo(Index).Rows = datax.RecordCount + 1
        i = 1
        Do While Not datax.EOF
            rec = ""
            Select Case Index
                Case 0
                    t = Format(datax!po_ponumb)
                Case 1
                    t = Format(datax!rec_recpnumb)
            End Select
            .TextMatrix(i, 0) = t
            i = i + 1
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
    End With
End Sub
Sub fillCOMBO2(ByRef Grid As MSHFlexGrid, Index As Integer)
On Error GoTo ErrHandler
Dim Sql
Dim i, n, Params, shot, x, spot, rec, clue
ReDim list(2) As String

Dim datax As New ADODB.Recordset
    With combo(Index)
        .Rows = 2
        For i = 0 To 0
            .TextMatrix(0, i) = list(i)
            Select Case Index
                Case 0
                    .TextMatrix(0, i) = "PO #"
                Case 1
                    .TextMatrix(0, i) = "Receipt #"
            End Select
            .TextMatrix(1, i) = ""
            .ColWidth(i) = 2000
            .ColAlignment(0) = 0
        Next
    End With
    
    Select Case Index
        Case 0
            If allPOs Then
                Sql = "SELECT po_ponumb, PO_Date, po_stas, po_freigforwr " _
                    & "FROM PO WHERE " _
                    & "po_npecode = '" + deIms.NameSpace + "' AND " _
                    & "po_docutype not in ('R' , 'Q', 'L') "
            Else
                Sql = "SELECT po_ponumb, PO_Date, po_stas, po_freigforwr " _
                    & "FROM PO WHERE " _
                    & "po_npecode = '" + deIms.NameSpace + "' AND " _
                    & "(po_stas='OP') and (po_stasdelv IN ('NR','RP') and " _
                    & "po_docutype not in ('R' , 'Q', 'L')) "
            End If
            Sql = Sql + "ORDER BY po_ponumb"
            cell(1) = ""
        Case 1
            Sql = "SELECT rec_recpnumb " _
                & "FROM RECEPTION WHERE " _
                & "rec_npecode = '" + deIms.NameSpace + "' "
            If cell(0) = "" Then
                Sql = Sql + "ORDER BY rec_recpnumb"
            Else
                Sql = Sql + " AND " _
                    & "rec_ponumb = '" + cell(0) + "' " _
                    & "ORDER BY rec_date DESC"
            End If
    End Select
    datax.Open Sql, deIms.cnIms, adOpenForwardOnly
    If datax.RecordCount < 1 Then Exit Sub
    Call doCOMBO(Index, datax, list, 2000)
    Set datax = New ADODB.Recordset
    Exit Sub
ErrHandler:
    Select Case Err.number
        
    End Select
End Sub

Sub makeResultsList()
Dim i As Integer
    With resultsList
        For i = 0 To 5
            If i = 1 Then
                .ColAlignment(i) = 7
                .ColWidth(i) = 400
            Else
                .ColAlignment(i) = 1
                .ColWidth(i) = 1140
            End If
        Next
        .TextMatrix(0, 0) = "PO #"
        .TextMatrix(0, 1) = "Line Item"
        .TextMatrix(0, 2) = "PO Date"
        .TextMatrix(0, 3) = "PO Status"
        .TextMatrix(0, 4) = "Receipt #"
        .TextMatrix(0, 5) = "Rec Date"
        allPOs.value = 1
    End With
End Sub


Private Sub Picture2_Click()

End Sub

Private Sub Picture3_Click()

End Sub


Private Sub allPOs_Click()
    Call fillCOMBO2(combo(0), 0)
End Sub

Private Sub cell_Change(Index As Integer)
Dim n As Integer
    If Not directCLICK Then
        Call alphaSEARCH(cell(Index), combo(Index), 0)
    Else
        directCLICK = False
    End If
End Sub

Private Sub cell_Click(Index As Integer)
    Call showCOMBO(combo(Index), Index)
End Sub

Sub showCOMBO(ByRef Grid As MSHFlexGrid, Index As Integer)
    With Grid
        Call fillCOMBO2(Grid, Index)
        If .Rows > 0 And .Text <> "" Then
            .Visible = True
            .ZOrder
            .Top = cell(Index).Top + 270
        End If
        .MousePointer = 0
    End With
End Sub

Private Sub cell_GotFocus(Index As Integer)
        With cell(Index)
            .BackColor = &H80FFFF
            .Appearance = 1
            .Refresh
            activeCELL = Index
            .SelLength = Len(.Text)
            .SelStart = 0
        End With
End Sub

Private Sub cell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    'justCLICK = False
    With cell(Index)
        If Not .locked Then
                Select Case KeyCode
                    Case 40
                        Call arrowKEYS("down", Index)
                    Case 38
                        Call arrowKEYS("up", Index)
                    Case Else
                    Dim Col
                End Select
        End If
    End With
End Sub


Sub arrowKEYS(direction As String, Index As Integer)
Dim Grid As MSHFlexGrid
    With cell(Index)
        Set Grid = combo(Index)
            Grid.Visible = True
            Call gridCOLORnormal(Grid, val(Grid.Tag))
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
            Grid.Tag = Grid.row
            If Not Grid.Visible Then
                Grid.Visible = True
            End If
            Grid.ZOrder
            Grid.topROW = IIf(Grid.row = 0, 1, Grid.row)
            usingARROWS = True
            Call gridCOLORdark(Grid, Grid.row)
            Grid.SetFocus
    End With
End Sub


Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i, t, n
Dim gotIT As Boolean
    With cell(Index)
        Select Case KeyAscii
            Case 13
                KeyAscii = 0
                t = UCase(combo(Index).TextMatrix(combo(Index).row, 0))
                If UCase(cell(Index)) = Left(t, Len(cell(Index))) Then
                    gotIT = True
                    i = combo(Index).row
                Else
                    For i = 1 To combo(Index).Rows - 1
                        If UCase(cell(Index)) = UCase(combo(Index).TextMatrix(i, 0)) Then
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
            Case 27
                combo(Index).Visible = False
                cell(Index) = cell(Index).Tag
        End Select
    End With
End Sub


Private Sub cell_LostFocus(Index As Integer)
Dim continue As Boolean
    If usingARROWS Then
        usingARROWS = False
    Else
        combo(activeCELL).Visible = False
    End If
    With cell(Index)
        .BackColor = vbWhite
    End With
    Screen.MousePointer = 0
End Sub


Private Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If findSTUFF(cell(Index), combo(Index), 0) = 0 Then cell(Index) = ""
End Sub


Function findSTUFF(toFIND, Grid As MSHFlexGrid, Col) As Integer
Dim i
Dim findIT As Boolean
    findSTUFF = 0
    With Grid
        If .Rows < 3 Then
            If .TextMatrix(1, 0) = "" Then
                findIT = False
            Else
                findIT = True
            End If
        Else
            findIT = True
        End If
        If findIT Then
            For i = 1 To .Rows - 1
                If UCase(Trim(.TextMatrix(i, Col))) = UCase(Trim(toFIND)) Then
                    findSTUFF = i
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Sub closeForm_Click()
    searchForm.Visible = False
End Sub



Private Sub combo_Click(Index As Integer)
    combo(Index).Visible = False
    directCLICK = True
    With combo(Index)
        If .row > 0 Then
            cell(Index) = .TextMatrix(.row, 0)
            .Refresh
            cell(Index).Tag = .TextMatrix(.row, 0)
        End If
    End With
End Sub

Private Sub combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    With cell(Index)
        If Not .locked Then
            Select Case KeyCode
                Case 40
                    Call arrowKEYS("down", Index)
                Case 38
                    Call arrowKEYS("up", Index)
                Case Else
                Dim Col
            End Select
        End If
    End With
End Sub


Private Sub combo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call combo_Click(Index)
        Case 27
    End Select
    combo(Index).Visible = False
    If Index > 0 Then
        If Index < 4 Then
            cell(Index + 1).SetFocus
            Call cell_Click(Index + 1)
        Else
            cell(Index).SetFocus
        End If
    End If
End Sub


Private Sub combo_LostFocus(Index As Integer)
    combo(Index).Visible = False
End Sub


Private Sub goSearch_Click()
On Error GoTo ErrHandler
Dim Sql
Dim i
Dim datax As New ADODB.Recordset
Dim list(6) As String
list(0) = "poi_ponumb"
list(1) = "poi_liitnumb"
list(2) = "poi_creadate"
list(3) = "status"
list(4) = "rec_recpnumb"
list(5) = "rec_date"
Sql = "SELECT poi_ponumb, poi_liitnumb, poi_creadate,rec_recpnumb,rec_date, " _
        & "poi_stasliit+'-'+poi_stasdlvy +'-'+poi_stasship +'-'+poi_stasinvt as status " _
        & "FROM poitem LEFT JOIN reception ON poi_npecode = rec_npecode AND poi_ponumb = rec_ponumb  " _
        & "WHERE 1 = 1 "
If cell(0) <> "" Then
    Sql = Sql + " AND poi_ponumb = '" + cell(0) + "' "
End If
If cell(1) <> "" Then
    Sql = Sql + " AND rec_recpnumb = '" + cell(1) + "' "
End If
Sql = Sql + " ORDER BY poi_ponumb,rec_recpnumb,poi_liitnumb"
datax.Open Sql, deIms.cnIms, adOpenForwardOnly
Call cleanGrid
Call doGrid(datax, list)
Set datax = New ADODB.Recordset
Exit Sub
ErrHandler:
    Select Case Err.number
        
    End Select



End Sub

Private Sub print_Click()
On Error GoTo ErrHandler
    Dim selectedReceipt As String
    selectedReceipt = resultsList.TextMatrix(resultsList.row, 4)
    If selectedReceipt = "" And resultsList.row > 0 Then
        MsgBox "Please select a Receipt Number"
        Exit Sub
    Else
        BeforePrint (selectedReceipt)
    End If
    MDI_IMS.CrystalReport1.Action = 1
Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

Private Sub resultsList_DblClick()
    If resultsList.TextMatrix(resultsList.row, 0) <> "" And resultsList.TextMatrix(resultsList.row, 3) <> "" Then
        Call fillUpReceipt(resultsList.TextMatrix(resultsList.row, 3), resultsList.TextMatrix(resultsList.row, 0))
        searchForm.Visible = False
    End If
End Sub


Private Sub searchButton_Click()
    Call makeResultsList
    searchForm.Visible = True
    searchForm.ZOrder
End Sub


Public Sub TextLINE_Validate(Cancel As Boolean)
        With TextLINE
            If POlist.Col = 6 Then
                If IsNumeric(.Text) Then
                    If CDbl(.Text) = 0 Then
                        POlist.TextMatrix(val(.Tag), 6) = "0"
                    Else
                        POlist.TextMatrix(val(.Tag), 6) = Format(.Text)
                        .Tag = ""
                        .Text = ""
                        POlist.SetFocus
                        .Visible = False
                        Exit Sub
                    End If
                End If
            End If
        End With
End Sub

Sub showTEXTline(column As Integer)
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
Sub fillPO()
    On Error Resume Next
    'On Error GoTo problems

    Dim str As String
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    str = "po_ponumb = '" & cboPoNumb & "'"
    
    With deIms.rsGETPONUMBERSFORRECEPTION_SP
        .Filter = 0
        .Filter = str
        Label1(5) = !po_buyr
    End With
        
    NavBar1.SaveEnabled = SaveEnabled
    NavBar1.PrintEnabled = False
    NavBar1.EMailEnabled = False
    Set rst = deIms.rsGETSUPPLIERINFOFROMPONUMBER_SP
    Set cmd = deIms.Commands("GETSUPPLIERINFOFROMPONUMBER_SP")
    If (rst.State And adStateOpen) = adStateOpen Then rst.Close
    
    cmd.parameters("@PONUMB").value = cboPoNumb
    cmd.parameters("@NAMESPACE").value = FNamespace
    
    Set rst = cmd.Execute
    
    If cmd.parameters("RETURN_VALUE") <> 0 Then
    
        Label1(14) = rst("sup_name") & ""
        Label1(15) = rst("sup_adr1") & ""
        Label1(16) = rst("sup_adr2") & ""
        Label1(17) = rst("sup_ctry") & ""
        Label1(18) = rst("sup_city") & ""
        Label1(19) = rst("sup_stat") & ""
        Label1(20) = rst("sup_zipc") & ""
    End If
    
    'Label1(3) = ""
    ReceptionDate = ""
    cboRecepTion = ""
    'dgDetl.DataMember = "" 'JCG 2008/6/25
    'Set dgDetl.DataSource = Nothing 'JCG 2008/6/25
    
    Set rst = deIms.rsGETPOITEMFORRECEPTION_SP 'JCG 2008/6/27
    
    'dgDetl.DataMember = "GETPOITEMFORRECEPTION_SP" 'JCG 2008/6/25
    'Added my Muzammil
    'Set dgDetl.DataSource = rst
    
    If (rst.State And adStateOpen) = adStateOpen Then rst.Close
    Set cmd = deIms.Commands("GETPOITEMFORRECEPTION_SP")
        
    cmd.parameters("@PONUMB").value = cboPoNumb
    cmd.parameters("@NAMESPACE").value = FNamespace
    
    Set rst = cmd.Execute
    
    'dgDetl.ZOrder 'JCG 2008/6/25
    'dgDetl.DataMember = "" 'JCG 2008/6/25
    'Set dgDetl.DataSource = Nothing 'JCG 2008/6/25
    'dgDetl.DataMember = "GETPOITEMFORRECEPTION_SP" 'JCG 2008/6/25

 
    'Set dgDetl.DataSource = deIms 'JCG 2008/6/25
    
    'Added by muzammil
  '  rst.Close
  '  rst.Open , , adOpenKeyset, adLockBatchOptimistic
  '  Set dgDetl.DataSource = rst
    
    'Label1(3) = Format(Date, "mm/dd/yyyy")
    ReceptionDate = Format(Date, "mm/dd/yyyy")
    
    
    'JCG 2008/6/25
    rst.MoveFirst
    Dim row
    row = 1
    POlist.Rows = 2
    Do While True
        With POlist
            .row = row
            .TextMatrix(row, 0) = rst!poi_liitnumb
            .TextMatrix(row, 1) = rst!poi_comm
            .TextMatrix(row, 2) = rst!poi_desc
            .TextMatrix(row, 3) = Format(rst!poi_unitprice, "###,##0.00") '  2010-10-26 Juan wrong value
            .TextMatrix(row, 3) = rst!poi_primuom
            .TextMatrix(row, 4) = rst!poi_primreqdqty
            .TextMatrix(row, 5) = rst!poi_qtydlvd
            .TextMatrix(row, 6) = 0
            .TextMatrix(row, 7) = Format(rst!poi_totaprice, "###,##0.00")
        End With
        rst.MoveNext
        row = row + 1
        If rst.EOF Then Exit Do
        POlist.AddItem ""
    Loop
    '----------------
    
    
    GetReceptions
    If Err Then Err.Clear
    'Call MakeGridReadonly(Not SaveEnabled) 'JCG 2008/6/27
    'Using t
    Set deIms.rsGETPOITEMFORRECEPTION_SP.ActiveConnection = Nothing  'JCG 2008/6/27
    
    
'problems:
    'MsgBox "errors->" + Err.Description
    'Resume Next
End Sub

Sub makePOlist() 'JCG 2008/6/25
Dim i
    With POlist
        .Clear
        .Rows = 2
        .row = 0
        
        .RowHeight(0) = 700
        .RowHeightMin = 240
        For i = 0 To 6
            .ColWidth(i) = 1005
            .ColAlignment(i) = 6
            .ColAlignmentFixed(i) = 4
            .Col = i
            .CellAlignment = 1
        Next
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
        .ColWidth(0) = 405
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(7) = 1160
        .WordWrap = True
        .TextMatrix(0, 0) = "Item"
        .TextMatrix(0, 1) = "Stock #"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Qty. PO"
        .TextMatrix(0, 5) = "Qty. Received to Date"
        .TextMatrix(0, 6) = "Qty. Received Reception"
        .TextMatrix(0, 7) = "Price"
        .row = 1
        .Col = 1
    End With
End Sub

Private Sub cboPoNumb_Change()
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


'dgDetl.Columns(5).Caption = "Qty. Received to Date" 'JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" 'JCG 2008/6/28


End Sub

'get po number recordset, and assign values to lable

Private Sub cboPoNumb_Click()

If cboPoNumb <> "" Then
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
 
End If

'dgDetl.Columns(5).Caption = "Qty. Received to Date" ' 'JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" ' 'JCG 2008/6/28



Screen.MousePointer = 11

    currentPO = cboPoNumb
    cboRecepTion.ListIndex = -1
    'dgDetl.Columns(6).Visible = True 'JCG 2008/6/28
    'dgDetl.Refresh 'JCG 2008/6/28
        Call fillPO
    'Set dgDetl.DataSource = Nothing
Screen.MousePointer = 0

If ((bl1 = True) And (cboRecepTion <> "") And (cboPoNumb <> "")) Then
EditButton.Visible = True
Else
EditButton.Visible = False
End If

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


NavBar1.SaveEnabled = False
 'dgDetl.Columns(6).locked = True 'JCG 2008/6/28
 POlist.Enabled = False 'JCG 2008/6/28

'NavBar1.SaveVisible = False
 Call MakeGridReadonly(True)
   cboPoNumb = ""
 '  cboRecepTion = ""
  'dgDetl Visible = False
  
    GetReceptions
  
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
locked = True
End If                                                       'without this End if the form will get compilation errors

'jawdat, end copy



End Sub

'unlock po number combe

Private Sub cboPoNumb_DropDown()
cboPoNumb.locked = False

'dgDetl.Columns(5).Caption = "Qty. Received to Date" 'JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" 'JCG 2008/6/28


End Sub

Private Sub cboPoNumb_GotFocus()
    cboPoNumb.BackColor = &HC0FFFF
End Sub

'do not allow enter data to ponumber combo

Private Sub cboPoNumb_KeyPress(KeyAscii As Integer)
'If NavBar1.SaveEnabled = False Then KeyAscii = 0

    'Added by Juan 11/17/2000
    Dim i, Text
    If KeyAscii = 13 Then
        Call cboPoNumb_Validate(False)
'        If cboPoNumb <> "" And cboPoNumb <> "Error" Then SendKeys ("{tab}")
        Exit Sub
    End If
    With cboRecepTion
        Text = .Text
        For i = 0 To .ListCount - 1
            If Text Like .list(i) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    '-------------------------
'If ((bl1 = True) And (cboRecepTion <> "") And (cboPoNumb <> "")) Then
'EditButton.Visible = True
'Else
EditButton.Visible = False

'dgDetl.Columns(5).Caption = "Qty. Received to Date" 'JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" 'JCG 2008/6/28


'End If
End Sub

Private Sub cboPoNumb_LostFocus()
    cboPoNumb.BackColor = &H80000005
End Sub

Private Sub cboPoNumb_Scroll()

Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

'dgDetl.Columns(5).Caption = "Qty. Received to Date" ''JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" 'JCG 2008/6/28
 
End Sub

Public Sub cboPoNumb_Validate(Cancel As Boolean)
    'Added by Juan 11/18/200
    On Error Resume Next
    Dim Text, i, Sql
    Dim searcher As ADODB.Recordset

    'If deIms.rsGETPOITEMFORRECEPTION_SP.State = 1 Then Exit Sub
    With cboPoNumb
        Text = .Text
        If Text <> "" Then
            For i = 0 To .ListCount - 1
                If Text Like .list(i) Then
                    Call fillPO
                    Exit Sub
                End If
            Next
            Sql = "SELECT po_ponumb, po_stas FROM PO WHERE po_npecode = '" + FNamespace + "' " _
                & "AND LTRIM(po_ponumb) = '" + Trim(.Text) + "'"
            Set searcher = New ADODB.Recordset
            searcher.Open Sql, deIms.cnIms, adOpenForwardOnly
            If searcher.RecordCount > 0 Then
                If Err.number > 0 Then
                    .Text = "Error"
                    .SetFocus
                    Exit Sub
                End If
                If searcher!po_stas = "OP" Then
                    Call fillPO
                    Exit Sub
                Else
                    msg1 = translator.Trans("M00697")
                    MsgBox IIf(msg1 = "", "This transaction is not open", msg1)
                End If
            Else
                msg1 = translator.Trans("M00698")
                MsgBox IIf(msg1 = "", "This transaction doesn't exist", msg1)
            End If
        End If
        .Text = ""
        .SetFocus
    End With
    '------------------------
End Sub

'call function get recordset for reception
'and populate data grid and format date data type

Private Sub cboRecepTion_Click()
receiptLabel = ""
'On Error Resume Next 'JCG 2008/6/30
On Error GoTo errorOptions 'JCG 2008/6/30

Dim rst As ADODB.Recordset
    If cboRecepTion = "" Then Exit Sub
    Screen.MousePointer = 11
    Set rst = deIms.rsGet_Reception_Info_From_PONumb
    Screen.MousePointer = 11
    If ((rst.State And adStateOpen) = adStateClosed) Then _
        Call deIms.Get_Reception_Info_From_PONumb(cboPoNumb, FNamespace)
    Screen.MousePointer = 11
    rst.Filter = 0
    rst.Filter = "recd_recpnumb = '" & cboRecepTion & "'"
    Screen.MousePointer = 11
    'dgDetl.DataMember = "" 'JCG 2008/6/28
    'Set dgDetl.DataSource = Nothing 'JCG 2008/6/28
    'dgDetl.DataMember = "Get_Reception_Info_From_PONumb" 'JCG 2008/6/28

    'If Rs.RecordCount = 0 Then Exit Sub 'JCG 2008/6/30
    If rst.RecordCount = 0 Then Exit Sub 'JCG 2008/6/30
    Screen.MousePointer = 11
    If ((Not IsNull(rst!rec_date)) Or IsEmpty(rst!rec_date)) Then
        'Label1(3) = Format(rst!rec_date, "mm/dd/yyyy")
        ReceptionDate = Format(rst!rec_date, "mm/dd/yyyy")
    End If
    Screen.MousePointer = 11
   'Set dgDetl.DataSource = deIms 'JCG 2008/6/28

 'JCG 2008/6/29
    rst.MoveFirst
    Dim row
    row = 1
    POlist.Rows = 2
    Do While True
        With POlist
            .row = row
            .TextMatrix(row, 0) = rst!poi_liitnumb
            .TextMatrix(row, 1) = rst!poi_comm
            .TextMatrix(row, 2) = rst!poi_desc
            .TextMatrix(row, 3) = Format(rst!poi_unitprice, "###,##0.00")
            .TextMatrix(row, 4) = rst!poi_primreqdqty
            .TextMatrix(row, 5) = rst!poi_qtydlvd
            .TextMatrix(row, 6) = rst!poi_qtytobedlvd
            .TextMatrix(row, 7) = Format(rst!poi_totaprice, "###,##0.00")
        End With
        rst.MoveNext
        row = row + 1
        If rst.EOF Then Exit Do
        POlist.AddItem ""
    Loop
    '----------------

   
   Call MakeGridReadonly(Not SaveEnabled)
 '  dgDetl.Columns(6).Visible = False
 
 'dgDetl.Columns(6).locked = True 'JCG 2008/6/28
POlist.Enabled = False  'JCG 2008/6/28

   'dgDetl.Refresh  'JCG 2008/6/28
  

    Screen.MousePointer = 11
    If Len(Trim$(cboRecepTion)) > 0 Then
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
        NavBar1.SaveEnabled = False
    End If
    currentRECEPTION = cboRecepTion
    If Err Then Err.Clear
    Screen.MousePointer = 0

bl1 = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
If ((bl1 = True) And (cboRecepTion <> "") And (cboPoNumb <> "")) Then
EditButton.Visible = True
Me.Refresh
 
  'dgDetl.Columns(5).Caption = "Qty. Received in Transaction" ' 'JCG 2008/6/28
    POlist.TextMatrix(0, 5) = "Qty. Received in Transaction" 'JCG 2008/6/28
Else
EditButton.Visible = False
'dgDetl.Columns(5).Caption = "Qty. Received to Date" 'JCG 2008/6/28
POlist.TextMatrix(0, 5) = "Qty. Received to Date" 'JCG 2008/6/28
End If

Exit Sub
errorOptions:
MsgBox Err.Description
Screen.MousePointer = 0
Resume Next
End Sub

'unlock reception data combo

Private Sub cboRecepTion_DropDown()
Screen.MousePointer = 11
cboRecepTion.locked = False
Screen.MousePointer = 0
End Sub

Private Sub cboRecepTion_GotFocus()
    Screen.MousePointer = 11
    cboRecepTion.BackColor = &HC0FFFF
    Screen.MousePointer = 0
End Sub

'do not allow enter data to reception data combo

Private Sub cboRecepTion_KeyPress(KeyAscii As Integer)
If NavBar1.SaveEnabled = False Then KeyAscii = 0
    
    'Added by Juan 11/17/2000
    Dim i, Text
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        Exit Sub
    End If
    With cboRecepTion
        Text = .Text
        For i = 0 To .ListCount - 1
            If Text Like .list(i) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    '-------------------------
End Sub

Private Sub cboRecepTion_LostFocus()
    cboRecepTion.BackColor = &H80000005
End Sub

'call function to add current reciptient to reciptient list

Private Sub cmd_Add_Click()
On Error Resume Next

    Imsmail.AddCurrentRecipient
    
    If Err Then Err.Clear
End Sub

'delete current recipient from reciptient list

Private Sub cmd_Remove_Click()
On Error Resume Next

'    rsReceptList.Delete
'    rsReceptList.Update
    
    If RecipientList.row > 0 Then
        If RecipientList.Rows > 2 Then
            RecipientList.RemoveItem (RecipientList.row)
        Else
            RecipientList.TextMatrix(1, 1) = ""
        End If
    End If
        
    If Err Then Err.Clear
End Sub

'Private Sub dgDetl_Error(ByVal DataError As Integer, response As Integer) 'JCG 2008/6/28
'    If DataError <> 0 Then
'        notREADY = True
'    End If
'End Sub

Private Sub EditButton_Click()



Load frm_EditReception
frm_EditReception.Show
End Sub

'load form and set navbar buttom

Private Sub Form_Load()
Dim datax As ADODB.Recordset
Dim Sql
On Error Resume Next
    notREADY = False
    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmReception")
    '------------------------------------------

    NavBar1.EditEnabled = False
    NavBar1.EditVisible = False
    NavBar1.SaveEnabled = SaveEnabled
    FNamespace = deIms.NameSpace: GetPoNumb
    
    Imsmail.NameSpace = FNamespace

    cboPoNumb_Click
    'IMSMail.Connected = True'M
    Imsmail.SetActiveConnection deIms.cnIms 'M
    Imsmail.Language = Language 'M
    Call DisableButtons(Me, NavBar1)
    Call makePOlist 'JCG 2008/6/25
    SaveEnabled = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    
  
    
    frmReception.Caption = frmReception.Caption + " - " + frmReception.Tag
'    dgDetl.Enabled = True ''JCG 2008/6/28
    POlist.Enabled = True 'JCG 2008/6/28
    
    With RecipientList
        .ColWidth(0) = 300
        .ColWidth(1) = 9095
        .Rows = 2
        .Clear
        .TextMatrix(0, 1) = "Recipient List"
    End With

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
    
    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        rsReceptList.Open
    End If
    
    With frmReception
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'set store procedure parameters values

Private Function PutDataInsert(row As Integer) As Boolean
Dim str As String
Dim msg, Style, Title
Dim CmdQuantityDelivered As New ADODB.Command

    Dim cmd As Command

    On Error GoTo errPutDataInsert

    PutDataInsert = False

    Set cmd = deIms.Commands("RECEPTIONDETLINSERT_SP")



    If Len(Trim$(sRecNum)) = 0 Then Err.Raise 2000, "Missing recep number"
    
  
    'Set the parameter values for the command to be executed.
    cmd.parameters("@recd_recpnumb") = sRecNum
    'cmd.Parameters("@recd_desc") = Rs!poi_desc 'JCG 2008/6/27
    cmd.parameters("@recd_desc") = POlist.TextMatrix(row, 2)
    cmd.parameters("@recd_npecode") = FNamespace
    'cmd.Parameters("@recd_liitnumb") = Rs!poi_liitnumb 'JCG 2008/6/27
    cmd.parameters("@recd_liitnumb") = POlist.TextMatrix(row, 0)
    
    'cmd.Parameters("@recd_primqtydlvd") = CDbl(Rs!poi_qtytobedlvd) 'JCG 2008/6/27
    cmd.parameters("@recd_primqtydlvd") = CDbl(POlist.TextMatrix(row, 6)) 'JCG 2008/6/28
    
    'cmd.Parameters("@recd_unitpric") = CDbl(Rs!poi_unitprice) 'JCG 2008/6/28
    cmd.parameters("@recd_unitpric") = CDbl(POlist.TextMatrix(row, 7))
    'cmd.Parameters("@recd_partnumb") = CStr(Rs!poi_comm) 'JCG 2008/6/27
    cmd.parameters("@recd_partnumb") = CStr(POlist.TextMatrix(row, 1))
    cmd.parameters("@user") = CurrentUser
    
    'Execute the command.
    'deIms.cnIms.Errors.Clear
    cmd.Execute
    
    'If CDbl(rs!poi_qtytobedlvd) > 0 Then 'JCG 2008/6/19
    If CDbl(POlist.TextMatrix(row, 6)) >= 0 Then
   
         CmdQuantityDelivered.CommandType = adCmdStoredProc
         CmdQuantityDelivered.CommandText = "UpdatePoitemQuantityDeliveredAndTbs"
         CmdQuantityDelivered.ActiveConnection = deIms.cnIms
         CmdQuantityDelivered.parameters.Append CmdQuantityDelivered.CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
         CmdQuantityDelivered.parameters.Append CmdQuantityDelivered.CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cboPoNumb)
         'CmdQuantityDelivered.Parameters.Append CmdQuantityDelivered.CreateParameter("@LINEITEM", adVarChar, adParamInput, 5, Rs!poi_liitnumb) 'JCG 2008/6/27
         CmdQuantityDelivered.parameters.Append CmdQuantityDelivered.CreateParameter("@LINEITEM", adVarChar, adParamInput, 5, val(POlist.TextMatrix(row, 0)))
         'CmdQuantityDelivered.Parameters.Append CmdQuantityDelivered.CreateParameter("@primqtydlvd", adDouble, adParamInput, 12, CDbl(Rs!poi_qtytobedlvd)) 'JCG 2008/6/27
         CmdQuantityDelivered.parameters.Append CmdQuantityDelivered.CreateParameter("@primqtydlvd", adDouble, adParamInput, 12, CDbl(POlist.TextMatrix(row, 6)))


''         CmdQuantityDelivered.Parameters("@NAMESPACE") = deIms.NameSpace
''         CmdQuantityDelivered.Parameters("@PONUMB") = Trim$(cboPoNumb)
''         CmdQuantityDelivered.Parameters("@LINEITEM") = deIms.rsGETPOITEMFORRECEPTION_SP!poi_liitnumb
          
         CmdQuantityDelivered.Execute
         
    End If
    
    PutDataInsert = True
    Exit Function

errPutDataInsert:
    MsgBox Err.Description
    Err.Clear
End Function

'validate data format and show messege

Private Function ValidateData(row As Integer) As Boolean 'JCG 2008/6/28
Dim str As Double
Dim msg, Style, Title

    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("L00123") 'J added
    msg2 = translator.Trans("M00359") 'J added
    'msg = IIf(msg1 = "", " Stock# ", msg1 + " ") & Trim$(Rs!poi_comm) 'J modified 'JCG 200/6/28
    msg = IIf(msg1 = "", " Stock# ", msg1 + " ") & Trim$(POlist.TextMatrix(row, 1))
    msg = msg & IIf(msg2 = "", " is being over received, Do you want to continue ?", " " + msg2) 'J modified
    '---------------------------------------------



Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Imswin"   ' Define title.

    'JCG 2008/6/28
    'Verify the field is not null.
    'If IsNull(Rs("poi_qtytobedlvd")) Then
    '    MsgBox "The field ' poi_qtytobedlvd ' cannot be null."
    '    Exit Function
    'End If
    '------------

    'Verify the integer field contains a valid value.
    'If Not IsNull(Rs("poi_qtytobedlvd")) Then 'JCG 2008/6/28
    Dim value As String
    value = POlist.TextMatrix(row, 6)
    If value <> "" Then
        'If Not IsNumeric(Rs("poi_qtytobedlvd")) And InStr(Rs("poi_qtytobedlvd"), ".") = 0 Then 'JCG 2008/6/28
        If Not IsNumeric(value) Then
            MsgBox "Qty being received does not contain a valid number."
        Exit Function
        End If
        'Added By muzammil 15/01/01
        'Reason - To make sure the user does not enter any value less than 0"
        
        'If Not CDbl(Rs("poi_qtytobedlvd")) >= 0 Then 'JCG 2008/6/28
        If Not CDbl(value) >= 0 Then
             MsgBox " Qty being received cannot be lower than 0"
             Exit Function
        End If
    End If
    
    'str = CDbl((Rs("poi_qtydlvd")) + CDbl(Rs("poi_qtytobedlvd"))) 'JCG 2008/6/28
    str = CDbl(POlist.TextMatrix(row, 5) + CDbl(value))

    'If Not IsNull(Rs("poi_qtytobedlvd")) Then 'JCG 2008/6/28
    If value <> "" Then
        'If (CDbl(Rs!poi_primreqdqty)) < str And CDbl(Rs("poi_qtytobedlvd")) <> 0 Then
        If (CDbl(POlist.TextMatrix(row, 4)) < str And CDbl(value)) <> 0 Then

            If MsgBox(msg, Style, Title) = vbNo Then
NavBar1.SaveEnabled = True             'jawdat added as bug fix 2.8.02
                'Exit Function: dgDetl.SetFocus 'JCG 2008/6/28
                POlist.row = row
                POlist.Col = 6
                Exit Function: POlist.SetFocus
                
            End If
         Else
         End If
    End If
        
    ValidateData = True

End Function

'function set recordset column values

Private Function GetPKValue(vBookMark As Variant, sColName As String) As Variant

    Dim i As Integer

    GetPKValue = Rs(sColName)

    For i = 1 To UBound(vPKValues, 2)
        If vPKValues(0, i) = vBookMark And LCase(vPKValues(1, i)) = LCase(sColName) Then
            GetPKValue = vPKValues(2, i)
            Exit Function
        End If
    Next i
End Function

'call function get recordset and set store procedure parameters
'validata data format, check reception number exist or not

Private Sub AddItems()
On Error GoTo ErrHandler
Screen.MousePointer = 11
Dim lng As Long
Dim cmd As ADODB.Command
'Dim rst As ADODB.Recordset 'JCG 2008/6/28
Dim Result As Boolean
Dim Check As Boolean

    Result = False
    
    
    If Len(Trim$(cboRecepTion)) > 0 Then Exit Sub
    'added by muzammil.17/01/01
      
    Screen.MousePointer = 11
    'Set rst = deIms.rsGETPOITEMFORRECEPTION_SP 'JCG 2008/6/27
    'If (rst.State And adStateOpen) = adStateOpen Then 'JCG 2008/6/27

        Screen.MousePointer = 11
        Set cmd = deIms.Commands("GetAutoNumber")
        
        cmd.parameters("@DOCUTYPE") = "REC"
        cmd.parameters("@RETVAL") = Null
        cmd.parameters("@NPECODE") = FNamespace
        cmd.parameters("@RETVAL").direction = adParamOutput
        Screen.MousePointer = 11
        Call cmd.Execute(lng)
        Screen.MousePointer = 11
        If deIms.cnIms.Errors.Count > 0 Then deIms.cnIms.Errors.Clear
        Screen.MousePointer = 11
        If IsNull(cmd.parameters("@RETVAL")) Then
            MsgBox "Auto numbering is not set up properly"
            Exit Sub
        End If
        sRecNum = cmd.parameters("@RETVAL")
        Screen.MousePointer = 11
        Set cmd = Nothing
        'Commented out by muzammil.17/01/01
        'deIms.cnIms.Errors.Clear
        
      'BEGIN A TRANSACTION
       'Added By muzammil 17/01/01
         deIms.cnIms.BeginTrans
        Screen.MousePointer = 11
        Set cmd = deIms.Commands("RECEPTIONINSERT_SP")
        Screen.MousePointer = 11
        cmd.parameters("@rec_recpnumb") = sRecNum
        cmd.parameters("@rec_ponumb") = cboPoNumb
        cmd.parameters("@rec_npecode") = FNamespace
        cmd.parameters("@user") = CurrentUser
        cmd.parameters("@rec_date") = CDate(ReceptionDate)
        Screen.MousePointer = 11
        Call cmd.Execute(lng)

'        If lng > 0 Then
'            cboRecepTion.AddItem sRecNum
'        End If
        Screen.MousePointer = 11
'        Call UpdateReceptiontable
        
        'rst.MoveFirst 'JCG 2008/6/2007
        'Do While Not (rst.EOF) 'JCG 2008/6/27
        Dim i As Integer
        For i = 1 To POlist.Rows - 1
           Screen.MousePointer = 11
            Check = ValidateData(i)
            'If (Not (Result = PutDataInsert)) And (Result = True) Then Exit Do
            If (Check = True) Then
                'rst.MoveNext 'JCG 2008/6/27
            Else
                GoTo ErrHandler
                'dgDetl.SetFocus 'JCG 2008/6/27
            End If
            'rst.MoveNext
        'Loop 'JCG 2008/6/27
        Next

        Screen.MousePointer = 11
        
        'rst.MoveFirst 'JCG 2008/6/28
        'Do While Not (rst.EOF) 'JCG 2008/6/28
        
        For i = 1 To POlist.Rows - 1 'JCG 2008/6/28
           Screen.MousePointer = 11
            Result = PutDataInsert(i)
            'If (Not (Result = PutDataInsert)) And (Result = True) Then Exit Do
            If (Result = True) Then
                'rst.MoveNext 'JCG 2008/6/28
            'Else: Exit Do 'JCG 2008/6/28
            Else 'JCG 2008/6/28
                Exit For 'JCG 2008/6/28
            End If
            'rst.MoveNext
        Next 'JCG 2008/6/28
        'Loop 'JCG 2008/6/28

       ' If Result = False Then Exit Sub: dgDetl.SetFocus
        If Result = False Then
            GoTo ErrHandler
            'dgDetl.SetFocus  'JCG 2008/6/28
            POlist.SetFocus 'JCG 2008/6/28
        End If
        
        Screen.MousePointer = 0
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00360") 'J added
        MsgBox IIf(msg1 = "", "Please note that your reception number is", msg1) + " " & sRecNum 'J modified
        '---------------------------------------------
        Screen.MousePointer = 11
        Me.Refresh
        Screen.MousePointer = 11
        'Commected out by Muzammil to make it a Boolean type Returning function
        'and to take care of errors
        'call UPDATEPOITENTOBE
        'Call Updaterepstatus
        If UPDATEPOITENTOBE = False Then GoTo ErrHandler
        If Updaterepstatus = False Then GoTo ErrHandler
        Screen.MousePointer = 11

        cboRecepTion.AddItem sRecNum
        cboRecepTion.ListIndex = cboRecepTion.ListCount - 1
        cboRecepTion.Tag = cboRecepTion
        
        BeforePrint
        Call SendWareHouseMessage(deIms.NameSpace, "Automatic Distribution", _
                                  "Freight Reciept", deIms.cnIms, CreateRpti)

        Screen.MousePointer = 11
    'End If
    deIms.cnIms.CommitTrans
    Screen.MousePointer = 0
    Exit Sub
ErrHandler:
    If Err Then MsgBox Err.Description: Err.Clear
    'Added by muzammil
    deIms.cnIms.RollbackTrans
    Screen.MousePointer = 0
    'Commencted Out by Muzammil
    'GoTo CleanUp
End Sub

Private Sub Form_Paint()
 '   cboPoNumb.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim closing

If MDI_IMS(idleStateEngagedFlag) = True Then
Exit Sub
Else
    If NavBar1.SaveEnabled Then
        closing = MsgBox("Do you really want to close and lose your last record?", vbYesNo)
        If closing = vbNo Then
            Cancel = True

        End If
    End If

End If
    
End Sub

'unload form and close recordsets

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
  
  
 
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
   
Unload frm_EditReception
    'Hide
    ' With deIms
        '.rsGETPOITEMFORRECEPTION_SP.Close
        '.rsGETPONUMBERSFORRECEPTION_SP.Close
'        .rsGETSUPPLIERINFOFROMPONUMBER_SP.Close
'        .rsGet_Reception_Info_From_PONumb.Close
    'End With

    'If Err Then Err.Clear

    'Set dgDetl.DataSource = Nothing
    'If Err Then Err.Clear

    'IMSMail.Connected = False 'M
    'Set rsReceptList = Nothing

    'If Err Then Err.Clear
    'If open_forms <= 5 Then ShowNavigator
    
    'Unload frmWarehouse
    
    
    
    
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnNewClick()
    notREADY = False
End Sub

'call function to print crystal report

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    BeforePrint
    
    MDI_IMS.CrystalReport1.Action = 1
    'MDI_IMS.CrystalReport1.Reset
            Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'before save call function to check data format

Private Sub NavBar1_OnSaveClick()
'If NavBar1.SaveEnabled = False Then
'Exit Sub

Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


    Screen.MousePointer = 11
    If Not notREADY Then
        'With deIms.Recordsets("GETPOITEMFORRECEPTION_SP") 'JCG 2008/6/27
            Screen.MousePointer = 11
            'If .RecordCount > 0 Then 'JCG 2008/6/27
            If POlist.Rows >= 2 And POlist.TextMatrix(1, 0) <> "" Then
                '.MoveFirst 'JCG 2008/6/27
                Screen.MousePointer = 11
                If Not notREADY Then
                    'JCG 2008/6/27
                    'Do While Not .EOF
                    '    Screen.MousePointer = 11
                    '    If Not IsNumeric(!poi_qtytobedlvd) Then
                    '        Screen.MousePointer = 0
                    '        MsgBox "Invalid Value"
                    '        dgDetl.Col = 6
                    '        dgDetl.SetFocus
                    '        Exit Sub
                    '    End If
                    '    .MoveNext
                    'Loop
                    Dim i As Integer
                    For i = 1 To POlist.Rows - 1
                        Screen.MousePointer = 11
                        If Not IsNumeric(POlist.TextMatrix(i, 6)) Then
                            Screen.MousePointer = 0
                            MsgBox "Invalid Value"
                            POlist.row = i
                            POlist.Col = 6
                            POlist.SetFocus
                            Exit Sub
                        End If
                    Next
                    notREADY = False
                End If
            Else
                Screen.MousePointer = 0
                MsgBox "Invalid Data"
                notREADY = True
                Exit Sub
            End If
        'End With 'JCG 2008/6/27
    End If
    
    If notREADY Then
    Else
        Screen.MousePointer = 11
        AddItems
        Screen.MousePointer = 11
        'cboPoNumb_Click
        
        'If Len(cboRecepTion.Tag) Then
         '   cboRecepTion.ListIndex = IndexOf(cboRecepTion, cboRecepTion.Tag)
        'End If
        If cboRecepTion.Enabled And cboRecepTion.Visible Then cboRecepTion.SetFocus
    End If
    notREADY = False
    Screen.MousePointer = 0
    'dgDetl.Refresh  'JCG 2008/6/28
'    End If
End Sub

'get po number recordset and populate data combo

Private Sub GetPoNumb()
On Error Resume Next

    deIms.rsGETPONUMBERSFORRECEPTION_SP.Close
    
    If Err Then Err.Clear: deIms.cnIms.Errors.Clear
    Call deIms.GETPONUMBERSFORRECEPTION_SP(FNamespace)
    
    Set Rs = deIms.rsGETPONUMBERSFORRECEPTION_SP
    
    Rs.Filter = 0
    Call PopuLateFromRecordSet(cboPoNumb, Rs, "po_ponumb", False)
    
    'DoEvents
    Set Rs = deIms.rsGETPOITEMFORRECEPTION_SP
    cboPoNumb.ListIndex = CB_ERR
    
    If Err Then Err.Clear
End Sub

'call function get reception number recordset

Private Sub GetReceptions()
'On Error Resume Next
Dim l As Long
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim i  As Integer
Dim goSAMEreception As Boolean

    If currentRECEPTION = cboRecepTion Then goSAMEreception = True
    
    'Set cmd = New ADODB.Command
    
    Set rst = deIms.rsGet_Reception_Info_From_PONumb
    Set cmd = deIms.Commands("Get_Reception_Info_From_PONumb")
    If rst.State And adStateOpen = adStateOpen Then rst.Close
    
    With cmd
        rst.Filter = 0
        Set rst = Nothing
        cboRecepTion.Clear
'        .CommandType = adCmdStoredProc
'        .ActiveConnection = deIms.cnIms
'        .CommandText = "Get_Reception_InfO_From_PONumb"
'
'        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
'        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cboPoNumb)
'        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, fNameSpace)
        
        .parameters("PONUMB").value = cboPoNumb
        .parameters("NAMESPACE").value = FNamespace
    
        Set rst = .Execute
    
        l = .parameters("RETURN_VALUE")
        
        If l Then Call PopuLateFromRecordSet(cboRecepTion, rst, "recd_recpnumb", False)
    End With
    If currentPO = cboPoNumb Then
        If goSAMEreception Then
            For i = 0 To cboRecepTion.ListCount - 1
                If cboRecepTion.list(i) = currentRECEPTION Then
                    cboRecepTion.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub

'lock data grid columns

Private Sub MakeGridReadonly(value As Boolean)
On Error Resume Next
    'dgDetl.Splits(0).locked = Value
    'dgDetl.Columns(6).locked = value  'JCG 2008/6/28
    POlist.Enabled = value  'JCG 2008/6/28
End Sub

'call function add current recipient to recipient list

Private Sub IMSMail_OnAddClick(ByVal address As String)

    
    If (InStr(1, address, "@") > 0) = 0 Then
        address = UCase(address)
    End If
    
    If Not IsInList(address, "Recipients", rsReceptList) Then _
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))

    'Set ssdbRecepientList.DataSource = rsReceptList
'    ssdbRecepientList.Columns(0).DataField = "Recipients"

    'Call getRECIPIENTSlist
    If InStr(UCase(address), "INTERNET") > 0 Then address = Mid(address, InStr(UCase(address), "INTERNET") + 8)
    If InStrRev(address, "!") > 0 Then address = Mid(address, InStrRev(address, "!") + 1)
    RecipientList.AddItem "" + vbTab + address
    If RecipientList.Rows > 2 And RecipientList.TextMatrix(1, 1) = "" Then RecipientList.RemoveItem (1)
    
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

Private Sub ssdbRecepientList_DblClick()
    RecipientList.RemoveItem RecipientList.row
End Sub

'call function send email fax

Private Sub NavBar1_OnEMailClick()
Dim Params(1) As String
Dim i As Integer
Dim Attachments() As String
Dim subject As String
Dim reports(0) As String
Dim Recepients() As String
Dim attention As String
Dim rptinfo As RPTIFileInfo
Dim FileName As String
Dim IFile As IMSFile
Screen.MousePointer = 11

'On Error GoTo errMESSAGE
    Set IFile = New IMSFile
     BeforePrint
     MDI_IMS.CrystalReport1.PrintFileType = crptRTF
    
    If RecipientList.TextMatrix(1, 1) <> "" Then
      
        subject = "Freight Forwarder Receipt #" + cboRecepTion
        reports(0) = "freception.rpt"
        
        attention = "Attention Please "
      
        With rptinfo
            Params(0) = "namespace;" + deIms.NameSpace + ";TRUE"
            Params(1) = "recnumb;" + cboRecepTion + ";TRUE"
            .ReportFileName = reportPath & "freception.rpt"
            Call translator.Translate_Reports("freception.rpt")
            .parameters = Params
        End With
        
        ' Call WriteRPTIFile(rptinfo, Left(MDI_IMS.CrystalReport1.ReportFileName, Len(MDI_IMS.CrystalReport1.ReportFileName) - 3) + "rtf") 'JCG 2008/8/1
        'Attachments = generateattachmentsPDF("freception.rpt", subject, Params, MDI_IMS.CrystalReport1, RTrim(cboRecepTion), "receipt") 'JCG 2008/8/1
        Attachments = generateattachmentswithCR11(Report_EmailFax_FreightReceipt_name, subject, Params, MDI_IMS.CrystalReport1)  'JCG 2008/8/1
        
        ReDim Recepients(RecipientList.Rows - 1)
        For i = 1 To RecipientList.Rows - 1
            Recepients(i) = RecipientList.TextMatrix(i, 1)
        Next
        'Recepients = ToArrayFromRecordset(rsReceptList)
        
        'Attachments(0) = "Freight Forwarder Receipt - " & cboRecepTion & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf" ''JCG 2008/8/1
        'FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
        
         'JCG 2008/8/1
        ' Filename = ConnInfo.EmailOutFolder & Attachments(0)
        'If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
        'If Not FileExists(Filename) Then MDI_IMS.SaveReport Filename, crptRTF
        '------------------
        
        Call WriteParameterFiles(Recepients, "", Attachments, subject, attention)
    Else
         MsgBox "No Recipients to Send", , "Imswin"
     
    End If
    Screen.MousePointer = 0


errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If





'Dim FileName As String
'
'
'    BeforePrint
'    Call WriteRPTIFile(CreateRpti, FileName)
'
'    'Modified by Juan (9/15/2000) for Multilingual
'    msg1 = translator.Trans("L00061") 'J added
'    Call SendEmailAndFax(rsReceptList, IIf(msg1 = "", "Recipients", msg1), "Freight Forwarder Receipt - " & cboRecepTion, "", FileName)   'J modified
'    '---------------------------------------------
'
'    On Error Resume Next
'    Call rsReceptList.Delete(adAffectAllChapters)
'    Set ssdbRecepientList.DataSource = rsReceptList
'
'    If Err Then Err.Clear
End Sub

'set crystal report parameters

Private Sub BeforePrint(Optional receiptNumber As String)
On Error GoTo ErrHandler

    MDI_IMS.CrystalReport1.Reset
    MDI_IMS.CrystalReport1.ReportFileName = reportPath & "freception.rpt"
    MDI_IMS.CrystalReport1.ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
    If IsNull(receiptNumber) Or receiptNumber = "" Then
        'Juan 2010-9-21
        If cboRecepTion = "" And receiptLabel <> "" Then
            MDI_IMS.CrystalReport1.ParameterFields(1) = "recnumb;" + receiptLabel + ";TRUE"
        Else
            MDI_IMS.CrystalReport1.ParameterFields(1) = "recnumb;" + cboRecepTion + ";TRUE" 'this line was the only one original before
        End If
        '------------------------------------
    Else
        MDI_IMS.CrystalReport1.ParameterFields(1) = "recnumb;" + receiptNumber + ";TRUE"
    End If
    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("L00465") 'J added
    MDI_IMS.CrystalReport1.WindowTitle = IIf(msg1 = "", "Freight Forwarder Receipt", msg1) 'J modified
    Call translator.Translate_Reports("freception.rpt") 'J added
    '---------------------------------------------
    
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'set store procedure parameters

Private Function InsertIntoTable(row As Integer) As Boolean
Set Recpdelt = New imsReceptionDetail

    Recpdelt.Npecode = FNamespace  '"' & deIms.Namespace & '"
    Recpdelt.Recpnumber = sRecNum
    Recpdelt.Recplineitem = Rs!poi_liitnumb
    'Recpdelt.Recppartnumb = CStr(Rs!poi_comm) 'JCG 2008/6/28
    Recpdelt.Recppartnumb = POlist.TextMatrix(row, 1)
    Recpdelt.Repdescription = Rs!poi_desc
    'Recpdelt.Recpriqtydelived = CDbl(Rs!poi_qtytobedlvd) 'JCG 2008/6/28
    Recpdelt.Recppartnumb = POlist.TextMatrix(row, 6)
    'Recpdelt.RecpUintprice = CDbl(Rs!poi_unitprice) 'JCG 2008/6/28
    Recpdelt.Recppartnumb = POlist.TextMatrix(row, 3)
    
'    Call Recpdelt.InsertReceptiondelt(deIms.cnIms)

End Function

'set store procedure parameters and call it to update po status

'Private Function UpdateReceptiontable() As String
'Dim cmd As ADODB.Command

'    Set cmd = New ADODB.Command
    
'    With cmd
'        .CommandType = adCmdStoredProc
'        .CommandText = "RECEPTION_UPDATE"
'        Set .ActiveConnection = deIms.cnIms
        
'        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
'        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cboPoNumb)
'        .Parameters.Append .CreateParameter("@LINENUMB", adVarChar, adParamInput, 6, dgDetl.Columns(0).Text)
'        .Execute
        
'    End With
        
'   Set cmd = Nothing

'End Function

'set store procedurer paratmeters and call it to update po line item
'status

Private Function UPDATEPOITENTOBE() As Boolean
UPDATEPOITENTOBE = False
On Error GoTo ErrHandler
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATEPOITENTOBE"
        Set .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cboPoNumb)
        .parameters.Append .CreateParameter("@RECPNUMB", adVarChar, adParamInput, 15, sRecNum)
'        .Parameters.Append .CreateParameter("@POLIITEM", adVarChar, adParamInput, 6, rs!poi_liitnumb)
        .Execute
        
    End With
        
   Set cmd = Nothing
   UPDATEPOITENTOBE = True
   Exit Function
ErrHandler:
  MsgBox Err.Description
  Err.Clear
End Function

'set store procedure parameters and call it to update
'reception statues

Private Function Updaterepstatus() As Boolean
On Error GoTo ErrHandler
Updaterepstatus = False
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATE_PO_RECEPSTATES"
        Set .ActiveConnection = deIms.cnIms
        
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, cboPoNumb)
        
'        .Parameters.Append .CreateParameter("@POLIITEM", adVarChar, adParamInput, 6, lblLineItem.Caption)
        .Execute
        
    End With
        
   Set cmd = Nothing
   Updaterepstatus = True
   Exit Function
ErrHandler:
   Err.Clear
End Function

Private Sub POlist_Click()
Dim i, currentCOL As Integer
        With POlist
            If .TextMatrix(.row, 1) <> "" Then
                POlist.Tag = .row
                currentCOL = .MouseCol
                Select Case currentCOL
                    Case 6
                        If POlist.Enabled Then
                            Call showTEXTline(currentCOL)
                        End If
                 End Select
            End If
        End With
End Sub

'set navbar print and email buttom

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        NavBar1.PrintEnabled = cboRecepTion.ListIndex <> CB_ERR
        NavBar1.EMailEnabled = NavBar1.PrintEnabled
    End If
End Sub

'get email parameters

Private Function CreateRpti() As RPTIFileInfo

    With CreateRpti
        ReDim .parameters(1)
        .ReportFileName = reportPath & "freception.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("freception.rpt") 'J added
        '---------------------------------------------
        
        .parameters(1) = "recnumb=" & cboRecepTion
        .parameters(0) = "namespace=" & deIms.NameSpace
        
    End With

End Function

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


