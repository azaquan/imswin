VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGlobalWH 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox savingLABEL 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3360
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
      Left            =   7575
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton saveBUTTON 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5895
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Print"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton newBUTTON 
      Caption         =   "&New Transaction"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CheckBox checkAll 
      Caption         =   "Transfer all items"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   9
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
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
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
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   3855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   2
      Left            =   120
      TabIndex        =   7
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Index           =   5
      Left            =   5400
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
      Left            =   5400
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
      Left            =   5400
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
      Height          =   3660
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6456
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      RowHeightMin    =   285
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid matrix 
      Height          =   735
      Left            =   9120
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
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
   Begin VB.Label label 
      Caption         =   "To Namespace"
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   20
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "To Warehouse"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   19
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label label 
      Caption         =   "To Company"
      Height          =   255
      Index           =   4
      Left            =   5400
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
Sub makeLists()
    With STOCKlist
        .cols = 6
        .TextMatrix(0, 0) = "#"
        .ColWidth(0) = 485
        .TextMatrix(0, 1) = "Commodity"
        .ColWidth(1) = 1400
        .TextMatrix(0, 2) = "Unit Price"
        .ColWidth(2) = 1000
        .TextMatrix(0, 3) = "Description"
        .ColWidth(3) = 3200
        .TextMatrix(0, 4) = "Unit"
        .ColWidth(4) = 1200
        .TextMatrix(0, 5) = "Qty"
        .ColWidth(5) = 1200
    End With
End Sub


Private Sub cell_Change(Index As Integer)
Dim n As Integer
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
Screen.MousePointer = 11
    With cell(Index)
        If saveBUTTON.Enabled Or Index = 0 Then
            If Index > 2 And Index <> 4 Then
                If combo(Index - 1) = "" And (Index - 1) > 1 Then
                    MsgBox "Please select " + Label(Index - 1) + " first"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            If Not (saveBUTTON.Enabled And Index = 0) Then
                Call showCOMBO(combo(Index), Index)
            End If
        End If
        Screen.MousePointer = 0

        .SelStart = 0
        .SelLength = Len(.text)
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
    With grid
        Call fillCOMBO(grid, Index)
        If .Rows > 0 And .text <> "" Then
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
        .cols = 2
        .TextMatrix(0, 0) = "Description"
        .TextMatrix(0, 1) = "Code"
        .ColWidth(0) = 2800
        .ColWidth(1) = 1400
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
        Case 4 'itwoulde zero as well
            sql = "select npce_name as namespaceName, npce_code as namespace from namespace " _
                + "order by namespaceName "
        Case 2, 5
            sql = "select com_name as companyName, com_compcode as company from company " _
                + "where com_npecode = '" + namespaceVal + "' " _
                + "order by companyName "
        Case 3, 6
            sql = "select loc_name as locationName, loc_locacode as location from location " _
                + "where loc_npecode = '" + namespaceVal + "' and loc_compcode = '" + companyVal + "' " _
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
            .AddItem datax.Fields(0) + vbTab + datax.Fields(1)
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

Private Sub combo_Click(Index As Integer)
Dim i, sql, t
Dim datax As New ADODB.Recordset
Dim currentformname, currentformname1
Dim MSGBOXReply As VbMsgBoxResult
Dim labelname As String
Dim ratio As Integer
    combo(Index).Visible = False
    DoEvents
    Screen.MousePointer = 11
    DoEvents
    directCLICK = True
    Set datax = New ADODB.Recordset
    DoEvents
    With combo(Index)
        cell(Index) = Trim(.TextMatrix(.row, 0))
        cell(Index).tag = Trim(.TextMatrix(.row, 1))
        Select Case Index
            Case 0
'                datax.Open sql, cn, adOpenForwardOnly
'                If datax.RecordCount > 0 Then
'                    Call fillTRANSACTION(datax)
'                End If
            Case 3
                sql = "select * from stockinfo where " _
                    + "NameSpace = '" + cell(1).tag + "' " _
                    + "and Company = '" + cell(2).tag + "' " _
                    + "and Location = '" + cell(3).tag + "' "
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
                    If savingLABEL.Visible Then
                        Label3 = "SAVING..."
                        savingLABEL.Visible = False
                    End If
                End If
        End Select
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
onDetailListInProcess = True
    With datax
        n = 0
        STOCKlist.Rows = 2
        STOCKlist.row = 1
        STOCKlist.col = 0
        STOCKlist.CellFontName = "MS Sans Serif"
        mainItemRow = 0
        Do While Not .EOF
            STOCKlist.ColAlignment(7) = 0
            rec = rec + Trim(!StockNumber) + vbTab
            rec = rec + Trim(!description) + vbTab
            rec = rec + Format(!qty, "0.00") + vbTab
            rec = rec + IIf(IsNull(!UnitName), "", !UnitName)
            STOCKlist.AddItem rec
            If n = 20 Then
                DoEvents
                STOCKlist.Refresh
            End If
            .MoveNext
            n = n + 1
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

Private Sub Form_Activate()
Dim rights As Boolean
    inProgress = False
    Screen.MousePointer = 0
    ''''' TODO SECURITY FUNCTIONALITY
    'rights = Getmenuuser(nameSP, CurrentUser, Me.tag, cn)
    'newBUTTON.Enabled = rights
    Me.Visible = True
    If newBUTTON.Enabled Then newBUTTON.SetFocus
    Me.Refresh
    'userNAMEbox = CurrentUser
    'dateBOX = Format(Now, "mm/dd/yyyy")
    Call makeLists
'    DoEvents
'    Call fillGRID(grid(1), logicBOX(0), 0)
'    DoEvents
'    Call fillGRID(grid(2), sublocaBOX(0), 0)
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
    cell(1).text = nameSPname + " (" + nameSP + ")"
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
    saveBUTTON.Enabled = True
    newBUTTON.Enabled = False
    cell(0).backcolor = &HFFFFC0
    cell(0) = ""
    ''''''
    '''Call hideREMARKS
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
