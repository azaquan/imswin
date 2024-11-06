VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_logical 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Tag             =   "01030500"
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid clonedGrid 
      Height          =   975
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      FirstEnabled    =   0   'False
      FirstVisible    =   0   'False
      LastEnabled     =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintVisible    =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid logwarGrid 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5953
      _Version        =   393216
      ForeColor       =   0
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   285
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      GridLinesFixed  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid combo 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
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
      Left            =   6720
      TabIndex        =   2
      Top             =   4080
      Width           =   2460
   End
   Begin VB.Label lbl_Logicals 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logical Warehouse"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frm_logical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim InUnload As Boolean
Dim Modify As String
Dim Create As String
Dim RecSaved As Boolean
Dim Visualize As String
Dim NVBAR_EDIT As Boolean
Dim NVBAR_ADD As Boolean
Dim NVBAR_SAVE As Boolean
Dim CAncelGrid As Boolean
Dim TableLocked As Boolean, currentformname As String   'jawdat
Dim newRecord As Boolean
Dim oldVALUE(2) As String
Dim originalVALUE(2) As String
Dim colSwitch As Boolean
Dim lastMark As Integer
Dim dontClose As Boolean
Dim oldRow, oldCol As Integer
Dim getOut As Boolean
Function changedRow(r As Integer) As Boolean
    Dim i As Integer
    Dim oldRow, newRow As String
    oldRow = ""
    newRow = ""
    changedRow = True
    If lblStatus.Caption = "Creation" Then Exit Function
    For i = 0 To logwarGrid.Cols - 1
        oldRow = oldRow + clonedGrid.TextMatrix(r, i)
        newRow = newRow + logwarGrid.TextMatrix(r, i)
    Next
    If oldRow = newRow And Len(oldRow) = Len(newRow) Then changedRow = False
End Function

Sub checkBox()
    With logwarGrid
        If .Text = "o" Then
                .TextMatrix(.row, 3) = "þ"
            Else
                .TextMatrix(.row, 3) = "o"
            End If
    End With
End Sub

    
Sub cleanGrid()
    Dim currentROW, currentCOL, i As Integer
    With logwarGrid
        If lastMark > 0 Then
            currentROW = .row
            .row = lastMark
            currentCOL = .Col
            .Col = 3
            .CellBackColor = vbWhite
            .row = currentROW
            .Col = currentCOL
        End If
        lastMark = .row
    End With
End Sub

Sub colorize(color)
    Dim r, c As Integer
    With logwarGrid
        For r = 1 To .Rows - 1
            For c = 0 To .Cols - 1
                .ForeColor = color
            Next
        Next
    End With
End Sub

Sub doClone()
    Dim r, c As Integer
    With logwarGrid
        clonedGrid.Rows = .Rows
        clonedGrid.Cols = .Cols
        For r = 1 To .Rows - 1
            For c = 0 To .Cols - 1
                clonedGrid.TextMatrix(r, c) = .TextMatrix(r, c)
            Next
        Next
    End With
End Sub

Sub fillCombo()
    If deIms.rslogwar_type.State > 0 Then
        deIms.rslogwar_type.Close
    End If
    Call deIms.logwar_type(deIms.NameSpace)
    With combo
        Dim i As Integer
        .Rows = 2
        For i = 0 To .Cols - 1
            .TextMatrix(1, i) = ""
        Next
        Do While Not deIms.rslogwar_type.EOF
            .AddItem deIms.rslogwar_type!type_code
            deIms.rslogwar_type.MoveNext
        Loop
        If .Rows > 2 Then .RemoveItem (1)
    End With
End Sub

Sub FillGrid()
Dim r As Integer
    With logwarGrid
        r = 1
        deIms.rsLOGWAR.Sort = "lw_code"
        Do While Not deIms.rsLOGWAR.EOF
            Dim i As Integer
            If r + 1 > .Rows Then .AddItem ""
            .row = r
            .TextMatrix(r, 0) = deIms.rsLOGWAR!lw_code
            .TextMatrix(r, 1) = deIms.rsLOGWAR!lw_desc
            .TextMatrix(r, 2) = IIf(IsNull(deIms.rsLOGWAR!lw_type), "", deIms.rsLOGWAR!lw_type)
            .Col = 3
            .CellFontName = "Wingdings"
            .CellFontSize = 16
            .ColAlignment(3) = 4
            If deIms.rsLOGWAR!lw_actvflag Then
                .TextMatrix(r, 3) = "þ"
            Else
                .TextMatrix(r, 3) = "o"
            End If
            deIms.rsLOGWAR.MoveNext
            r = r + 1
        Loop
    End With
End Sub
Sub makeGrid()
    With logwarGrid
        .TextMatrix(0, 0) = "Code"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "Active"
        .ColWidth(0) = 1600
        .ColWidth(1) = 4200
        .ColWidth(2) = 2000
        .ColWidth(3) = 700
    End With
End Sub



Sub showCOMBO()
    With logwarGrid
        Dim i As Integer
        Call fillCombo
        If .TextMatrix(.row, .Col) <> "" Then
            For i = 1 To combo.Rows - 1
                If .TextMatrix(.row, .Col) = combo.TextMatrix(i, 0) Then
                    combo.row = i
                End If
            Next
        End If
        combo.Top = .CellTop + .CellHeight + .Top
        combo.Left = .ColPos(2) + .Left
        combo.Width = .CellWidth + 400
        combo.ColWidth(0) = combo.Width
        combo.Visible = True
        combo.ZOrder
        combo.SetFocus
    End With
End Sub

Private Sub box_Change()
    If colSwitch Then
        colSwitch = False
    Else
        oldVALUE(val(box.Tag)) = box.Text
        logwarGrid.TextMatrix(logwarGrid.row, logwarGrid.Col) = box.Text
    End If
End Sub

Private Sub box_KeyPress(KeyAscii As Integer)
    With logwarGrid
        Select Case KeyAscii
            Case 13
                .TextMatrix(.row, .Col) = box
            Case 27
                box = originalVALUE(.Col)
            Case Else
                Exit Sub
        End Select
    End With
End Sub


Private Sub box_LostFocus()
    If dontClose Then
        logwarGrid.row = oldRow
        logwarGrid.Col = oldCol
        dontClose = False
        box.Visible = True
        box.SetFocus
    Else
        box.Visible = False
    End If
End Sub

Private Sub box_Validate(Cancel As Boolean)
    If Not getOut Then
        If box.Text = "" Then
            MsgBox "Please enter a valid Warehouse Code"
            dontClose = True
        End If
    End If
End Sub


Private Sub combo_Click()
    combo.Visible = False
    logwarGrid.TextMatrix(logwarGrid.row, logwarGrid.Col) = combo
End Sub

Private Sub combo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call combo_Click
        Case 27
            combo.Visible = False
            Exit Sub
    End Select
    combo.Visible = False
    logwarGrid.SetFocus
End Sub

Private Sub combo_LostFocus()
    combo.Visible = False
End Sub


Private Sub Form_Load()
Dim ctl As Control
    lastMark = 0
    Screen.MousePointer = vbHourglass
    Call makeGrid
    msg1 = translator.Trans("M00126")
    Modify = IIf(msg1 = "", "Modification", msg1)
    msg1 = translator.Trans("M00092")
    Visualize = IIf(msg1 = "", "Visualization", msg1)
    msg1 = translator.Trans("M00125")
    Create = IIf(msg1 = "", "Creation", msg1)
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    Caption = Caption + " - " + Tag
    RecSaved = True
    dontClose = False
    getOut = False
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    If deIms.rsLOGWAR.State <> 0 Then deIms.rsLOGWAR.Close
    Call deIms.LOGWAR(deIms.NameSpace)
    Set NavBar1.Recordset = deIms.rsLOGWAR
    Call FillGrid
    deIms.rsLOGWAR.Close
    Call doClone
    logwarGrid.Enabled = False
    
    NVBAR_EDIT = NavBar1.EditEnabled
    NVBAR_ADD = NavBar1.NewEnabled
    NVBAR_SAVE = NavBar1.SaveEnabled
    If logwarGrid.Rows >= 2 Then
        NavBar1.EditEnabled = True
    Else
        If logwarGrid.TextMatrix(1, 0) <> "" Then
            NavBar1.EditEnabled = True
        End If
    End If
    
    NavBar1.EditVisible = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    NavBar1.CloseEnabled = True
    NavBar1.Width = 5050
    With frm_logical
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
    Screen.MousePointer = vbDefault
End Sub

Sub showBOX(Col As Integer, Optional bottom As Boolean)
Dim x As Integer
Dim y As Integer
    With logwarGrid
        colSwitch = True
        box.Text = ""
        box.BackColor = vbYellow
        box.ForeColor = vbBlack
        If .row = 0 And .FixedRows > 0 Then .row = 1
        box.Height = .RowHeight(.row)
        If .row = 1 Then
            box.Height = box.Height - 20
        Else
            box.Height = box.Height + 10
        End If
        Select Case Col
            Case 0
                box.MaxLength = 10
            Case 1
                box.MaxLength = 40
        End Select
        x = leftCOL(Col)
        box.Left = x
        Dim n As Integer
        If .Rows > 10 Then
            n = 3
        Else
            n = 0
        End If
        ''y = topROW(.row - n)
        y = .RowPos(.row) + 20
        If (y > .Height) Then
            y = .RowPos(.row) + .Top - .row + 30
        Else
            y = y + .Top
        End If
        box.Top = y
        box.Width = .ColWidth(Col) - 20
        box.Visible = True
        oldVALUE(Col) = .TextMatrix(.row, Col)
        originalVALUE(Col) = .TextMatrix(.row, Col)
        
        box.Text = .TextMatrix(.row, Col)


        box.Tag = Col
        box.SetFocus
        box.ZOrder
    End With
End Sub
Function leftCOL(Col) As Integer
Dim x As Integer
Dim i As Integer
    With logwarGrid
        x = .Left + 30
        If Col > 0 Then
            For i = 0 To Col - 1
                x = x + .ColWidth(i)
            Next
        End If
    End With
    leftCOL = x + 10
End Function
Function topROW(row, Optional bottom As Boolean) As Integer
Dim y As Integer
Dim i As Integer
Dim n As Integer
    With logwarGrid
        If bottom Then
            n = row
        Else
            n = row - 1
        End If
        y = 20
        For i = 0 To n
            y = y + .RowHeight(row)
        Next
    End With
    If bottom Then
        If row = 1 Then
            y = y + 20
        Else
            y = y + 30
        End If
    Else
        If row = 1 Then
            y = y + 10
        End If
    End If
    topROW = y
End Function
Private Sub Form_Unload(Cancel As Integer)
If TableLocked = True Then    'jawdat
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    currentformname = Forms(3).Name
    Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

Dim response As String
On Error Resume Next
InUnload = True
deIms.rsLOGWAR.Close
 If RecSaved = True Then
    Hide
    deIms.rsLOGWAR.Close
    If open_forms <= 5 Then ShowNavigator
    If Err Then Err.Clear
Else
    Cancel = True
End If
End Sub

Sub logwarGrid_Click()
    With logwarGrid
        If lblStatus.Caption = "Creation" Then
            If .row < .Rows - 1 Then Exit Sub
        End If
        If dontClose Then
            .row = oldRow
            .Col = oldCol
        Else
            box.Visible = False
            Call cleanGrid
            oldRow = .row
            oldCol = .Col
            Select Case .Col
                Case 0
                Case 1
                    Call showBOX(.Col)
                Case 2
                    Call showCOMBO
                Case 3
                    .CellBackColor = vbYellow
                    Call checkBox
            End Select
        End If
    End With
End Sub

Private Sub logwarGrid_GotFocus()
    If dontClose Then box.SetFocus
End Sub

Private Sub logwarGrid_Scroll()
    box.Visible = False
End Sub

Private Sub NavBar1_BeforeNewClick()
    Dim i As Integer
    NavBar1.CancelEnabled = True
    NavBar1.EditEnabled = False
    NavBar1.NewEnabled = False
    NavBar1.SaveEnabled = True
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = Create
    With logwarGrid
        Call colorize(&H80000011)
        .AddItem ""
        .row = .Rows - 1
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = vbBlack
        Next
        .Col = 3
        .CellFontName = "Wingdings"
        .CellFontSize = 16
        .ColAlignment(3) = 4
        .TextMatrix(.row, 3) = "þ"
        .Col = 0
        .Enabled = True
        oldRow = .row
        oldCol = .Col
        .topROW = .Rows - 1
        Call showBOX(0, True)
    End With
End Sub

Private Sub NavBar1_OnCancelClick()
    With logwarGrid
        getOut = True
        If lblStatus.Caption = "Creation" Then
            .RemoveItem .Rows - 1
        End If
        .Enabled = False
        lblStatus.ForeColor = &HFF00&
        lblStatus.Caption = Visualize
        box.Visible = False
        NavBar1.EditEnabled = True
        NavBar1.NewEnabled = True
        NavBar1.CancelEnabled = False
        NavBar1.SaveEnabled = False
    End With
End Sub


Private Sub NavBar1_OnCloseClick()
    If TableLocked = True Then    'jawdat
        Dim imsLock As imsLock.Lock
        Set imsLock = New imsLock.Lock
        currentformname = Forms(3).Name
        Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
    End If
    newRecord = False
    Unload Me
End Sub


Private Sub NavBar1_OnEditClick()
    Call colorize(&HFFFFFFFF)
    NavBar1.CancelEnabled = True
    NavBar1.EditEnabled = False
    NavBar1.SaveEnabled = True
    NavBar1.NewEnabled = False
    lblStatus.ForeColor = &HFF0000
    lblStatus.Caption = "Modify"
    logwarGrid.Enabled = True
End Sub


Private Sub NavBar1_OnSaveClick()
    Dim i As Integer
    Dim sql, Code, Description, codeType, Active As String
    On Error GoTo Err
    Call box_LostFocus
    With logwarGrid
        Dim startPoint As Integer
        If lblStatus.Caption = "Creation" Then
            startPoint = .Rows - 1
        Else
            startPoint = 1
        End If
        For i = startPoint To .Rows - 1
            If changedRow(i) Then
                Code = .TextMatrix(.row, 0)
                Description = .TextMatrix(.row, 1)
                codeType = .TextMatrix(.row, 2)
                Active = IIf(.TextMatrix(.row, 3) = "þ", "1", "0")
                Select Case lblStatus.Caption
                    Case "Creation"
                        sql = "INSERT INTO logwar (lw_code, lw_npecode,  lw_desc, lw_actvflag, lw_type) VALUES (" _
                            + "'" + Code + "', " _
                            + "'" + deIms.NameSpace + "', " _
                            + "'" + Description + "', " _
                            + "" + Active + ", " _
                            + "'" + codeType + "' ) "
                    Case "Modify"
                        sql = "UPDATE logwar SET " _
                            + "lw_desc = '" + Description + "', " _
                            + "lw_actvflag = " + Active + ", " _
                            + "lw_type = '" + codeType + "'  " _
                            + "WHERE lw_npecode = '" + deIms.NameSpace + "'  " _
                            + "AND lw_code = '" + Code + "'"
                End Select
                Dim cmd As New ADODB.Command
                cmd.ActiveConnection = deIms.cnIms
                cmd.CommandText = sql
                Call cmd.Execute(, , adExecuteNoRecords)
            End If
        Next
        .Enabled = False
    End With
    
Err:
    If Err.number > 0 Then MsgBox Err.descriptio
    Err.Clear
    
    box.Visible = False
    NavBar1.EditEnabled = True
    NavBar1.NewEnabled = True
    NavBar1.CancelEnabled = False
    NavBar1.SaveEnabled = False
    lblStatus.ForeColor = &HFF00&
    lblStatus.Caption = Visualize
    logwarGrid.Enabled = False
End Sub


