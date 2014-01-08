VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmCurrency 
   Caption         =   "Currency"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Tag             =   "01010500"
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid history 
      Height          =   3120
      Left            =   285
      TabIndex        =   6
      Top             =   1170
      Visible         =   0   'False
      Width           =   5800
      _ExtentX        =   10239
      _ExtentY        =   5503
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton HistoryButton 
      Caption         =   "&History"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4440
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      FirstVisible    =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      SaveEnabled     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid currencyLIST 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   12
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   285
      AllowBigSelection=   0   'False
      GridLinesFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.MonthView calendar 
      Height          =   2370
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60227585
      CurrentDate     =   36972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Actual Currency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mode As String
Dim currentROW As Integer
Dim bypassFOCUS As Boolean
Dim oldVALUE As String
Dim originalVALUE As String
Dim SaveEnabled As Boolean
Dim lastDATE As Date
Dim TableLocked As Boolean, currentformname As String   'jawdat
Function topROW(row, Optional bottom As Boolean) As Integer
Dim Y As Integer
Dim i As Integer
Dim n As Integer
    With currencyLIST
        If bottom Then
            n = row
        Else
            n = row - 1
        End If
        Y = 20
        For i = 0 To n
            Y = Y + .RowHeight(row)
        Next
    End With
    If bottom Then
        If row = 1 Then
            Y = Y + 20
        Else
            Y = Y + 30
        End If
    Else
        If row = 1 Then
            Y = Y + 10
        End If
    End If
    topROW = Y
End Function

Sub fillHISTORY()
Dim Sql As String
Dim mark As String
Dim rec As String
Dim currDATA As New ADODB.Recordset
Dim currCODE
Screen.MousePointer = 11
    With history
        .Rows = 0
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .Left = currencyLIST.Left + 30
        .Top = currencyLIST.Top + 330
        .Width = currencyLIST.Width - 60
        .Height = currencyLIST.Height - 375
    End With
    currCODE = currencyLIST.TextMatrix(currencyLIST.row, 0)
    Sql = "SELECT curr_code, curr_desc, " _
        & "curd_from, curd_to, curd_value " _
        & "FROM CURRENCY LEFT OUTER JOIN CURRENCYDETL ON " _
        & "curr_code = curd_code AND " _
        & "curr_npecode = curd_npecode WHERE " _
        & "curr_npecode = '" + deIms.NameSpace + "' AND " _
        & "curr_code = '" + currCODE + "' " _
        & "ORDER BY curr_code ASC, curd_to DESC"
    Set currDATA = New ADODB.Recordset
    With currDATA
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If .RecordCount > 0 Then
            Do While Not .EOF
                rec = !curr_code & vbTab
                rec = rec + !curr_desc & vbTab
                rec = rec + Format(!curd_from, "mm/dd/yyyy") & vbTab
                rec = rec + Format(!curd_to, "mm/dd/yyyy") & vbTab
                rec = rec + Format(!curd_value, "0.00000")
                history.AddItem rec
                .MoveNext
            Loop
            With history
                If .Rows > 10 Then
                    .Width = .Width + 285
                    .Left = Int((Me.Width - .Width) / 2)
                    .Height = currencyLIST.Height - 330
                    .Appearance = flex3D
                    .BorderStyle = flexBorderSingle
                End If
            End With
        End If
    End With
Screen.MousePointer = 0
End Sub
Sub fillLIST()
Dim Sql As String
Dim mark As String
Dim rec As String
Dim currDATA As New ADODB.Recordset
Screen.MousePointer = 11
    currencyLIST.Rows = 2
    Sql = "SELECT curr_code, curr_desc, " _
        & "curd_from, curd_to, curd_value " _
        & "FROM CURRENCY LEFT OUTER JOIN CURRENCYDETL ON " _
        & "curr_code = curd_code AND " _
        & "curr_npecode = curd_npecode WHERE " _
        & "curr_npecode = '" + deIms.NameSpace + "'" _
        & "ORDER BY curr_code ASC, curd_creadate DESC"
        'Juan 2010-5-1 replaced the line below by the one above
        '& "ORDER BY curr_code ASC, curd_to DESC"
        '-------------------
    Set currDATA = New ADODB.Recordset
    With currDATA
        .Open Sql, deIms.cnIms, adOpenForwardOnly
        If .RecordCount > 0 Then
            mark = ""
            Do While Not .EOF
                If mark <> !curr_code Then
                    mark = !curr_code
                    rec = !curr_code & vbTab
                    rec = rec + !curr_desc & vbTab
                    rec = rec + Format(!curd_from, "mm/dd/yyyy") & vbTab
                    rec = rec + Format(!curd_to, "mm/dd/yyyy") & vbTab
                    rec = rec + Format(!curd_value, "0.00000")
                    currencyLIST.AddItem rec
                End If
                .MoveNext
            Loop
            currencyLIST.RemoveItem 1
            If currencyLIST.Rows > 10 Then
                currencyLIST.Width = currencyLIST.Width + 285
                currencyLIST.Left = Int((Me.Width - currencyLIST.Width) / 2)
            End If
        End If
    End With
Screen.MousePointer = 0
End Sub

Function leftCOL(Col) As Integer
Dim x As Integer
Dim i As Integer
    With currencyLIST
        x = .Left + 10
        If Col > 0 Then
            For i = 0 To Col - 1
                x = x + .ColWidth(i)
            Next
        End If
    End With
    leftCOL = x + 10
End Function

Sub SETcurrencyLIST()
Dim i As Integer
    With currencyLIST
        For i = 0 To 4
            .ColAlignmentFixed(i) = 4
        Next
        .ColWidth(0) = 800
        .TextMatrix(0, 0) = "Code"
        .ColAlignment(0) = 0
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Description"
        .ColAlignment(1) = 0
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "From Date"
        .ColAlignment(2) = 3
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "To Date"
        .ColAlignment(3) = 3
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Value"
        .ColAlignment(4) = 6
    End With
    With history
        .ColWidth(0) = 800
        .ColAlignment(0) = 0
        .ColWidth(1) = 2000
        .ColAlignment(1) = 0
        .ColWidth(2) = 1000
        .ColAlignment(2) = 3
        .ColWidth(3) = 1000
        .ColAlignment(3) = 3
        .ColWidth(4) = 1000
        .ColAlignment(4) = 6
    End With
End Sub

Sub Coloring(dye)
Dim currentCOL As Integer
Dim i As Integer
    With currencyLIST
        currentCOL = .Col
        For i = 0 To 4
            .Col = i
            .CellBackColor = dye
        Next
        .Col = currentCOL
    End With
End Sub

Private Sub box_Change()
    If val(box.Tag) = 4 Then
        If IsNumeric(box) Then
            oldVALUE = box
        Else
            box = oldVALUE
        End If
    Else
        oldVALUE = box
    End If
End Sub

Private Sub box_KeyPress(KeyAscii As Integer)
    With currencyLIST
        Select Case KeyAscii
            Case 13
                .TextMatrix(currentROW, val(box.Tag)) = box
                If .Col = 0 Or .Col = 1 Then
                    Call box_Validate(True)
                    If box = "" Then Exit Sub
                    bypassFOCUS = True
                    .Col = .Col + 1
                    Call currencyLIST_Click
                End If
            Case 27
                box = originalVALUE
            Case Else
                Exit Sub
        End Select
        Call box_LostFocus
    End With
End Sub


Private Sub box_LostFocus()
Dim Flag As Integer
    If bypassFOCUS Then
        bypassFOCUS = False
    Else
        With currencyLIST
            Flag = .Col
            If currentROW < .Rows Then .row = currentROW
            If box.Tag <> "" Then
                .Col = val(box.Tag)
            End If
            If box <> "" Then .Text = box
            If mode = "" Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = &HC0C0FF
            End If
            .Col = Flag
            box = ""
            box.Tag = ""
            box.Visible = False
            box.Refresh
        End With
        Call box_Validate(True)
    End If
End Sub


Public Sub box_Validate(Cancel As Boolean)
Dim Sql As String
Dim currDATA As New ADODB.Recordset
        
    With currencyLIST
        If .Col = 0 Then
            Sql = "SELECT curr_code  FROM CURRENCY WHERE " _
                & "curr_code = '" + Trim(box) + "' AND " _
                & "curr_npecode = '" + deIms.NameSpace + "'"
            Set currDATA = New ADODB.Recordset
            currDATA.Open Sql, deIms.cnIms, adOpenForwardOnly
            If currDATA.RecordCount > 0 Then
                currencyLIST.Col = 0
                box = ""
                MsgBox "Currency Code already exists"
                currencyLIST.row = currentROW
            End If
        End If
    End With
End Sub


Private Sub calendar_DateClick(ByVal DateClicked As Date)
    currencyLIST.TextMatrix(currentROW, val(calendar.Tag)) = Format(calendar.value, "mm/dd/yyyy")
    calendar.Visible = False
End Sub

Private Sub calendar_DateDblClick(ByVal DateDblClicked As Date)
    Call calendar_KeyPress(13)
End Sub

Private Sub calendar_GotFocus()
Dim dateFROM As Date
Dim dateTO As Date
On Error GoTo repairDATES
    With currencyLIST
        dateFROM = CDate(Format(lastDATE, "mm/dd/yyyy"))
        If .TextMatrix(currentROW, 3) = "" Then
            dateTO = dateFROM
        Else
            dateTO = CDate(.TextMatrix(currentROW, 3))
        End If
        If dateTO < dateFROM Then
            dateTO = dateFROM
        End If
                        
        If .Col = 2 Then
            If calendar.value < dateFROM Then
                calendar.value = dateFROM
            Else
                calendar.value = .TextMatrix(currentROW, 2)
            End If
            calendar.MinDate = dateFROM
            calendar.MaxDate = dateTO
        Else
            calendar.MinDate = dateFROM
            calendar.MaxDate = CDate("12/31/9999")
            
            'Juan 2010-5-1
            'If calendar.value < dateFROM Then
            '    calendar.value = dateFROM
            'Else
            '    If IsDate(.TextMatrix(currentROW, 3)) Then
            '        If calendar.MinDate <= .TextMatrix(currentROW, 3) Then
            '    '        calendar.value = .TextMatrix(currentROW, 3)
            '            calendar.value = .TextMatrix(currentROW, 2)
            '        End If
            '    Else
            '        calendar.value = dateFROM
            '    End If
            'End If
            '----------------------
        End If
    End With
    calendar.MaxDate = CDate("12/31/9999") 'Juan 2010-9-15 to open the max date limit
    calendar.Year = Year(calendar.value)
    calendar.Month = Month(calendar.value)
    Exit Sub
    
repairDATES:
    Select Case Err.number
        Case 35773
            If calendar.MinDate > dateFROM Then
                calendar.MinDate = dateFROM - 1
            End If
            If currencyLIST.Col = 2 Then
                calendar.value = calendar.MinDate
            Else
                If Year(calendar.MaxDate) < 9999 Then
                    If calendar.MaxDate + 1 <= dateTO Then
                        calendar.MaxDate = calendar.MaxDate + 1
                    Else
                    End If
                End If
            End If
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub calendar_KeyPress(KeyAscii As Integer)
Dim Flag As Integer
    Select Case KeyAscii
        Case 13
            Flag = val(calendar.Tag)
            currencyLIST.TextMatrix(currentROW, Flag) = Format(calendar.value, "mm/dd/yyyy")
            Call calendar_LostFocus
            bypassFOCUS = True
            currencyLIST.Col = Flag + 1
            'Juan 2010-5-1
            'Call currencyLIST_Click
            '----------------
        Case 27
            Call calendar_LostFocus
    End Select
End Sub

Public Sub calendar_LostFocus()
Dim Flag As Integer
    If bypassFOCUS Then
        bypassFOCUS = False
    Else
        With currencyLIST
            calendar.Visible = False
            If currentROW < .Rows Then .row = currentROW
            Flag = .Col
            .Col = val(calendar.Tag)
            calendar.Tag = ""
            If mode = "" Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = &HC0C0FF
            End If
            .Col = Flag
        End With
    End If
End Sub

Private Sub calendar_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    With currencyLIST
        If calendar.Tag <> "" Then
            .TextMatrix(currentROW, val(calendar.Tag)) = Format(calendar.value, "mm/dd/yyyy")
        End If
    End With
End Sub

Private Sub Command1_Click()
End Sub

Public Sub currencyLIST_Click()
    With currencyLIST
        .RowSel = .row
        .ColSel = .Col
        Select Case mode
            Case "editing"
                If currentROW = .row Then
                    Select Case .Col
                        Case 2, 3
                            Call showCALENDAR(.Col)
                        Case 4
                            Call showBOX(.Col)
                    End Select
                End If
            Case "new"
                Select Case .Col
                    Case 0, 1, 4
                        Call showBOX(.Col)
                    Case 2, 3
                        Call showCALENDAR(.Col)
                End Select
            Case Else
                .Tag = .row
        End Select
    End With
End Sub

Private Sub currencyLIST_DblClick()
    If SaveEnabled Then
        With currencyLIST
            Select Case mode
                Case "editing"
                Case "new"
                Case Else
                    Call Coloring(&HC0C0FF)
                    mode = "editing"
                    currentROW = .row
                    
                    If IsDate(.TextMatrix(currentROW, 3)) Then
                        lastDATE = CDate(.TextMatrix(currentROW, 2))
                    Else
                        lastDATE = Now
                    End If
'                    .TextMatrix(currentROW, 2) = Format(lastDATE + 1, "mm/yy/yyyy")
'                    .TextMatrix(currentROW, 3) = ""
                    NavBar1.NewEnabled = False
                    NavBar1.SaveEnabled = SaveEnabled
                    Select Case .Col
                        Case 0, 1
                            .Col = 2
                            Call showCALENDAR(.Col)
                        Case 2, 3
                            Call showCALENDAR(.Col)
                        Case 4
                            Call showBOX(.Col)
                    End Select
            End Select
        End With
    End If
End Sub

Sub showBOX(Col As Integer)
Dim x As Integer
Dim Y As Integer
    With currencyLIST
        If .row = 0 And .FixedRows > 0 Then .row = 1
        box.Height = .RowHeight(.row)
        If .row = 1 Then
            box.Height = box.Height - 20
        Else
            box.Height = box.Height + 10
        End If
        x = leftCOL(Col)
        box.Left = x
        Y = topROW(.row)
        box.Top = Y + .Top
        box.Width = .ColWidth(Col) + 10
        box.Visible = True
        box.Text = .TextMatrix(.row, Col)
        oldVALUE = box.Text
        originalVALUE = box.Text
        Select Case .ColAlignment(Col)
            Case 0 To 2
                box.Alignment = 0
            Case 3 To 5
                box.Alignment = 2
            Case 6 To 8
                box.Alignment = 1
        End Select
        Select Case Col
            Case 0
                box.MaxLength = 3
            Case 1
                box.MaxLength = 30
            Case 4
                box.MaxLength = 12
        End Select
        box.Tag = Col
        box.SetFocus
        box.ZOrder
    End With
End Sub

Sub showCALENDAR(Col As Integer)
Dim x As Integer
Dim Y As Integer
    With currencyLIST
        .Col = Col
        If .row = 0 And .FixedRows > 0 Then .row = 1
        x = leftCOL(Col)
        If Col > 2 Then
            x = x + .ColWidth(Col)
            x = x - calendar.Width + 10
        End If
        calendar.Left = x
        Y = topROW(.row, True)
        If (frmCurrency.Height - Y) <= (calendar.Height + 1200) Then
            Y = Y - calendar.Height - .RowHeight(.row)
        End If
        .CellBackColor = &HC0FFFF
        calendar.Top = Y + .Top - 30
        calendar.Visible = True
        calendar.Tag = Col
        calendar.SetFocus
        calendar.ZOrder
    End With
End Sub

Private Sub currencyLIST_EnterCell()
    Select Case mode
        Case "editing"
        Case "new"
        Case Else
            Call Coloring(&HFFC0C0)
    End Select
End Sub

Private Sub currencyLIST_RowColChange()
Dim currentROW
    Select Case mode
        Case "editing"
        Case "new"
        Case Else
            With currencyLIST
                If IsNumeric(.Tag) Then
                    If val(.Tag) <> .row Then
                        currentROW = .row
                        .row = val(.Tag)
                        If .CellBackColor <> vbWhite Then
                            Call Coloring(vbWhite)
                        End If
                        .row = currentROW
                        .Tag = .row
                    End If
                End If
            End With
    End Select
End Sub


Private Sub Form_Load()

'copy begin here

'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   currencyLIST.Enabled = False
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
'    Exit Sub
    Else
      TableLocked = True
Dim rights
      rights = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
  
          SaveEnabled = rights
    NavBar1.NewEnabled = SaveEnabled
    NavBar1.SaveEnabled = SaveEnabled
    mode = ""
    End If
'End If

'end copy





    Me.Caption = Me.Caption + " - " + Me.Tag
    
    

    Call SETcurrencyLIST
    Call fillLIST

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    
    
    Unload Me
End Sub


Private Sub HistoryButton_Click()
    With HistoryButton
        If .Caption = "&History" Then
            .Caption = "&Current"
            Label1 = "Currency History"
            history.Visible = True
            Call fillHISTORY
        Else
            .Caption = "&History"
            Label1 = "Actual Currency"
            history.Visible = False
        End If
    End With
End Sub

Private Sub NavBar1_BeforeSaveClick()
Dim i As Integer
Dim id As Integer
Dim Sql As String
Dim currDATA As New ADODB.Recordset
Dim dateFROM As Date
Dim dateTO As Date
On Error GoTo errCLOSE
    deIms.cnIms.BeginTrans
    Screen.MousePointer = 11

    'Juan 2010-5-1
    If box.Visible Then
        If val(box.Tag) = 4 Then
            Call box_LostFocus
        End If
    End If
    '--------------------

    For i = 0 To 4
        If currencyLIST.TextMatrix(currentROW, i) = "" Then
            Screen.MousePointer = 0
            MsgBox "Incomplete data"
            currencyLIST.Col = i
            Call currencyLIST_Click
            NavBar1.SaveEnabled = SaveEnabled
            Exit Sub
        End If
    Next
    FrmShowApproving.Show
    FrmShowApproving.Label2 = "Saving Currency"
    FrmShowApproving.Refresh
    
    Select Case mode
        Case "editing"
            Sql = "UPDATE CURRENCY SET " _
                & "curr_modidate = '" + Format(Now, "yyyy/mm/dd") + "', " _
                & "curr_modiuser = '" + CurrentUser + "' WHERE " _
                & "curr_code = '" + currencyLIST.TextMatrix(currentROW, 0) + "' AND " _
                & "curr_npecode = '" + deIms.NameSpace + "' "
        Case "new"
            Sql = "INSERT INTO CURRENCY (curr_code, curr_npecode, curr_desc, curr_creauser) " _
                & "VALUES ( " _
                & "'" + currencyLIST.TextMatrix(currentROW, 0) + "', " _
                & "'" + deIms.NameSpace + "', " _
                & "'" + currencyLIST.TextMatrix(currentROW, 1) + "', " _
                & "'" + CurrentUser + "')"
    End Select
    dateFROM = Format(CDate(currencyLIST.TextMatrix(currentROW, 2)), "yyyy/mm/dd")
    dateTO = Format(CDate(currencyLIST.TextMatrix(currentROW, 3)), "yyyy/mm/dd")
    deIms.cnIms.Execute Sql
    
    Sql = "SELECT max(curd_id) as id FROM CURRENCYDETL WHERE " _
        & "curd_code = '" + currencyLIST.TextMatrix(currentROW, 0) + "' AND " _
        & "curd_npecode = '" + deIms.NameSpace + "' "
    Set currDATA = New ADODB.Recordset
    currDATA.Open Sql, deIms.cnIms, adOpenForwardOnly
    If currDATA.RecordCount > 0 Then
        If IsNull(currDATA!id) Then
            id = 1
        Else
            id = currDATA!id + 1
        End If
    Else
        id = 1
    End If
    
    Sql = "INSERT INTO CURRENCYDETL (" _
        & "curd_id, " _
        & "curd_npecode, " _
        & "curd_code, " _
        & "curd_from, " _
        & "curd_to, " _
        & "curd_value, " _
        & "curd_creauser) " _
        & "VALUES ( " _
        & Format(id) + ", " _
        & "'" + deIms.NameSpace + "', " _
        & "'" + currencyLIST.TextMatrix(currentROW, 0) + "', " _
        & "'" + Format(dateFROM, "yyyy/mm/dd") + "', " _
        & "'" + Format(dateTO, "yyyy/mm/dd") + "', " _
        & currencyLIST.TextMatrix(currentROW, 4) + ", " _
        & "'" + CurrentUser + "')"
    deIms.cnIms.Execute Sql
    Call NavBar1_OnCancelClick
    Unload FrmShowApproving
    Me.ZOrder
    deIms.cnIms.CommitTrans
    Screen.MousePointer = 0
    Exit Sub
    
errCLOSE:
    Err.Clear
    deIms.cnIms.RollbackTrans
    Call NavBar1_OnCancelClick
End Sub

Public Sub NavBar1_OnCancelClick()
    mode = ""
    box.Visible = False
    calendar.Visible = False
    NavBar1.NewEnabled = SaveEnabled
    NavBar1.SaveEnabled = SaveEnabled
    Call fillLIST
End Sub


Private Sub NavBar1_OnCloseClick()
     
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

Unload Me
End Sub

Private Sub NavBar1_OnNewClick()
    With currencyLIST
        If mode = "" Then
            .AddItem ""
            mode = "new"
            .row = .Rows - 1
            currentROW = .row
            Call Coloring(&HC0C0FF)
            Call showBOX(0)
            lastDATE = Now
            NavBar1.SaveEnabled = SaveEnabled
        End If
    End With
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
Screen.MousePointer = 11
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Currency.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        msg1 = translator.Trans("L00047")
        .WindowTitle = IIf(msg1 = "", "Currency", msg1)
        Call translator.Translate_Reports("Currency.rpt")
        .Action = 1
        Screen.MousePointer = 0
        .Reset
    End With
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub


