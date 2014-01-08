VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.0#0"; "LRNAVI~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStandardWarehouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Supplier"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Tag             =   "02050700"
   Begin VB.CommandButton Command2 
      Caption         =   "&All Columns"
      Height          =   375
      Left            =   8640
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   10080
      TabIndex        =   24
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox userLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   8160
      TabIndex        =   22
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox TextLINE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   11040
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show Only Selection"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox nomPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox nomPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox cell 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   6960
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelEnabled   =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      FirstVisible    =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      Mode            =   3
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   16777215
      CustomFormat    =   "MMMM/dd/yyyy"
      Format          =   23003139
      CurrentDate     =   36867
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid STOCKlist 
      Height          =   4035
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   18
      RowHeightMin    =   285
      GridColorFixed  =   0
      HighLight       =   0
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox remark 
      Height          =   915
      Left            =   120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1080
      Width           =   11775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   975
      Left            =   4440
      TabIndex        =   20
      Top             =   630
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
      ScrollBars      =   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid TransactionComboList 
      Height          =   975
      Left            =   2280
      TabIndex        =   7
      Top             =   630
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
      ScrollBars      =   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid POComboList 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   630
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
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   10080
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label nomLabel 
      Caption         =   "Purchase Unit"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label nomLabel 
      Caption         =   "Already Invoiced"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label 
      Caption         =   "Transaction #"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label 
      Caption         =   "Company"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label 
      Caption         =   "Warehouse"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   1815
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
      Left            =   7560
      TabIndex        =   1
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmStandardWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Form As FormMode
Dim readyFORsave As Boolean
Dim Rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim colorsROW(12)
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    If cellACTIVE <> "" Then
        With gridACTIVE
            If Not .Visible Then .Visible = True
            If IsNumeric(.Tag) Then
                .ROW = Val(.Tag)
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
                    .ROW = i
                    .CellBackColor = &H800000 'Blue
                    .CellForeColor = &HFFFFFF 'White
                    .Tag = .ROW
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                .ROW = 0
                .Tag = ""
            End If
            If IsNumeric(.Tag) Then .TopRow = Val(.Tag)
        End With
    End If
End Sub

Sub arrowKEYS(direction As String, Index As Integer)
Dim Grid As MSHFlexGrid
    With cell(Index)
        Select Case Index
            Case 0
                Set Grid = POComboList
            Case 1
                Set Grid = TransactionComboList
                
        End Select
        
        Select Case Index
            Case 0, 1
                If IsNumeric(Grid.Tag) Then
                    Grid.ROW = Val(Grid.Tag)
                    Grid.CellBackColor = &HFFFF00   'Cyan
                    Grid.CellForeColor = &H80000008 'Default Window Text
                End If
                Select Case direction
                Case "down"
                    If Grid.ROW < (Grid.Rows - 1) Then
                        If Grid.ROW = 0 And .Text = "" Then
                            .Text = Grid.Text
                        Else
                            Grid.ROW = Grid.ROW + 1
                        End If
                    Else
                        Grid.ROW = Grid.Rows - 1
                    End If
                Case "up"
                    If Grid.ROW > 0 Then
                        Grid.ROW = Grid.ROW - 1
                    Else
                        Grid.ROW = 1
                    End If
            End Select
            If Not Grid.Visible Then
                Grid.Visible = True
            End If
            Grid.ZOrder
            Grid.TopRow = Grid.ROW
            Grid.SetFocus
        End Select
    End With
End Sub

Sub BeforePrint()
    With MDI_IMS.CrystalReport1
        .Reset
        'msg1 = translator.Trans("L00176")
        .WindowTitle = IIf(msg1 = "", "transaction", msg1)
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        If cell(1) = "" Then
            .ReportFileName = FixDir(App.Path) + "CRreports\transactionGlobal.rpt"
            .ParameterFields(1) = "ponumb;" + cell(0) + ";TRUE"
            'call translator.Translate_Reports("transactionGlobal.rpt")
        Else
            .ReportFileName = FixDir(App.Path) + "CRreports\transaction.rpt"
            .ParameterFields(1) = "invnumb;" + cell(1) + ";TRUE"
            .ParameterFields(2) = "ponumb;" + cell(0) + ";TRUE"
            'Call translator.Translate_Reports("transaction.rpt")
            'Call translator.Translate_SubReports
        End If
    End With
End Sub

Sub begining()
Dim i
'    With supplierDATA
'        .ColWidth(0) = 900
'        .ColWidth(1) = 3000
'        .ColAlignmentFixed(0) = 6
'        .ColAlignment(1) = 1
'        .TextMatrix(0, 0) = "Supplier"
'        .TextMatrix(1, 0) = "Address"
'        .TextMatrix(3, 0) = "City"
'        .TextMatrix(4, 0) = "State"
'        .TextMatrix(5, 0) = "Country"
'        .TextMatrix(6, 0) = "Zip"
'    End With
End Sub


Sub colorCOLS()
Dim i As Integer
    With STOCKlist
        .ROW = STOCKlist.Rows - 1
        .Col = 3
        .CellBackColor = &HE0E0E0
        .Col = 7
        .CellBackColor = &HE0E0E0
        .Col = 11
        .CellBackColor = &HE0E0E0
        For i = 8 To 10
            .Col = i
            If Val(.TextMatrix(.ROW, 17)) = 0 Then
                .CellBackColor = &HC0FFFF 'Very Light Yellow
            Else
                .CellBackColor = &HFFFFC0 'Very Light Green
            End If
        Next
    End With
End Sub

Sub differences(ROW As Integer)
Dim d1, d2 As Double
Dim s1, s2 As String
Dim Col, currentROW As Integer
    s1 = STOCKlist.TextMatrix(ROW, 6)
    s2 = STOCKlist.TextMatrix(ROW, 10)
    
    Select Case s1
        Case Is = "", 0
            d1 = 0
        Case Else
            If IsNull(s1) Then
                d1 = 0
            Else
                d1 = CDbl(s1)
            End If
    End Select
    
    Select Case s2
        Case "", 0
            d2 = 0
        Case Else
            If IsNull(s2) Then
                d2 = 0
            Else
                d2 = CDbl(s2)
            End If
    End Select
    
    If IsNumeric(s1) And IsNumeric(s2) Then
        STOCKlist.TextMatrix(ROW, 12) = FormatNumber((d2 - d1), 2)
        Col = STOCKlist.Col
        STOCKlist.Col = 12
        currentROW = STOCKlist.ROW
        STOCKlist.ROW = ROW
        If (d2 - d1) >= 0 Then
            STOCKlist.CellForeColor = vbBlack
        Else
            STOCKlist.CellForeColor = vbRed
        End If
        STOCKlist.Col = Col
        STOCKlist.ROW = currentROW
    End If
End Sub

Sub drawLINEcol(ByVal Grid As MSHFlexGrid, Col As Integer)
    With Grid
        .ColWidth(Col) = 50 'Line
        .Col = Col
        .CellBackColor = &H808080
    End With
End Sub

Sub getCOLORSrow()
Dim i, currentCOL As Integer
    currentCOL = STOCKlist.Col
    For i = 1 To 12
        STOCKlist.Col = i
        colorsROW(i) = STOCKlist.CellBackColor
    Next
    STOCKlist.Col = currentCOL
End Sub

Sub gettransaction(transaction As String)
On Error Resume Next
Dim datatransaction  As New ADODB.Recordset
Dim sql As String
        
    Screen.MousePointer = 11
    Call clearDOCUMENT
    
    'Header
    If Left(cell(0), 1) <> "(" And Right(cell(0), 1) <> ")" Then cell(0) = UCase(cell(0))
    If transaction = "*" Then
        cell(1) = ""
        sql = "SELECT * from PO_Header_for_transaction WHERE NameSpace = '" + deIms.NameSpace + "' " _
        & "AND PO = '" + cell(0) + "'"
    Else
        If cell(1) = "" Then
            sql = "SELECT * from PO_Header_for_transaction WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "'"
        Else
            sql = "SELECT * from transaction_Header WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND transaction = '" + Trim(cell(1).Text) + "'"
        End If
    End If
    Set datatransaction = New ADODB.Recordset
    datatransaction.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.Number <> 0 Then Exit Sub
        
    With datatransaction
        If .RecordCount > 0 Then
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
            cell(0) = !po
            cell(2) = IIf(IsNull(!UserName), "", !UserName)
            cell(3) = IIf(IsNull(!transactiondDate), "", !transactiondDate)
            cell(4) = IIf(IsNull(!CreatedDate), "", !CreatedDate)
            cell(5) = IIf(IsNull(!Currency), "", !Currency)
            cell(6) = IIf(IsNull(!DateIssued), "", !DateIssued)
            cell(7) = IIf(IsNull(!DateRequested), "", !DateRequested)
            cell(8) = IIf(IsNull(!Buyer), "", !Buyer)
            cell(9) = IIf(IsNull(!BuyerPhone), "", !BuyerPhone)
            remark = IIf(IsNull(!remarks), "", !remarks)
                        
'            supplierDATA.TextMatrix(0, 1) = IIf(IsNull(!Supplier), "", !Supplier)
'            supplierDATA.TextMatrix(1, 1) = IIf(IsNull(!address1), "", !address1)
'            supplierDATA.TextMatrix(2, 1) = IIf(IsNull(!address2), "", !address2)
'            supplierDATA.TextMatrix(3, 1) = IIf(IsNull(!City), "", !City)
'            supplierDATA.TextMatrix(4, 1) = IIf(IsNull(!State), "", !State)
'            supplierDATA.TextMatrix(5, 1) = IIf(IsNull(!Country), "", !Country)
'            supplierDATA.TextMatrix(6, 1) = IIf(IsNull(!Zip), "", !Zip)
'            supplierDATA.TextMatrix(7, 1) = IIf(IsNull(!Telephone), "", !Telephone)
            
            'Details
            Err.Clear
            If transaction = "*" Then
                Call getLINEitems("*")
                cell(0).SelStart = 0
                cell(0).SelLength = Len(cell(0))
                cell(0).SetFocus
                cell(1) = ""
                POComboList.Visible = True
            Else
                Call getLINEitems(cell(1))
                cell(1).SelStart = 0
                cell(1).SelLength = Len(cell(1))
                cell(1).SetFocus
            End If
            NavBar1.NewEnabled = True
        Else
            NavBar1.PrintEnabled = False
            NavBar1.EMailEnabled = False
            Screen.MousePointer = 0
            'msg1 = translator.Trans("M00088")
            MsgBox IIf(msg1 = "", "Does not exist yet", msg1)
            cell(0) = ""
        End If
    End With
    
    Screen.MousePointer = 0
End Sub

Sub gettransactionComboList()
Dim sql As String
Dim dataLIST As ADODB.Recordset
    Err.Clear
    Set dataLIST = New ADODB.Recordset
    sql = "SELECT inv_invcnumb FROM transaction " _
        & "WHERE inv_npecode = '" + deIms.NameSpace + "'"
    If cell(0) <> "(By transaction)" And cell(0) <> "" Then
        sql = sql + " AND inv_ponumb = '" + Trim(cell(0).Text) + "' "
    End If
    sql = sql + " ORDER BY inv_invcnumb"
    dataLIST.Open sql, deIms.cnIms, adOpenForwardOnly
    
    With TransactionComboList
        .Visible = False
        .ColWidth(0) = 1600
        .Clear
        .Rows = 0
        .ColAlignment(0) = 1
    End With
    If Err.Number = 0 Then
        If dataLIST.RecordCount > 0 Then
            Do While Not dataLIST.EOF
                TransactionComboList.AddItem " " + Trim(dataLIST!inv_invcnumb)
                dataLIST.MoveNext
            Loop
            TransactionComboList.ROW = 0
            TransactionComboList.RowHeightMin = 240
        End If
    End If
End Sub


Function isOPEN(po As String) As Boolean
Dim sql As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    isOPEN = False
    po = Trim(cell(0))
    sql = "SELECT po_ponumb, po_stas from PO WHERE po_npecode = '" + deIms.NameSpace + "' " _
        & "AND po_ponumb = '" + cell(0) + "'"
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.Number <> 0 Then Exit Function
    If dataPO.RecordCount > 0 Then
        If dataPO!po_stas = "OP" Then
            isOPEN = True
        Else
            isOPEN = False
        End If
    Else
        isOPEN = False
    End If
End Function

Sub markROW()
Dim nextROW, originalROW, purchaseUNIT As String
Dim i  As Integer
    With STOCKlist
        If Val(.TextMatrix(.ROW, 17)) > 0 Then
            MsgBox "You have already transactiond this line item.  Please print a report before continue"
            Exit Sub
        End If
        originalROW = .ROW
        Select Case .TextMatrix(.ROW, 1)
            Case ""
                Exit Sub
            Case "§"
                nextROW = "UP"
            Case Else
                If .ROW < .Rows - 1 Then
                    .ROW = .ROW + 1
                    If .TextMatrix(.ROW, 1) = "§" Then
                        nextROW = "DOWN"
                    Else
                        nextROW = "NO"
                    End If
                    .ROW = .ROW - 1
                Else
                    nextROW = "NO"
                End If
        End Select
        .Col = 0
        For i = 1 To 2
            .CellFontName = "Wingdings 3"
            .CellFontSize = 10
            If .Text = "" Then
                .Text = "Æ"
                .TextMatrix(.ROW, 9) = .TextMatrix(.ROW, 5)
                purchaseUNIT = Trim(.TextMatrix(.ROW, 15))
                If purchaseUNIT = "P" Or purchaseUNIT = "" Then
                    If i = 1 Then
                        .TextMatrix(.ROW, 10) = .TextMatrix(.ROW, 6)
                        .TextMatrix(.ROW, 12) = "00.0"
                    Else
                        .TextMatrix(.ROW, 8) = .TextMatrix(.ROW, 4)
                    End If
                Else
                    If i = 1 Then
                        .TextMatrix(.ROW, 8) = .TextMatrix(.ROW, 4)
                    Else
                        .TextMatrix(.ROW, 10) = .TextMatrix(.ROW, 6)
                        .TextMatrix(.ROW, 12) = "00.0"
                    End If
                End If
            Else
                .Text = ""
                .TextMatrix(.ROW, 8) = ""
                .TextMatrix(.ROW, 9) = ""
                .TextMatrix(.ROW, 10) = ""
                .TextMatrix(.ROW, 12) = ""
            End If

            Select Case nextROW
                Case "UP"
                    .ROW = .ROW - 1
                Case "DOWN"
                    If .ROW < .Rows - 1 Then
                        .ROW = .ROW + 1
                    End If
                Case "NO"
                    Exit For
            End Select
        Next
        .ROW = originalROW
    End With
End Sub

Sub clearDOCUMENT()
Dim i As Integer
    readyFORsave = False
    For i = 2 To 9
        cell(i) = ""
        cell(i).BackColor = remark.BackColor
    Next
'    For i = 0 To 6
'        supplierDATA.TextMatrix(i, 1) = ""
'    Next
    POComboList.Visible = False
    TransactionComboList.Visible = False
    remark = ""
    nomPicture(0).Visible = False
    nomLabel(0).Visible = False
    Command1.Caption = "&Show Only Selection"
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
        If c.ROW = 0 Then
            .Top = c.Top
            .Height = .Height - 80
        Else
            h = 20
            For i = 0 To c.ROW - 1
                h = h + c.RowHeight(i)
            Next
            .Top = h + c.Top - 30
            .Height = .Height + 10
        End If
        .Visible = True
        .Value = IIf(IsDate(c.Text), c.Text, Now)
        .SetFocus
        Call DTPicker1_DropDown
    End With
End Sub

Sub getPOComboList()
On Error Resume Next
Dim sql As String
Dim datPO As New ADODB.Recordset

    Err.Clear
    With POComboList
        .Visible = False
        .ColWidth(0) = 1600
        .ColAlignment(0) = 1
    End With
    
    Set datPO = New ADODB.Recordset
        
    sql = "SELECT po_ponumb FROM PO WHERE po_npecode = '" + deIms.NameSpace + "' " _
        & "ORDER BY po_ponumb"
    
    POComboList.Rows = 0
    With datPO
        .Open sql, deIms.cnIms, adOpenForwardOnly
        If Err.Number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            POComboList.AddItem "(By transaction)"
            Do While Not .EOF
                POComboList.AddItem Trim(!PO_PONUMB)
                .MoveNext
            Loop
        End If
        POComboList.ROW = 0
        POComboList.RowHeightMin = 240
    End With
End Sub

Sub getLINEitems(transaction As String)
Dim dataPO As New ADODB.Recordset
Dim sql, rowTEXT, stock As String
Dim i As Integer
Dim qty As Double

    On Error Resume Next
    Screen.MousePointer = 11
    Call makeDETAILgrid
    If transaction = "*" Then
        sql = "SELECT * from PO_Details_For_transaction WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "' ORDER BY PO, CONVERT(integer, LineItem)"
    Else
        transaction = Trim(transaction)
        sql = "SELECT * from transaction_Details WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "' AND transaction = '" + transaction + "' ORDER BY PO, CONVERT(integer, LineItem)"
    End If
    STOCKlist.RowHeightMin = 0
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.Number <> 0 Then Exit Sub
    With dataPO
        If .RecordCount > 0 Then
            Do While Not .EOF
                rowTEXT = "" + vbTab
                rowTEXT = rowTEXT + IIf(IsNull(!LineItem), "", !LineItem) + vbTab 'PO Line Item
                stock = IIf(IsNull(!StockNumber), "", Trim(!StockNumber)) + " - " + IIf(IsNull(!Description), "", !Description)
                rowTEXT = rowTEXT + stock + vbTab 'Stock Number + Description
                rowTEXT = rowTEXT + "" + vbTab 'Line
                
                'Purchase
                rowTEXT = rowTEXT + FormatNumber(!Quantity1, 2) + vbTab 'Primary Quantity
                rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice1), 0, !UnitPrice1), 2) + vbTab 'Primary Unit Price
                
                'transaction
                rowTEXT = rowTEXT + "" + vbTab 'Line
                If transaction = "*" Then
                    If IsNumeric(!SumQty1) Then
                        qty = !SumQty1
                    Else
                        qty = 0
                    End If
                    rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Sumary Primary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!SumUnitPrice1), "", !SumUnitPrice1), 2) + vbTab 'Sumary Primary Unit Price
                Else
                    If IsNumeric(!QuantityI1) Then
                        qty = !QuantityI1
                    Else
                        qty = 0
                    End If
                    rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Primary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPriceI1), 0, !UnitPriceI1), 2) + vbTab 'Primary Unit Price
                End If
                
                STOCKlist.AddItem rowTEXT
                STOCKlist.ROW = STOCKlist.Rows - 1
                STOCKlist.TextMatrix(STOCKlist.ROW, 16) = !Unit1Code
                STOCKlist.TextMatrix(STOCKlist.ROW, 17) = IIf(IsNull(!transactions), 0, !transactions)
                Call colorCOLS
                Call differences(STOCKlist.ROW)
                If !unit1 = !unit2 Then
                    STOCKlist.TextMatrix(STOCKlist.ROW, 15) = ""
                Else
                    STOCKlist.TextMatrix(STOCKlist.ROW, 15) = !UnitSwitch
                    nomPicture(0).Visible = True
                    nomLabel(0).Visible = True
                    STOCKlist.RowHeight(STOCKlist.ROW) = 240
                    rowTEXT = "" + vbTab + "" + vbTab + "" + vbTab
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    
                    'Purchase
                    rowTEXT = rowTEXT + FormatNumber(!Quantity2, 2) + vbTab 'Secundary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Secundary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice2), 0, !UnitPrice2), 2) + vbTab 'Secundary Unit Price
                    
                    'transaction
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    If transaction = "*" Then
                        If IsNumeric(!SumQty2) Then
                            qty = !SumQty2
                        Else
                            qty = 0
                        End If
                        rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Sumary Primary Quantity
                        rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!SumUnitPrice2), "", !SumUnitPrice2), 2) + vbTab 'Sumary Primary Unit Price
                    Else
                        If IsNumeric(!QuantityI2) Then
                            qty = !QuantityI2
                        Else
                            qty = 0
                        End If
                        rowTEXT = rowTEXT + IIf(qty = 0, "", FormatNumber(qty, 2)) + vbTab   'Primary Quantity
                        rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPriceI2), 0, !UnitPriceI2), 2) + vbTab 'Primary Unit Price
                    End If
                    
                    STOCKlist.AddItem rowTEXT
                    STOCKlist.ROW = STOCKlist.Rows - 1
                    STOCKlist.TextMatrix(STOCKlist.ROW, 15) = !UnitSwitch
                    STOCKlist.TextMatrix(STOCKlist.ROW, 16) = !Unit2Code
                    STOCKlist.TextMatrix(STOCKlist.ROW, 17) = IIf(IsNull(!transactions), 0, !transactions)
                    Call colorCOLS
                    STOCKlist.Col = 1
                    STOCKlist = "§"
                    STOCKlist.CellFontName = "Wingdings"
                    'stocklist.CellFontSize = 8
                    Call differences(STOCKlist.ROW)
                    If UCase(Trim(!UnitSwitch)) = "P" Or IsNull(!UnitSwitch) Then STOCKlist.ROW = STOCKlist.Rows - 2
                    For i = 4 To 6
                        STOCKlist.Col = i
                        STOCKlist.CellBackColor = &HC0C0FF
                    Next
                    
                    STOCKlist.ROW = STOCKlist.Rows - 1
                End If
                
                STOCKlist.RowHeight(STOCKlist.ROW) = 240
                STOCKlist.AddItem ""
                STOCKlist.ROW = STOCKlist.Rows - 1
                For i = 0 To STOCKlist.Cols - 1
                    STOCKlist.Col = i
                    If i = 0 Then
                        STOCKlist.CellBackColor = &H808080
                    Else
                        STOCKlist.CellBackColor = &HE0E0E0
                    End If
                Next
                STOCKlist.RowHeight(STOCKlist.ROW) = 50
                STOCKlist.TextMatrix(STOCKlist.ROW, 13) = 50
                .MoveNext
            Loop
            STOCKlist.RemoveItem (1)
            STOCKlist.RemoveItem (STOCKlist.Rows - 1)
            STOCKlist.ROW = 0
        End If
    End With
    Screen.MousePointer = 0
End Sub


Sub gridLIST(ByVal mainGRID As MSHFlexGrid, ByVal childGRID As MSHFlexGrid)
Dim h, i As Integer
    
    With childGRID
        .Left = mainGRID.Left + mainGRID.ColWidth(0)
        h = 20
        For i = 0 To mainGRID.ROW
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
        x = .Col
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = &H800000   'Blue
            .CellForeColor = &HFFFFFF   'White
        Next
        .Col = x
        .Tag = .ROW
    End With
End Sub

Sub lockDOCUMENT(locked As Boolean)
Dim i As Integer
    
    If locked Then
        cell(3).locked = True
    Else
        cell(3).locked = False
    End If
    
    If locked Then
        remark.locked = True
    Else
        remark.locked = False
    End If
End Sub

Sub makeDETAILgrid()
Dim i, Col As Integer
    With STOCKlist
        .Clear
        .Rows = 2
        
        'Col 0
        .ColAlignment(0) = 4
        .ColWidth(0) = 285
        .ROW = 0
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .TextMatrix(0, 0) = "®"
                
        For i = 1 To .Cols - 1
            .ColAlignment(i) = 0
            .ColAlignmentFixed(i) = 4
        Next
                
        .TextMatrix(0, 1) = "Commodity"
        .ColWidth(1) = 1000
        .TextMatrix(0, 2) = "Description"
        .ColWidth(2) = 3000
        .TextMatrix(0, 3) = "Purchase Qty"
        .ColWidth(3) = 800
        .TextMatrix(0, 4) = "Country of Origin"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Own"
        .ColWidth(5) = 500
        .TextMatrix(0, 6) = "Lease Company"
        .ColWidth(6) = 1200
        .TextMatrix(0, 7) = "Logical Warehouse"
        .ColWidth(7) = 1200
        .TextMatrix(0, 8) = "Sub Location"
        .ColWidth(8) = 1000
        .TextMatrix(0, 9) = "Unit Price"
        .ColWidth(9) = 800
        .TextMatrix(0, 10) = "Currency"
        .ColWidth(10) = 800
        .TextMatrix(0, 11) = "Currency Value"
        .ColWidth(11) = 800
        .TextMatrix(0, 12) = "Quantity"
        .ColWidth(12) = 800
        .TextMatrix(0, 13) = "Unit"
        .ColWidth(13) = 800
        .TextMatrix(0, 14) = "Pool"
        .ColWidth(14) = 500
        .TextMatrix(0, 15) = "Serial Number"
        .ColWidth(15) = 800
        
'        Call drawLINEcol(stocklist, 3)
'        For i = 0 To 2
'            Col = i * 4
'            Call drawLINEcol(stocklist, 3 + Col)
'        Next
'        .TextMatrix(0, 12) = "Unit Price Difference"
'
'        'Invisible columns
'        For i = 13 To 17
'            .ColWidth(i) = 0
'        Next
'        .TextMatrix(0, 13) = "Real Height"
'        .TextMatrix(0, 14) = "Old value"
'        .TextMatrix(0, 15) = "Switch"
'        .TextMatrix(0, 16) = "Unit of Mesure Code"
'        .TextMatrix(0, 16) = "transactions"
'        .ROW = 1
'        .Col = 1
        .RowHeight(0) = 500
        .RowHeightMin = 240
        .WordWrap = True
        .Tag = ""
    End With
        
'    With POtitles
'        .ColAlignmentFixed(0) = 4
'        .ColAlignmentFixed(2) = 4
'        .ColAlignmentFixed(4) = 4
'        .ROW = 0
'        Call drawLINEcol(POtitles, 1)
'        Call drawLINEcol(POtitles, 3)
'        .ROW = 1
'        Call fixPOtitles(0)
'    End With

End Sub

Function Iexists() As Boolean
Dim sql, transaction As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    Iexists = True
    transaction = Trim(cell(0))
    sql = "SELECT inv_invcnumb from transaction WHERE inv_npecode = '" + deIms.NameSpace + "' " _
        & "AND inv_ponumb = '" + cell(0) + "' AND inv_invcnumb = '" + cell(1) + "'"
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.Number <> 0 Then
        Iexists = False
        Exit Function
    End If
    If dataPO.RecordCount < 1 Then
        Iexists = False
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
        If .Rows > 0 And .Text <> "" Then
            .ZOrder
            .Visible = True
        End If
    End With
End Sub

Sub showTEXTline()
Dim positionX, positionY, i, currentCOL, currentROW As Integer
    With STOCKlist
        currentCOL = .Col
        currentROW = .ROW
        If .TextMatrix(.ROW, 0) <> "" Then
            If Trim(.TextMatrix(.ROW, 15)) = "P" Then
                If .TextMatrix(.ROW, 1) = "§" Then
                    If .Col = 10 Then Exit Sub
                End If
            Else
                If .TextMatrix(.ROW, 1) <> "§" Then
                    If .Col = 10 Then Exit Sub
                End If
            End If
                positionX = .Left + 20
                For i = 0 To .Col - 1
                    positionX = positionX + .ColWidth(i)
                Next
                positionY = .Top + 20
                For i = .TopRow - 1 To .ROW - IIf(.TopRow = 1, 1, 0)
                    positionY = positionY + .RowHeight(i)
                Next
                TextLINE.Text = .Text
                TextLINE.Left = positionX
                TextLINE.Width = .ColWidth(.Col) + 10
                TextLINE.Top = positionY
                TextLINE.Height = .RowHeight(.ROW) + 10
                TextLINE.Tag = .ROW
                TextLINE.SelStart = 0
                TextLINE.SelLength = Len(TextLINE.Text)
                TextLINE.Visible = True
                TextLINE.SetFocus
        End If
        .Col = currentCOL
        .ROW = currentROW
    End With
End Sub

Sub textBOX(ByVal mainCONTROL As MSHFlexGrid, standard As Boolean)
Dim h, i As Integer
Dim box As textBOX

    With mainCONTROL
        box.Height = .RowHeight(i)
        box.Height = box.Height + 10
        If .ROW = 0 And .FixedRows > 0 Then
            box.Top = .Top
            box.Height = box.Height - 80
        Else
            If standard Then
                box.Left = .Left + .ColWidth(0)
                h = 20
                For i = 0 To .ROW - 1
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
                    If Form = mdVisualization Then
                        If cell(Index) = "" Then
                            Call clearDOCUMENT
                            NavBar1.NewEnabled = False
                        Else
                            If Me.ActiveControl.Name = "cell" Then
                                If Me.ActiveControl.Index = 0 Then Call alphaSEARCH(cell(Index), POComboList, 0)
                            End If
                        End If
                    Else
                        If cell(0) = "" Then
                        End If
                    End If
                Case 1
                    If Form <> mdVisualization Then
                        If Index = 1 Then Exit Sub
                        If cell(Index) <> "" Then Call alphaSEARCH(cell(Index), TransactionComboList, 0)
                    End If
            End Select
        End With
    End If
End Sub

Private Sub cell_Click(Index As Integer)
    Select Case Index
        Case 0
            If Form = mdVisualization Then
                Call showLIST(POComboList)
            Else
                POComboList.Visible = False
            End If
        Case 1
            If Form = mdVisualization Then
                Call showLIST(TransactionComboList)
            Else
                TransactionComboList.Visible = False
            End If
    End Select
End Sub

Private Sub cell_GotFocus(Index As Integer)
    With cell(Index)
        If Not .locked Then
            .BackColor = vbYellow
            .Appearance = 1
            .Refresh
            .Tag = .Text
            Select Case Index
                Case 0
                    If Form = mdVisualization Then
                        If POComboList.Visible Then
                            POComboList.Visible = False
                        Else
                            Call showLIST(POComboList)
                        End If
                    End If
                Case 1
                    If Form = mdVisualization Then
                        If TransactionComboList.Visible Then
                            TransactionComboList.Visible = False
                        Else
                            Call showLIST(TransactionComboList)
                        End If
                    End If
                Case 3
                    If IsDate(cell(Index)) Then DTPicker1.Value = CDate(cell(Index))
                    If Form <> mdVisualization Then
                        Call showDTPicker1(Index)
                    End If
            End Select
        End If
    End With
End Sub

Private Sub cell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    With cell(Index)
        If Not .locked Then
            activeARROWS = False
            If Index <= 2 And Form = mdVisualization Then activeARROWS = True
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
                    If cell(Index) <> "" Then
                        Select Case Index
                            Case 0
                                If KeyAscii = 13 Then
                                    Select Case Form
                                        Case mdVisualization
                                            cell(0) = POComboList
                                            POComboList.Visible = False
                                            Call gettransaction("*")
                                            Call gettransactionComboList
                                        Case mdCreation
                                    End Select
                                End If
                                POComboList.Visible = False
                                cell(1).SetFocus
                            Case 1
                                If KeyAscii = 13 Then
                                    Select Case Form
                                        Case mdVisualization
                                            If cell(1) <> "" Then
                                                Call gettransaction("*")
                                            End If
                                            TransactionComboList.Visible = False
                                            cell(1).SetFocus
                                        Case mdCreation
                                            If Iexists Then
                                                'msg1 = translator.Trans("M00282")
                                                MsgBox IIf(msg1 = "", "Transaction Number is already exist", msg1)
                                                Exit Sub
                                            Else
                                                cell(3).SetFocus
                                            End If
                                    End Select
                                End If
                            Case 7
                        End Select
                    End If
                Case 27
                    .Text = cell(Index).Tag
                    Select Case Index
                        Case 0
                            POComboList.Visible = False
                        Case 1
                            TransactionComboList.Visible = False
                        Case 7

                    End Select
            End Select
        End If
    End With
End Sub

Private Sub cell_LostFocus(Index As Integer)
On Error Resume Next
    With cell(Index)
        If Not .locked Then
            .BackColor = remark.BackColor
            Select Case Index
                Case 0
                    Select Case Form
                        Case mdVisualization
                        Case mdCreation
                            POComboList.Visible = False
                            Exit Sub
                    End Select
                Case 1
                    Select Case Form
                        Case mdVisualization

                            If TransactionComboList.Visible Then
                                TransactionComboList.Visible = False
                            End If
                        Case mdCreation
                            If Iexists Then
                                'msg1 = translator.Trans("M00282")
                                MsgBox IIf(msg1 = "", "Transaction Number is already exist", msg1)
                                cell(1).SelStart = 0
                                cell(1).SelLength = Len(cell(1))
                                cell(1).SetFocus
                                Exit Sub
                            Else
                                cell(3).SetFocus
                            End If
                    End Select
                Case 2, 8, 9
                    .Text = .Tag
                    If Me.ActiveControl.Name <> "DTPicker1" Then
                        DTPicker1.Visible = False
                    End If
                Case 3
                    If Me.ActiveControl.Name <> "transactionComboList" Then
                        TransactionComboList.Visible = False
                    End If
                Case 7
                    If Me.ActiveControl.Name <> "destinationList" Then

                    End If
            End Select
        End If
    End With
End Sub



Public Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If Form <> mdVisualization Then
        With cell(Index)
            If Not .locked Then
                If .Text <> "" Then
                    If Form = mdCreation Then
                        Select Case Index
                            Case 0, 1
                            Case 2, 8, 9
                                If Not IsDate(.Text) Then
                                    .Text = ""
                                End If
                            Case 3
                                If .Text <> TransactionComboList Then
                                    .Text = ""
                                End If
                            Case 4
                            Case 5
                            Case 6
                            Case 7
                        End Select
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub Command1_Click()
Dim showALL As Boolean
Dim i As Integer
    If Command1.Caption = "&Show Only Selection" Then
        Command1.Caption = "&Show All Records"
        showALL = False
    Else
        Command1.Caption = "&Show Only Selection"
        showALL = True
    End If
    
    With STOCKlist
        .Col = 0
        If showALL Then
            .RowHeightMin = 50
            .RowHeight(-1) = 240
        Else
            For i = 1 To .Rows - 1
                If .RowHeight(i) > 240 Then
                    .TextMatrix(i, 13) = .RowHeight(i)
                End If
            Next
            .RowHeightMin = 0
            .RowHeight(-1) = 0
            For i = .Rows - 1 To 1 Step -1
                .ROW = i
                If .Text <> "" Then
                    .RowHeight(i) = 240
                End If
            Next
        End If
        .RowHeight(0) = 500
        For i = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(i, 13)) Then
                If Val(.TextMatrix(i, 13)) > 240 Then
                    .RowHeight(i) = Val(.TextMatrix(i, 13))
                End If
            End If

            If showALL Then
                If Val(.TextMatrix(i, 13)) = 50 Then .RowHeight(i) = 50
            Else
                If .TextMatrix(i, 0) <> "" And Not IsNumeric(.TextMatrix(i, 1)) Then
                    If .Rows > i + 1 Then .RowHeight(i + 1) = 50
                End If
            End If
        Next
    End With
End Sub



Private Sub Command2_Click()
    With STOCKlist
        If Command2.Caption = "&All Columns" Then
            Command2.Caption = "&Some Columns"
            .ColWidth(2) = 0
        Else
            Command2.Caption = "&All Columns"
            .ColWidth(2) = 3000
        End If
    End With
End Sub

Public Sub DTPicker1_DropDown()
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    With DTPicker1
        Select Case KeyCode
            Case 13
                cell(Val(.Tag)).Text = Format(.Value, "MMMM/dd/yyyy")
                cell(Val(.Tag) + 1).SetFocus
        End Select
    End With
End Sub

Private Sub DTPicker1_LostFocus()
Dim indexCELL As Integer
    With DTPicker1
        If IsNumeric(.Tag) Then
            cell(Val(.Tag)).Text = Format(.Value, "MMMM/dd/yyyy")
            indexCELL = Val(.Tag)
            If Me.ActiveControl.Name = "cell" Then
                If Me.ActiveControl.Index <> Val(.Tag) Then .Visible = False
                indexCELL = Me.ActiveControl.Index
            End If
            If Me.ActiveControl.Name = "cell" Then
                cell(indexCELL).SetFocus
            Else
                .Visible = False
            End If
        End If
        .Value = Now
    End With
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = 0
    If Form = mdVisualization Then
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        NavBar1.NewEnabled = False
        If Iexists Then
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
        End If
    End If
    Call makeDETAILgrid
'    frmStandardWarehouse.Left = Int((MDI_IMS.Width - frmStandardWarehouse.Width) / 2)
'    frmStandardWarehouse.Top = Int((MDI_IMS.Height - frmStandardWarehouse.Height) / 2) - 500
    cell(0).SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
    'Call translator.Translate_Forms("frmStandardWarehouse")
    NavBar1.Language = Language
    Call begining
    Form = mdVisualization
    Screen.MousePointer = 11
    Call lockDOCUMENT(True)
    Call getPOComboList
    frmStandardWarehouse.Caption = frmStandardWarehouse.Caption + " - " + frmStandardWarehouse.Tag
    Screen.MousePointer = 0
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.Number, True)
End Sub

Private Sub NavBar1_BeforeSaveClick()
Dim wrong, wrong2 As Boolean
Dim i, ii, position, Col As Integer
Screen.MousePointer = 11
    
    'Revision for Header
    wrong = False
    For i = 0 To 3
        If cell(i) = "" Then
            Screen.MousePointer = 0
            'msg1 = translator.Trans("M00016")
            MsgBox IIf(msg1 = "", "Cannot be left empty", msg1)
            cell(i).SetFocus
            Exit Sub
        End If
    Next
    If wrong Then
        Screen.MousePointer = 0
        'msg1 = translator.Trans("M00122")
        MsgBox IIf(msg1 = "", "Invalid Value", msg1)
        cell(position).SetFocus
        Exit Sub
    End If

    'Revision for Details
    wrong = True
    wrong2 = False
    position = 0
    For i = 1 To STOCKlist.Rows - 1
        If STOCKlist.TextMatrix(i, 0) <> "" Then
            For ii = 0 To 1
                Col = 8 + (ii * 2)
                If IsNumeric(STOCKlist.TextMatrix(i, Col)) Then
                    If CDbl(STOCKlist.TextMatrix(i, Col)) > 0 Then
                        wrong = False
                        readyFORsave = True
                    Else
                        readyFORsave = False
                        wrong = True
                        position = i
                        Exit For
                    End If
                Else
                    wrong = True
                    position = i
                    Exit For
                End If
            Next
            If wrong2 Then
                wrong = True
                Exit For
            End If
        End If
    Next
    If wrong Then

        If position > 0 Then
            Screen.MousePointer = 0
            'msg1 = translator.Trans("M00122")
            MsgBox IIf(msg1 = "", "Invalid Value", msg1)
            STOCKlist.ROW = position
            STOCKlist.Col = Col
            STOCKlist.SetFocus
        Else
            Screen.MousePointer = 0
            'msg1 = translator.Trans("M00707")
            MsgBox IIf(msg1 = "", "You have to select at least one line item.", msg1)
        End If
    Else
        Call SAVE
        'Call ChangeMode(mdVisualization)
        Call getPOComboList
        Call gettransactionComboList
        Call gettransaction(cell(0))
        cell(0).locked = False
        cell(0).SelLength = Len(cell(0))
        cell(0).SelStart = 0
        'msg1 = translator.Trans("M00306")
        MsgBox IIf(msg1 = "", "Insert into Supplier transaction List is completed successfully", msg1)
        NavBar1.CancelEnabled = False
        POComboList.Visible = True
        cell(0).SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As String
     'msg1 = translator.Trans("M00706")
    'msg2 = translator.Trans("L00441")
    response = MsgBox(IIf(msg1 = "", "Are you sure you want to cancel changes?", msg1), vbYesNo, IIf(msg2 = "", "Cancel", msg2))
    If response = vbYes Then
        With NavBar1
            cell(0).locked = False
            'Call ChangeMode(mdVisualization)
            Call lockDOCUMENT(True)
            Call clearDOCUMENT
            If cell(0) <> "" Then
                .NewEnabled = True
                Call gettransaction("*")
            End If
            .CancelEnabled = False
            .SaveEnabled = False
            .PrintEnabled = False
        End With
    End If
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnEMailClick()
Dim Params(1) As String
Dim rptinfo As RPTIFileInfo
Screen.MousePointer = 11
On Error Resume Next
    Call BeforePrint
    
    With rptinfo
        Params(0) = "namespace=" + deIms.NameSpace
        Params(1) = "manifestnumb=" + cell(0)
        .ReportFileName = ReportPath & "transaction.rpt"
        'Call translator.Translate_Reports("transaction.rpt")
        .Parameters = Params
    End With
    
    Params(0) = ""
    Call WriteRPTIFile(rptinfo, Params(0))
    Call SendEmailAndFax(rsReceptList, "Recipient", "Transaction " & cell(0), "", Params(0))
    Screen.MousePointer = 0
If Err Then Call LogErr(Name & "::NavBar1_OnEMailClick", Err.Description, Err.Number, True)
End Sub

Private Sub NavBar1_OnNewClick()
Dim i As Integer
Dim sql, response As String
Dim dataUSER As ADODB.Recordset

    Screen.MousePointer = 11
    With NavBar1
        If cell(0) = "" Then
            Screen.MousePointer = 0
            MsgBox "Invalid Transaction Number"
        Else
            If isOPEN(cell(0)) Then
                POComboList.Visible = False
                TransactionComboList.Visible = False
                For i = 1 To 3
                    cell(i) = ""
                Next
                cell(4) = Format(Now, "MMMM/dd/yyyy")
                'Call ChangeMode(mdCreation)
                Call begining
                Set dataUSER = New ADODB.Recordset
                sql = "SELECT usr_username FROM XUSERPROFILE WHERE usr_npecode = '" + deIms.NameSpace + "' AND usr_userid = '" + CurrentUser + "'"
                dataUSER.Open sql, deIms.cnIms, adOpenForwardOnly
                If dataUSER.RecordCount > 0 Then
                    cell(2) = dataUSER!usr_username
                End If
                Screen.MousePointer = 0
                .NewEnabled = False
                .CancelEnabled = True
                .SaveEnabled = True
                .PrintEnabled = False
                
                Screen.MousePointer = 11
                Call getLINEitems("*")
                Call lockDOCUMENT(False)
            Else
                Screen.MousePointer = 0
                MsgBox "This PO is already closed"
                cell(0).SetFocus
                Exit Sub
            End If
        End If
    End With
    Screen.MousePointer = 0
    cell(1).SetFocus
End Sub

'Private Function ChangeMode(FMode As FormMode) As Boolean
'On Error Resume Next
'    Select Case FMode
'        Case mdCreation
'            lblStatu.ForeColor = vbRed
'            'msg1 = translator.Trans("L00125")
'            lblStatu.Caption = IIf(msg1 = "", "Creation", msg1)
'            lblStatu.Tag = "Creation"
'            ChangeMode = True
'        Case mdVisualization
'            lblStatu.ForeColor = vbGreen
'            'msg1 = translator.Trans("L00092") 'J added
'            lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
'            lblStatu.Tag = "Visualization"
'            ChangeMode = True
'    End Select
'    Form = FMode
'End Function

Private Sub NavBar1_OnPrintClick()
On Error Resume Next
Screen.MousePointer = 11
    With MDI_IMS.CrystalReport1
        Call BeforePrint
        'msg1 = translator.Trans("L00213")
        .WindowTitle = IIf(msg1 = "", "transaction", msg1)
        .Action = 1
    End With
Screen.MousePointer = 0
End Sub

Sub SAVE()
Dim header As New ADODB.Recordset
Dim details As New ADODB.Recordset
Dim remarks As New ADODB.Recordset

Dim INVitem As New ADODB.Recordset

Dim i, ROW As Integer
Dim sql As String
Dim Q, Quantity, PRICE As Double
On Error Resume Next
    
    If readyFORsave Then
        'Header routine
        'msg1 = translator.Trans("M00708")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Header", msg1)
        deIms.cnIms.BeginTrans
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM transaction WHERE inv_ponumb = ''"
        header.Open sql, deIms.cnIms, adOpenDynamic, adLockPessimistic
        With header
            .AddNew
            !inv_creauser = CurrentUser
            !inv_npecode = deIms.NameSpace
            
            !inv_ponumb = cell(0)
            !inv_invcnumb = cell(1)
            !inv_invcdate = CDate(cell(3))
            !inv_creadate = CDate(cell(4))
            .Update
        End With
        
        'Remarks routine
        'msg1 = translator.Trans("M00719")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Remarks", msg1)
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM transactionREM WHERE invr_ponumb = ''"
        remarks.Open sql, deIms.cnIms, adOpenDynamic, adLockPessimistic
        With remarks
            .AddNew
            !invr_creauser = CurrentUser
            !invr_npecode = deIms.NameSpace
            !invr_creadate = CDate(cell(4))
            
            !invr_ponumb = cell(0)
            !invr_invcnumb = cell(1)
            !invr_rem = remark
            !invr_linenumb = 1
            .Update
        End With
                
        'Details routine
        'msg1 = translator.Trans("M00710")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Details", msg1)
        Set details = New ADODB.Recordset
        sql = "SELECT * FROM transactionDETL WHERE invd_ponumb = ''"
        details.Open sql, deIms.cnIms, adOpenKeyset, adLockPessimistic
        With details
            For i = 1 To STOCKlist.Rows - 1
                If STOCKlist.TextMatrix(i, 0) <> "" Then
                    If IsNumeric(STOCKlist.TextMatrix(i, 1)) Then
                        .AddNew
                        !invd_npecode = deIms.NameSpace
                        !invd_creauser = CurrentUser
                        !invd_creadate = CDate(cell(4))
                        
                        !invd_ponumb = cell(0)
                        !invd_invcnumb = cell(1)
                        !invd_liitnumb = STOCKlist.TextMatrix(i, 1)
                        
                        Quantity = IIf(IsNumeric(STOCKlist.TextMatrix(i, 8)), CDbl(STOCKlist.TextMatrix(i, 8)), 0)
                        !invd_primreqdqty = Quantity
                        !invd_primuom = STOCKlist.TextMatrix(i, 16)
                        PRICE = IIf(IsNumeric(STOCKlist.TextMatrix(i, 10)), CDbl(STOCKlist.TextMatrix(i, 10)), 0)
                        !invd_unitpric = PRICE
                        !invd_totapric = Quantity * PRICE
                                                
                        If Trim(STOCKlist.TextMatrix(i, 15)) = "" Then
                            ROW = i
                        Else
                            ROW = i + 1
                        End If
                        Quantity = IIf(IsNumeric(STOCKlist.TextMatrix(ROW, 8)), CDbl(STOCKlist.TextMatrix(ROW, 8)), 0)
                        !invd_secoreqdqty = Quantity
                        !invd_secouom = STOCKlist.TextMatrix(ROW, 16)
                        PRICE = IIf(IsNumeric(STOCKlist.TextMatrix(ROW, 10)), CDbl(STOCKlist.TextMatrix(ROW, 10)), 0)
                        !invd_secounitprice = PRICE
                        !invd_secototaprice = Quantity * PRICE
                    End If
                End If
            Next
            'msg1 = translator.Trans("M00714")
            'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Saving Transaction", msg1)
            .UpdateBatch
        End With
        'msg1 = translator.Trans("M00715")
        'MDI_IMS.StatusBar1.Panels(1).Text = IIf(msg1 = "", "Commiting Transaction", msg1)
        deIms.cnIms.CommitTrans
        'MDI_IMS.StatusBar1.Panels(1).Text = ""
        Screen.MousePointer = 0
        Screen.MousePointer = 11
        Call lockDOCUMENT(True)
        Call clearDOCUMENT
        Call getPOComboList
    End If
End Sub

Private Sub POComboList_Click()
    Select Case Form
        Case mdVisualization
            POComboList.Tag = POComboList.ROW
            cell(0) = Trim(POComboList)
            If Left(cell(0), 1) = "(" And Right(cell(0), 1) = ")" Then
                Call clearDOCUMENT
                POComboList.Visible = True
                cell(0).SetFocus
            Else
                Call gettransaction("*")
            End If
            Call gettransactionComboList
            cell(0).SetFocus
        Case mdCreation
            cell(1).SetFocus
    End Select
End Sub

Private Sub POComboList_KeyPress(KeyAscii As Integer)
    With POComboList
        Select Case KeyAscii
            Case 13
                Select Case Form
                    Case mdVisualization
                        cell(0) = .Text
                        Call gettransaction("*")
                        Call gettransactionComboList
                    Case mdCreation
                        cell(1).SetFocus
                End Select
            Case 27
                POComboList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(0) = Chr(KeyAscii)
                Call alphaSEARCH(cell(0), POComboList, 0)
                .Tag = ""
                cell(0).SetFocus
                cell(0).SelStart = Len(cell(0))
                cell(0).SelLength = 0
        End Select
    End With
End Sub

Private Sub stocklist_Click()
Dim i, currentCOL As Integer
    If Form <> mdVisualization Then
        With STOCKlist
            If .TextMatrix(.ROW, 1) <> "" Then
                If .ROW > 0 Then
                    Select Case .MouseCol
                        Case 0, 1
                            Call markROW
                        Case 8, 10
                            Call showTEXTline
                    End Select
                End If
            End If
        End With
    End If
End Sub

Private Sub POComboList_EnterCell()
    With POComboList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
    End With
End Sub

Private Sub POComboList_GotFocus()
    Call gridONfocus(POComboList)
End Sub

Private Sub POComboList_LeaveCell()
    With POComboList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub


Private Sub POComboList_LostFocus()
    With POComboList
        cell(0).Text = Trim(.Text)
    End With
End Sub

Public Sub POComboList_Validate(Cancel As Boolean)
    cell(0) = Trim(POComboList)
End Sub

Private Sub stocklist_EnterCell()
Dim changeCOLORS As Boolean
    If Form <> mdVisualization Then
        Dim i, currentCOL, currentROW As Integer
        With STOCKlist
            currentCOL = .Col
            If IsNumeric(.Tag) Then
                If Val(.Tag) = .ROW Then
                    changeCOLORS = False
                Else
                    currentROW = .ROW
                    .ROW = Val(.Tag)
                    If colorsROW(1) <> "" Then
                        For i = 1 To 12
                            .Col = i
                            .CellBackColor = colorsROW(i)
                        Next
                        .Col = currentCOL
                    End If
                    .ROW = currentROW
                    .Tag = currentROW
                    Call getCOLORSrow
                    changeCOLORS = True
                End If
            Else
                STOCKlist.Tag = .ROW
                Call getCOLORSrow
                changeCOLORS = True
            End If
            
            If .TextMatrix(.ROW, 1) <> "" Then
                currentCOL = .Col
                If changeCOLORS Then
                    For i = 1 To 12
                        .Col = i
                        Select Case .CellBackColor
                            Case &HC0FFFF 'Very Light Yellow
                                .CellBackColor = &HFFC0C0 'Very Light Blue
                            Case &HC0C0FF 'Very Light Red
                                .CellBackColor = &HFFC0FF 'Very Light Magenta
                            Case &HE0E0E0 'Very Light Gray
                            Case Else
                                .CellBackColor = &HFFC0C0 'Very Light Blue
                        End Select
                    Next
                    Select Case .Col
                        Case 8, 10
                            Call showTEXTline
                    End Select
                End If
            End If
            .Col = currentCOL
        End With
    End If
End Sub

Private Sub stocklist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With STOCKlist
        If .TextMatrix(.MouseRow, 1) = "" Then
            If IsNumeric(.TextMatrix(.MouseRow, 13)) Then
                .RowHeight(.MouseRow) = Val(.TextMatrix(.MouseRow, 13))
            End If
        End If
    End With
End Sub

Private Sub stocklist_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim ROW, Col As Integer
    With STOCKlist
        ROW = .MouseRow
        Col = .MouseCol
        If Col = 0 Then
            If .TextMatrix(ROW, 1) = "" Then
                If IsNumeric(.TextMatrix(ROW, 13)) Then
                    .RowHeight(ROW) = Val(.TextMatrix(ROW, 13))
                Else
                    .RowHeight(ROW) = 240
                End If
            End If
        End If
    End With
End Sub

Private Sub stocklist_Scroll()
    If Form <> mdVisualization Then TextLINE.Visible = False
End Sub

Private Sub stocklist_SelChange()
    With STOCKlist
        If Form <> mdVisualization Then
            If .TextMatrix(.ROW, 1) <> "" Then
                If .RowHeight(STOCKlist.ROW) > 240 Then
                    .TextMatrix(STOCKlist.ROW, 13) = .RowHeight(STOCKlist.ROW)
                End If
            End If
        End If
    End With
End Sub

Private Sub POtitles_Click()

End Sub

Private Sub transactionComboList_Click()
    Select Case Form
        Case mdVisualization
            TransactionComboList.Tag = TransactionComboList.ROW
            cell(1) = Trim(TransactionComboList)
            If Left(cell(0), 1) = "(" And Right(cell(0), 1) = ")" Then
                Call gettransaction(cell(1))
                cell(1).SetFocus
            Else
                Call gettransaction(cell(1))
            End If
        Case mdCreation
            cell(2).SetFocus
    End Select
End Sub

Private Sub transactionComboList_EnterCell()
    With TransactionComboList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(1) = .Text
    End With
End Sub


Private Sub transactionComboList_GotFocus()
    Call gridONfocus(TransactionComboList)
End Sub

Private Sub transactionComboList_KeyPress(KeyAscii As Integer)
    With TransactionComboList
        Select Case KeyAscii
            Case 13
                cell(2).SetFocus
            Case 27
                TransactionComboList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(1) = Chr(KeyAscii)
                Call alphaSEARCH(cell(1), TransactionComboList, 0)
                .Tag = ""
                cell(1).SetFocus
                cell(1).SelStart = Len(cell(1))
                cell(1).SelLength = 0
        End Select
    End With
End Sub

Private Sub transactionComboList_LeaveCell()
    With TransactionComboList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub transactionComboList_LostFocus()
    With TransactionComboList
        cell(1).Text = Trim(.Text)
        cell(1).SetFocus
        cell(1).SelStart = Len(cell(1))
        cell(1).SelLength = 0
    End With
End Sub

Private Sub transactionComboList_Validate(Cancel As Boolean)
    cell(1) = TransactionComboList
    TransactionComboList.Visible = False
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
            Call TextLINE_Validate(True)
        End If
    End With
End Sub

Public Sub TextLINE_Validate(Cancel As Boolean)
Dim i, Col, ROW As Integer
Dim qty, switch As String
Dim newPRICE, QTY1, QTY2, uPRICE1, uPRICE2 As Double
Dim newPRICEok As Boolean
    With TextLINE
        If STOCKlist.Col = 8 Or STOCKlist.Col = 10 Then
            Col = STOCKlist.Col
            If IsNumeric(.Text) Then
                If Val(.Text) > 0 Then
                     STOCKlist.TextMatrix(STOCKlist.ROW, Col) = FormatNumber(.Text, 2)
                    switch = Trim(STOCKlist.TextMatrix(STOCKlist.ROW, 15))
                    Select Case switch
                        Case ""
                            Call differences(STOCKlist.ROW)
                        Case "P", "S"
                            If STOCKlist.TextMatrix(STOCKlist.ROW, 1) = "§" Then
                                ROW = STOCKlist.ROW - 1
                            Else
                                ROW = STOCKlist.ROW
                            End If
                            newPRICEok = True
                            If IsNumeric(STOCKlist.TextMatrix(ROW, 8)) Then
                                QTY1 = CDbl(STOCKlist.TextMatrix(ROW, 8))
                            Else
                                QTY1 = 0
                                newPRICEok = False
                            End If
                            If IsNumeric(STOCKlist.TextMatrix(ROW + 1, 8)) Then
                                QTY2 = CDbl(STOCKlist.TextMatrix(ROW + 1, 8))
                            Else
                                QTY2 = 0
                                newPRICEok = False
                            End If
                            If switch = "P" Then
                                If IsNumeric(STOCKlist.TextMatrix(ROW, 10)) Then
                                    uPRICE1 = CDbl(STOCKlist.TextMatrix(ROW, 10))
                                Else
                                    uPRICE1 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE2 = (QTY1 * uPRICE1) / QTY2
                                    STOCKlist.TextMatrix(ROW + 1, 10) = FormatNumber(uPRICE2, 2)
                                End If
                            Else
                                If IsNumeric(STOCKlist.TextMatrix(ROW + 1, 10)) Then
                                    uPRICE2 = CDbl(STOCKlist.TextMatrix(ROW + 1, 10))
                                Else
                                    uPRICE2 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE1 = (QTY2 * uPRICE2) / QTY1
                                    STOCKlist.TextMatrix(ROW, 10) = FormatNumber(uPRICE1, 2)
                                End If
                            End If
                            Call differences(ROW)
                            Call differences(ROW + 1)
                    End Select
                    
                    .Tag = ""
                    .Text = ""
                    .Visible = False
                    Exit Sub
                End If
            End If
            If .Text <> "" Then
                'msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                TextLINE = ""
            End If
        End If
    End With
End Sub


