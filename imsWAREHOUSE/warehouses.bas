Attribute VB_Name = "warehouses"
Public Enum FormMode
    mdNa = 0
    mdCreation
    mdModified
    mdModification
    mdVisualization
End Enum

Public Type DefaultFQA
    Company As String
    Location As String
    CamChart As String
    UsChart As String
    StockType As String
End Type

Global readyFORsave As Boolean
Global rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Global colorsROW(22)
Global activeCELL As Integer
Global nodeSEL
Global nodeONtop As Integer
Global tabindex
Global currentBOX As Integer
Global currentNODE As Integer
Global totalNODE As Integer
Global Transnumb
Global directCLICK As Boolean
Global noRETURN As Boolean
Global justCLICK As Boolean
Global SecUnit As Double
Global summaryPOSITION As Integer
Global cn As ADODB.Connection
Global nameSP As String
Global CurrentUser As String
Global msg1, msg2
Global repoPATH As String
Global dsnF, dsnDSQ, dsnUID, dsnPWD
Global lang As String
Global rowguid As String
Global cTT As New cTreeTips
Global GFQAComboFilled As Boolean
Global GDefaultFQA As DefaultFQA
Global GDefaultValue As Boolean
Global GExtendedCurrency As String 'M
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Global direction, activeBOX

Sub hideCOMBOS()
Dim i
    For i = 0 To 4
        frmWarehouse.combo(i).Visible = False
    Next
End Sub

Sub unlockCELLS()
Dim i
    For i = 1 To 4
        frmWarehouse.cell(i).Enabled = True
    Next
End Sub
Sub lockCELLS()
Dim i
    For i = 1 To 4
        frmWarehouse.cell(i).Enabled = False
        frmWarehouse.cell(i).backcolor = vbWhite
    Next
End Sub
Sub arrowKEYS(Index As Integer, box As textBOX, grid As MSHFlexGrid)
    With box
        grid.Visible = True
        Select Case direction
            Case "down"
                If grid.row <= (grid.Rows - 1) Then
                    If grid.row = 0 And .text = "" Then
                        grid.row = grid.row + 1
                    End If
                    .text = grid.text
                Else
                    grid.row = grid.Rows - 1
                End If
            Case "up"
                If grid.row > 0 Then
                Else
                    grid.row = 1
                End If
        End Select
        
        grid.tag = grid.row
        If Not grid.Visible Then
            grid.Visible = True
        End If
        grid.ZOrder
        usingARROWS = True
        grid.col = 0
        grid.ColSel = grid.cols - 1
        'If frmWarehouse.ActiveControl.name = "combo" Then
            grid.SetFocus
        'End If
    End With
End Sub


Sub fillTRANSACTION(datax As ADODB.Recordset)
Dim i, n, rec, condition, key, conditionCODE, fromlogic
Dim fromSUBLOCA, unitCODE, unit, StockNumber, unitprice
Dim shot
    Call cleanDETAILS
    Call hideDETAILS
    frmWarehouse.STOCKlist.Visible = False
    
    frmWarehouse.searchFIELD(0).Visible = False
    frmWarehouse.searchFIELD(1).Visible = False
    
    frmWarehouse.Tree.Height = 2000
    frmWarehouse.SUMMARYlist.Top = frmWarehouse.searchFIELD(0).Top
    frmWarehouse.SUMMARYlist.Height = (frmWarehouse.remarks.Top + frmWarehouse.remarks.Height) - frmWarehouse.SUMMARYlist.Top
    
    'SUMMARYlist.Height = 1980 + 2340 + 1740 'M
    
    frmWarehouse.SUMMARYlist.ZOrder
    frmWarehouse.summaryLABEL.Top = frmWarehouse.SUMMARYlist.Top - 240
    'summaryLABEL.Visible = True M
    frmWarehouse.summaryLABEL.Visible = False
    
    frmWarehouse.remarks.Top = frmWarehouse.Tree.Top + 2000 + 400
    frmWarehouse.remarks.Height = frmWarehouse.Height - frmWarehouse.remarks.Top - 990
    frmWarehouse.remarks.Visible = True
    'remarks.ZOrder M
    'remarksLABEL.Top = remarks.Top - 240 M
    'remarksLABEL.Visible = True M
    
    
    frmWarehouse.remarks.locked = True
    frmWarehouse.Refresh
    
    frmWarehouse.dateBOX = Format(datax!Date, "Short Date")
    frmWarehouse.userNAMEbox = getUSERname(datax!userCODE)
    frmWarehouse.remarks = IIf(IsNull(datax!remarks), "", datax!remarks)
    With frmWarehouse.SUMMARYlist
        .Rows = 2
        i = 0
        
        frmWarehouse.cell(1).tag = datax!Company
        directCLICK = True
        frmWarehouse.cell(1) = getCOMPANYdescription(frmWarehouse.cell(1).tag)
        Select Case frmWarehouse.tag
            Case "02040400", "02040100", "02040300" 'ReturnFromRepair, 'WarehouseReceipt, 'Return from Well
                frmWarehouse.cell(2).tag = datax!FromPlace
                directCLICK = True
                frmWarehouse.cell(2) = getLOCATIONdescription(frmWarehouse.cell(2).tag)
                frmWarehouse.cell(3).tag = datax!Warehouse
                directCLICK = True
                frmWarehouse.cell(3) = getLOCATIONdescription(frmWarehouse.cell(3).tag)
                If frmWarehouse.tag = "02040100" Then
                    frmWarehouse.cell(4).tag = datax!PO
                    directCLICK = True
                    frmWarehouse.cell(4) = frmWarehouse.cell(4).tag
                End If
            Case "02050200" 'AdjustmentEntry
                frmWarehouse.cell(2).tag = datax!FromPlace
                directCLICK = True
                frmWarehouse.cell(2) = getLOCATIONdescription(frmWarehouse.cell(2).tag)
            Case "02040200", "02040500", "02040600", "02050400"  'WarehouseIssue, 'WellToWell, 'WarehouseToWarehouse, 'Sales
                frmWarehouse.cell(2).tag = datax!Warehouse
                directCLICK = True
                frmWarehouse.cell(2) = getLOCATIONdescription(frmWarehouse.cell(2).tag)
                frmWarehouse.cell(3).tag = datax!IssueToPlace
                directCLICK = True
                frmWarehouse.cell(3) = getLOCATIONdescription(frmWarehouse.cell(3).tag)
            Case "02040700", "02050300" 'InternalTransfer, 'AdjustmentIssue
                frmWarehouse.cell(2).tag = datax!Warehouse
                directCLICK = True
                frmWarehouse.cell(2) = getLOCATIONdescription(frmWarehouse.cell(2).tag)
        End Select
        Do While Not datax.EOF
            shot = ImsDataX.GetConditions(nameSP, IIf(IsNull(datax!OriginalCondition), "", datax!OriginalCondition), True, cn)
            condition = shot(0)
            conditionCODE = shot(1)
            StockNumber = datax!StockNumber
            rec = Format(datax!TransactionLine) + vbTab
            rec = rec + StockNumber + vbTab
            If datax!serialnumber <> "" Then
                If frmWarehouse.newBUTTON.Enabled Then
                    rec = rec + Trim(datax!serialnumber) + vbTab
                Else
                    rec = rec + Trim(datax!serial) + vbTab
                End If
            Else
                rec = rec + "Pool" + vbTab
            End If
            rec = rec + condition + vbTab
            rec = rec + Format(datax!unitprice, "0.00") + vbTab
            rec = rec + IIf(IsNull(datax!StockDescription), "", datax!StockDescription) + vbTab
            unitCODE = getUNIT(StockNumber)
            unit = getUNITdescription(unitCODE)
            rec = rec + unit + vbTab
            rec = rec + Format(datax!QTY1) + vbTab
            rec = rec + Format(i) + vbTab
            rec = rec + Trim(IIf(IsNull(datax!fromlogic), "", datax!fromlogic)) + vbTab
            rec = rec + Trim(IIf(IsNull(datax!fromSUBLOCA), "", datax!fromSUBLOCA)) + vbTab
            rec = rec + IIf(IsNull(datax!toLOGIC), "", Trim(datax!toLOGIC)) + vbTab
            rec = rec + IIf(IsNull(datax!toSUBLOCA), "", Trim(datax!toSUBLOCA)) + vbTab
            rec = rec + IIf(IsNull(datax!OriginalCondition), "", datax!OriginalCondition) + vbTab
            rec = rec + unit
            .addITEM rec
            .TextMatrix(.Rows - 1, 20) = conditionCODE
            datax.MoveNext
            i = i + 1
        Loop
        If .Rows > 2 Then .RemoveItem 1
        Call reNUMBER(frmWarehouse.SUMMARYlist)
    End With
    directCLICK = False
End Sub
Sub reNUMBER(grid As MSHFlexGrid)
Dim i
    With grid
        For i = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(i, 0)) Or .TextMatrix(i, 0) = "" Then
                .TextMatrix(i, 0) = Format(i)
            End If
        Next
    End With
End Sub

Sub hideDETAILS()
Dim i
    With frmWarehouse
        .SUMMARYlist.Visible = True
        .SUMMARYlist.ZOrder
        .hideDETAIL.Visible = False
        .submitDETAIL.Visible = False
        .removeDETAIL.Visible = False
        .Label4(0).Visible = False
        .Label4(1).Visible = False
    End With
End Sub

Public Function RollbackTransaction(cn As ADODB.Connection)
On Error Resume Next
    With MakeCommand(cn, adCmdText)
        .CommandText = "ROLLBACK TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    If Err Then Err.Clear
End Function

Sub gridCOLORdark(grid As MSHFlexGrid, row)
    With grid
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
Sub gridCOLORnormal(grid As MSHFlexGrid, row)
    With grid
        .row = row
        .CellBackColor = &HFFFFC0      'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Sub setupBOXES(n, datax As ADODB.Fields, serial As Boolean, Optional QTYpo)
Dim x, cond, logic, subloca, newCOND
On Error Resume Next

    With frmWarehouse
        Load .quantity(n)
        If Not .newBUTTON.Enabled Then Call putBOX(.quantity(n), .detailHEADER.ColWidth(0) + 140, topNODE(n), .detailHEADER.ColWidth(1) - 40, vbWhite)
        Load .balanceBOX(n)
        .balanceBOX(n) = .quantity(n)
        Load .quantityBOX(n)
        .quantityBOX(n).tabindex = tabindex + 2
        Load .priceBOX(n)
        Load .NEWconditionBOX(n)
        Select Case .tag
            'ReturnFromRepair WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales,ReturnFromWell
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                If serial Then
                    .quantity(n) = 1
                Else
                    If .newBUTTON.Enabled Then
                        .quantity(n) = Format(datax!QTY1, "0.00")
                        cond = Trim(datax!OriginalCondition)
                        logic = Trim(datax!fromlogic)
                        subloca = Trim(datax!fromSUBLOCA)
                        newCOND = IIf(IsNull(datax!NEWcondition), "", datax!NEWcondition)
                    Else
                        .quantity(n) = Format(datax!qty, "0.00")
                        cond = Trim(datax!condition)
                        logic = Trim(datax!logic)
                        subloca = Trim(datax!subloca)
                        newCOND = datax!condition
                    End If
                End If
                .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), cond, logic, subloca, IIf(IsNull(datax!serialnumber), "POOL", Trim(datax!serialnumber)), n), "0.00")
                .priceBOX(n) = Format(datax!unitprice, "0.00")
                .NEWconditionBOX(n).tag = newCOND
            Case "02050200" 'AdjustmentEntry
                .quantity(n) = Format(summaryQTY(Trim(datax!StockNumber), Left(.Tree.Nodes(n).key, 2), "", "", "POOL", n), "0.00")
                .quantityBOX(n) = .quantity(n)
                If summaryPOSITION > 0 Then
                    .priceBOX(n) = Format(.SUMMARYlist.TextMatrix(summaryPOSITION, 4), "0.00")
                Else
                    .priceBOX(n) = "0.00"
                End If
                .NEWconditionBOX(n).tag = Left(.Tree.Nodes(n).key, 2)
            Case "02040100" 'WarehouseReceipt
                .quantity(n) = Format(QTYpo, "0.00")
                If .newBUTTON.Enabled Then
                    newCOND = datax!NEWcondition
                    .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "GENERAL", "GENERAL", "POOL", n), "0.00")
                Else
                    newCOND = datax!condition
                    .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "unique", "unique", "POOL", n), "0.00")
                End If
                .priceBOX(n) = Format(datax!unitprice, "0.00")
                .NEWconditionBOX(n).tag = newCOND
                Load .repairBOX(n)
                .repairBOX(n) = Format(datax!POitem)
        End Select
        .NEWconditionBOX(n) = .NEWconditionBOX(n).tag
        
        Load .logicBOX(n)
        .logicBOX(n).tabindex = tabindex
        Load .sublocaBOX(n)
        .sublocaBOX(n).tabindex = tabindex + 1
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                .logicBOX(n) = datax!toLOGIC
                .sublocaBOX(n) = datax!toSUBLOCA
            Else
                .logicBOX(n) = "GENERAL"
                .sublocaBOX(n) = "GENERAL"
            End If
        Else
            .logicBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 11)
            .logicBOX(n).tag = .logicBOX(n)
            .sublocaBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 12)
            .sublocaBOX(n).tag = .sublocaBOX(n)
            .logicBOX(n).ToolTipText = getWAREHOUSEdescription(.logicBOX(n))
            .sublocaBOX(n).ToolTipText = getSUBLOCATIONdescription(.sublocaBOX(n))
        End If
        
        Load .unitBOX(n)
        
        If .newBUTTON.Enabled Then
            .unitBOX(n) = "" '*************************************************************************
        Else
            .unitBOX(n) = datax!unit
        End If
        
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                newCOND = datax!NEWcondition
            Else
                newCOND = datax!condition
                .NEWconditionBOX(n).ToolTipText = datax!ConditionName
            End If
        
            .NEWconditionBOX(n).tag = newCOND
            .NEWconditionBOX(n) = Format(newCOND, "00")
        Else
            .NEWconditionBOX(n).tag = .SUMMARYlist.TextMatrix(summaryPOSITION, 13)
            .NEWconditionBOX(n) = Format(.NEWconditionBOX(n).tag, "00")
            .NEWconditionBOX(n).ToolTipText = .SUMMARYlist.TextMatrix(summaryPOSITION, 14)
        End If
        
        Select Case .tag
            Case "02040200", "02040500" 'WarehouseIssue, WellToWell
                'If Not .newBUTTON.Enabled Then
                    '.logicBOX(n).Enabled = False
                    '.sublocaBOX(n).Enabled = False
                'End If
            Case "02040400" 'ReturnFromRepair
                Load .repairBOX(n)
                If summaryPOSITION = 0 Then
                    If .newBUTTON.Enabled Then
                        .repairBOX(n) = Format(datax!repairCOST, "0.00")
                        .cell(5) = Trim(datax!NewStockNumber)
                        .cell(5).tag = .cell(5)
                        .unitLABEL(1) = getUNIT(.cell(5).tag)
                        .newDESCRIPTION = Trim(datax!NewStockDescription)
                    Else
                        .repairBOX(n) = "0.00"
                    End If
                Else
                    If .newBUTTON.Enabled Then
                        .repairBOX(n) = Format(datax!repairCOST, "0.00")
                        .cell(5) = Trim(datax!NewStockNumber)
                        .cell(5).tag = .cell(5)
                        .unitLABEL(1) = getUNIT(.cell(5).tag)
                        .newDESCRIPTION = Trim(datax!NewStockDescription)
                    Else
                        .repairBOX(n) = Format(.SUMMARYlist.TextMatrix(summaryPOSITION, 17), "0.00")
                        .cell(5) = .SUMMARYlist.TextMatrix(summaryPOSITION, 18)
                        .cell(5).tag = .cell(5)
                        .unitLABEL(1) = getUNIT(.cell(5))
                        .newDESCRIPTION = .SUMMARYlist.TextMatrix(summaryPOSITION, 19)
                    End If
                End If
            Case "02040100" 'WarehouseReceipt
                If Not .newBUTTON.Enabled Then
                    .NEWconditionBOX(n).Enabled = True
                End If
            Case Else
                If Not .newBUTTON.Enabled Then
                    .NEWconditionBOX(n).Enabled = True
                    .logicBOX(n).Enabled = True
                    .sublocaBOX(n).Enabled = True
                    .repairBOX(n).Enabled = True
                End If
        End Select
        If .newBUTTON.Enabled Then
            .quantityBOX(n).Enabled = False
            .priceBOX(n).Enabled = False
            .NEWconditionBOX(n).Enabled = False
            .logicBOX(n).Enabled = False
            .sublocaBOX(n).Enabled = False
            .repairBOX(n).Enabled = False
        Else
            .quantityBOX(n).Enabled = True
            .priceBOX(n).Enabled = True
        End If
    End With
End Sub

Sub fillDETAILlist(StockNumber, description, unit, Optional QTYpo)
Dim i, n, sql, rec, cond, loca, subloca, stock, total, key, lastLINE, thick, condNAME, currentLOGIC, currentSUBloca
Dim sublocaNAME, logicNAME, currentCOND, firstSET
Dim pool As Boolean
Dim moreSERIAL As Boolean
Dim datax As ADODB.Recordset
Dim datay As ADODB.Recordset
Dim dataz As ADODB.Recordset
Dim docTYPE As ADODB.Recordset

On Error Resume Next
    With frmWarehouse
        firstSET = 0
        Screen.MousePointer = 11
        .STOCKlist.MousePointer = Screen.MousePointer
        tabindex = 1
        .commodityLABEL = StockNumber
        .unitLABEL(0) = unit
        .unitLABEL(1) = ""
        .descriptionLABEL = description
        If StockNumber + description + unit = "" Then
            Call cleanDETAILS
            Screen.MousePointer = 0
            frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
            Exit Sub
        End If
        directCLICK = True
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
                .cell(5).locked = False
                .cell(5) = .commodityLABEL
                .cell(5).tag = .cell(5)
                .unitLABEL(1) = .unitLABEL(0)
                .newDESCRIPTION = .descriptionLABEL
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        
        If .newBUTTON.Enabled Then
            Select Case .tag
                'WarehouseIssue,WellToWell,InternalTransfer,
                'AdjustmentIssue,WarehouseToWarehouse,Sales
                Case "02040200", "02040500", "02040700", "02050300", "02040600", "02050400"
                    sql = "SELECT * FROM StockInfoIssues_New WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Transaction# = '" + .cell(0).tag + "' AND " _
                        & "Stocknumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY OriginalCondition, LogicName, SubLocaName"
                'AdjustmentEntry, WarehouseReceipt, ReturnFromRepair, Return from Well
                Case "02050200", "02040100", "02040400", "02040300"
                    sql = "SELECT * FROM StockInfoReceptions_New WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Transaction# = '" + .cell(0).tag + "' AND " _
                        & "Stocknumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY OriginalCondition, LogicName, SubLocaName"
            End Select
        Else
            Select Case .tag
                'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
                'AdjustmentIssue,WarehouseToWarehouse,Sales
                Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                    sql = "SELECT  * FROM StockInfoQTYST4_New WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Company = '" + .cell(1).tag + "' AND " _
                        & "Warehouse = '" + .cell(2).tag + "' AND " _
                        & "StockNumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY Condition, LogicName, SubLocaName"
                Case "02050200" 'AdjustmentEntry
                    sql = "SELECT stk_stcknumb as StockNumber, stk_desc as StockDescription, stk_poolspec as Pool FROM STOCKMASTER WHERE " _
                        & "(stk_npecode = '" + nameSP + "') AND " _
                        & "(stk_stcknumb = '" + .commodityLABEL + "')"
                Case "02040100" 'WarehouseReceipt
                    Dim response
                    Set datax = getDATA("statusFREIGHT", Array(nameSP, Format(.cell(4)), Format(StockNumber)))
                    If datax.RecordCount = 0 Then
                        Screen.MousePointer = 0
                        MsgBox "Error on Warehouse Module about statusFREIGHT"
                        Exit Sub
                    End If
                    
                    Set docTYPE = getDATA("getDOCTYPE", Array(nameSP, Format(.cell(4))))
                    If docTYPE.RecordCount = 0 Then
                        Screen.MousePointer = 0
                        MsgBox "Error on Warehouse Module about getDOCTYPE"
                        Exit Sub
                    End If
                    
                    If datax!poi_stasdlvy = "NR" And datax!po_freigforwr Then
                        Set docTYPE = getDATA("getDOCTYPE", Array(nameSP, Format(.cell(4))))
                        If docTYPE!doc_invcreqd Then
                            Screen.MousePointer = 0
                            MsgBox "There is no Freight Reception entered against selected line item."
                            NavBar1.SaveEnabled = False
                            Exit Sub
                        End If
                    End If
                    
                    If docTYPE!doc_invcreqd Then
                        Set datax = getDATA("statusINVOICE", Array(nameSP, Format(.cell(4)), Format(StockNumber)))
                        If datax.RecordCount = 0 Then
                            Screen.MousePointer = 0
                            MsgBox "Error on Warehouse Module"
                            Exit Sub
                        Else
                            If IsNull(datax(1)) Then
                                If docTYPE!doc_invcreqd Then
                                    Screen.MousePointer = 0
                                    MsgBox "There is no Supplier Invoice entered against selected line item."
                                    NavBar1.SaveEnabled = False
                                    Exit Sub
                                Else
                                    Msg = "There is no Supplier Invoice entered against selected line item.  Do you want to continue?"
                                    response = MsgBox(Msg, 1)
                                    If response = 1 Then
                                    Else
                                        Screen.MousePointer = 0
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    sql = "SELECT TOP 1 * FROM StockInfoPO WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "StockNumber = '" + StockNumber + "' AND " _
                        & "PO = '" + .cell(4) + "' AND " _
                        & "POitem = '" + .STOCKlist.TextMatrix(.STOCKlist.row, 6) + "'"
            End Select
        End If
    End With
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        cleanDETAILS
    Else
        datax.MoveLast
        Call workBOXESlist("clean")
        datax.MoveFirst
        total = CDbl(0)
        With frmWarehouse.Tree
            .Nodes.Clear
            moreSERIAL = False
            Do While Not datax.EOF
                If frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
'                    If frmWarehouse.newBUTTON.Enabled Then
'                        currentLOGIC = IIf(IsNull(datax!fromlogic), "", Trim(datax!fromlogic))
'                        currentSUBloca = IIf(IsNull(datax!fromSUBLOCA), "", Trim(datax!fromSUBLOCA))
'                        logicNAME = IIf(IsNull(datax!logicNAME), "", datax!logicNAME)
'                        sublocaNAME = IIf(IsNull(datax!sublocaNAME), "", datax!sublocaNAME)
'                    Else
'                        Set dataz = getDATA("getLOGICAL", NameSpace)
'                        If dataz.RecordCount > 0 Then
'                            do while
'                            logicNAME = IIf(IsNull(datax!logicNAME), "", datax!logicNAME)
'                        Else
'                            MsgBox "Error detected in the Warehouse Module"
'                            Exit Sub
'                        End If
'                    End If
                Else
                    If frmWarehouse.newBUTTON.Enabled Then
                        If frmWarehouse.tag = "02040100" Then 'WarehouseReceipt
                            currentCOND = IIf(IsNull(datax!NEWcondition), "", Trim(datax!NEWcondition))
                        Else
                            currentCOND = IIf(IsNull(datax!OriginalCondition), "", Trim(datax!OriginalCondition))
                        End If
                    Else
                        currentCOND = Trim(datax!condition)
                    End If
                    If cond <> currentCOND Then
                        moreSERIAL = False
                        If frmWarehouse.newBUTTON.Enabled Then
                            If frmWarehouse.tag = "02040100" Then 'WarehouseReceipt
                                cond = Trim(datax!NEWcondition)
                                condNAME = Trim(datax!NewConditionName)
                            Else
                                cond = Trim(datax!OriginalCondition)
                                condNAME = Trim(datax!OriginalConditionName)
                            End If
                        Else
                            cond = Trim(datax!condition)
                            condNAME = Trim(datax!ConditionName)
                        End If
                        loca = ""
                        subloca = ""
                        If frmWarehouse.tag = "02040100" Then 'WarehouseReceipt
                        Else
                            .Nodes.Add , tvwChild, "@" + cond, "Condition " + cond + " - " + condNAME, "thing"
                            .Nodes("@" + cond).Bold = True
                            .Nodes("@" + cond).backcolor = &HE0E0E0
                        End If
                    End If
                    Err.Clear
                    If frmWarehouse.newBUTTON.Enabled Then
                        currentLOGIC = IIf(IsNull(datax!fromlogic), "", Trim(datax!fromlogic))
                        currentSUBloca = IIf(IsNull(datax!fromSUBLOCA), "", Trim(datax!fromSUBLOCA))
                        logicNAME = IIf(IsNull(datax!logicNAME), "", datax!logicNAME)
                        sublocaNAME = IIf(IsNull(datax!sublocaNAME), "", datax!sublocaNAME)
                    Else
                        Select Case frmWarehouse.tag
                            Case "02040100" 'WarehouseReceipt
                                currentLOGIC = "General"
                                currentSUBloca = ""
                            Case Else
                                currentLOGIC = Trim(datax!logic)
                                currentSUBloca = Trim(datax!subloca)
                        End Select
                    End If
                End If
                Select Case frmWarehouse.tag
                    'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
                    'AdjustmentIssue,WarehouseToWarehouse,Sales,AdjustmentEntry"Special"
                    Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                        If loca <> currentLOGIC Then
                            loca = currentLOGIC
                            subloca = ""
                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + loca, "Logical Warehouse: " + datax!logicNAME, "thing 0"
                        End If
                        If subloca <> currentSUBloca Then
                            subloca = currentSUBloca
                            logicNAME = IIf(IsNull(datax!logicNAME), "", datax!logicNAME)
                            sublocaNAME = IIf(IsNull(datax!sublocaNAME), "", datax!sublocaNAME)
                            key = cond + "-" + condNAME + "{{" + loca + "{{" + subloca
                            If IsNull(datax!serialnumber) Or datax!serialnumber = "" Or UCase(datax!serialnumber) = "POOL" Then
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaNAME, "thing 1"
                                Call setupBOXES(.Nodes.Count, datax.Fields, False)
                            Else
                                moreSERIAL = True
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaNAME, "thing 0"
                            End If
                            If frmWarehouse.newBUTTON.Enabled Then
                                total = total + datax!QTY1
                            Else
                                total = total + datax!qty
                            End If
                        End If
                        If loca <> currentLOGIC Then
                            loca = currentLOGIC
                            subloca = ""
                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + loca, "Logical Warehouse: " + logicNAME, "thing 0"
                        End If
                        If subloca <> currentSUBloca Then
                            subloca = currentSUBloca
                            key = cond + "-" + datax!ConditionName + "{{" + loca + "{{" + subloca
                            If IsNull(datax!serialnumber) Or datax!serialnumber = "" Or UCase(datax!serialnumber) = "POOL" Then
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaNAME, "thing 1"
                                firstSET = .Nodes.Count
                                Call setupBOXES(.Nodes.Count, datax.Fields, False)
                            Else
                                moreSERIAL = True
                                firstSET = .Nodes.Count
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaNAME, "thing 0"
                            End If
                            total = total + datax!qty
                        End If
                        If moreSERIAL Then
                            firstSET = .Nodes.Count
                            .Nodes.Add key, tvwChild, key + "#" + datax!serialnumber, "Serial #: " + datax!serialnumber, "thing 1"
                            Call setupBOXES(.Nodes.Count, datax.Fields, True)
                        End If
                    Case "02050200" 'AdjustmentEntry
                        If frmWarehouse.newBUTTON.Enabled Then
                            Set datay = datax
                        Else
                            sql = "SELECT cond_condcode, cond_desc From condition WHERE " _
                                & "cond_npecode = '" + nameSP + "' " _
                                & "ORDER BY cond_condcode"
                            Set datay = New ADODB.Recordset
                            datay.Open sql, cn, adOpenForwardOnly
                        End If
                        If datay.RecordCount > 0 Then
                            If Err.Number = 0 Then
                                Do While Not datay.EOF
                                    If frmWarehouse.newBUTTON.Enabled Then
                                        cond = datay!NEWcondition
                                        condNAME = datay!NewConditionName
                                        pool = IIf(IsNull(datay!serialnumber), True, IIf(datay!serialnumber = "", True, False))
                                    Else
                                        cond = datay!cond_condcode
                                        condNAME = datay!cond_desc
                                        pool = datax!pool
                                    End If
                                    .Nodes.Add , tvwChild, "@" + cond, cond + "-" + condNAME, "thing"
                                    .Nodes("@" + cond).Bold = True
                                    .Nodes("@" + cond).backcolor = &HE0E0E0
                                    key = cond + "-" + condNAME + "{{"
                                    If pool Then
                                        .Nodes.Add "@" + cond, tvwChild, key, "Pool", "thing 1"
                                        firstSET = .Nodes.Count
                                        Call setupBOXES(.Nodes.Count, datax.Fields, False)
                                        frmWarehouse.addITEM.Enabled = False
                                    Else
                                        firstSET = .Nodes.Count
                                        frmWarehouse.addITEM.Enabled = True
                                    End If
                                    datay.MoveNext
                                Loop
                            End If
                        End If
                        Exit Do
                    Case "02040100" 'WarehouseReceipt
'                        If newBUTTON.Enabled Then
'                            .Nodes.Add "@" + cond, tvwChild, cond + "-" + condNAME + "{{" + "unique", "New Inventory", "thing 1"
'                        Else
'                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + "unique", "New Inventory", "thing 1"
'                        End If
'                        firstSET = .Nodes.Count
'                        Call setupBOXES(.Nodes.Count, datax.Fields, False, QTYpo)
'                        total = QTYpo

                        If frmWarehouse.newBUTTON.Enabled Then
                            Set datay = datax
                        Else
                            sql = "SELECT cond_condcode, cond_desc From condition WHERE " _
                                & "cond_npecode = '" + nameSP + "' " _
                                & "ORDER BY cond_condcode"
                            Set datay = New ADODB.Recordset
                            datay.Open sql, cn, adOpenForwardOnly
                        End If
                        If datay.RecordCount > 0 Then
                            If Err.Number = 0 Then
                                Do While Not datay.EOF
                                    If frmWarehouse.newBUTTON.Enabled Then
                                        cond = datay!NEWcondition
                                        condNAME = datay!NewConditionName
                                        pool = IIf(IsNull(datay!serialnumber), True, IIf(datay!serialnumber = "", True, False))
                                    Else
                                        cond = datay!cond_condcode
                                        condNAME = datay!cond_desc
                                    End If
                                    .Nodes.Add , tvwChild, "@" + cond, cond + "-" + condNAME, "thing"
                                    .Nodes("@" + cond).Bold = True
                                    .Nodes("@" + cond).backcolor = &HE0E0E0
                                    key = cond + "-" + condNAME + "{{"
                                    .Nodes.Add "@" + cond, tvwChild, key, "Serial:", "thing 1"
                                    firstSET = .Nodes.Count
                                    Call setupBOXES(.Nodes.Count, datax.Fields, False, QTYpo)
                                    total = QTYpo
                                    frmWarehouse.addITEM.Enabled = False
                                    datay.MoveNext
                                Loop
                            End If
                        End If
                        Exit Do
                End Select
                datax.MoveNext
            Loop
            Select Case frmWarehouse.tag
                Case "02040100" 'WarehouseReceipt
                    .Nodes.Add , , "Total", Space(90) + IIf(frmWarehouse.newBUTTON.Enabled, Space(24), "")
                Case Else
                    .Nodes.Add , , "Total", Space(53) + IIf(frmWarehouse.newBUTTON.Enabled, Space(24), "Total Available:")
            End Select
            .Nodes("Total").Bold = True
            .Nodes("Total").backcolor = &HC0C0C0
            
            'Scrolling stuff
            With cTT
                Set .Tree = frmWarehouse.Tree
            End With
            
            totalNODE = .Nodes.Count
            lastLINE = 6
            thick = 2
            
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                    frmWarehouse.combo(5).Visible = False
                    lastLINE = 8
                Case "02050200" 'AdjustmentEntry
                    lastLINE = 5
                    thick = 1
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = Space(148) + "Total to Adjust:"
                Case "02040200" 'WarehouseIssue
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(57) + "Total to Issue:"
                Case "02040500" 'WellToWell
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(53) + "Total to Transfer:"
                Case "02040700" 'InternalTransfer
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(53) + "Total to Transfer:"
                Case "02050300" 'AdjustmentIssue
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(56) + "Total to Adjust:"
                Case "02040600" 'WarehouseToWarehouse
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(53) + "Total to Transfer:"
                Case "02040100" 'WarehouseReceipt
                    lastLINE = 6
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(53) + "Total to Receive:"
                Case "02050400" 'Sales
                    If Not frmWarehouse.newBUTTON.Enabled Then .Nodes("Total").text = .Nodes("Total").text + Space(59) + "Total to Sell:"
                Case "02040300" 'Return from Well
                    lastLINE = 7
            End Select
            
            Load frmWarehouse.quantity(totalNODE)
            If Err.Number = 360 Then
                Err.Clear
                frmWarehouse.quantity(totalNODE) = ""
            End If
            frmWarehouse.quantity(totalNODE).Enabled = True
            frmWarehouse.quantity(totalNODE) = total
            
            Load frmWarehouse.NEWconditionBOX(totalNODE)
            If Err.Number = 360 Then
                Err.Clear
                frmWarehouse.NEWconditionBOX(totalNODE) = ""
            End If
            frmWarehouse.NEWconditionBOX(totalNODE).Enabled = True
            
            Load frmWarehouse.quantityBOX(totalNODE)
            If Err.Number = 360 Then
                Err.Clear
                frmWarehouse.quantityBOX(totalNODE) = ""
            End If
            frmWarehouse.quantityBOX(totalNODE).locked = True
            
            Load frmWarehouse.balanceBOX(totalNODE)
            If Err.Number = 360 Then
                Err.Clear
                frmWarehouse.balanceBOX(totalNODE) = ""
            End If
            frmWarehouse.balanceBOX(totalNODE).Enabled = True
            
            Call calculations
                                    
            For i = 1 To totalNODE
                .Nodes(i).Expanded = True
            Next
            If Not .Visible Then Call SHOWdetails
            .ZOrder
            If Not frmWarehouse.newBUTTON.Enabled Then frmWarehouse.SUMMARYlist.Visible = False
            Call SHOWdetails
                        
            'Lines stuff
            n = 0
            For i = 1 To lastLINE
                Load frmWarehouse.linesV(i)
                If Err.Number = 360 Then Err.Clear
                If i = thick Then
                    frmWarehouse.linesV(i).width = 40
                End If
                frmWarehouse.linesV(i).Top = .Top + 30
                frmWarehouse.linesV(i).Height = ((totalNODE - 1) * 240)
                frmWarehouse.linesV(i).Left = frmWarehouse.detailHEADER.ColWidth(i - 1) + 150 + n
                n = n + frmWarehouse.detailHEADER.ColWidth(i - 1)
                frmWarehouse.linesV(i).Visible = True
                frmWarehouse.linesV(i).ZOrder
            Next
            frmWarehouse.linesV(lastLINE).BorderStyle = 0
            frmWarehouse.linesV(lastLINE).Appearance = 0
            frmWarehouse.linesV(lastLINE).backcolor = &HE0E0E0
            frmWarehouse.linesV(lastLINE).width = frmWarehouse.detailHEADER.ColWidth(lastLINE) + 10
            frmWarehouse.linesV(lastLINE).Height = .Height - 60
            
            Call workBOXESlist("fix")
            If .Nodes.Count > 15 Then
                frmWarehouse.linesV(lastLINE).Visible = False
                .Nodes(1).EnsureVisible
            End If
            
            If firstSET > 0 Then
                If frmWarehouse.logicBOX(firstSET).Visible Then
                    frmWarehouse.logicBOX(firstSET).SetFocus
                    frmWarehouse.grid(1).ToolTipText = Format(firstSET, "00") + "logicBOX"
                    Call showGRID(frmWarehouse.grid(1), firstSET, frmWarehouse.logicBOX(firstSET), True)
                End If
            End If
        End With
    End If
    directCLICK = False
    Screen.MousePointer = 0
    frmWarehouse.MousePointer = 0
    frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
End Sub

Sub showGRID(ByRef grid As MSHFlexGrid, Index, box As textBOX, Optional noFILLING As Boolean)
Dim n
    If frmWarehouse.Tree.Visible = False Then Exit Sub
    If frmWarehouse.Tree.Nodes.Count < 2 Then Exit Sub
    If frmWarehouse.SUMMARYlist.Visible Then Exit Sub
    With grid
        If .Rows > 2 And .text <> "" Then
            n = box.Left + .width
            If n >= frmWarehouse.width Then
                .Left = box.Left - (n - frmWarehouse.width) - 100
            Else
                .Left = box.Left
            End If
            .Top = box.Top + box.Height + 10
            .ZOrder
            .Visible = True
        End If
    End With
End Sub

Sub doCOMBO(Index, datax As ADODB.Recordset, list, totalwidth)
Dim rec, i, r, extraW
Dim t As String
    Err.Clear
    With frmWarehouse.combo(Index)
        .Rows = .Rows + datax.RecordCount
        r = 1
        Do While Not datax.EOF
            rec = ""
            For i = 0 To frmWarehouse.matrix.TextMatrix(1, Index) - 1
                If list(i) = "error" Then
                    MsgBox "Definition error, please contact IMS"
                    Exit Sub
                Else
                    t = IIf(IsNull(datax(list(i))), "", datax(list(i)))
                    .TextMatrix(r, i) = t
                End If
            Next
            r = r + 1
            datax.MoveNext
        Loop
        ''''''
        If .TextMatrix(.Rows - 1, 0) = "" Then .RemoveItem (.Rows - 1)
        .row = 1
        If .Rows < 6 Then
            extraW = 0
            .Height = (240 * .Rows)
            .ScrollBars = flexScrollBarNone
        Else
            extraW = 270
            .Height = 1455
            .ScrollBars = flexScrollBarVertical
        End If
        If frmWarehouse.cell(Index).width > (totalwidth + extraW) Then
            .width = frmWarehouse.cell(Index).width
            .ColWidth(0) = .ColWidth(0) + (.width - totalwidth) - extraW
        Else
            .width = totalwidth + extraW
        End If
        If (frmWarehouse.cell(Index).Left + .width) > frmWarehouse.width Then
            .Left = frmWarehouse.width - .width - 100
        Else
            .Left = frmWarehouse.cell(Index).Left
        End If
    End With
End Sub

Function getFROMgrid(grid As MSHFlexGrid, column, text) As String
Dim i
    For i = 1 To grid.Rows - 1
        If text = grid.TextMatrix(i, column) Then
            grid.row = i
        End If
    Next
End Function

Function InvtReceipt_Insert2a(NameSpace As String, PONumb As String, TranType As String, CompanyCode As String, Warehouse As String, user As String, cn As ADODB.Connection, Optional ManufacturerNumb As String, Optional TranFrom As String, Optional TransNum As String) As Integer

Dim v As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceipt_Insert"
    
        If Len(TransNum) = 0 Then _
         Err.Raise 1000, "Transaction Number missing" 'TansNum = GetTransNumb(NameSpace, cn)
        
        If Len(Trim$(NameSpace)) = 0 Then Err.Raise 5000, "Namespace is empty"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@COMPANYCODE", adChar, adParamInput, 10, RTrim$(CompanyCode))
        .Parameters.Append .CreateParameter("@WHAREHOUSE", adChar, adParamInput, 10, RTrim$(Warehouse))
        
        .Parameters.Append .CreateParameter("@TRANS", adVarChar, adParamInput, 15, RTrim$(TransNum))
        .Parameters.Append .CreateParameter("@TRANTYPE", adChar, adParamInput, 2, RTrim$(TranType))
        
        v = RTrim$(TranFrom)
        If Len(Trim$(TranFrom)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("@TRANFROM", adVarChar, adParamInput, 10, v)
        
        v = RTrim$(ManufacturerNumb)
        If Len(Trim$(ManufacturerNumb)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, v)
        
        v = RTrim$(PONumb)
        If Len(Trim$(PONumb)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, v)
        
        .Parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, user)
        
        Call .Execute(Options:=adExecuteNoRecords)
        InvtReceipt_Insert2a = .Parameters("RV") = 0
    End With
        
    If InvtReceipt_Insert2a Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function
Public Function MakeCommand(cn As ADODB.Connection, CommandType As ADODB.CommandTypeEnum) As ADODB.Command
    Set MakeCommand = Nothing
    Set MakeCommand = New ADODB.Command
    Set MakeCommand.ActiveConnection = cn
    MakeCommand.CommandType = CommandType
End Function
Sub colorCOLS()
Dim i As Integer
    With frmWarehouse.STOCKlist
        .row = STOCKlist.Rows - 1
        .col = 3
        .CellBackColor = &HE0E0E0
        .col = 7
        .CellBackColor = &HE0E0E0
        .col = 11
        .CellBackColor = &HE0E0E0
        For i = 8 To 10
            .col = i
            If Val(.TextMatrix(.row, 17)) = 0 Then
                .CellBackColor = &HC0FFFF 'Very Light Yellow
            Else
                .CellBackColor = &HFFFFC0 'Very Light Green
            End If
        Next
    End With
End Sub

Sub differences(row As Integer)
Dim d1, d2 As Double
Dim s1, s2 As String
Dim col, currentROW As Integer

With frmWarehouse.STOCKlist
        s1 = .TextMatrix(row, 6)
        s2 = .TextMatrix(row, 10)
        
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
            .TextMatrix(row, 12) = FormatNumber((d2 - d1), 2)
            col = .col
            .col = 12
            currentROW = .row
            .row = row
            If (d2 - d1) >= 0 Then
                .CellForeColor = vbBlack
            Else
                .CellForeColor = vbRed
            End If
            .col = col
            .row = currentROW
        End If
    End With
End Sub

Sub drawLINEcol(ByVal grid As MSHFlexGrid, col As Integer)
    With frmWarehouse.grid(0)
        .ColWidth(col) = 50 'Line
        .col = col
        .CellBackColor = &H808080
    End With
End Sub
Sub fillSTOCKlist(datax As ADODB.Recordset)
Dim n, rec, i
    With datax
        n = 1
        frmWarehouse.STOCKlist.Rows = .RecordCount + 1
        frmWarehouse.STOCKlist.row = 1
        frmWarehouse.STOCKlist.col = 0
        frmWarehouse.STOCKlist.CellFontName = "MS Sans Serif"
        Do While Not .EOF
            Select Case frmWarehouse.tag
                'ReturnFromRepair, AdjustmentEntry,WarehouseIssue,WellToWell,InternalTransfer,
                'AdjustmentIssue,WarehouseToWarehouse,Sales
                Case "02040400", "02050200", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                    frmWarehouse.STOCKlist.TextMatrix(n, 0) = Format(n)
                    frmWarehouse.STOCKlist.TextMatrix(n, 1) = Trim(!StockNumber)
                    frmWarehouse.STOCKlist.TextMatrix(n, 2) = IIf(IsNull(!unitprice), "0.00", Format(!unitprice, "0.00"))
                    frmWarehouse.STOCKlist.TextMatrix(n, 3) = IIf(IsNull(!description), "", !description)
                    frmWarehouse.STOCKlist.TextMatrix(n, 4) = IIf(IsNull(!unit), "", !unit)
                    frmWarehouse.STOCKlist.TextMatrix(n, 5) = Format(!qty, "0.00")
                    frmWarehouse.STOCKlist.TextMatrix(n, 6) = IIf(IsNull(!unit), "", !unit)
                    frmWarehouse.STOCKlist.TextMatrix(n, 7) = Format(!qty, "0.00")
                Case "02040100" 'WarehouseReceipt
                    frmWarehouse.STOCKlist.TextMatrix(n, 0) = Format(!POitem)
                    frmWarehouse.STOCKlist.TextMatrix(n, 1) = Trim(!StockNumber)
                    frmWarehouse.STOCKlist.TextMatrix(n, 2) = IIf(IsNull(!QTYpo), "0.00", Format(!QTYpo, "0.00"))
                    frmWarehouse.STOCKlist.TextMatrix(n, 3) = Format(!QTY1, "0.00")
                    frmWarehouse.STOCKlist.TextMatrix(n, 4) = IIf(IsNull(!unit), "", !unit)
                    frmWarehouse.STOCKlist.TextMatrix(n, 5) = IIf(IsNull(!description), "", !description)
                    frmWarehouse.STOCKlist.TextMatrix(n, 6) = Format(!POitem)
                    frmWarehouse.STOCKlist.TextMatrix(n, 7) = Format(!QTY1, "0.00")
            End Select
            If n = 20 Then
                DoEvents
                frmWarehouse.STOCKlist.Refresh
            End If
            n = n + 1
            .MoveNext
        Loop
        frmWarehouse.STOCKlist.RowHeightMin = 240
        frmWarehouse.STOCKlist.row = 1
        frmWarehouse.STOCKlist.col = 0
        frmWarehouse.STOCKlist.ColSel = frmWarehouse.STOCKlist.cols - 1
    End With
End Sub

Function getLOCATIONdescription(Location) As String
Dim sql
Dim datax As New ADODB.Recordset
    With frmWarehouse
        sql = "SELECT loc_name FROM LOCATION WHERE " _
            & "loc_npecode = '" + nameSP + "' AND " _
            & "loc_compcode = '" + .cell(1).tag + "' AND " _
            & "loc_locacode = '" + Location + "'"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenForwardOnly
        If datax.RecordCount > 0 Then
            getLOCATIONdescription = datax!loc_name
        Else
            getLOCATIONdescription = ""
        End If
    End With
End Function

Function getCOMPANYdescription(Company) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "SELECT com_compcode FROM COMPANY WHERE " _
        & "com_npecode = '" + nameSP + "' AND " _
        & "com_compcode = '" + Company + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getCOMPANYdescription = datax!com_compcode
    Else
        getCOMPANYdescription = ""
    End If
End Function

Function getUNITdescription(unit) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "SELECT uni_desc FROM UNIT WHERE " _
        & "uni_npecode = '" + nameSP + "' AND " _
        & "uni_code = '" + unit + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getUNITdescription = datax!uni_desc
    Else
        getUNITdescription = ""
    End If
End Function
Function getUNIT(StockNumber) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "SELECT stk_primuon FROM STOCKMASTER WHERE " _
        & "stk_npecode = '" + nameSP + "' AND " _
        & "stk_stcknumb = '" + StockNumber + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getUNIT = datax!stk_primuon
    Else
        getUNIT = ""
    End If
End Function

Function getUSERname(userCODE) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "SELECT usr_username FROM XUSERPROFILE WHERE " _
        & "usr_npecode = '" + nameSP + "' AND " _
        & "usr_userid = '" + userCODE + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getUSERname = datax!usr_username
    Else
        getUSERname = ""
    End If
End Function

Function howMANY(text, toSEARCH) As Integer
Dim i As Integer
Dim slice
    i = 0
    slice = text
    Do While True
        If InStr(slice, toSEARCH) > 0 Then
            i = i + 1
            slice = Mid(slice, InStr(slice, toSEARCH) + 1)
        Else
            Exit Do
        End If
    Loop
    howMANY = i
End Function

Function isOPEN(PO As String) As Boolean
Dim sql As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    With frmWarehouse
        isOPEN = False
        PO = Trim(.cell(0))
        sql = "SELECT po_ponumb, po_stas from PO WHERE po_npecode = '" + nameSP + "' " _
            & "AND po_ponumb = '" + .cell(0) + "'"
        Set dataPO = New ADODB.Recordset
        dataPO.Open sql, cn, adOpenForwardOnly
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
    End With
End Function

Sub cleanSTOCKlist()
Dim i
    With frmWarehouse.STOCKlist
        .Rows = 2
        For i = 0 To .cols - 1
            .TextMatrix(1, i) = ""
        Next
        .RowHeightMin = 0
        .RowHeight(1) = 0
    End With
End Sub
Sub cleanSUMMARYlist()
Dim i
    With frmWarehouse.SUMMARYlist
        .Rows = 2
        For i = 0 To .cols - 2
            .TextMatrix(1, i) = ""
        Next
        .RowHeightMin = 0
        .RowHeight(1) = 0
    End With
End Sub

Sub clearDOCUMENT()
Dim i As Integer
    With frmWarehouse
        readyFORsave = False
        For i = 2 To 9
            .cell(i) = ""
            .cell(i).backcolor = .remarks.backcolor
        Next
        .remarks = ""
        .Command1.Caption = "&Show Only Selection"
    End With
End Sub

Function controlOBJECT(controlNAME As String) As Control
Dim c As Control
    With frmWarehouse
        For Each c In .Controls
            If c.name = controlNAME Then
                Exit For
            End If
            Set c = Nothing
        Next
        Set controlOBJECT = c
    End With
End Function
Sub markROW(grid As MSHFlexGrid)
Dim nextROW, purchaseUNIT As String
Dim i  As Integer
Dim stock
Screen.MousePointer = 11
frmWarehouse.Refresh
    With grid
        .col = 0
        Dim currentformname, currentformname1
        Dim imsLock As imsLock.Lock
        Dim ListOfPrimaryControls() As String
        Set imsLock = New imsLock.Lock
        stock = .TextMatrix(.row, 1)
        currentformname = frmWarehouse.tag + "stock"
        currentformname1 = currentformname
        
        If IsNumeric(.text) Then
            'Lock
            Call imsLock.Check_Lock(STOCKlocked, cn, CurrentUser, Array("", stock, nameSP, "", "", "", "", ""), currentformname1, rowguid, "STOCKMASTER")
            If Err.Number <> 0 Then
                Err.Clear
                Screen.MousePointer = 0
                Exit Sub
            End If
            If STOCKlocked = True Then
                Screen.MousePointer = 0
                Exit Sub
            Else
                STOCKlocked = True
            End If
            '----
            
            .CellFontName = "Wingdings 3"
            .CellFontSize = 10
            .text = "Æ"
            If .name = frmWarehouse.STOCKlist.name Then
                Call PREdetails
            Else
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6), .TextMatrix(.row, 2))
            End If
        Else
            .CellFontName = "MS Sans Serif"
            .CellFontSize = 8.5
            .text = .row
            Call fillDETAILlist("", "", "")
            
            'Unlock
            Set imsLock = New imsLock.Lock
            Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, True, "STOCKMASTER", stock, False)
            Set imsLock = Nothing
            '------
        End If
    End With
Screen.MousePointer = 0
End Sub
Function getSUBLOCATIONdescription(sublocation) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "select sb_desc as Description from SUBLOCATION WHERE " _
        & "sb_npecode = '" + nameSP + "' AND " _
        & "sb_code = '" + sublocation + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getSUBLOCATIONdescription = datax!description
    Else
        getSUBLOCATIONdescription = ""
    End If
End Function
Function getWAREHOUSEdescription(Warehouse) As String
Dim sql
Dim datax As New ADODB.Recordset
    sql = "select lw_desc as Description from LOGWAR WHERE " _
        & "lw_npecode = '" + nameSP + "' AND " _
        & "lw_code = '" + Warehouse + "'"
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenForwardOnly
    If datax.RecordCount > 0 Then
        getWAREHOUSEdescription = datax!description
    Else
        getWAREHOUSEdescription = ""
    End If
End Function
Sub SHOWdetails()
    With frmWarehouse
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
                .detailHEADER.Top = .newDESCRIPTION.Top + .newDESCRIPTION.Height + 100
                .Tree.Top = .detailHEADER.Top + 315
                .hideDETAIL.Top = .Tree.Top - .hideDETAIL.Height - 410
                .removeDETAIL.Top = .hideDETAIL.Top
                .submitDETAIL.Top = .hideDETAIL.Top
                .Tree.Height = .Command5.Top - .Tree.Top - 150
                .cell(5).Visible = True
                .newDESCRIPTION.Visible = True
                If .newBUTTON.Enabled Then
                    .cell(5).Enabled = False
                Else
                    .cell(5).Enabled = True
                End If
                Call workBOXESlist("FIX")
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
            Case "02050400" 'Sales
            Case "02040300" 'Return from Repair
        End Select
        .otherLABEL(0).Visible = True
        .commodityLABEL.Visible = True
        .descriptionLABEL.Visible = True
        .remarksLABEL.Visible = False
        .remarks.Visible = False
        .SUMMARYlist.Visible = False
        .hideDETAIL.Visible = True
        .submitDETAIL.Visible = True
        .removeDETAIL.Visible = True
        .Label4(0).Visible = True
        .Label4(1).Visible = True
        .hideDETAIL.Visible = True
        .submitDETAIL.Visible = True
        .removeDETAIL.Visible = True
    End With
End Sub
Sub putBOX(box As textBOX, Left, Top, width, backcolor)
    With box
        .Left = Left
        .width = width
        .Top = Top
        .Height = 210
        If (frmWarehouse.Tree.Nodes.Count > 15 And .Index < 16) Or frmWarehouse.Tree.Nodes.Count < 16 Then
            .ZOrder
            .Visible = True
        End If
        .backcolor = backcolor
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200", "02040100" 'AdjustmentEntry, WarehouseReceipt
                If .name = "quantity" Then .Visible = False
                If .name = "balanceBOX" Then .Visible = False
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
    End With
End Sub

Function topNODE(Index) As Integer
    topNODE = frmWarehouse.Tree.Top + 45 + (240 * (Index - nodeONtop))
End Function

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
                box.width = .ColWidth(1)
            Else
                box.Left = .Left
                box.Top = .Top - box.Height
                box.width = .ColWidth(0)
            End If
        End If
        box.Visible = True
        box.text = .text
        If standard Then
            box.SetFocus
        End If
    End With
End Sub

Sub unlockBUNCH()
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    Dim grid1 As Boolean
    Dim grid2 As Boolean
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, grid1, "STOCKMASTER", , grid2)
End Sub

Sub validateQTY(box As textBOX, Index)
Dim n
Dim d As Integer
    noRETURN = True
    With box
        If Index <> totalNODE Then
            If IsNumeric(.text) Then
                If .name = "priceBOX" Then
                    d = 2
                Else
                    d = 0
                End If
                n = FormatNumber(CDbl(.text), d)
                If Right(.text, 1) = "." Then
                Else
                    If InStr(.text, ".") = 0 Then
                        If CDbl(.text) > 0 Then
                            If Len(n) > 8 Then .text = Left(.text, 8)
                        Else
                            .text = "0"
                        End If
                    Else
                        n = Mid(.text, InStr(.text, ".") + 1)
                        If Len(n) > 8 Then .text = Left(.text, Len(.text) - 1)
                    End If
                End If
            Else
                .text = "0"
            End If
        End If
        .SelStart = Len(.text)
    End With
End Sub

Function PutReturnData(prefix As String) As Boolean
Dim NP As String
Dim WH As String
Dim From As String
On Error GoTo errPutReturnData
    With frmWarehouse
        PutReturnData = False
        NP = nameSP
        Transnumb = prefix + "-" & GetTransNumb(NP, cn)
        WH = .cell(3).tag
        From = .cell(2).tag
        PutReturnData = InvtReceipt_Insert2a(NP, "", prefix, .cell(1).tag, WH, Format(CurrentUser), cn, , From, Format(Transnumb))
        Exit Function
    End With
errPutReturnData:
    MsgBox Err.description: Err.Clear
End Function
Function PutReturnData2() As Boolean
Dim NP As String
Dim WH As String
Dim cmd As Command
Dim From As String
On Error GoTo errPutReturnData
    With frmWarehouse
        PutReturnData2 = False
        'Set cmd = deIms.Commands("InvtIssue_Insert")
        NP = nameSP
        Transnumb = "AE-" & GetTransNumb(NP, cn)
        WH = .cell(2).tag
        From = WH
        PutReturnData2 = InvtReceipt_Insert(NP, "", "AE", .cell(1).tag, WH, Format(CurrentUser), cn, , From, Format(Transnumb))
        Exit Function
    End With

errPutReturnData:
    MsgBox Err.description: Err.Clear
End Function

Function summaryQTY(StockNumber, conditionCODE, fromlogic, sublocation, serial, Node) As Integer
Dim i, condition, key
    With frmWarehouse.SUMMARYlist
        For i = 1 To .Rows - 1
            summaryPOSITION = i
            If Trim(.TextMatrix(i, 1)) = Trim(StockNumber) And .TextMatrix(i, 20) = conditionCODE And .TextMatrix(i, 9) = fromlogic And .TextMatrix(i, 10) = sublocation Then
                If IsNull(serial) Or serial = "" Or UCase(serial) = "POOL" Then
                    key = frmWarehouse.Tree.Nodes(Node).key
                    condition = Mid(key, InStr(key, "-") + 1, InStr(key, "{{") - InStr(key, "-") - 1)
                    If condition = .TextMatrix(i, 3) Then
                        summaryQTY = .TextMatrix(i, 7)
                        Exit Function
                    End If
                Else
                    If .TextMatrix(i, 2) = serial Then
                        summaryQTY = .TextMatrix(i, 7)
                        Exit Function
                    End If
                End If
            End If
        Next
        summaryPOSITION = 0
        summaryQTY = 0
    End With
End Function

Sub getCOLORSrow(grid As MSHFlexGrid, columns)
Dim i, currentCOL As Integer
    currentCOL = STOCKlist.col
    With frmWarehouse.grid(0)
        For i = 1 To columns
            .col = i
            colorsROW(i) = .CellBackColor
        Next
        .col = currentCOL
    End With
End Sub
Sub workBOXESlist(work)
Dim i, size, point
On Error Resume Next
    With frmWarehouse
        size = .Tree.Nodes.Count
        If size > 0 Then
            For i = 1 To size
                Err.Clear
                If .quantity(i) <> "" Then
                    If .quantity(i) <> "" Then
                        If Err.Number = 0 Then
                            Select Case UCase(work)
                                Case "CLEAN"
                                    If i > 0 Then
                                        Unload .quantity(i)
                                        Unload .logicBOX(i)
                                        Unload .sublocaBOX(i)
                                        Unload .balanceBOX(i)
                                        Unload .NEWconditionBOX(i)
                                        Unload .quantityBOX(i)
                                        Unload .priceBOX(i)
                                        Unload .unitBOX(i)
                                        Unload .repairBOX(i)
                                    End If
                                Case "FIX"
                                    Select Case frmWarehouse.tag
                                        Case "02040400" 'ReturnFromRepair
                                            point = 2
                                        Case "02040300" 'ReturnFromWell
                                            point = 1
                                        Case Else
                                            point = 0
                                    End Select
                                    If i = size Then
                                        Select Case .tag
                                            Case "02040100" 'WarehouseReceipt
                                                Call putBOX(.quantityBOX(totalNODE), .linesV(3 + point).Left + 30, topNODE(size), .detailHEADER.ColWidth(3 + point) - 50, &HC0C0C0)
                                            Case Else
                                                If Not .newBUTTON.Enabled Then Call putBOX(.quantity(totalNODE), .linesV(1).Left + 20, topNODE(size), .detailHEADER.ColWidth(4 + point) - 50, &HC0C0C0)
                                                Call putBOX(.quantityBOX(totalNODE), .linesV(4 + point).Left + 30, topNODE(size), .detailHEADER.ColWidth(4 + point) - 50, &HC0C0C0)
                                                If Not .newBUTTON.Enabled Then Call putBOX(.balanceBOX(totalNODE), .linesV(5 + point).Left + 30, topNODE(size), .detailHEADER.ColWidth(5 + point) - 50, &HC0C0C0)
                                        End Select
                                    Else
                                        If Not .newBUTTON.Enabled Then
                                            Select Case .tag
                                                Case "02040100" 'WarehouseReceipt
                                                    Call putBOX(.logicBOX(i), .linesV(1).Left + 55, topNODE(i), .detailHEADER.ColWidth(1) - 80, vbWhite)
                                                    Call putBOX(.sublocaBOX(i), .linesV(2).Left + 30, topNODE(i), .detailHEADER.ColWidth(2) - 50, vbWhite)
                                                    Call putBOX(.quantityBOX(i), .linesV(3 + point).Left + 30, topNODE(i), .detailHEADER.ColWidth(3 + point) - 50, vbWhite)
                                                Case Else
                                                    Call putBOX(.quantity(i), .linesV(1).Left + 40, topNODE(i), .detailHEADER.ColWidth(1) - 80, vbWhite)
                                                    Call putBOX(.balanceBOX(i), .linesV(5 + point).Left + 30, topNODE(i), .detailHEADER.ColWidth(5 + point) - 50, vbWhite)
                                                    
                                                    Call putBOX(.logicBOX(i), .linesV(2).Left + 55, topNODE(i), .detailHEADER.ColWidth(2) - 80, vbWhite)
                                                    Call putBOX(.sublocaBOX(i), .linesV(3).Left + 30, topNODE(i), .detailHEADER.ColWidth(3) - 50, vbWhite)
                                                    Call putBOX(.quantityBOX(i), .linesV(4 + point).Left + 30, topNODE(i), .detailHEADER.ColWidth(4 + point) - 50, vbWhite)
                                            End Select
                                        End If
                                        Select Case .tag
                                            Case "02040400", "02040300" 'ReturnFromRepair, ReturnFromWell
                                                Call putBOX(.NEWconditionBOX(i), .linesV(4).Left + 30, topNODE(i), .detailHEADER.ColWidth(4) - 50, vbWhite)
                                                If frmWarehouse.tag = "02040400" Then
                                                    Call putBOX(.repairBOX(i), .linesV(5).Left + 30, topNODE(i), .detailHEADER.ColWidth(5) - 50, vbWhite)
                                                End If
                                            Case "02050200" 'AdjustmentEntry
                                                Call putBOX(.logicBOX(i), .linesV(1).Left + 50, topNODE(i), .detailHEADER.ColWidth(1) - 80, vbWhite)
                                                Call putBOX(.sublocaBOX(i), .linesV(2).Left + 30, topNODE(i), .detailHEADER.ColWidth(2) - 50, vbWhite)
                                                Call putBOX(.priceBOX(i), .linesV(3).Left + 30, topNODE(i), .detailHEADER.ColWidth(3 + point) - 50, vbWhite)
                                                Call putBOX(.quantityBOX(i), .linesV(4).Left + 30, topNODE(i), .detailHEADER.ColWidth(4 + point) - 50, vbWhite)
                                            Case "02040100" 'WarehouseReceipt
                                        End Select
                                    End If
                            End Select
                        Else
                            Err.Clear
                        End If
                    End If
                End If
            Next
        End If
    End With
    Err.Clear
End Sub

Sub calculations()
Dim i, this, r, balance
Dim col
Dim firstTIME As Boolean
On Error Resume Next
    firstTIME = True
    With frmWarehouse
        this = 0
        balance = 0
        For i = 1 To .Tree.Nodes.Count
            If i <> totalNODE Then
                If .quantity(i) <> "" Then
                    If Err.Number = 0 Then
                        .balanceBOX(i) = Format(.quantity(i) - .quantityBOX(i), "0.00")
                        this = this + CDbl(.quantityBOX(i))
                        Select Case frmWarehouse.tag
                            Case "02040100" 'WarehouseReceipt
                                If firstTIME Then
                                    balance = balance + CDbl(.balanceBOX(i))
                                    firstTIME = False
                                End If
                            Case Else
                                balance = balance + CDbl(.balanceBOX(i))
                        End Select
                        If balance < 0 Then balance = 0
                        If .quantityBOX(i) >= 0 Then
                            r = findSTUFF(.commodityLABEL, .STOCKlist, 1)
                            If r > 0 Then
                                    Select Case frmWarehouse.tag
                                        Case "02040100" 'WarehouseReceipt
                                            col = 3
                                        Case Else
                                            col = 5
                                    End Select
                                .STOCKlist.TextMatrix(r, col) = Format(balance, "0.00")
                            End If
                        End If
                    Else
                        Err.Clear
                    End If
                End If
            End If
        Next
        .quantityBOX(totalNODE) = Format(this, "0.00")
        .balanceBOX(totalNODE) = Format(balance, "0.00")
    End With
End Sub

Function findSTUFF(toFIND, grid As MSHFlexGrid, col) As Integer
Dim i
Dim findIT As Boolean
    findSTUFF = 0
    With grid
        If .Rows > 1 Then
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
                    If UCase(Trim(.TextMatrix(i, col))) = UCase(Trim(toFIND)) Then
                        findSTUFF = i
                        Exit For
                    End If
                Next
            End If
        End If
    End With
End Function


Sub cleanDETAILS()
Dim i
On Error Resume Next
    With frmWarehouse
        nodeONtop = 1
        For i = 1 To 10
            Unload .linesV(i)
            If Err.Number <> 0 Then Err.Clear
        Next
        .cell(5).Visible = False
        .combo(5).Visible = False
        .newDESCRIPTION.Visible = False
        .otherLABEL(2).Visible = False
        Call workBOXESlist("clean")
        .Tree.Nodes.Clear
    End With
    Err.Clear
End Sub

Sub BeforePrint()
Set translatorFORM = imsTranslator
    
    With MDI_IMS.CrystalReport1
        .Reset
        'msg1 = translator.Trans("L00176")
        .WindowTitle = IIf(msg1 = "", "transaction", msg1)
        .ParameterFields(0) = "namespace;" + nameSP + ";TRUE"
        If frmWarehouse.cell(1) = "" Then
            '*******************
            'CHECK THIS PATH
            .ReportFileName = App.Path + "CRreports\transactionGlobal.rpt"
            .ParameterFields(1) = "ponumb;" + frmWarehouse.cell(0) + ";TRUE"
            'call translator.Translate_Reports("transactionGlobal.rpt")
        Else
            '*******************
            'CHECK THIS PATH
            .ReportFileName = App.Path + "CRreports\transaction.rpt"
            .ParameterFields(1) = "invnumb;" + frmWarehouse.cell(1) + ";TRUE"
            .ParameterFields(2) = "ponumb;" + frmWarehouse.cell(0) + ";TRUE"
            'Call translator.Translate_Reports("transaction.rpt")
            'Call translator.Translate_SubReports
        End If
    End With
End Sub
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    If cellACTIVE <> "" Then
        With gridACTIVE
            If Not .Visible Then .Visible = True
            If .Rows < Val(.tag) Then .tag = 1
            If IsNumeric(.tag) Then
                .col = column
            End If
            If .cols <= column Then Exit Sub
            .col = column
            .tag = ""
            found = False
            
            For i = 1 To .Rows - 1
                word = Trim(UCase(.TextMatrix(i, column)))
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    .row = i
                    .tag = .row
                    .RowSel = i
                    .ColSel = .cols - 1
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                .row = 0
                .tag = ""
            End If
            If IsNumeric(.tag) Then
                If .tag = "0" Then
                    .TopRow = 1
                Else
                    Do While True
                        If .RowIsVisible(.row) Then
                            Exit Do
                        Else
                            Select Case UCase(direction)
                                Case "UP"
                                    If .TopRow < .Rows - 1 Then .TopRow = .row + 1
                                Case "DOWN"
                                    If .TopRow > 2 And .row > 1 Then .TopRow = .row - 1
                                Case Else
                                    If .RowIsVisible(.row) Then
                                    Else
                                        .TopRow = .row
                                    End If
                            End Select
                            Exit Do
                        End If
                    Loop
                End If
            End If
        End With
    End If
End Sub

Sub doARRAYS(kind, text, tempARRAY)
    Dim n, chain, shot
    If InStr(text, ",") > 0 Then
        n = -1
        chain = text
        ReDim tempARRAY(Len(text))
        Do While True
            If InStr(chain, ",") > 0 Then
                shot = Left(chain, InStr(chain, ",") - 1)
                chain = Trim(Mid(chain, InStr(chain, ",") + 1))
            Else
                shot = chain
                chain = ""
            End If
            n = n + 1
            If UCase(kind) = "S" Then
                tempARRAY(n) = Trim(shot)
            Else
                If IsNumeric(Trim(shot)) Then
                    tempARRAY(n) = Val(Trim(shot))
                Else
                    tempARRAY(n) = 0
                End If
            End If
            If chain = "" Then Exit Do
        Loop
        ReDim Preserve tempARRAY(n)
    Else
        ReDim tempARRAY(0)
        If Len(text) > 0 Then
            tempARRAY(0) = Trim(text)
        Else
            tempARRAY(0) = "error"
        End If
    End If
End Sub

Public Function putDATA(Access, Parameters) As Variant
Dim cmd As New ADODB.Command
Dim data As New ADODB.Recordset
On Error GoTo errTRACK
    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = Access
        .Execute , Parameters
    End With
    
errTRACK:
    If Err.Number = 0 Then
        putDATA = True
    Else
        putDATA = False
    End If
    Err.Clear
End Function

Public Function GetTransNumb(NameSpace As String, cn As ADODB.Connection) As Long
    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_InvtNumb"
        
        .Parameters.Append .CreateParameter("@RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NameSpace", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@numb", adInteger, adParamOutput, 4, Null)
        
        Call .Execute(Options:=adExecuteNoRecords)
        GetTransNumb = .Parameters("@numb").Value
    End With
End Function
Public Function getDATA(Access, Parameters) As ADODB.Recordset
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = Access
        Set getDATA = .Execute(, Parameters)
    End With
End Function
Public Function getCOMMAND(Access) As ADODB.Command
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = Access
        Set getCOMMAND = cmd
    End With
End Function

Public Function BeginTransaction(cn As ADODB.Connection)
    With MakeCommand(cn, adCmdText)
        .CommandText = "BEGIN TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
End Function



Sub PREdetails()
Screen.MousePointer = 11
    frmWarehouse.Refresh
    With frmWarehouse.STOCKlist
        Select Case frmWarehouse.tag
            'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 3), .TextMatrix(.row, 4))
            Case "02050200" 'AdjustmentEntry
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 2), .TextMatrix(.row, 3))
            Case "02040100" 'WarehouseReceipt
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 4), .TextMatrix(.row, 3))
        End Select
    End With
Screen.MousePointer = 0
End Sub


Public Function loadFQA(CompanyCode As String, Optional LocationCode As String) As Boolean

On Error GoTo ErrHand
loadFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Level ='C'"

RsCompany.Open , cn

Do While RsCompany.EOF

    SSOleDBFQA.addITEM RsCompany("FQA")
    RsCompany.MoveNext
    
Loop


'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='L'"

RsLocation.Open , cn


Do While RsLocation.EOF

    SSOleDBFQA.addITEM RsCompany("FQA")
    RsLocation.MoveNext
    
Loop


'Get US Chart FQA

RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='UC'"

RsUC.Open , cn


Do While RsUC.EOF

    SSOleDBFQA.addITEM RsUC("FQA")
    RsUC.MoveNext
    
Loop

'Get Cam Chart FQA

RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & CompanyCode & "' and Locationcode='" & LocationCode & "' and Level ='CC'"

RsCC.Open , cn


Do While RsCC.EOF

    SSOleDBFQA.addITEM RsCompany("FQA")
    RsCC.MoveNext
    
Loop

Set RsCompany = Nothing
Set RsLocation = Nothing
Set RsUC = Nothing
Set RsCC = Nothing

loadFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes.", vbCritical, "Ims"

Err.Clear

End Function

