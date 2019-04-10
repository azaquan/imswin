Attribute VB_Name = "warehousesFabrication"
Global fabricationFirst As Boolean
Global newFabricatedStock As Boolean
Global firstNewMultipleNode As Boolean
Global newStockCount As Integer
Global finalCostNode As Integer
Sub addFabricationMultipleNode()
Dim datax As ADODB.Recordset
Set datax = New ADODB.Recordset
On Error GoTo ErrHandler
    With frmFabrication.Tree
        If firstNewMultipleNode Then
            .Nodes.Add "Fabrication", tvwChild, "@" + "processCost", "Process Cost", "thing 1"
            Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
            Call fabWorkBOXESlist
            firstNewMultipleNode = False
        End If
        .Nodes.Add "Fabrication", tvwChild, "@" + "finalCost", "Final Unit Cost", "thing 0"
        finalCostNode = .Nodes.Count
        Call fabSetupBOXES(.Nodes.Count + 1, datax.Fields, False)
        Call fabWorkBOXESlist

        newStockCount = newStockCount + 1
        .Nodes.Add "@finalCost", tvwChild, "@" + "newStock-" + Format(newStockCount), "New Stock - " + Format(newStockCount) + ":", "thing 1"
        Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
        Call fabWorkBOXESlist
    End With

    With frmFabrication
        totalNode = .Tree.Nodes.Count
        .combo(5).Visible = False
        lastLine = 8
        For i = 1 To totalNode
            .Tree.Nodes(i).Expanded = True
        Next
        If Not .Visible Then
            Call fabSHOWdetails
        End If
        .ZOrder
        Dim newStocks As Integer
        If .many(1).Value Or .many(2).Value Then
            For i = 1 To .Tree.Nodes.Count
                key = .Tree.Nodes(i).key
                If InStr(key, "@newStock") Then key = "@newStock"
                Select Case key
                    Case "@newStock"
                        newStocks = newStocks + 1
                End Select
            Next
            Dim totalCost As Double
            totalCost = CDbl(.fabCostBOX(3)) + .priceBOX(2)
            For i = 1 To .Tree.Nodes.Count
                key = .Tree.Nodes(i).key
                If InStr(key, "@newStock") Then key = "@newStock"
                Select Case key
                    Case "@newStock"
                        If newStocks > 0 Then
                            If .many(0).Value Then
                                .priceBOX(i) = Format((totalCost / newStocks), "0.00")
                            Else
                                If .priceBOX(i) = "" Then
                                    .priceBOX(i) = "0.00"
                                End If
                            End If
                        End If
                End Select
            Next
        End If
'        Call FabLineStuff(lastLine, thick)
'        Call calculationsFabrication(False, totalNode)
    End With
    frmFabrication.Tree.Nodes(1).EnsureVisible
    Err.Clear
    baseFrame.Refresh
    Exit Sub
    
ErrHandler:
If Err.Number > 0 Then
    'MsgBox Err.description
    Err.Clear
End If
Resume Next
End Sub


Sub addFabricationNode()
Dim datax As ADODB.Recordset
Set datax = New ADODB.Recordset
Dim nodePosition As Integer
On Error GoTo ErrHandler
    With frmFabrication.Tree
        .Nodes.Add "Fabrication", tvwChild, "@" + "processCost", "Process Cost", "thing 1"
        Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
        Call fabWorkBOXESlist
        
        .Nodes.Add "Fabrication", tvwChild, "@" + "finalCost", "Final Unit Cost", "thing 0"
        finalCostNode = .Nodes.Count
        Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
        Call fabWorkBOXESlist

        .Nodes.Add "@finalCost", tvwChild, "@" + "newStock", "New Stock:", "thing 1"
        '.Nodes.Add "@processCost", tvwChild, "@" + "newStock", "New Stock#:", "thing 1"
        Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
        Call fabWorkBOXESlist
        

        nodePosition = .Nodes.Count
    End With

    With frmFabrication
        totalNode = .Tree.Nodes.Count
        .combo(5).Visible = False
        lastLine = 8
        For i = 1 To totalNode
            .Tree.Nodes(i).Expanded = True
        Next
        If .many(2).Value = False Then
            .priceBOX(nodePosition).Enabled = False
        End If
        If Not .Visible Then
            Call fabSHOWdetails
        End If
        .ZOrder
        Call FabLineStuff(lastLine, thick)
        Call calculationsFabrication(False, totalNode)
    End With
    frmFabrication.Tree.Nodes(1).EnsureVisible
    Err.Clear
    baseFrame.Refresh
    
    Exit Sub
    
ErrHandler:
If Err.Number > 0 Then
    'MsgBox Err.description
    Err.Clear
End If
Resume Next
End Sub
Sub fillDetailListFabrication(datax As ADODB.Recordset)
On Error GoTo ErrHandler
    With frmFabrication.Tree
        .width = frmFabrication.detailHEADER.width
        '.Nodes.Clear
        .Nodes.Add , tvwChild, "Fabrication", "Fabrication", "thing 0"
        .Nodes("Fabrication").Bold = True
        .Nodes("Fabrication").backcolor = &HE0E0E0
        fabricationFirst = False
        Do While Not datax.EOF
            If frmFabrication.newBUTTON.Enabled Then
            'TODO
            Else
                currentStock = IIf(IsNull(datax!StockNumber), "", Trim(datax!StockNumber))
            End If
            'TODO check this
            If frmFabrication.newBUTTON.Enabled Then
            Else
            End If
            '----------
            .Nodes.Add "Fabrication", tvwChild, "@" + currentStock, currentStock, "thing 1"
            frmFabrication.invoiceLabel.Visible = flse
            frmFabrication.invoiceLineLabel.Visible = False
            frmFabrication.invoiceNumberLabel.Visible = False
            frmFabrication.commodityLABEL.Visible = False
            frmFabrication.descriptionLABEL.Visible = False
            frmFabrication.Label4(0).Visible = False
            frmFabrication.Label4(1).Visible = False
            frmFabrication.otherLABEL(0).Visible = False
            frmFabrication.otherLABEL(1).Visible = False
            frmFabrication.unitLABEL(0).Visible = False
            Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)
            datax.MoveNext
        Loop
        frmFabrication.quantity(0).Visible = False
        frmFabrication.addFinalStock.Enabled = True
    End With
    
    With frmFabrication
        .addFinalStock.Visible = True
        totalNode = .Tree.Nodes.Count
        .combo(5).Visible = False
        lastLine = 8
        Call calculationsFabrication(False, .Tree.Nodes.Count)
        For i = 1 To totalNode
            .Tree.Nodes(i).Expanded = True
        Next
        If Not .Tree.Visible Then
            Call fabSHOWdetails
        End If
        .ZOrder
        .SUMMARYlist.Visible = .newBUTTON.Enabled
        .SUMMARYlist.width = frmFabrication.detailHEADER.width
        Call FabLineStuff(lastLine, thick)
        Call fabWorkBOXESlist
    End With
    frmFabrication.Tree.Nodes(1).EnsureVisible
    Err.Clear
    treeTimes = treeTimes + 1
    frmFabrication.treeFrame.Top = 0
    treeFrame.Refresh
    baseFrame.Refresh
Exit Sub
    
ErrHandler:
If Err.Number > 0 Then
    'MsgBox Err.description
    Err.Clear
End If
Resume Next
End Sub


Sub fabWorkBOXESlist()
Dim i, size, point, balanceCol
On Error Resume Next
'On Error GoTo errHandler
    With frmFabrication
        topvalue = 120
        topvalue2 = 0
        xPos = 0
        i = .Tree.Nodes.Count
        If i > 0 Then
            Err.Clear
            If .quantity(i) <> "" Then
                If Err.Number = 0 Then
                    balanceCol = 7
                    point = 0
                    Dim qtyCol
                    qtyCol = 6
                    If Not .newBUTTON.Enabled Then
                        Call fabPutBOX(.quantity(i), .linesV(1).Left + 20, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(4 + point) - 50, &HC0C0C0)
                        Dim key As String
                        key = .Tree.Nodes(i).key
                        If InStr(key, "@newStock") Then key = "@newStock"
                        Select Case key
                            Case "@finalCost"
                                .quantity(i).Visible = False
                                .quantityBOX(i).Visible = False
                                Call fabPutBOX(.priceBOX(i), .linesV(5).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(5) - 50, vbWhite)
                                Call fabPutBOX(.balanceBOX(i), .linesV(7).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(balanceCol + point) - 50, vbWhite)
                                .priceBOX(i).backcolor = &HC0FFC0
                                .balanceBOX(i) = "0.00"
                                .balanceBOX(i).Visible = False
                                .priceBOX(i).locked = True
                            Case "@processCost"
                                Call fabPutBOX(.fabCostBOX(i), .linesV(5).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(5) - 50, vbWhite)
                                .fabCostBOX(i).backcolor = &HFFFF80
                                .fabCostBOX(i).Enabled = True
                                .quantity(i).Visible = False
                            Case "@newStock"
                                If Not .quantityBOX(i).Visible Then
                                    xPos = 2100
                                    Call fabPutBOX(.searchStock(i), xPos, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(2) - 80, &HC0C0FF)
                                    Call fabPutBOX(.logicBOX(i), .linesV(2).Left + 55, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(2) - 80, &HC0C0FF)
                                    Call fabPutBOX(.sublocaBOX(i), .linesV(3).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(3) - 50, &HC0C0FF)
                                    Call fabPutBOX(.NEWconditionBOX(i), .linesV(4).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(4) - 50, vbWhite)
                                    Call fabPutBOX(.priceBOX(i), .linesV(5).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(5) - 50, vbWhite)
                                    Call fabPutBOX(.quantityBOX(i), .linesV(6).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(qtyCol + point) - 50, &HC0C0C0)
                                    Call fabPutBOX(.balanceBOX(i), .linesV(7).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(balanceCol + point) - 50, vbWhite)
                                    .balanceBOX(i) = "0.00"
                                    '.quantityBOX(i).Enabled = True
                                    .quantityBOX(i) = "1.00"
                                    .NEWconditionBOX(i) = "01"
                                    .priceBOX(i).Enabled = True
                                    .quantityBOX(i).Enabled = True
                                End If
                            Case Else
                                Call fabPutBOX(.quantityBOX(i), .linesV(qtyCol + point).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(qtyCol + point) - 50, &HC0C0C0)
                                Call fabPutBOX(.balanceBOX(i), .linesV(balanceCol + point).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(balanceCol + point) - 50, &HC0C0C0)
                                Call fabPutBOX(.priceBOX(i), .linesV(5).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(5) - 50, vbWhite)
                                Call fabPutBOX(.NEWconditionBOX(i), .linesV(4).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(4) - 50, vbWhite)
                                Call fabPutBOX(.balanceBOX(i), .linesV(balanceCol + point).Left + 30, fabTopNODE(i) + topvalue2, .detailHEADER.ColWidth(balanceCol + point) - 50, vbWhite)
                                .balanceBOX(i) = "0.00"
                                .quantityBOX(i).Enabled = True
                                .quantityBOX(i) = "1.00"
                        End Select
                        If .addFinalStock.Enabled Then
'                            .priceBOX(i).Enabled = False
                        Else
                            If .many(2).Value Then
'                                .priceBOX(i).Enabled = False
                            Else
                                .priceBOX(i).Enabled = True
                            End If
                        End If
                        .quantity(i).backcolor = vbWhite
                        .quantityBOX(i).backcolor = vbWhite
                    End If
                Else
                    Err.Clear
                End If
            End If
        End If
    End With
    Err.Clear
End Sub


Sub fabWorkBOXESlistClean()
Dim i, size, point, balanceCol
On Error Resume Next
    With frmFabrication
        size = .Tree.Nodes.Count
        If size > 0 Then
            For i = 1 To size
                Unload .poItemBox(i)
                Unload .positionBox(i)
                Unload .quantity(i)
                Unload .logicBOX(i)
                Unload .sublocaBOX(i)
                Unload .balanceBOX(i)
                Unload .NEWconditionBOX(i)
                Unload .quantityBOX(i)
                Unload .quantity2BOX(i)
                Unload .priceBOX(i)
                Unload .unitBOX(i)
                Unload .unit2BOX(i)
                Unload .repairBOX(i)
                Unload .fabCostBOX(i)
                Unload .searchStock(i)
            Next
        End If
    End With
End Sub

Sub fabRePositionThings(yPosition As Integer) 'Juan 2014-01-29, for scrolling placement
Dim c As textBOX
Dim i, size, newY, distance
On Error Resume Next

With frmFabrication
    
    size = .Tree.Nodes.Count
    If size > 0 Then
        distance = .Tree.Top + 320
        For i = 2 To size
            newY = fabTopNODE(yPosition) + distance 'for moving check if + distance is necessary, otherwise  comment it back
            Err.Clear
            .quantity(i).Top = .quantity(i).Top - newY
            If Err.Number = 0 Then
                .poItemBox(i).Top = fabTopNODE(i) - newY
                .positionBox(i).Top = fabTopNODE(i) - newY
                .quantity(i).Top = fabTopNODE(i) - newY
                .logicBOX(i).Top = fabTopNODE(i) - newY
                .sublocaBOX(i).Top = fabTopNODE(i) - newY
                .quantityBOX(i).Top = fabTopNODE(i) - newY
                .quantity2BOX(i).Top = fabTopNODE(i) - newY
                .balanceBOX(i).Top = fabTopNODE(i) - newY
                .NEWconditionBOX(i).Top = fabTopNODE(i) - newY
                .priceBOX(i).Top = fabTopNODE(i) - newY
                .unitBOX(i).Top = fabTopNODE(i) - newY
                .unit2BOX(i).Top = fabTopNODE(i) - newY
                .repairBOX(i).Top = fabTopNODE(i) - newY
                .fabCostBOX(i).Top = fabTopNODE(i) - newY
                .linesH(0).Top = .quantityBOX(totalNode).Top
            End If
        Next
    End If
End With
Err.Clear
End Sub
Sub fabPutThingsInsideExtension(Index As Integer) 'Juan 2014-02-02, for scrolling placement
With frmFabrication
    .quantity(Index).Visible = False
    .poItemBox(Index).Visible = False
    .positionBox(Index).Visible = False
    .quantity(Index).Visible = False
    .logicBOX(Index).Visible = False
    .sublocaBOX(Index).Visible = False
    .quantityBOX(Index).Visible = False
    .quantity2BOX(Index).Visible = False
    .balanceBOX(Index).Visible = False
    .NEWconditionBOX(Index).Visible = False
    .priceBOX(Index).Visible = False
    .unitBOX(Index).Visible = False
    .unit2BOX(Index).Visible = False
    .repairBOX(Index).Visible = False
    .searchStock(Index).Visible = False
End With
End Sub
Sub fabShowBoxes(Index)
On Error Resume Next
With frmFabrication
    .quantity(Index).Visible = True
    .positionBox(Index).Visible = True
    .logicBOX(Index).Visible = True
    .sublocaBOX(Index).Visible = True
    .quantityBOX(Index).Visible = True
    .quantity2BOX(Index).Visible = True
    .balanceBOX(Index).Visible = True
    .NEWconditionBOX(Index).Visible = True
    .priceBOX(Index).Visible = True
    .unitBOX(Index).Visible = True
    .unit2BOX(Index).Visible = True
    .repairBOX(Index).Visible = True
End With
End Sub

Sub fabPutThingsInside() 'Juan 2014-01-12, putting inside the tree container the controls
Dim c As textBOX
Dim i, size, distance
On Error Resume Next

With frmFabrication
    size = .Tree.Nodes.Count
    If size > 0 Then
        For i = 0 To 5
            .cell(i).Container = .treeFrame
            Err.Clear
        Next
        distance = .Tree.Top
        distance = distance + 360
        For i = 2 To size
            Err.Clear
            Set .quantity(i).Container = .treeFrame
            If Err.Number = 0 Then
                Set .poItemBox(i).Container = .treeFrame
                Set .positionBox(i).Container = .treeFrame
                .quantity(i).Left = 40
                .quantity(i).Top = fabTopNODE(i) - distance
                
                Set .logicBOX(i).Container = .treeFrame
                .logicBOX(i).Left = .logicBOX(i).Left - .baseFrame.Left
               .logicBOX(i).Top = fabTopNODE(i) - distance
               
                Set .sublocaBOX(i).Container = .treeFrame
                .sublocaBOX(i).Left = .sublocaBOX(i).Left - .baseFrame.Left
                .sublocaBOX(i).Top = fabTopNODE(i) - distance
                
                Set .quantityBOX(i).Container = .treeFrame
                .quantityBOX(i).Left = .quantityBOX(i).Left - .baseFrame.Left
                .quantityBOX(i).Top = fabTopNODE(i) - distance
                
'                Set .quantity2BOX(i).Container = .treeFrame
'                .quantity2BOX(i).Left = .quantity2BOX(i).Left - .baseFrame.Left
'                .quantity2BOX(i).Top = fabTopNODE(i) - distance
                
                Set .NEWconditionBOX(i).Container = .treeFrame
                .NEWconditionBOX(i).Left = .NEWconditionBOX(i).Left - .baseFrame.Left
                .NEWconditionBOX(i).Top = fabTopNODE(i) - distance

                Set .priceBOX(i).Container = .treeFrame
                .priceBOX(i).Left = .priceBOX(i).Left - .baseFrame.Left
                .priceBOX(i).Top = fabTopNODE(i) - distance
                
                Set .unitBOX(i).Container = .treeFrame
                .unitBOX(i).Left = .unitBOX(i).Left - .baseFrame.Left
                .unitBOX(i).Top = fabTopNODE(i) - distance
                
                Set .unit2BOX(i).Container = .treeFrame
                .unit2BOX(i).Left = .unit2BOX(i).Left - .baseFrame.Left
                .unit2BOX(i).Top = fabTopNODE(i) - distance
                
                Set .fabCostBOX(i).Container = .treeFrame
                .fabCostBOX(i).Left = .fabCostBOX(i).Left - .baseFrame.Left
                .fabCostBOX(i).Top = fabTopNODE(i) - distance
                
                Set .balanceBOX(i).Container = .treeFrame
                .balanceBOX(i).Left = .balanceBOX(i).Left - .baseFrame.Left
                .balanceBOX(i).Top = fabTopNODE(i) - distance
                .baseFrame.width = .balanceBOX(i).Left + .balanceBOX(i).width + 20
                .treeFrame.width = .baseFrame.width
                
                Set .searchStock(i).Container = .treeFrame
                .searchStock(i).Left = 2500
                .searchStock(i).Top = fabTopNODE(i) - distance
            End If
        Next
        .treeFrame.Height = .baseFrame.Height
        .baseFrame.Visible = True
    End If
End With
Err.Clear
End Sub
Public Function fabWriteParameterFiles(Recepients As String, sender As String, Attachments() As String, subject As String, attention As String)
Dim l
Dim x
Dim y
Dim i
Dim Email As String
Dim fax() As String
Dim rs As New ADODB.Recordset
 If Len(Trim(sender)) = 0 Then
    rs.source = "select com_name from company where com_compcode = ( select psys_compcode from pesys where psys_npecode ='" & nameSP & "')"
    rs.ActiveConnection = cn
    rs.Open
    If rs.RecordCount > 0 Then
        If Len(rs("com_name") & "") > 0 Then sender = rs("com_name")
    End If
    rs.Close
End If

On Error GoTo errMESSAGE
    Email = frmFabrication.emailRecepient.text
    If Not Email = "" Then
        Call WriteParameterFileEmail(Attachments, Email, subject, sender, attention)
    End If
errMESSAGE:
    If Err.Number <> 0 And Err.Number <> 9 Then
        MsgBox "Process fabWriteParameterFiles " + Err.description
    Else
        Err.Clear
    End If
End Function

Sub fabUpdateStockListStatus()
'This checks and updates each line on stockList depending if there is a corresponding value on the summary list
Dim i, j As Integer
Dim StockNumber As String
Dim hasMark As Boolean
Dim imsLock As imsLock.Lock

On Error GoTo errorHandler
    With frmFabrication
        For i = 1 To .STOCKlist.Rows - 1
            StockNumber = .STOCKlist.TextMatrix(i, 1)
            hasMark = False
            For j = 1 To .SUMMARYlist.Rows - 1
                'Look for possible movements within summary list
                If StockNumber = .SUMMARYlist.TextMatrix(j, 1) Then
                    If Not hasMark Then
                        .STOCKlist.row = i
                        .STOCKlist.col = 0
                        .STOCKlist.CellFontName = "Wingdings 3"
                        .STOCKlist.CellFontSize = 10
                        .STOCKlist.text = "Æ"
                        hasMark = True
                        Exit For
                    End If
                End If
            Next
            If Not hasMark Then
                If .tag = "02040100" Then 'WarehouseReceipt
                    .STOCKlist.row = i
                    .STOCKlist.col = 0
                    .STOCKlist.text = .STOCKlist.TextMatrix(0, 8) 'to recover original line item
                Else
                    .STOCKlist.TextMatrix(i, 0) = Format(i)
                    'Unlock
                    Set imsLock = New imsLock.Lock
                    Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, True, "STOCKMASTER", StockNumber, False)
                    Set imsLock = Nothing
                    '------
                End If
            End If
        Next
    End With
    Exit Sub
errorHandler:
    MsgBox Err.description
    Err.Clear
    Resume Next
End Sub




Sub fabBottomLine(totalNode, total, pool As Boolean, StockNumber, dofabRecalculate As Boolean, lastLine, ctt As cTreeTips)
Dim thick
On Error Resume Next

With frmFabrication
    totalNode = .Tree.Nodes.Count
    'Juan 2010-6-10
    'lastLINE = 6
    lastLine = 7

    thick = 2

        Select Case .tag
            Case "02040400" 'ReturnFromRepair
                .combo(5).Visible = False
                lastLine = 8
            Case "02050200" 'AdjustmentEntry
                'Juan 2010-11-20 to modify it to be similar to retorun from well
                'lastLine = 5 '
                lastLine = 7
                ''thick = 1
                'If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = Space(148) + "Total to Adjust:"
                '--------------------
            Case "02040200" 'WarehouseIssue
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(57) + "Total to Issue:"
                lastLine = 6
            Case "02040500" 'WellToWell
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(53) + "Total to Transfer:"
            Case "02040700" 'InternalTransfer
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(53) + "Total to Transfer:"
            Case "02050300" 'AdjustmentIssue
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(56) + "Total to Adjust:"
            Case "02040600" 'WarehouseToWarehouse
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(53) + "Total to Transfer:"
            Case "02040100" 'WarehouseReceipt
                lastLine = 9
                'Juan 2010-6-30
                'If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(53) + "Total to Receive:"
                '----------------------
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(43) + "Total to Receive:"
            Case "02050400" 'Sales
                If Not .newBUTTON.Enabled Then .Tree.Nodes("Total").text = .Tree.Nodes("Total").text + Space(59) + "Total to Sell:"
            Case "02040300" 'Return from Well
                lastLine = 7
        End Select
        
        Load .quantity(totalNode)
        If Err.Number = 360 Then
            Err.Clear
            .quantity(totalNode) = ""
        End If
        .quantity(totalNode).Enabled = True
        .quantity(totalNode) = Format(total, "0.00")
        .quantity(totalNode) = vbGreen
        
        
        Load .NEWconditionBOX(totalNode)
        If Err.Number = 360 Then
            Err.Clear
            .NEWconditionBOX(totalNode) = ""
        End If
        .NEWconditionBOX(totalNode).Enabled = True
        
        Load .quantityBOX(totalNode)
        If Err.Number = 360 Then
            Err.Clear
            .quantityBOX(totalNode) = ""
        End If
        .quantityBOX(totalNode).locked = True
        
        Load .quantity2BOX(totalNode)
        If Err.Number = 360 Then
            Err.Clear
            .quantity2BOX(totalNode) = ""
        End If
        .quantity2BOX(totalNode).locked = True
        
        Load .balanceBOX(totalNode)
        If Err.Number = 360 Then
            Err.Clear
            .balanceBOX(totalNode) = ""
        End If
        .balanceBOX(totalNode).Enabled = True
        
        If isFirstSubmit Then
    
                If pool Then
                    Call calculations(True, , True)
                Else
                    Call calculations(True, False, False)
                End If

        Else
            Call calculations2(.SUMMARYlist.row, .Tree.Nodes(.Tree.Nodes.Count - 1), .Tree.Nodes.Count - 1)
        End If
        For i = 1 To totalNode
            .Tree.Nodes(i).Expanded = True
        Next
        If Not .Visible Then
            Call fabSHOWdetails
        End If
        If Not pool Then
            If dofabRecalculate Then
                Call fabRecalculate(StockNumber)
            End If
        End If
        .ZOrder
        
        If Not .newBUTTON.Enabled Then .SUMMARYlist.Visible = False
'        Call fabSHOWdetails
        
        Call FabLineStuff(lastLine, thick)
        Call fabWorkBOXESlist
        If .Tree.Nodes.Count > 15 Then
            .linesV(lastLine).Visible = False
            .Tree.Nodes(1).EnsureVisible
            
            'Scrolling stuff
            Err.Clear
            
            Select Case treeTimes
                Case 0
                    Set ctt.Tree = frmFabrication.Tree
                Case 1
                    Set ctt1.Tree = frmFabrication.Tree
                Case 2
                    Set ctt2.Tree = frmFabrication.Tree
                Case 3
                    Set ctt3.Tree = frmFabrication.Tree
            End Select
            treeTimes = treeTimes + 1
            .treeFrame.Top = 0
            treeFrame.Refresh
            baseFrame.Refresh
        End If
End With

End Sub

Function fabControlExists(controlNAME As String, controlIndex As Integer) As Boolean
fabControlExists = False
Dim ctl As Control

For Each ctl In frmFabrication.Controls
    If ctl.name = controlNAME Then
        If ctl.Index = controlIndex Then
            fabControlExists = True
            Exit For
        End If
    End If
Next
End Function

Function fabControlOBJECT(controlNAME As String) As Control
Dim c As Control
    With frmFabrication
        For Each c In .Controls
            If c.name = controlNAME Then
                Exit For
            End If
            Set c = Nothing
        Next
        Set fabControlOBJECT = c
    End With
End Function
Sub FabLineStuff(lastLine, thick)
On Error Resume Next
    With frmFabrication
        n = 0
        For i = 1 To lastLine
            Load .linesV(i)
            Set .linesV(n).Container = .treeFrame 'Juan 2014-01-12, putting inside the tree container the controls
            If Err.Number = 360 Then Err.Clear
            If i = thick Then
                .linesV(i).width = 40
            End If
            .linesV(i).Top = .Tree.Top + 30
            .linesV(i).Height = ((totalNode) * 325)
            .linesV(i).Left = .detailHEADER.ColWidth(i - 1) + 150 + n
            n = n + .detailHEADER.ColWidth(i - 1)
            If i > 1 Then .linesV(i).Visible = True
            .linesV(i).ZOrder
        Next
    End With
End Sub

Sub fabRecalculate(StockNumber) 'Juan 2010-7-26
    Dim totalCount As Integer
    Dim qtyToReceive As Integer
    Dim r As Integer
    With frmFabrication
        totalCount = 0
        r = .STOCKlist.row
        For i = 1 To .SUMMARYlist.Rows - 1
            If .SUMMARYlist.TextMatrix(i, 1) = StockNumber Then
                totalCount = totalCount + 1
            End If
        Next
        If IsNumeric(.STOCKlist.TextMatrix(r, 9)) Then
            qtyToReceive = Val(.STOCKlist.TextMatrix(r, 9))
            totalCount = totalCount  'This discounts the current record itself which was already discounted
            qtyToReceive = qtyToReceive - totalCount
            .STOCKlist.TextMatrix(r, 5) = Format(qtyToReceive, "0.00")
            
'            If computerFactorValue > 0 Then
'                balance2 = 10000 / computerFactorValue
'            Else
'                balance2 = 1
'            End If
'            balance2 = balance * balance2
'            .STOCKlist.TextMatrix(r, col + 2) = Format(balance2, "0.00")
        End If
    End With
End Sub

Sub fabfabColorCOLS()
Dim i As Integer
    With frmFabrication.STOCKlist
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
Sub fabCleanSTOCKlist()
Dim i
    With frmFabrication.STOCKlist
        .Rows = 2
        For i = 0 To .cols - 1
            .TextMatrix(1, i) = ""
        Next
        .RowHeightMin = 0
        .RowHeight(1) = 0
    End With
End Sub
Sub fabCleanSUMMARYlist()
Dim i
    With frmFabrication.SUMMARYlist
        .Rows = 2
        For i = 0 To .cols - 2
            .TextMatrix(1, i) = ""
        Next
        .RowHeightMin = 0
        .RowHeight(1) = 0
    End With
    inProgress = False 'Juan 2010-7-22
End Sub
Sub fabClearDOCUMENT()
Dim i As Integer
    With frmFabrication
        readyFORsave = False
        For i = 2 To 9
            .cell(i) = ""
            .cell(i).backcolor = .remarks.backcolor
        Next
        .remarks = ""
        .Command1.Caption = "&Show Only Selection"
    End With
End Sub
Sub fabSetupBOXES(n, datax As ADODB.Fields, serial As Boolean, Optional QTYpo)
Dim x, cond, logic, subloca, newCOND, serialPool
serialPool = IIf(serial, "SERIAL", "POOL") 'Juan 2010-5-14
Dim newButtonEnabled As Boolean
On Error GoTo ErrHandler:

    With frmFabrication
        newButtonEnabled = .newBUTTON.Enabled
        Load .quantity(n)
        If Not .newBUTTON.Enabled Then Call fabPutBOX(.quantity(n), .detailHEADER.ColWidth(0) + 140, fabTopNODE(n), .detailHEADER.ColWidth(1) - 40, vbWhite)
        Load .balanceBOX(n)
        .balanceBOX(n) = Format(.quantity(n), "0.00")
        Load .quantityBOX(n)
        .quantityBOX(n).tabindex = tabindex + 2
        'Juan 2010-6-13
        Load .quantity2BOX(n)
        .quantity2BOX(n).tabindex = tabindex + 2
        '---------------------
        Load .priceBOX(n)
        Load .NEWconditionBOX(n)
        Load .invoiceBOX(n)
        Load .invoiceLineBOX(n)
        If serial Then
            .quantity(n) = 1
        Else
            If .newBUTTON.Enabled Then
                .quantity(n) = Format(datax!qty1, "0.00")
                cond = Trim(datax!originalcondition)
                logic = Trim(datax!fromlogic)
                subloca = Trim(datax!fromSubLoca)
                newCOND = IIf(IsNull(datax!NEWcondition), "", datax!NEWcondition)
            Else
                If datax.Count = 0 Then
                    .quantity(n) = "1.00"
                    cond = "2"
                    logic = ""
                    subloca = ""
                    newCOND = ""
                Else
                    .quantity(n) = Format(datax!qty, "0.00")
                    cond = Trim(datax!condition)
                    logic = Trim(datax!logic)
                    subloca = Trim(datax!subloca)
                    newCOND = datax!condition
                End If
            End If
        End If
        If datax.Count = 0 Then
            .quantityBOX(n) = "1.00"
            .priceBOX(n) = "0.00"
        Else
            .quantityBOX(n) = Format(fabSummaryQTY(Trim(datax!StockNumber), cond, logic, subloca, IIf(IsNull(datax!serialNumber), "POOL", Trim(datax!serialNumber)), n), "0.00")
            .priceBOX(n) = Format(datax!unitPRICE, "0.00")
        End If
        .NEWconditionBOX(n).tag = newCOND
        .NEWconditionBOX(n) = .NEWconditionBOX(n).tag
        Load .poItemBox(n)
        .poItemBox(n) = .poItemLabel
        Load .positionBox(n)
        Load .logicBOX(n)
        .logicBOX(n).tabindex = tabindex
        Load .sublocaBOX(n)
        .sublocaBOX(n).tabindex = tabindex + 1
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                .logicBOX(n) = datax!toLOGIC
                .sublocaBOX(n) = datax!toSUBLOCA
            Else
                .logicBOX(n) = ""
                .logicBOX(n).backcolor = &HC0C0FF
                .logicBOX(n).ToolTipText = "Select a Logic Wareshouse"
                .sublocaBOX(n) = ""
                .sublocaBOX(Index).backcolor = &HC0C0FF
                .sublocaBOX(n).ToolTipText = "Select a Sub Location"
            End If
        Else
            .logicBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 11)
            .sublocaBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 12)
            .grid(2).Visible = False
            .logicBOX(n).ToolTipText = getWAREHOUSEdescription(.logicBOX(n))
            .sublocaBOX(n).ToolTipText = getSUBLOCATIONdescription(.sublocaBOX(n))
        End If
        .logicBOX(n).tag = .logicBOX(n)
        .sublocaBOX(n).tag = .sublocaBOX(n)
        Load .unitBOX(n)
        Load .unit2BOX(n)
        .unitBOX(n).Enabled = False
        .unit2BOX(n).Enabled = False
        If .newBUTTON.Enabled Then
            .unitBOX(n) = ""
            .unit2BOX(n) = ""
        Else
            .unitBOX(n) = ""
            .unit2BOX(n) = ""
        End If
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                newCOND = datax!NEWcondition
            Else
                If datax.Count = 0 Then
                    newCOND = ""
                    .NEWconditionBOX(n).ToolTipText = ""
                Else
                    newCOND = datax!condition
                    .NEWconditionBOX(n).ToolTipText = datax!conditionName
                End If
            End If
            .NEWconditionBOX(n).tag = newCOND
            .NEWconditionBOX(n) = Format(newCOND, "00")
        Else
            .NEWconditionBOX(n).tag = .SUMMARYlist.TextMatrix(summaryPOSITION, 13)
            .NEWconditionBOX(n) = Format(.NEWconditionBOX(n).tag, "00")
            .NEWconditionBOX(n).ToolTipText = .SUMMARYlist.TextMatrix(summaryPOSITION, 14)
        End If
        Load .fabCostBOX(n)
        .fabCostBOX(n).Visible = newFabricatedStock
        
        Load .searchStock(n)
        .searchStock(n).Visible = True
        Load .stockCombo(n)
        
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                .fabCostBOX(n) = Format(datax!ird_fabrication_cost, "0.00")
                .cell(5) = Trim(datax!NewStockNumber)
                .cell(5).tag = .cell(5)
                .unitLABEL(1) = getUNIT(.cell(5).tag)
                .newDESCRIPTION = Trim(datax!NewStockDescription)
            Else
                .fabCostBOX(n) = "0.00"
            End If
        Else
            If .newBUTTON.Enabled Then
                .fabCostBOX(n) = Format(datax!ird_fabrication_cost, "0.00")
                .cell(5) = Trim(datax!NewStockNumber)
                .cell(5).tag = .cell(5)
                .unitLABEL(1) = getUNIT(.cell(5).tag)
                .newDESCRIPTION = Trim(datax!NewStockDescription)
            Else
                .fabCostBOX(n) = SUMMARYlist.TextMatrix(summaryPOSITION, 17)
                .cell(5) = SUMMARYlist.TextMatrix(summaryPOSITION, 18)
                .cell(5).tag = .cell(5)
                .unitLABEL(1) = getUNIT(.cell(5))
                .newDESCRIPTION = .SUMMARYlist.TextMatrix(summaryPOSITION, 19)
            End If
        End If
        .NEWconditionBOX(n).Enabled = False

        If .newBUTTON.Enabled Then
            .quantityBOX(n).Enabled = False
            .quantity2BOX(n).Enabled = False
            .priceBOX(n).Enabled = False
            .NEWconditionBOX(n).Enabled = False
            .sublocaBOX(n).Enabled = True
            .repairBOX(n).Enabled = False
            .fabCostBOX(n).Enabled = False
        Else
            'Juan 2010-5-17
            If serialPool = "SERIAL" Then
                .quantityBOX(n) = "1.00"
                .quantity2BOX(n) = "1.00"
                .quantityBOX(n).Enabled = False
                .quantity2BOX(n).Enabled = False
             
            Else
                .quantityBOX(n).Enabled = True
                .quantity2BOX(n).Enabled = False 'Juan 2014-03-06, changed to false because is what Alain wants
            End If
            '---------------------
            .priceBOX(n).Enabled = True
        End If
        .fabCostBOX(n).Enabled = newFabricatedStock
        .fabCostBOX(n).Visible = newFabricatedStock
    End With
    
'Juan 2010-5-17
ErrHandler:
    Select Case Err.Number
        Case 360, 340, 30
            Resume Next
        Case 0
        Case Else
            'MsgBox "Error: " + Format(Err.Number) + "/" + Err.description
            Resume Next
    End Select
    Err.Clear
'------------------------
End Sub


Sub fabFillDETAILlist(StockNumber, description, unit, Optional QTYpo, Optional stockListRow, Optional serialNum, Optional hasInvoice As Boolean, Optional ctt As cTreeTips)
Dim i, n, sql, rec, cond, loca, subloca, stock, total, key, lastLine, thick, condName, currentLOGIC, currentSUBloca
Dim sublocaname, logicname, currentCOND, currentStock
Dim pool As Boolean
Dim moreSerial As Boolean
Dim datax As ADODB.Recordset
Dim datay As ADODB.Recordset
Dim dataz As ADODB.Recordset
Dim docTYPE As ADODB.Recordset
Dim sNumber As String
Dim thisSubLoca, thisLogic As String
Dim multipleLine As Boolean
mutipleLine = False
serialStockNumber = False
sqlKey = ""
'On Error Resume Next
On Error GoTo ErrHandler
    With frmFabrication
    
    
    If fabSummaryQTYshort(StockNumber) > 0 Then Exit Sub
        
        .STOCKlist.MousePointer = Screen.MousePointer
        tabindex = 1
        .commodityLABEL = StockNumber
        Screen.MousePointer = 11
        
        If IsMissing(serialNum) Then
            serialLabel = ""
        Else
            serialLabel = serialNum
        End If

        If IsMissing(stockListRow) Then
            .poItemLabel = ""
        Else
            .poItemLabel = stockListRow
        End If

        .unitLABEL(0) = unit
        .unitLABEL(1) = ""
        .descriptionLABEL = description
        If StockNumber + description + unit = "" Then
            Call fabCleanDETAILS
            Screen.MousePointer = 0
            frmFabrication.STOCKlist.MousePointer = Screen.MousePointer
            'ctt.enable (False)
            Exit Sub
        End If
 
       isFirstSubmit = True
       doCalculations = False
       directCLICK = True

        .cell(5).locked = False
        .cell(5) = .commodityLABEL
        .cell(5).tag = .cell(5)
        .unitLABEL(1) = .unitLABEL(0)



        If .newBUTTON.Enabled Then
            sql = "SELECT * FROM StockInfoReceptions WHERE " _
                & "NameSpace = '" + nameSP + "' AND " _
                & "Transaction# = '" + .cell(0).tag + "' AND " _
                & "Stocknumber = '" + .commodityLABEL + "' " _
                & "ORDER BY OriginalCondition, LogicName, SubLocaName"
        Else
            If Not IsNull(StockNumber) Then
                If StockNumber <> "" Then
                    sNumber = StockNumber
                    'Juan 2010-9-4 implementing ratio rather than computer factor
                    'computerFactorValue = ImsDataX.ComputingFactor(nameSP, sNumber, cn)
                    Set datax = getDATA("getStockRatio", Array(nameSP, sNumber, .cell(2).tag))
                    If datax.RecordCount > 0 Then
                        'Juan 2014-8-26 new ratio valuation
                        'ratioValue = datax!stk_ratio2
                        ratioValue = datax!realratio
                    Else
                        ratioValue = 1
                    End If
                    stock = ""
                End If
            End If
        
            sql = "SELECT  * FROM StockInfoQTYST4 WHERE " _
                & "NameSpace = '" + nameSP + "' AND " _
                & "Company = '" + .cell(1).tag + "' AND " _
                & "Warehouse = '" + .cell(2).tag + "' AND " _
                & "StockNumber = '" + .commodityLABEL + "' " _
                & "ORDER BY Condition, LogicName, SubLocaName"
            sqlKey = sql
        End If
    End With
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        Call fabCleanDETAILS
    Else
        ReDim qtyArray(datax.RecordCount)
        ReDim subLocationArray(datax.RecordCount)

                Call fillDetailListFabrication(datax)
                Exit Sub

        datax.MoveLast
        Call fabWorkBOXESlistClean
        datax.MoveFirst
        total = CDbl(0)
        With frmFabrication.Tree
            .width = frmFabrication.detailHEADER.width
            'ctt.enable (False)
            .Nodes.Clear
            moreSerial = False
            Dim r As Integer
            r = 0
            Do While Not datax.EOF
                currentStock = IIf(IsNull(datax!StockNumber), "", Trim(datax!StockNumber))
                If cond <> currentCOND Then
                    moreSerial = False
                    If frmFabrication.newBUTTON.Enabled Then
                        cond = Trim(datax!originalcondition)
                        condName = Trim(datax!OriginalConditionName)
                    Else
                        cond = Trim(datax!condition)
                        condName = Trim(datax!conditionName)
                    End If
                    loca = ""
                    subloca = ""

                    .Nodes.Add , tvwChild, "@" + cond, "Condition " + cond + " - " + condName, "thing"
                    .Nodes("@" + cond).Bold = True
                    .Nodes("@" + cond).backcolor = &HE0E0E0

                End If
                Err.Clear
                If frmFabrication.newBUTTON.Enabled Then
                    currentLOGIC = IIf(IsNull(datax!fromlogic), "", Trim(datax!fromlogic))
                    currentSUBloca = IIf(IsNull(datax!fromSubLoca), "", Trim(datax!fromSubLoca))
                    logicname = IIf(IsNull(datax!logicname), "", datax!logicname)
                    sublocaname = IIf(IsNull(datax!sublocaname), "", datax!sublocaname)
                Else

                End If
                'Juan 2014-03-13, to get a real logicname
                 sublocaname = getSUBLOCATIONdescription(currentSUBloca)
                 logicname = gettLOGICdescription(currentLOGIC)
                 '----------------------------------------------

                .Nodes.Add , tvwChild, "@" + currentStock, currentStock, "thing 1"
                frmFabrication.invoiceLabel.Visible = flse
                frmFabrication.invoiceLineLabel.Visible = False
                frmFabrication.invoiceNumberLabel.Visible = False
                frmFabrication.commodityLABEL.Visible = False
                frmFabrication.descriptionLABEL.Visible = False
                frmFabrication.Label4(0).Visible = False
                frmFabrication.Label4(1).Visible = False
                frmFabrication.otherLABEL(0).Visible = False
                frmFabrication.otherLABEL(1).Visible = False
                frmFabrication.unitLABEL(0).Visible = False
                Call fabSetupBOXES(.Nodes.Count, datax.Fields, False)

                datax.MoveNext
            Loop

            .Nodes.Add , , "Total", Space(53) + IIf(frmFabrication.newBUTTON.Enabled, Space(24), "Total Available:")

            .Nodes("Total").Bold = True
            .Nodes("Total").backcolor = &HC0C0C0
            originalQty = total
            Call bottomLine(totalNode, total, pool, StockNumber, False, lastLine, ctt)
        End With
        'Juan 2014-02-28, horizontal line stuff
        With frmFabrication
            .linesH(0).Height = 240
            .linesH(0).Top = .quantityBOX(totalNode).Top
            .linesH(0).Visible = True
        End With
    End If
    '--------------------------------------------------
    frmFabrication.baseFrame.Visible = True
    frmFabrication.treeFrame.Top = 0
    directCLICK = False
    Screen.MousePointer = 0
    frmFabrication.MousePointer = 0
    frmFabrication.STOCKlist.MousePointer = Screen.MousePointer
    Exit Sub
    
ErrHandler:
If Err.Number > 0 Then
    'MsgBox Err.description
    Err.Clear
End If
Resume Next
End Sub


Sub fabDoCombo(Index, datax As ADODB.Recordset, list, totalwidth)
Dim rec, i, extraW
Dim t As String
    Err.Clear
    With frmFabrication.combo(Index)
        Do While Not datax.EOF
            rec = ""
            For i = 0 To frmFabrication.matrix.TextMatrix(1, Index) - 1
                If list(i) = "error" Then
                    MsgBox "Definition error, please contact IMS"
                    Exit Sub
                Else
                    If datax(list(i)).Type = adDate Or datax(list(i)).Type = adDBDate Or datax(list(i)).Type = adDBTime Or datax(list(i)).Type = adDBTimeStamp Then
                        t = Format(datax(list(i)), "yyyy-mm-dd")
                    Else
                        t = IIf(IsNull(datax(list(i))), "", datax(list(i)))
                    End If
                    rec = rec + t
                    If i < (datax.Fields.Count - 1) Then
                        rec = rec + vbTab
                    End If
                End If
            Next
            .addITEM rec
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
            If (350 * .Rows) > 4680 Then
                .Height = 4680
            Else
                .Height = 350 * .Rows
            End If
            .ScrollBars = flexScrollBarVertical
        End If
        If frmFabrication.cell(Index).width > (totalwidth + extraW) Then
            .width = frmFabrication.cell(Index).width
            .ColWidth(0) = .ColWidth(0) + (.width - totalwidth) - extraW
        Else
            .width = totalwidth + extraW
        End If
        If (frmFabrication.cell(Index).Left + .width) > frmFabrication.width Then
            .Left = frmFabrication.width - .width - 100
        Else
            .Left = frmFabrication.cell(Index).Left
        End If
        .RowHeightMin = 240
    End With
End Sub

Function fabInvtReceipt_Insert2a(NameSpace As String, PONumb As String, TranType As String, Companycode As String, Warehouse As String, user As String, cn As ADODB.Connection, Optional ManufacturerNumb As String, Optional TranFrom As String, Optional TransNum As String) As Integer

Dim v As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceipt_Insert"
    
        If Len(TransNum) = 0 Then _
         Err.Raise 1000, "Transaction Number missing" 'TansNum = GetTransNumb(NameSpace, cn)
        
        If Len(Trim$(NameSpace)) = 0 Then Err.Raise 5000, "Namespace is empty"
        
        .parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .parameters.Append .CreateParameter("@COMPANYCODE", adChar, adParamInput, 10, RTrim$(Companycode))
        .parameters.Append .CreateParameter("@WHAREHOUSE", adChar, adParamInput, 10, RTrim$(Warehouse))
        
        .parameters.Append .CreateParameter("@TRANS", adVarChar, adParamInput, 15, RTrim$(TransNum))
        .parameters.Append .CreateParameter("@TRANTYPE", adChar, adParamInput, 2, RTrim$(TranType))
        
        v = RTrim$(TranFrom)
        If Len(Trim$(TranFrom)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@TRANFROM", adVarChar, adParamInput, 10, v)
        
        v = RTrim$(ManufacturerNumb)
        If Len(Trim$(ManufacturerNumb)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 10, v)
        
        v = RTrim$(PONumb)
        If Len(Trim$(PONumb)) = 0 Then v = Null
        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, v)
        
        .parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, user)
        
        Call .Execute(Options:=adExecuteNoRecords)
        fabInvtReceipt_Insert2a = .parameters("RV") = 0
    End With
        
    If fabInvtReceipt_Insert2a Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function
Sub fabColorCOLS()
Dim i As Integer
    With frmFabrication.STOCKlist
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

Sub fabDifferences(row As Integer)
Dim d1, d2 As Double
Dim s1, s2 As String
Dim col, currentROW As Integer

With frmFabrication.STOCKlist
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

Sub fabFillSTOCKlist(datax As ADODB.Recordset)
On errror GoTo errorHandler
Dim n, rec, i, qty2Value, lineNumber
Dim firstTime As Boolean
'Dim mainItemRow, mainItemToReceive
stockReference = ""
firstTime = True
lineNumber = 0
onDetailListInProcess = True

    With datax
        n = 0
        'Juan 2010-5-21
        'frmFabrication.STOCKlist.Rows = .RecordCount + 1
        frmFabrication.STOCKlist.Rows = 2
        frmFabrication.STOCKlist.row = 1
        frmFabrication.STOCKlist.col = 0
        frmFabrication.STOCKlist.CellFontName = "MS Sans Serif"
        mainItemRow = 0

        Do While Not .EOF
            Select Case frmFabrication.tag
                'Fabrication
                Case "02040800"
                    n = n + 1
                    rec = Format(n) + vbTab
                    rec = rec + Trim(!StockNumber) + vbTab
                    rec = rec + IIf(IsNull(!unitPRICE), "0.00", Format(!unitPRICE, "0.00")) + vbTab
                    rec = rec + IIf(IsNull(!description), "", !description) + vbTab
                    rec = rec + IIf(IsNull(!unit), "", !unit) + vbTab
                    'Juan 2010-6-5
                    'rec = rec + Format(!qty) + vbTab
                    rec = rec + Format(!qty, "0.00") + vbTab
                    rec = rec + Format(!qty, "0.00")
            End Select
            frmFabrication.STOCKlist.addITEM rec
            If n = 20 Then
                DoEvents
                frmFabrication.STOCKlist.Refresh
            End If
            .MoveNext
        Loop
        'Call calculateMainItem(stockReference)
        If frmFabrication.STOCKlist.Rows > 2 Then frmFabrication.STOCKlist.RemoveItem (1)
        frmFabrication.STOCKlist.RowHeightMin = 240
        frmFabrication.STOCKlist.row = 0
        If frmFabrication.STOCKlist.topROW = 0 Then ' uh oh needs fixing .
            If frmFabrication.STOCKlist.Rows > 1 Then
                frmFabrication.STOCKlist.FixedRows = 0
                frmFabrication.STOCKlist.FixedRows = 1
                frmFabrication.STOCKlist.topROW = 1
            End If
        End If
    End With
    
errorHandler:
If Err.Number > 0 Then
    'MsgBox "fabFillSTOCKlist " + Err.description
    Err.Clear
    Resume Next
End If
End Sub

Function fabGetLOCATIONdescription(Location) As String
Dim sql
Dim datax As New ADODB.Recordset
    With frmFabrication
        sql = "SELECT loc_name FROM LOCATION WHERE " _
            & "loc_npecode = '" + nameSP + "' AND " _
            & "loc_compcode = '" + .cell(1).tag + "' AND " _
            & "loc_locacode = '" + Location + "'"
        Set datax = New ADODB.Recordset
        datax.Open sql, cn, adOpenForwardOnly
        If datax.RecordCount > 0 Then
            fabGetLOCATIONdescription = datax!loc_name
        Else
            fabGetLOCATIONdescription = ""
        End If
    End With
End Function

Function fabIsOPEN(PO As String) As Boolean
Dim sql As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    With frmFabrication
        fabIsOPEN = False
        PO = Trim(.cell(0))
        sql = "SELECT po_ponumb, po_stas from PO WHERE po_npecode = '" + nameSP + "' " _
            & "AND po_ponumb = '" + .cell(0) + "'"
        Set dataPO = New ADODB.Recordset
        dataPO.Open sql, cn, adOpenForwardOnly
        If Err.Number <> 0 Then Exit Function
        If dataPO.RecordCount > 0 Then
            If dataPO!po_stas = "OP" Then
                fabIsOPEN = True
            Else
                fabIsOPEN = False
            End If
        Else
            fabIsOPEN = False
        End If
    End With
End Function

Sub fabMarkROW(grid As MSHFlexGrid, Optional editing As Boolean, Optional ctt As cTreeTips)
On Error Resume Next
Dim nextROW, purchaseUNIT As String
Dim i  As Integer
Dim stock
Screen.MousePointer = 11
frmFabrication.Refresh
submitted = False
    With grid
        .col = 0
        If frmFabrication.tag <> "02040800" Then 'Fabrication
            Call fabCleanDETAILS
        End If
        Dim currentformname, currentformname1
        Dim imsLock As imsLock.Lock
        Dim ListOfPrimaryControls() As String
        Set imsLock = New imsLock.Lock
        stock = .TextMatrix(.row, 1)
        currentformname = frmFabrication.tag + "stock"
        currentformname1 = currentformname

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
            inProgress = True 'Juan 2010-7-23
            If .text = "Æ" Then
                'Juan 2010-5-25
                .col = 3
                If IsNumeric(.text) Then
                    If Val(.text) <= 0 Then
                        Exit Sub
                    End If
                End If
                .col = 0
                '------
            Else
                If frmFabrication.STOCKlist.TextMatrix(frmFabrication.STOCKlist.row, 1) = "" Then
                Else
                    previousItemMark = .text
                    .CellFontName = "Wingdings 3"
                    .CellFontSize = 10
                    .text = "Æ"
                End If
            End If

            If frmFabrication.STOCKlist.TextMatrix(frmFabrication.STOCKlist.row, 1) = "" Then
            Else
                If frmFabrication.many(0).Value Then
                    frmFabrication.oneStock = ""
                    frmFabrication.addFinalStock.Caption = "&Add Final Stock #"
                Else
                    frmFabrication.STOCKlist.Enabled = False
                    frmFabrication.oneStock = .TextMatrix(.row, 1) + "-- " + .TextMatrix(.row, 3)
                    frmFabrication.oneStock.Top = addFinalStock.Top
                    frmFabrication.oneStock.Visible = True
                    frmFabrication.addFinalStock.Caption = "&Add New Stock #"
                    firstNewMultipleNode = True
                End If
                Call fabPREdetails(ctt)
            End If
            
    End With
    For i = 0 To 2
        frmFabrication.grid(i).Visible = False
    Next
    
Screen.MousePointer = 0
End Sub
Sub fabSHOWdetails()
    With frmFabrication
        Call fabWorkBOXESlist
        .otherLABEL(0).Visible = True
        .commodityLABEL.Visible = True
        .descriptionLABEL.Visible = True
        .remarksLABEL.Visible = False
        .remarks.Visible = False
        .SUMMARYlist.Visible = False
        .hideDETAIL.Visible = True
        .submitDETAIL.Visible = True
        'juan 2012-1-8 commented the line until edition mode works well
        ' .removeDETAIL.Visible = True
        .Label4(0).Visible = True
        .Label4(1).Visible = True
        .hideDETAIL.Visible = True
        .submitDETAIL.Visible = True
    End With
End Sub
Sub fabPutBOX(box As textBOX, Left, Top, width, backcolor)
    With box
        .Left = Left
        .width = width
        .Top = Top
        .Height = 180
        .ZOrder
        .Visible = True
        .backcolor = backcolor
    End With
End Sub

Function fabTopNODE(Index) As Integer
Dim heightFactor, spaceFactor, firstSpacer As Integer
    heightFactor = 240
    spaceFactor = 0
    If Index = 2 Then
        firstSpacer = 40
    Else
        firstSpacer = 0
    End If
    fabTopNODE = frmFabrication.Tree.Top + spaceFactor + (heightFactor * (Index - nodeONtop - 1)) + firstSpacer
End Function

Sub fabUnlockBUNCH()
    Dim imsLock As imsLock.Lock
    Set imsLock = New imsLock.Lock
    Dim grid1 As Boolean
    Dim grid2 As Boolean
    grid2 = True
    grid1 = False
    Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, grid1, "STOCKMASTER", , grid2)
End Sub

Sub fabUnmarkAllRows(grid As MSHFlexGrid)
Dim i  As Integer
Dim stock
Dim imsLock As imsLock.Lock

Screen.MousePointer = 11
frmFabrication.Refresh
    With grid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "" Then
            Else
                stock = .TextMatrix(.row, 1)
                'Unlock
                Set imsLock = New imsLock.Lock
                Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, True, "STOCKMASTER", stock, False)
                Set imsLock = Nothing
                '------
                .TextMatrix(i, 0) = ""
            End If
        Next
    End With
Screen.MousePointer = 0
End Sub

Sub fabUnMarkROW(stock, Optional unmarkIt As Boolean, Optional ctt As cTreeTips)
    'Juan 2010-7-4
    Dim imsLock As imsLock.Lock
    Dim tempMove As Boolean
    Dim commodity
    tempMove = False
    Dim row As Integer
    With frmFabrication.STOCKlist
        For row = 1 To .Rows - 1
            Dim markChar
            markChar = .TextMatrix(row, 0)
            If markChar = "?" Or markChar = "Æ" Then
                .col = 0
                .row = row
                .CellFontName = "MS Sans Serif"
                .CellFontSize = 8.5
                .TextMatrix(row, 0) = Format(row)
            End If
        Next
    End With

    Call fabFillDETAILlist("", "", "", , , , , ctt)
    
    'Unlock
    Set imsLock = New imsLock.Lock
    Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, True, "STOCKMASTER", stock, False)
    Set imsLock = Nothing
    '------
End Sub

Sub fabUpdateStockListBalance() 'Juan 2010-9-19 to re-load the proper values of the stocklist, specially after have removed a row
    Dim i, ii As Integer
    Dim StockNumber As String
    Dim balance, qtySummaryList, qtyStockLIst As Double
    With frmFabrication
        'Reloading original values for the column 3 & 9(qty to receive)
        For i = 1 To .STOCKlist.Rows - 1
            .STOCKlist.TextMatrix(i, 3) = .STOCKlist.TextMatrix(i, 9)
            .STOCKlist.TextMatrix(i, 5) = .STOCKlist.TextMatrix(i, 10)
        Next
        'Taking one by one each line on summaryList to update stockList col 3
        For i = 1 To .SUMMARYlist.Rows - 1
            StockNumber = .SUMMARYlist.TextMatrix(i, 1)
            If StockNumber <> "" Then
                qtySummaryList = CDbl(.SUMMARYlist.TextMatrix(i, 7))
                'Localizing and updating the corresponding row based on its stock or commodity number
                For ii = 1 To .STOCKlist.Rows - 1
                    If StockNumber = .STOCKlist.TextMatrix(ii, 1) Then
                        qtyStockLIst = CDbl(.STOCKlist.TextMatrix(ii, 3))
                        balance = qtyStockLIst - qtySummaryList
                        .STOCKlist.TextMatrix(ii, 3) = Format(balance, "0.00")
                        Exit For
                    End If
                Next
            End If
        Next
    End With
End Sub

Function fabPutReturnData(prefix As String) As Boolean
Dim NP As String
Dim WH As String
Dim From As String
On Error GoTo errfabPutReturnData
    With frmFabrication
        fabPutReturnData = False
        NP = nameSP
        Transnumb = prefix + "-" & GetTransNumb(NP, cn)
        WH = .cell(3).tag
        From = .cell(2).tag
        fabPutReturnData = fabInvtReceipt_Insert2a(NP, "", prefix, .cell(1).tag, WH, Format(CurrentUser), cn, , From, Format(Transnumb))
        Exit Function
    End With
errfabPutReturnData:
    MsgBox Err.description: Err.Clear
End Function
Function fabPutReturnData2() As Boolean
Dim NP As String
Dim WH As String
Dim cmd As Command
Dim From As String
On Error GoTo errfabPutReturnData
    With frmFabrication
        fabPutReturnData2 = False
        'Set cmd = deIms.Commands("InvtIssue_Insert")
        NP = nameSP
        Transnumb = "AE-" & GetTransNumb(NP, cn)
        WH = .cell(2).tag
        From = WH
        fabPutReturnData2 = InvtReceipt_Insert(NP, "", "AE", .cell(1).tag, WH, Format(CurrentUser), cn, , From, Format(Transnumb))
        Exit Function
    End With

errfabPutReturnData:
    MsgBox Err.description: Err.Clear
End Function

Function fabSummaryQTYshort(StockNumber, Optional invoice As String) As Integer
fabSummaryQTYshort = 0
    With frmFabrication
        For i = 1 To .SUMMARYlist.Rows - 1
            If Trim(.SUMMARYlist.TextMatrix(i, 1)) = Trim(StockNumber) Then
                If invoice = "" Then
                    fabSummaryQTYshort = 1
                    Exit Function
                Else
                    If Trim(.summaryValues.TextMatrix(i, 2)) = Trim(invoice) Then
                        fabSummaryQTYshort = 1
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Function fabSummaryQTY(StockNumber, conditionCODE, fromlogic, sublocation, serial, node) As Integer
Dim i, condition, key
    With frmFabrication.SUMMARYlist
        For i = 1 To .Rows - 1
            summaryPOSITION = i
            If Trim(.TextMatrix(i, 1)) = Trim(StockNumber) And .TextMatrix(i, 20) = conditionCODE And .TextMatrix(i, 9) = fromlogic And .TextMatrix(i, 10) = sublocation Then
                If IsNull(serial) Or serial = "" Or UCase(serial) = "POOL" Then
                    key = frmFabrication.Tree.Nodes(node).key
                    condition = Mid(key, InStr(key, "-") + 1, InStr(key, "{{") - InStr(key, "-") - 1)
                    If condition = .TextMatrix(i, 3) Then
                        fabSummaryQTY = .TextMatrix(i, 7)
                        Exit Function
                    End If
                Else
                    If .TextMatrix(i, 2) = serial Then
                        fabSummaryQTY = .TextMatrix(i, 7)
                        Exit Function
                    End If
                End If
            End If
        Next
        summaryPOSITION = 0
        fabSummaryQTY = 0
    End With
End Function
Sub fabGetCOLORSrow(grid As MSHFlexGrid, columns)
Dim i, currentCOL As Integer
    currentCOL = STOCKlist.col
    With frmFabrication.grid(0)
        For i = 1 To columns
            .col = i
            colorsROW(i) = .CellBackColor
        Next
        .col = currentCOL
    End With
End Sub
Sub fabSelectROW(grid As MSHFlexGrid, Optional clean As Boolean)
On Error GoTo getOUT
Dim changeCOLORS As Boolean
Dim i, currentCOL, currentROW As Integer
    Screen.MousePointer = 11
    With frmFabrication.grid(0)
        currentCOL = .col
        If IsNumeric(.tag) Then
            If Val(.tag) = .row Then
                changeCOLORS = False
            Else
                currentROW = .row
                .row = Val(.tag)
                If colorsROW(1) <> "" And .row > 0 Then
                    For i = 1 To .cols - 1
                        .col = i
                        .CellBackColor = colorsROW(i)
                    Next
                    .col = currentCOL
                End If
                .row = currentROW
                .tag = currentROW
                Call fabGetCOLORSrow(grid, .cols - 1)
                changeCOLORS = True
            End If
        Else
            .tag = .row
            Call fabGetCOLORSrow(grid, .cols - 1)
            changeCOLORS = True
        End If
        
        For i = 1 To .cols - 1
            .col = i
            If clean Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = &HFFC0C0 'Very Light Blue
            End If
        Next
    End With
    Screen.MousePointer = 0
    Exit Sub
    
getOUT:
    Screen.MousePointer = 0
    Exit Sub
End Sub

Sub fabCalculations(updateStockList As Boolean, Optional isDynamic As Boolean, Optional isPool As Boolean)
Dim this, r, summary, balance, balance2, balanceTotal, col
Dim sumByQtyBox, sumByLine, sumByQty, sumByLines
Dim qtyBoxTotal  As Double
Dim i As Integer
Dim once As Boolean
Dim originalQty As Double
once = True
balanceTotal = 0
Dim goAhead As Boolean
goAhead = False

'On Error GoTo errorHandler
On Error Resume Next
    With frmFabrication
        'Global declarations
        Dim colRef, colRef2, colTot As Integer
        Dim fromStockList As Boolean
        isDynamic = True
        fromStockList = False
        colRef = 2
        colRef2 = 7
        colTot = 5
        Select Case .tag
            Case "02040800" 'Fabrication
                colTot = 5
                colRef = 7
        End Select
        r = .STOCKlist.row
            r = findSTUFF(.commodityLABEL, .STOCKlist, 1)
        
        If r > 0 Then
            'When isDynamic variable is false means we are taking the values from the stockList
            If isDynamic Then
            Else
                If IsNumeric(.STOCKlist.TextMatrix(r, colRef)) Then
                    originalQty = CDbl(.STOCKlist.TextMatrix(r, colRef))
                    If originalQty > 0 Then
                        this = 0
                        'qtyBoxTotal = 0
                        'originalQty = 0
                        balance = originalQty
                    Else
                        Exit Sub
                    End If
                    .STOCKlist.row = r
                    Call fabSelectROW(.STOCKlist)
                End If
            End If
        End If
        
        'Main cycle to scan the active tree nodes-------------
        sumByQty = 0
        sumByQtyBox = 0
        sumByLine = 0
        sumByLines = 0
'        If isPool Then
            For i = 1 To .Tree.Nodes.Count
                If i <> totalNode Then
                    If Err.Number = 0 Then
                        'This allows to get original qty's based on its business logics
                        If isDynamic Then
                            originalQty = .quantity(i)
                            If Err.Number = 0 Then
                                balance = originalQty
                            Else
                                Err.Clear
                            End If
                        End If
                        Dim qBoxExists As Boolean
                        qBoxExists = False
                        If controlExists("quantityBOX", i) Then
                            qBoxExists = True
                        End If
                        If once And qBoxExists Then  'This is to count what is on summaryList
                            once = False
                            Dim subTot, lineQty As Double
                            Dim position As Integer
                            subTot = 0
                            lineQty = 0
                            If controlExists("positionBox", i) Then
                                If IsNumeric(.positionBox(i).text) Then
                                    position = Val(.positionBox(i).text)
                                Else
                                    position = 0
                                End If
                            Else
                                position = 0
                            End If
                            
                            For j = 1 To .SUMMARYlist.Rows - 1
                                'The reason for this select case is to manage if there is difrerences on
                                If frmFabrication.invoiceNumberLabel = "" Then
                                    goAhead = True
                                Else
                                    If frmFabrication.invoiceNumberLabel = .SUMMARYlist.TextMatrix(j, 12) And frmFabrication.invoiceLineLabel = .SUMMARYlist.TextMatrix(j, 13) Then
                                        goAhead = True
                                    End If
                                End If
                                If goAhead Then
                                    If .SUMMARYlist.TextMatrix(j, 1) = .commodityLABEL.Caption Then
                                        If position = j Then 'This leaves the current line without summarizing
                                        Else
                                            If IsNumeric(.SUMMARYlist.TextMatrix(j, colRef2)) Then
                                                lineQty = CDbl(.SUMMARYlist.TextMatrix(j, colRef2))
                                            Else
                                                lineQty = 0
                                            End If
                                            subTot = subTot + lineQty
                                        End If
                                    End If
                                End If
                            Next
                            balance = balance - subTot
                        End If '-----------------------------------
                        'Step to update cells on screen-----------------
                        If qBoxExists Then
                            'new
                            sumByLine = .quantity(i) - .quantityBOX(i)
                            sumByLines = sumByLines + sumByLine
                            '.quantity(i) = Format(sumByLine, "0.00")
                            sumByQty = sumByQty + .quantity(i)
                            sumByQtyBox = sumByQtyBox + .quantityBOX(i)
                            '-------------
                              
                            balance = balance - .quantityBOX(i)

                            .balanceBOX(i) = Format(sumByLine, "0.00")
                            '------------------
                        End If
                    Else
                        Err.Clear
                    End If
                End If
            Next
        submitted = True
        .quantityBOX(totalNode) = Format(sumByQtyBox, "0.00")
        .quantity(totalNode) = Format(sumByQty, "0.00")
        '------------------
        If isDynamic Then
            .balanceBOX(totalNode) = Format(balanceTotal, "0.00")
            .balanceBOX(totalNode) = Format(sumByLines, "0.00")
            balance = balanceTotal
        Else
            .balanceBOX(totalNode) = Format(balance, "0.00")
        End If
        If ratioValue > 1 Then
            balance2 = balance * ratioValue
        Else
            balance2 = balance
        End If
        If updateStockList Then
            Select Case .tag
                Case "02040100" 'WarehouseReceipt
                Case Else
                    .STOCKlist.TextMatrix(r, colTot) = Format(balance, "0.00")
            End Select
            Select Case .tag
                Case "02040100" 'WarehouseReceipt
                Case Else
                    .STOCKlist.TextMatrix(r, colTot) = Format(sumByLines, "0.00") 'new juan 2015-10-3
            End Select

            stockReference = .STOCKlist.TextMatrix(mainItemRow, 1)
            'Juan 2014-07-05 it does calculate the total to be received for the main item
            Call calculateMainItem(stockReference, True) 'r before next
            '----------------------
        End If
    End With
    Exit Sub
errorHandler:
    If Err.Number = 340 Then
    Else
        MsgBox Err.description
        Err.Clear
    End If
    Resume Next
End Sub

Sub calculationsFabrication(updateStockList As Boolean, i As Integer)
Dim this, r, summary, balance, balance2, balanceTotal, col
Dim sumByQtyBox, sumByLine, sumByQty, sumByLines
Dim qtyBoxTotal  As Double
Dim once As Boolean
Dim originalQty As Double
once = True
balanceTotal = 0
Dim goAhead As Boolean
goAhead = False

On Error Resume Next
    With frmFabrication
        Dim colRef, colRef2, colTot As Integer
        Dim fromStockList As Boolean
        
        fromStockList = False
        colRef = 6
        colTot = 5
        r = .STOCKlist.row
        r = fabFindSTUFF(.commodityLABEL, .STOCKlist, 1)
        
        If r > 0 Then
            If IsNumeric(.STOCKlist.TextMatrix(r, colRef)) Then
                originalQty = CDbl(.STOCKlist.TextMatrix(r, colRef))
                If originalQty > 0 Then
                    this = 0
                    balance = originalQty
                Else
                    Exit Sub
                End If
                .STOCKlist.row = r
                Call fabSelectROW(.STOCKlist)
            End If
        End If
        sumByQty = 0
        sumByQtyBox = 0
        sumByLine = 0
        sumByLines = 0
        If Err.Number = 0 Then
            Dim qBoxExists As Boolean
            qBoxExists = False
            If fabControlExists("quantityBOX", i) Then
                qBoxExists = True
            End If
            'Step to update cells on screen-----------------
            If qBoxExists Then
                If .quantityBOX(i) = 0 Then
                    .quantity(i) = Format(originalQty, "0.00")
                Else
                    key = .Tree.Nodes(i).key
                    If InStr(key, "@newStock") Then
                        .quantity(i) = .quantityBOX(i)
                    Else
                        .quantity(i) = Format(originalQty - CDbl(.quantityBOX(i)), "0.00")
                    End If
                End If
                If CDbl(.quantity(i)) > 0 Then
                    sumByLine = .quantity(i) - .quantityBOX(i)
                Else
                    sumByLine = CDbl(.quantity(i))
                End If
                .balanceBOX(i) = Format(sumByLine, "0.00")
            End If
        Else
            Err.Clear
        End If
        If updateStockList Then
            .STOCKlist.TextMatrix(r, colTot) = .quantity(i)
            stockReference = .STOCKlist.TextMatrix(mainItemRow, 1)
        End If
        Dim finalPrice As Double
        finalPrice = 0
        Dim stockTotalPrice As Double
        stockTotalPrice = 0
        Dim newStocks As Integer
        newStocks = 0
        'summarizing
        For i = 1 To .Tree.Nodes.Count
            key = .Tree.Nodes(i).key
            If InStr(key, "@newStock") Then key = "@newStock"
            Select Case key
                Case "@finalCost"
                    'does nothing but just to not get it into the sumes
                Case "@processCost"
                    finalPrice = finalPrice + CDbl(.fabCostBOX(i))
                Case "@newStock"
                    newStocks = newStocks + CDbl(.quantityBOX(i))
                    stockTotalPrice = stockTotalPrice + CDbl(.priceBOX(i))
                Case Else
                    finalPrice = finalPrice + (CDbl(.priceBOX(i)) * CDbl(.quantityBOX(i)))
            End Select
        Next
        
        
        If .Tree.Nodes(.Tree.Nodes.Count).key = "@newStock" Then
            If IsNumeric(.Tree.Nodes.Count) Then
                If .many(0).Value Then
                    .priceBOX(.Tree.Nodes.Count) = Format(finalPrice / CDbl(.quantityBOX(.Tree.Nodes.Count)), "0.00")
                End If
            Else
                .priceBOX(.Tree.Nodes.Count) = Format(finalPrice, "0.00")
            End If
        End If
        Dim totalCost As Double
        Dim totalCostBalance As Double
        totalCost = finalPrice
        totalCostBalance = finalPrice
        For i = 1 To .Tree.Nodes.Count
            key = .Tree.Nodes(i).key
            If InStr(key, "@newStock") Then key = "@newStock"
            If InStr(key, "@finalCost") Then key = "@finalCost"
            Select Case key
                Case "@finalCost"
                    .priceBOX(i) = Format((totalCost), "0.00")
'                    .balanceBOX(i) = Format((totalCost - stockTotalPrice), "0.00")
                Case "@newStock"
                    If newStocks > 0 Then
                        If .many(2).Value = False Then
                            .priceBOX(i) = Format((totalCost / newStocks), "0.00")
                        Else
                            totalCostBalance = totalCostBalance - (CDbl(.priceBOX(i)) * .quantityBOX(i))
                            .balanceBOX(i) = Format((totalCostBalance), "0.00")
                            If (CDbl(.balanceBOX(i))) < 0 Then
                                .balanceBOX(i).ForeColor = vbRed
                            Else
                                .balanceBOX(i).ForeColor = vbBlack
                            End If
                        End If
                    End If
                Case Else
                    If .Tree.Nodes(key).Parent.key = "Fabrication" Then
                        Dim baseBalance As Double
                        baseBalance = (CDbl(.priceBOX(i)) * .quantityBOX(i))
                        .balanceBOX(i) = Format((baseBalance), "0.00")
                        If (CDbl(.balanceBOX(i))) < 0 Then
                            .balanceBOX(i).ForeColor = vbRed
                        Else
                            .balanceBOX(i).ForeColor = vbBlack
                        End If
                    End If
            End Select
        Next
    End With
    Exit Sub
errorHandler:
    If Err.Number = 340 Then
    Else
        MsgBox Err.description
        Err.Clear
    End If
    Resume Next
End Sub


Function fabFindSTUFF(toFIND, grid As MSHFlexGrid, col, Optional toFIND2, Optional col2 As Integer) As Integer
Dim i
Dim invoice
invoice = frmFabrication.invoiceNumberLabel
LineItem = frmFabrication.invoiceLineLabel
Dim findIT As Boolean
    fabFindSTUFF = 0
    mainItemRow = 0
    With grid
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
                    If frmFabrication.invoiceNumberLabel = "" Then
                    Else
                        If .cols > 11 Then
                            If frmFabrication.invoiceNumberLabel = .TextMatrix(i, 12) And frmFabrication.invoiceLineLabel = .TextMatrix(i, 13) Then
                                
                            End If
                        End If
                    End If
                    If IsMissing(toFIND2) Or IsMissing(col2) Then
                        fabFindSTUFF = i
                        Exit For
                    Else
                         If UCase(Trim(.TextMatrix(i, col2))) = UCase(Trim(toFIND2)) Then
                            If mainItemRow = 0 Then
                                mainItemRow = i
                                fabFindSTUFF = mainItemRow
                            End If
                            If invoice = "" Then
                                fabFindSTUFF = i
                                Exit For
                            Else
                                If UCase(Trim(.TextMatrix(i, 12))) = UCase(Trim(invoice)) And Trim(.TextMatrix(i, 13)) = Trim(LineItem) Then 'Juan 2014-8-30 added linitem
                                    fabFindSTUFF = i
                                    Exit For
                                End If
                            End If
                         Else
                         End If
                    End If
                End If
            Next
        End If
    End With
End Function





Sub fabCleanDETAILS()
Dim i
On Error Resume Next
    With frmFabrication
        nodeONtop = 0
        For i = 1 To 10
            Unload .linesV(i)
            If Err.Number <> 0 Then Err.Clear
        Next
        .cell(5).Visible = False
        .combo(5).Visible = False
        
        .newDESCRIPTION.Visible = False
        .otherLABEL(2).Visible = False
        Call fabWorkBOXESlistClean
        .Tree.Nodes.Clear
    End With
    Err.Clear
End Sub

Sub fabBeforePrint()
Set translatorFORM = imsTranslator
    
    With MDI_IMS.CrystalReport1
        .Reset
        'msg1 = translator.Trans("L00176")
        .WindowTitle = IIf(msg1 = "", "transaction", msg1)
        .ParameterFields(0) = "namespace;" + nameSP + ";TRUE"
        If frmFabrication.cell(1) = "" Then
            '*******************
            'CHECK THIS PATH
            .ReportFileName = App.path + "CRreports\transactionGlobal.rpt"
            .ParameterFields(1) = "ponumb;" + frmFabrication.cell(0) + ";TRUE"
            'call translator.Translate_Reports("transactionGlobal.rpt")
        Else
            '*******************
            'CHECK THIS PATH
            .ReportFileName = App.path + "CRreports\transaction.rpt"
            .ParameterFields(1) = "invnumb;" + frmFabrication.cell(1) + ";TRUE"
            .ParameterFields(2) = "ponumb;" + frmFabrication.cell(0) + ";TRUE"
            'Call translator.Translate_Reports("transaction.rpt")
            'Call translator.Translate_SubReports
        End If
    End With
End Sub
Sub fabPREdetails(ctt As cTreeTips)
Screen.MousePointer = 11
    frmFabrication.Refresh
    With frmFabrication.STOCKlist
        Call fabFillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 3), .TextMatrix(.row, 4), .row, , , , ctt)
    End With
Screen.MousePointer = 0
End Sub



