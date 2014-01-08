Attribute VB_Name = "warehouses"
 Public Enum FormMode
    mdNa = 0
    mdCreationho
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
Global totalNode As Integer
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
Global computerFactorValue As Double
Global ratioValue As Integer
Global qtyArray() As Double
Global subLocationArray() As String
Global latestStockNumberQty As String
Global isFirstSubmit As Boolean
Global sqlKey As String
Global uid As String
Global pwd As String
Global InitCatalog As String
Global dsnName As String
Global emailOutFolder As String
Global skipAlphaSearch As Boolean
Global skipExistance As Boolean
Global originalQty

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Public Function generateattachmentswithCR11(fileName As String, reportCaption As String, ParamsForCrystalReport() As String, reportName As String, path As String) As String()
Dim Attachments(0) As String
Dim IFile As IMSFile
Dim file As String
Dim i As Integer
Set IFile = New IMSFile

On Error GoTo errMESSAGE
    Attachments(0) = reportName & "-" & nameSP & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".Pdf"
    file = cEmailOutFolder & Attachments(0)
    Dim x As New clsexport
    x.ExportFilePath = emailOutFolder + file
    x.reportName = fileName
    If IFile.FileExists(file) Then IFile.DeleteFile (file)
    Attachments(0) = emailOutFolder + file
    
    Call x.GeneratePdf(ParamsForCrystalReport, emailOutFolder)
    generateattachmentswithCR11 = Attachments
Exit Function

errMESSAGE:
    If Err.Number <> 0 Then
        MsgBox "Process generateattachments " + Err.description
    End If
End Function

Public Function GeneratePdf(ParamsForCrystalReport() As String) As String
Dim Report As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSubreport As CRAXDRT.Report
Dim Param As CRAXDRT.ParameterFieldDefinition
Dim arrparam() As String
On Error GoTo ErrHandler
    Set crxApplication = New CRAXDRT.Application
    Set Report = crxApplication.OpenReport(reportPATH + reportName, 1)
    Set Report = InitializeReport(Report, ParamsForCrystalReport())
    Call Export(Report)
Exit Function
            
ErrHandler:
    GeneratePdf = "Errors Occurred while trying to generate a PDF, please try again." + Err.description
Err.Clear
End Function

Public Sub LogErr(RoutineName As String, ErrorDescription As String, ErrorNumber As Long, Optional Clear As Boolean = False)
Dim i As IMSFile
Dim ms As imsmisc
Dim fileName As String
Dim FileNumb As Integer
On Error Resume Next
    If Len(Trim$(ErrorDescription)) = 0 Then Exit Sub
    Set i = New IMSFile
    Set ms = New imsmisc
    If Not i.DirectoryExists(LogPath) Then Call MkDir(LogPath)
    FileNumb = FreeFile
    fileName = LogPath + i.ChangeFileExt(App.EXEName + Format$(Date, "ddmmyy"), "imserrlog")
    Open fileName For Append As 1
        Print #FileNumb, "Module:             " & App.EXEName
        Print #FileNumb, "Routine:            " & RoutineName
        Print #FileNumb, "Error Number:       " & ErrorNumber
        Print #FileNumb, "Error Source:       " & Err.source
        Print #FileNumb, "Error Description:  " & ErrorDescription
        Print #FileNumb, "Error Date:         " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        Print #FileNumb, "": Print #FileNumb, ""
    Close #FileNumb
    Set i = Nothing
    Set ms = Nothing
    If Err Then Err.Clear
End Sub


Public Function InitializeReport(Report As CRAXDRT.Report, ParamsForCrystalReport() As String) As CRAXDRT.Report
Dim crxSubreport As CRAXDRT.Report
Dim arrparam() As String
On Error GoTo ErrHand

        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
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

If reportName = Report_EmailFax_PO_name Then

             Call FixDB(Report.Database.Tables)
            
            '‘Set crxSubreport to the subreport ‘Sub1’ of the main report. The subreport name needs to be known to use this ‘method.
            Set crxSubreport = Report.OpenSubreport("porem.rpt")
            
            
            Call FixDB(crxSubreport.Database.Tables)
            
            Set crxSubreport = Report.OpenSubreport("poclause.rpt")
            Call FixDB(crxSubreport.Database.Tables)
            
             arrparam = Split(ParamsForCrystalReport(1), ";")
            Report.ParameterFields.Item(1).AddCurrentValue nameSP
            Report.ParameterFields.Item(2).AddCurrentValue arrparam(1)
End If

Set InitializeReport = Report
Exit Function
ErrHand:
'Call LogErr("InitializeReport ", Err.description, Err.Number)
MsgBox "InitializeReport function : " + Err.description
Err.Clear
End Function
        

Private Function FixDB(crxDatabaseTableS As CRAXDRT.DatabaseTables)
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
For Each crxDatabaseTable In crxDatabaseTableS
    crxDatabaseTable.SetLogOnInfo ConnInfo.dsnName, ConnInfo.InitCatalog, ConnInfo.uid, ConnInfo.pwd    ' "imsO", "pecten_dev", "sa", "scms"
    crxDatabaseTable.Location = crxDatabaseTable.name
Next crxDatabaseTable
End Function
        

Public Function WriteParameterFiles(Recepients As String, sender As String, Attachments() As String, subject As String, attention As String)
Dim l
Dim x
Dim Y
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
    Email = frmWarehouse.emailRecepient.text
    If Not Email = "" Then
        Call WriteParameterFileEmail(Attachments, Email, subject, sender, attention)
    End If
errMESSAGE:
    If Err.Number <> 0 And Err.Number <> 9 Then
        MsgBox "Process WriteParameterFiles " + Err.description
    Else
        Err.Clear
    End If
End Function

Public Function WriteParameterFileEmail(Attachments() As String, Recipients As String, subject As String, sender As String, attention As String) As Integer
On Error GoTo errMESSAGE
    Dim fileName As String
    Dim FileNumb As Integer
    Dim i As Integer, l As Integer
    Dim reports As String
    Dim recepientSTR As String
    i = 0
    If UBound(Attachments) > 0 Then
        For i = 0 To UBound(Attachments)
                 reports = reports & Trim$(Attachments(i) & ";")
        Next
    ElseIf UBound(Attachments) = 0 Then
        reports = reports & Trim$(Attachments(i))
    End If

    If Len(Recipients) > 0 Then
        Call sendProcess(Recipients, reports, subject, attention)
    End If
    Recepients = ""
    reports = ""
    WriteParameterFileEmail = 1
Exit Function
errMESSAGE:
    If Err.Number <> 0 Then
        MsgBox Err.description
    End If
End Function

Public Function IsArrayLoaded(ArrayToTest() As String) As Boolean
Dim x As Integer
On Error GoTo ErrHandler
    IsArrayLoaded = False
    x = UBound(ArrayToTest)
    IsArrayLoaded = True
    Exit Function
ErrHandler:
Err.Clear
End Function
Public Sub sendProcess(recipientList As String, Attachments As String, subject As String, messageText As String)
'Save the Email/ request to the Database
On Error GoTo errorHandler
    Dim strOut As String
    Dim programName As String
    Dim parameters As String
    Dim cmd As ADODB.Command
    Set cmd = MakeCommand(cn, ADODB.CommandTypeEnum.adCmdStoredProc)
    With cmd
        .CommandText = "InsertEmailFax"
        .parameters.Append .CreateParameter("@Subject", adVarChar, adParamInput, 4000, subject)
        .parameters.Append .CreateParameter("@Body", adVarChar, adParamInput, 8000, messageText)
        .parameters.Append .CreateParameter("@AttachmentFile", adVarChar, adParamInput, 2000, Attachments)
        .parameters.Append .CreateParameter("@recepientStr", adVarChar, adParamInput, 8000, recipientList)
        .parameters.Append .CreateParameter("@creauser", adVarChar, adParamInput, 100, CurrentUser)
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    Set cmd = Nothing
    LogExec ("Successfully saved email\ Fax request with Subject " & subject & " to the Database.")
Exit Sub

errorHandler:
    Call LogErr("sendProcess", "Érror Occured while trying to save Email request to the DB for Subject " + subject + " Body " + messageText + " Attachment " + Attachments + " Recepient List " + recipientList + ". " + Err.description, Err.Number, False)
    MsgBox "Errors Occured while trying to generate email request. Please dont send any more emails and faxes and call the Administrator. " + Err.description
    Err.Clear
    
End Sub



Private Sub Export(Report As CRAXDRT.Report)
    Report.ExportOptions.FormatType = crEFTPortableDocFormat
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.DiskFileName = ExportFilePath
    Report.Export False
End Sub
Sub calculationsFlat(Optional selectedStockNumber As String)
'This is an alternate calculations procedure to recalculate after hiding the edition tree
Dim originalQTY1(), originalQTY2()
Dim balance1(), balance2() As Double
Dim i, j As Integer
Dim StockNumber As String

On Error GoTo errorHandler
    With frmWarehouse
        'Global declarations
        Dim colRef, colRef2, colTot As Integer
        colRef = 5 'stocklist qty column
        colRef2 = 7 'summaryList qty column
        colTot = 5
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                colRef = 9
                colTot = 3
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        'Get initial values
        ReDim originalQTY1(.STOCKlist.Rows)
        ReDim originalQTY2(.STOCKlist.Rows)
        ReDim balance1(UBound(originalQTY1))
        ReDim balance2(UBound(originalQTY2))
        For i = 1 To .STOCKlist.Rows - 1
            originalQTY1(i) = .STOCKlist.TextMatrix(i, colRef)
            balance1(i) = CDbl(originalQTY1(i))
            If .tag = "02040100" Then 'Receipt
                originalQTY2(i) = .STOCKlist.TextMatrix(i, colRef + 1)
                balance2(i) = CDbl(originalQTY2(i))
            Else
                originalQTY2(i) = originalQTY1(i)
                balance2(i) = balance1(i)
            End If
        Next

        For i = 1 To .STOCKlist.Rows - 1
            StockNumber = .STOCKlist.TextMatrix(i, 1)
            For j = 1 To .SUMMARYlist.Rows - 1
                'Look for possible movements within summary list
'                If StockNumber = .SUMMARYlist.TextMatrix(j, 1) Then
'                    balance1(i) = balance1(i) - .SUMMARYlist.TextMatrix(j, colRef2)
'                    If .tag = "02040100" Then 'Receipt
'                        balance2(i) = balance2(i) - .SUMMARYlist.TextMatrix(j, colRef2)
'                    Else
'                        balance2(i) = balance1(i)
'                    End If
'                Else
                If Not IsMissing(selectedStockNumber) Then
                    If StockNumber = selectedStockNumber Then
                        If IsNumeric(.Tree.Nodes.Count) Then
                            balance1(i) = originalQty
                            balance2(i) = balance1(i)
                        End If
                    End If
                End If
'                End If
' juan 2012-1-17 commented to fix bug
'                'Write final values on stocklist
                .STOCKlist.TextMatrix(i, colTot) = Format(balance1(i), "0.00")
                If .tag = "02040100" Then 'Receipt
                    .STOCKlist.TextMatrix(i, colTot + 2) = Format(balance2(i), "0.00")
                Else
                End If
                
            Next
        Next
    End With
    Exit Sub
errorHandler:
    'MsgBox Err.description
    Err.Clear
    Resume Next
End Sub


Sub updateStockListStatus()
'This checks and updates each line on stockList depending if there is a corresponding value on the summary list
Dim i, j As Integer
Dim StockNumber As String
Dim hasMark As Boolean
Dim imsLock As imsLock.Lock

On Error GoTo errorHandler
    With frmWarehouse
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




Sub bottomLine(totalNode, total, pool As Boolean, StockNumber, doRecalculate As Boolean)
Dim lastLine, thick
On Error Resume Next
'Scrolling stuff
With cTT
    Set .Tree = frmWarehouse.Tree
End With

With frmWarehouse
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
                Call calculations(True)
            Else
                Call calculations(False)
            End If
        Else
            Call calculations2(.SUMMARYlist.row, .Tree.Nodes(.Tree.Nodes.Count - 1), .Tree.Nodes.Count - 1)
        End If
        For i = 1 To totalNode
            .Tree.Nodes(i).Expanded = True
        Next
        If Not .Visible Then
            Call SHOWdetails
        End If
        If Not pool Then
            If doRecalculate Then
                Call recalculate(StockNumber)
            End If
        End If
        .ZOrder
        
        If Not .newBUTTON.Enabled Then .SUMMARYlist.Visible = False
'        Call SHOWdetails
        
        Call lineStuff(lastLine, thick)
        Call workBOXESlist("fix")
        If .Tree.Nodes.Count > 15 Then
            .linesV(lastLine).Visible = False
            .Tree.Nodes(1).EnsureVisible
        End If
End With
End Sub

Function controlExists(controlNAME As String, controlIndex As Integer) As Boolean
controlExists = False
Dim ctl As Control

For Each ctl In frmWarehouse.Controls
    If ctl.name = controlNAME Then
        If ctl.Index = controlIndex Then
            controlExists = True
            Exit For
        End If
    End If
Next
End Function

Sub lineStuff(lastLine, think)
On Error Resume Next
    With frmWarehouse
        n = 0
        For i = 1 To lastLine
            Load .linesV(i)
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
        '.linesV(lastLine).BorderStyle = 0
        '.linesV(lastLine).Appearance = 0
        '.linesV(lastLine).backcolor = &HE0E0E0
        '.linesV(lastLine).width = .detailHEADER.ColWidth(lastLine) + 10
        '.linesV(lastLine).Height = .Height - 60
    End With
End Sub

Sub recalculate(StockNumber) 'Juan 2010-7-26
    Dim totalCount As Integer
    Dim qtyToReceive As Integer
    Dim r As Integer
    With frmWarehouse
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

Sub setupBoxes2(n, row, serial As Boolean, Optional QTYpo)
Dim x, cond, logic, subloca, newCOND, serialPool, StockNumber, unitPRICE, unit, unit2, conditionName, qty, qty2, quantity
serialPool = IIf(serial, "SERIAL", "POOL") 'Juan 2010-5-14
Dim newButtonEnabled As Boolean
On Error GoTo ErrHandler:

    With frmWarehouse
        '1 "Commodity"
        '2 "Serial"
        '3 "Condition"
        '4 "Unit Price"
        '5 "Description"
        '6 "Unit"
        '7 "Qty"
        '8 "node"
        '9 "From Logical"
        '10 "From Subloca"
        '11 "To Logical"
        '12 "To Subloca"
        '13 "New Condition Code"
        '14 "New Condition Description"
        '15 "Unit Code"
        '16 "Computer Factor"
        '20 "Original Condition Code"
        '21 "Secundary QTY"
        '17 "repaircost"
        '18 "newcomodity"
        '19 "newdescription"

        StockNumber = .SUMMARYlist.TextMatrix(row, 1)
        unitPRICE = .SUMMARYlist.TextMatrix(row, 4)
        logic = .SUMMARYlist.TextMatrix(row, 11)
        subloca = .SUMMARYlist.TextMatrix(row, 12)
        cond = .SUMMARYlist.TextMatrix(row, 3)
        newCOND = .SUMMARYlist.TextMatrix(row, 13)
        unit = .SUMMARYlist.TextMatrix(row, 6)
        'Juan 2010-8-19
        'unit2 = .SUMMARYlist.TextMatrix(row, 23)
        'qty2 = .SUMMARYlist.TextMatrix(row, 21)
        unit2 = .SUMMARYlist.TextMatrix(row, 21)
        qty2 = .SUMMARYlist.TextMatrix(row, 23)
        '--------------------------------------------------
        conditionName = .SUMMARYlist.TextMatrix(row, 14)
        qty = .SUMMARYlist.TextMatrix(row, 7)
        
        Load .quantity(n)
        Call putBOX(.quantity(n), .detailHEADER.ColWidth(0) + 140, topNODE(n), .detailHEADER.ColWidth(1) - 40, vbWhite)
        Load .balanceBOX(n)
        .balanceBOX(n) = Format(.quantity(n), "0.00")
        Load .quantityBOX(n)
        .quantityBOX(n).tabindex = tabindex + 2
        Load .quantity2BOX(n)
        .quantity2BOX(n).tabindex = tabindex + 2
        Load .priceBOX(n)
        Load .NEWconditionBOX(n)
        Load .positionBox(n)
        .positionBox(n).text = .SUMMARYlist.row
        
        Load .logicBOX(n)
        .logicBOX(n).tabindex = tabindex
        Load .sublocaBOX(n)
        .sublocaBOX(n).tabindex = tabindex + 1
        .priceBOX(n) = unitPRICE
        .NEWconditionBOX(n).tag = newCOND
        Select Case .tag
            'ReturnFromRepair WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales,ReturnFromWell,AdjustmentEntry
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300", "02050200"
                If serial Then
                    .quantity(n) = 1
                Else
                    .quantity(n) = QTYpo
                End If
                .quantityBOX(n) = qty

            'Case "02050200" 'AdjustmentEntry
'                .quantity(n) = summaryQTY(Trim(datax!stockNumber), Left(.Tree.Nodes(n).key, 2), "", "", "POOL", n)
'                .quantityBOX(n) = .quantity(n)
'                If summaryPOSITION > 0 Then
'                    .priceBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 4)
'                Else
'                    .priceBOX(n) = "0.00"
'                End If
'                .NEWconditionBOX(n).tag = Left(.Tree.Nodes(n).key, 2)
            Case "02040100" 'WarehouseReceipt
                .quantity(n) = Format(QTYpo, "0.00")
                newCOND = "01"
                If serialPool = "SERIAL" Then 'Juan 2010-5-17
                    .quantityBOX(n) = "1.00"
                    .quantity2BOX(n) = "1.00"
                Else
                    .quantityBOX(n) = qty
                    .quantity2BOX(n) = qty2
                End If
                Load .repairBOX(n)
                .repairBOX(n) = poItem
                Load .poItemBox(n)
                'Juan 2010-9-25 This was a bug because it was not giving the correct row from the stocklist so now it comes from the poItemBox directly
                '.poItemBox(n) = frmWarehouse.STOCKlist.TextMatrix(frmWarehouse.STOCKlist.row, 8)
                .poItemBox(n) = .SUMMARYlist.TextMatrix(row, 22)
                '--------

        End Select
        .NEWconditionBOX(n) = .NEWconditionBOX(n).tag
        

        If summaryPOSITION = 0 Then
            .logicBOX(n) = logic
            .sublocaBOX(n) = subloca
        Else
            .logicBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 11)
            .logicBOX(n).tag = .logicBOX(n)
            .sublocaBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 12)
            .sublocaBOX(n).tag = .sublocaBOX(n)
            .logicBOX(n).ToolTipText = getWAREHOUSEdescription(.logicBOX(n))
            .sublocaBOX(n).ToolTipText = getSUBLOCATIONdescription(.sublocaBOX(n))
        End If
        
        Load .unitBOX(n)
        Load .unit2BOX(n)
        .unitBOX(n).Enabled = False
        .unit2BOX(n).Enabled = False
        .unitBOX(n) = unit
        .unit2BOX(n) = unit2
        
        If summaryPOSITION = 0 Then
            .NEWconditionBOX(n).ToolTipText = conditionName
            .NEWconditionBOX(n).tag = newCOND
            .NEWconditionBOX(n) = Format(newCOND, "00")
        Else
            .NEWconditionBOX(n).tag = .SUMMARYlist.TextMatrix(summaryPOSITION, 13)
            .NEWconditionBOX(n) = Format(.NEWconditionBOX(n).tag, "00")
            .NEWconditionBOX(n).ToolTipText = .SUMMARYlist.TextMatrix(summaryPOSITION, 14)
        End If
        
        Select Case .tag
            Case "02040200", "02040500" 'WarehouseIssue, WellToWell
                .logicBOX(n).Enabled = False
                .sublocaBOX(n).Enabled = True
                .grid(2).Visible = False
            Case "02040400" 'ReturnFromRepair
'                Load .repairBOX(n)
'                If summaryPOSITION = 0 Then
'                    If .newBUTTON.Enabled Then
'                        .repairBOX(n) = Format(datax!repairCOST, "0.00")
'                        .cell(5) = Trim(datax!NewStockNumber)
'                        .cell(5).tag = .cell(5)
'                        .unitLABEL(1) = getUNIT(.cell(5).tag)
'                        .newDESCRIPTION = Trim(datax!NewStockDescription)
'                    Else
'                        .repairBOX(n) = "0"
'                    End If
'                Else
'                    If .newBUTTON.Enabled Then
'                        .repairBOX(n) = Format(datax!repairCOST, "0.00")
'                        .cell(5) = Trim(datax!NewStockNumber)
'                        .cell(5).tag = .cell(5)
'                        .unitLABEL(1) = getUNIT(.cell(5).tag)
'                        .newDESCRIPTION = Trim(datax!NewStockDescription)
'                    Else
'                        .repairBOX(n) = SUMMARYlist.TextMatrix(summaryPOSITION, 17)
'                        .cell(5) = SUMMARYlist.TextMatrix(summaryPOSITION, 18)
'                        .cell(5).tag = .cell(5)
'                        .unitLABEL(1) = getUNIT(.cell(5))
'                        .newDESCRIPTION = .SUMMARYlist.TextMatrix(summaryPOSITION, 19)
'                    End If
'                End If
            Case "02040100" 'WarehouseReceipt
            Case Else
                .NEWconditionBOX(n).Enabled = True
                .logicBOX(n).Enabled = True
                .sublocaBOX(n).Enabled = True
                .repairBOX(n).Enabled = True
        End Select
        If serialPool = "SERIAL" Then
            .quantityBOX(n).Enabled = False
            .quantity2BOX(n).Enabled = False
        Else
            .quantityBOX(n).Enabled = True
            .quantity2BOX(n).Enabled = True
        End If
        .priceBOX(n).Enabled = True
    End With
    
ErrHandler:
    Select Case Err.Number
        Case 360, 340, 30, 438
            Resume Next
        Case 0
        Case Else
            'MsgBox "setupBoxes2 Error: " + Format(Err.Number) + "/" + Err.description
            Resume Next
    End Select
    Err.Clear
End Sub

Sub setupBOXES(n, datax As ADODB.Fields, serial As Boolean, Optional QTYpo)
Dim x, cond, logic, subloca, newCOND, serialPool
serialPool = IIf(serial, "SERIAL", "POOL") 'Juan 2010-5-14
Dim newButtonEnabled As Boolean
On Error GoTo ErrHandler:

    With frmWarehouse
        newButtonEnabled = .newBUTTON.Enabled
        Load .quantity(n)
        If Not .newBUTTON.Enabled Then Call putBOX(.quantity(n), .detailHEADER.ColWidth(0) + 140, topNODE(n), .detailHEADER.ColWidth(1) - 40, vbWhite)
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
        
        Select Case .tag
            'ReturnFromRepair WarehouseIssue,WellToWell,InternalTransfer,
            'AdjustmentIssue,WarehouseToWarehouse,Sales,ReturnFromWell, AdjustmentEntry
            Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300", "02050200"
                If serial Then
                    .quantity(n) = 1
                Else
                    If .newBUTTON.Enabled Then
                        .quantity(n) = Format(datax!qty1, "0.00")
                        cond = Trim(datax!OriginalCondition)
                        logic = Trim(datax!fromlogic)
                        subloca = Trim(datax!fromSubLoca)
                        newCOND = IIf(IsNull(datax!NEWcondition), "", datax!NEWcondition)
                    Else
                        .quantity(n) = Format(datax!qty, "0.00")
                        cond = Trim(datax!condition)
                        logic = Trim(datax!logic)
                        subloca = Trim(datax!subloca)
                        newCOND = datax!condition
                    End If
                End If
                'Juan 2013-12-29 Added to fix AE and qty for both serial and pool
                If .tag = "02050200" Then ' "02050200" adjustement entry
                    If serial Then
                        .quantityBOX(n) = "1.00"
                    Else
                        .quantityBOX(n) = "0.00"
                    End If
                Else
                        .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), cond, logic, subloca, IIf(IsNull(datax!serialNumber), "POOL", Trim(datax!serialNumber)), n), "0.00")
                End If
                '-------------------------------------------
                'Juan 2010-10-4 These transactions should not use currency rate because USD is the standard
                '.priceBOX(n) = Format(datax!unitPRICE * currencyRate, "0.00")
                .priceBOX(n) = Format(datax!unitPRICE, "0.00")
                '--------------------
                .NEWconditionBOX(n).tag = newCOND
              Case "02040100" 'WarehouseReceipt
                .quantity(n) = Format(QTYpo, "0.00")
                If newButtonEnabled = True Then
                    newCOND = datax!NEWcondition
                    .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "GENERAL", "GENERAL", serialPool, n), "0.00")
                    'Juan 2010-6-6
                    .quantity2BOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "GENERAL", "GENERAL", serialPool, n), "0.00")
                    '----------------------
                Else
                    'newCOND = datax!condition 'Juan 2010-5-15
                    newCOND = "01"
                    doChanges = False
                    If serialPool = "SERIAL" Then 'Juan 2010-5-17
                        .quantityBOX(n) = "1.00"
                        .quantity2BOX(n) = "1.00"
                        Call calculations(True, True) ' juan 2012-3-9
                    Else 'Juan 2010-5-17
                        .quantityBOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "unique", "unique", serialPool, n), "0.00")
                        .quantity2BOX(n) = Format(summaryQTY(Trim(datax!StockNumber), "01", "unique", "unique", serialPool, n), "0.00")
                    End If '-------------
                    doChanges = True
                End If
                .priceBOX(n) = Format(datax!unitPRICE, "0.00")
                .NEWconditionBOX(n).tag = newCOND
                Load .repairBOX(n)
                .repairBOX(n) = Format(datax!poItem)
        End Select
        .NEWconditionBOX(n) = .NEWconditionBOX(n).tag
        
        Load .poItemBox(n)
        'WarehouseReceipt
        If .tag = "02040100" Then
            .poItemBox(n) = datax!poItem
            .poItemLabel = datax!poItem
        Else
            .poItemBox(n) = .poItemLabel
        End If
        
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
                .logicBOX(n) = "GENERAL"
                .sublocaBOX(n) = "GENERAL"
            End If
        Else
            .logicBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 11)
            .logicBOX(n).tag = .logicBOX(n)
            .sublocaBOX(n) = .SUMMARYlist.TextMatrix(summaryPOSITION, 12)
            .sublocaBOX(n).tag = .sublocaBOX(n)
            .grid(2).Visible = False
            .logicBOX(n).ToolTipText = getWAREHOUSEdescription(.logicBOX(n))
            .sublocaBOX(n).ToolTipText = getSUBLOCATIONdescription(.sublocaBOX(n))
        End If
        
        Load .unitBOX(n)
        Load .unit2BOX(n)
        .unitBOX(n).Enabled = False
        .unit2BOX(n).Enabled = False
        If .newBUTTON.Enabled Then
            .unitBOX(n) = ""
            .unit2BOX(n) = ""
        Else
            .unitBOX(n) = datax!unit
            .unit2BOX(n) = datax!unit2
        End If
        
        If summaryPOSITION = 0 Then
            If .newBUTTON.Enabled Then
                newCOND = datax!NEWcondition
            Else
                newCOND = datax!condition
                .NEWconditionBOX(n).ToolTipText = datax!conditionName
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
                If Not .newBUTTON.Enabled Then
                    .logicBOX(n).Enabled = False
                    '.sublocaBOX(n).Enabled = False
                    .sublocaBOX(n).Enabled = True
                End If
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
                        .repairBOX(n) = "0"
                    End If
                Else
                    If .newBUTTON.Enabled Then
                        .repairBOX(n) = Format(datax!repairCOST, "0.00")
                        .cell(5) = Trim(datax!NewStockNumber)
                        .cell(5).tag = .cell(5)
                        .unitLABEL(1) = getUNIT(.cell(5).tag)
                        .newDESCRIPTION = Trim(datax!NewStockDescription)
                    Else
                        .repairBOX(n) = SUMMARYlist.TextMatrix(summaryPOSITION, 17)
                        .cell(5) = SUMMARYlist.TextMatrix(summaryPOSITION, 18)
                        .cell(5).tag = .cell(5)
                        .unitLABEL(1) = getUNIT(.cell(5))
                        .newDESCRIPTION = .SUMMARYlist.TextMatrix(summaryPOSITION, 19)
                    End If
                End If
            Case "02040100" 'WarehouseReceipt
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
            .quantity2BOX(n).Enabled = False
            .priceBOX(n).Enabled = False
            .NEWconditionBOX(n).Enabled = False
            .logicBOX(n).Enabled = False
            .sublocaBOX(n).Enabled = False
            .repairBOX(n).Enabled = False
        Else
            'Juan 2010-5-17
            If serialPool = "SERIAL" Then
                If frmWarehouse.tag = "02040300" Or frmWarehouse.tag = "02040200" Then  'Return from Well, 'WarehouseIssue

                Else
                    '.quantityBOX(n) = "1.00"
                    '.quantity2BOX(n) = "1.00"
                    .quantityBOX(n).Enabled = False
                    .quantity2BOX(n).Enabled = False
                End If
            Else
                Select Case frmWarehouse.tag
                    Case "02050200" 'AdjustmentEntry
                        .quantityBOX(n) = "1.00"
                        .quantity2BOX(n) = "1.00"
                        .quantityBOX(n).Enabled = True
                        .quantity2BOX(n).Enabled = False
                    Case Else
                        .quantityBOX(n).Enabled = True
                        .quantity2BOX(n).Enabled = True
                End Select
            End If
            '---------------------
            .priceBOX(n).Enabled = True
        End If
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


Sub fillDETAILlist(StockNumber, description, unit, Optional QTYpo, Optional stockListRow, Optional serialNum)
Dim i, n, sql, rec, cond, loca, subloca, stock, total, key, lastLine, thick, condName, currentLOGIC, currentSUBloca
Dim sublocaname, logicname, currentCOND
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
sqlKey = ""
'On Error Resume Next
On Error GoTo ErrHandler
    With frmWarehouse
    
        'juan 2012-1-3 need to validate if the stocknumber already taken into the list
        If .tag = "02040100" Or .tag = "02050200" Then 'receipt 2012-3-8, added to allow receipt to be re entered; Juan 2013-12-29 also adde AE
            If summaryQTYshort(StockNumber) > 0 Then
                multipleLine = True
            End If
        Else
            If summaryQTYshort(StockNumber) > 0 Then Exit Sub
        End If
        Screen.MousePointer = 11
        .STOCKlist.MousePointer = Screen.MousePointer
        tabindex = 1
        .commodityLABEL = StockNumber
        If IsMissing(serialNum) Then
            serialLabel = ""
        Else
            serialLabel = serialNum
        End If
        'WarehouseReceipt
        If .tag = "02040100" Then
            .poItemLabel = .STOCKlist.TextMatrix(Val(stockListRow), 8)
        Else
            If IsMissing(stockListRow) Then
                .poItemLabel = ""
            Else
                .poItemLabel = stockListRow
            End If
        End If
        .unitLABEL(0) = unit
        .unitLABEL(1) = ""
        .descriptionLABEL = description
        If StockNumber + description + unit = "" Then
            Call cleanDETAILS
            Screen.MousePointer = 0
            frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
            Exit Sub
        End If
 
       isFirstSubmit = True
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
                    sql = "SELECT * FROM StockInfoIssues WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Transaction# = '" + .cell(0).tag + "' AND " _
                        & "Stocknumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY OriginalCondition, LogicName, SubLocaName"
                        
                'AdjustmentEntry, WarehouseReceipt, ReturnFromRepair, Return from Well
                Case "02050200", "02040100", "02040400", "02040300"
                    sql = "SELECT * FROM StockInfoReceptions WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Transaction# = '" + .cell(0).tag + "' AND " _
                        & "Stocknumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY OriginalCondition, LogicName, SubLocaName"
            End Select
        Else
        
            If Not IsNull(StockNumber) Then
                If StockNumber <> "" Then
                    sNumber = StockNumber
                    'Juan 2010-9-4 implementing ratio rather than computer factor
                    computerFactorValue = ImsDataX.ComputingFactor(nameSP, sNumber, cn)
                    Set datax = getDATA("getStockRatio", Array(nameSP, sNumber))
                    If datax.RecordCount > 0 Then
                        ratioValue = datax!stk_ratio2
                    Else
                        ratioValue = 1
                    End If
                    datax.Close
                    '----------------------
                    stock = ""
                End If
            End If
        
            Select Case .tag
                'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
                'AdjustmentIssue,WarehouseToWarehouse,Sales
                Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                    sql = "SELECT  * FROM StockInfoQTYST4 WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "Company = '" + .cell(1).tag + "' AND " _
                        & "Warehouse = '" + .cell(2).tag + "' AND " _
                        & "StockNumber = '" + .commodityLABEL + "' " _
                        & "ORDER BY Condition, LogicName, SubLocaName"
                Case "02050200" 'AdjustmentEntry
                    sql = "SELECT stk_stcknumb as StockNumber, stk_desc as StockDescription, stk_poolspec, '01' as condition,   " _
                        & "'GENERAL' AS logic, 'GENERAL' AS LogicName, 'GENERAL' AS subloca,  'GENERAL' AS SubLocaName, 'NEW' AS conditionName, " _
                        & "1 as qty, 'Pool' as serialnumber, 0 as unitPRICE  FROM STOCKMASTER WHERE " _
                        & "(stk_npecode = '" + nameSP + "') AND " _
                        & "(stk_stcknumb = '" + .commodityLABEL + "')"
                Case "02040100" 'WarehouseReceipt
                    Dim response
                    Set datax = getDATA("statusFREIGHT", Array(nameSP, Format(.cell(4)), StockNumber))
                    If datax.RecordCount = 0 Then
                        Screen.MousePointer = 0
                        MsgBox "Error on Warehouse Module about statusFREIGHT"
                        Exit Sub
                    Else
                        If datax!po_freigforwr = 1 Then
                            If datax!poi_stasdlvy <> "RC" Then
                                Screen.MousePointer = 0
                                MsgBox "Freight Reciept has not been completed, please receive it first."
                                Exit Sub
                            End If
                        End If
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
                            MsgBox "Error on Warehouse Module referent to a non existing invoice"
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
                    'sql = "SELECT TOP 1 * FROM StockInfoPO WHERE " _
                      '  & "NameSpace = '" + nameSP + "' AND " _
                      '  & "StockNumber = '" + StockNumber + "' AND " _
                      ' & "POitem = '" + .STOCKlist.TextMatrix(.STOCKlist.row, 6) + "'"
                        
                        'muzammil 10/20/05  'BUG1
                        'since it did not have the pono as one of the qualifying criterias
                        'it would get any po lineitem record with this stockno with ths line item.
                        'and usually these were reqs. with unitprice 0 and so some of the receipts would have up as 0
                        'for some pos
                    sql = "SELECT TOP 1 * FROM StockInfoPO WHERE " _
                        & "NameSpace = '" + nameSP + "' AND " _
                        & "StockNumber = '" + StockNumber + "' AND " _
                        & "PO = '" + .cell(4).text + "' AND  " _
                        & "POitem = '" + .STOCKlist.TextMatrix(.STOCKlist.row, 8) + "' " _
                        & "ORDER BY curd_creadate Desc"
                        'Juan 2010-5-2010
                        '& "POitem = '" + .STOCKlist.TextMatrix(.STOCKlist.row, 6) + "'"
                        '-----------------
                        
            End Select
            sqlKey = sql
        End If
    End With
    Set datax = New ADODB.Recordset
    datax.Open sql, cn, adOpenStatic
    If datax.RecordCount = 0 Then
        cleanDETAILS
    Else
        ReDim qtyArray(datax.RecordCount)
        ReDim subLocationArray(datax.RecordCount)
        datax.MoveLast
        Call workBOXESlist("clean")
        datax.MoveFirst
        total = CDbl(0)
        With frmWarehouse.Tree
            .width = frmWarehouse.detailHEADER.width
            .Nodes.Clear
            moreSerial = False
            Dim r As Integer
            r = 0
            Do While Not datax.EOF
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
                    moreSerial = False
                    If frmWarehouse.newBUTTON.Enabled Then
                        If frmWarehouse.tag = "02040100" Then 'WarehouseReceipt
                            cond = Trim(datax!NEWcondition)
                            condName = Trim(datax!NewConditionName)
                        Else
                            cond = Trim(datax!OriginalCondition)
                            condName = Trim(datax!OriginalConditionName)
                        End If
                    Else
                        cond = Trim(datax!condition)
                        condName = Trim(datax!conditionName)
                    End If
                    loca = ""
                    subloca = ""
                    If frmWarehouse.tag = "02040100" Or frmWarehouse.tag = "02050200" Then 'AdjustmentEntry
                    Else
                        .Nodes.Add , tvwChild, "@" + cond, "Condition " + cond + " - " + condName, "thing"
                        .Nodes("@" + cond).Bold = True
                        .Nodes("@" + cond).backcolor = &HE0E0E0
                    End If
                End If
                Err.Clear
                If frmWarehouse.newBUTTON.Enabled Then
                    currentLOGIC = IIf(IsNull(datax!fromlogic), "", Trim(datax!fromlogic))
                    currentSUBloca = IIf(IsNull(datax!fromSubLoca), "", Trim(datax!fromSubLoca))
                    logicname = IIf(IsNull(datax!logicname), "", datax!logicname)
                    sublocaname = IIf(IsNull(datax!sublocaname), "", datax!sublocaname)
                Else
                    If frmWarehouse.tag = "02040100" Then 'WarehouseReceipt then
                    Else
                        currentLOGIC = Trim(datax!logic)
                        currentSUBloca = Trim(datax!subloca)
                    End If
                End If
                Select Case frmWarehouse.tag
                    Case "02050200" 'AdjustmentEntry
                        pool = IIf(datax!stk_poolspec = True, True, False)
                        If Not pool Then ' To allow the stocknumber to be selected more than once
                            frmWarehouse.STOCKlist.col = 0
                            frmWarehouse.STOCKlist.text = "È"
                            QTYpo = 1 'Juan: 2013-12-29 Added to make sure there is always one for serial
                        End If

                        If frmWarehouse.newBUTTON.Enabled = True Then
                            .Nodes.Add "@AE", tvwChild, "Adjustement Entry {{" + "unique", "New Inventory", "thing 1"
                        Else
                            .Nodes.Add , tvwChild, "@AE", "Adjustement Entry ", "thing"
                            .Nodes("@AE").Bold = True
                            .Nodes("@AE").backcolor = &HE0E0E0
                            key = "AE-NEW{"
                            If pool Then
                                .Nodes.Add "@AE", tvwChild, key + "{{Pool", "Pool", "thing 1"
                            Else
                                .Nodes.Add "@AE", tvwChild, key + "{{Serial", "Serial:", "thing 1"
                            End If
                        End If
                        Call setupBOXES(.Nodes.Count, datax.Fields, Not pool, QTYpo)
 
                    'ReturnFromRepair, WarehouseIssue,WellToWell,InternalTransfer,
                    'AdjustmentIssue,WarehouseToWarehouse,Sales,ReturnFromWell
                    Case "02040400", "02040200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                        If loca <> currentLOGIC Then
                            loca = currentLOGIC
                            subloca = ""
                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + loca, "Logical Warehouse: " + datax!logicname, "thing 0"
                        End If
                        'muzammil 10/20/2005    'BUG1
                        'commented this IF line of code and added the IF after this one.
                        ' what was happening was we took out the code which would remove "General" which was a default sublocation for any kind of transactions
                        'it would not feed any value to sublocations when an Issue was being done.
                        'so sulocation field in QTYST4 table has only an empty string
                        'and then the if condition would check to see if the currentSUBloca( which is the sublocation from db)
                        'is empty, if it is empty then does not add the quantity line at all.
                        ' I added the code of line beneath it and added exception for ReturnFromRepair and Wellto well
                        
                        'ReturnFromRepair,WellToWell,ReturnFromWell,AdjustmentEntry
                        'If subloca <> currentSUBloca Then
                         
                        '2011-5-22 Juan - modified to optimize and add adjustementIssue and sales
                        'If subloca <> currentSUBloca(frmWarehouse.tag = "02040400" Or frmWarehouse.tag = "02040500" Or frmWarehouse.tag = "02040300" Or frmWarehouse.tag = "02050200") Then   'MUZAMMIL 10/20/05
                        If subloca <> currentSUBloca Then
                            Select Case frmWarehouse.tag
                                Case "02040400", "02040500", "02040300", "02050200", "02050300", "02050400", "02040700", "02040200"
                                    subloca = currentSUBloca
                                    logicname = IIf(IsNull(datax!logicname), "", datax!logicname)
                                    sublocaname = IIf(IsNull(datax!sublocaname), "", datax!sublocaname)
                                    key = cond + "-" + condName + "{{" + loca + "{{" + subloca
                                    qtyArray(r) = datax!qty
                                    subLocationArray(r) = subloca
                                    thisSubLoca = subloca
                                    thisLogic = datax!logic
                                    r = r + 1
                                    If IsNull(datax!serialNumber) Or datax!serialNumber = "" Or UCase(datax!serialNumber) = "POOL" Then
                                        .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaname, "thing 1"
                                        Call setupBOXES(.Nodes.Count, datax.Fields, False)
                                    Else
                                        moreSerial = True
                                        .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaname, "thing 0"
                                        Dim bookMark
                                        
                                        Select Case frmWarehouse.tag
                                            Case "02040300", "02040200", "02050300", "02050400"  'Return from Well, 'WarehouseIssue, AdjustementIssue, Sales
                                                bookMark = datax.bookMark
                                                datax.MoveFirst
                                                Do While Not datax.EOF
                                                    If RTrim(thisSubLoca) = RTrim(datax!subloca) And RTrim(thisLogic) = RTrim(datax!logic) Then
                                                        .Nodes.Add key, tvwChild, key + "{{#" + datax!serialNumber, "Serial #: " + datax!serialNumber, "thing 1"
                                                        Call setupBOXES(.Nodes.Count, datax.Fields, True)
                                                    End If
                                                    datax.MoveNext
                                                Loop
                                                datax.bookMark = bookMark
                                                
                                            End Select
                                    End If
                                    If frmWarehouse.newBUTTON.Enabled Then
                                        total = total + datax!qty1
                                    Else
                                        total = total + datax!qty
                                    End If
                            End Select
                        End If
                        If loca <> currentLOGIC Then
                            loca = currentLOGIC
                            subloca = ""
                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + loca, "Logical Warehouse: " + logicname, "thing 0"
                        End If
                        
                        'muzammil 10/20/2005
                        'what is the purpose of these statements here
                        'the above comparision of  subloca <> currentSUBloca would make them the same if they are different
                        'almost as if this line of code is never being executed at all
                        If subloca <> currentSUBloca Then
                            subloca = currentSUBloca
                            key = cond + "-" + datax!conditionName + "{{" + loca + "{{" + subloca
                            If IsNull(datax!serialNumber) Or datax!serialNumber = "" Or UCase(datax!serialNumber) = "POOL" Then
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaname, "thing 1"
                                Call setupBOXES(.Nodes.Count, datax.Fields, False)
                            Else
                                moreSerial = True
                                .Nodes.Add cond + "{{" + loca, tvwChild, key, "Sublocation: " + sublocaname, "thing 0"
                            End If
                            total = total + datax!qty
                        End If

                    Case "02040100" 'WarehouseReceipt
                        '1 "Commodity"
                        '2 "Serial"
                        '3 "Condition"
                        '4 "Unit Price"
                        '5 "Description"
                        '6 "Unit"
                        '7 "Qty"
                        '8 "node"
                        '9 "From Logical"
                        '10 "From Subloca"
                        '11 "To Logical"
                        '12 "To Subloca"
                        '13 "New Condition Code"
                        '14 "New Condition Description"
                        '15 "Unit Code"
                        '16 "Computer Factor"
                        '17 "repaircost"
                        '18 "newcomodity"
                        '19 "newdescription"
                        '20 "Original Condition Code"
                        '21 "Unit 2"
                        '22 "po line item
                        '23 "Secundary QTY"
                        '24 "po
                        If Null = QTYpo Then
                        Else
                            If QTYpo = "0" Or QTYpo = "0.00" Then Exit Sub
                        End If
                        '2010-5-13 Juan
                        Set datay = getDATA("getPoolSpec", Array(nameSP, StockNumber))
                        If datay.RecordCount = 0 Then
                            pool = True
                        Else
                            pool = IIf(datay!stk_poolspec = True, True, False)
                            If Not pool Then
                                frmWarehouse.STOCKlist.col = 0
                                frmWarehouse.STOCKlist.text = "È"
                            End If
                            'pool = IIf(IsNull(datay!serialnumber), True, IIf(datay!serialnumber = "", True, False))
                        End If
                        If pool = True And multipleLine = True Then Exit Sub 'juan 2012-3-7
                        '---------------
                        
                        'Juan 2010-5-13
                        If frmWarehouse.newBUTTON.Enabled = True Then
                            .Nodes.Add "@" + cond, tvwChild, cond + "{{" + "unique", "New Inventory", "thing 1"
                        Else
                            .Nodes.Add , tvwChild, "@" + cond, cond + "-" + condName, "thing"
                            .Nodes("@" + cond).Bold = True
                            .Nodes("@" + cond).backcolor = &HE0E0E0
                            key = cond + "-" + condName + "{{"
                            If pool Then
                                .Nodes.Add "@" + cond, tvwChild, key + "{{Pool", "Pool", "thing 1"
                            Else
                                .Nodes.Add "@" + cond, tvwChild, key + "{{Serial", "Serial:", "thing 1"
                            End If
                        End If
                        'Call setupBOXES(.Nodes.Count, datax.Fields, False, QTYpo)
                        Call setupBOXES(.Nodes.Count, datax.Fields, Not pool, QTYpo)
                        total = QTYpo
                        
'                       .Nodes(key + "{{Serial").Selected = True
'                       .StartLabelEdit
                '--------------------
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
            originalQty = total
            Call bottomLine(totalNode, total, pool, StockNumber, False)
            frmWarehouse.sublocaBOX(frmWarehouse.sublocaBOX.Count).backcolor = &H80FFFF
        End With
    End If
    directCLICK = False
    Screen.MousePointer = 0
    frmWarehouse.MousePointer = 0
    frmWarehouse.STOCKlist.MousePointer = Screen.MousePointer
    'Juan 2010-5-14
    If frmWarehouse.tag = "02040100" Or frmWarehouse.tag = "02050200" Then 'WarehouseReceipt, AdjustmentEntry
        If Not pool Then
            With frmWarehouse.Tree
                If .Nodes.Count > 1 Then
                    .Nodes(.Nodes.Count - 1).Selected = True
                    .StartLabelEdit
                End If
            End With
        End If
    End If
    Err.Clear
    '----------------
    Exit Sub
    
ErrHandler:
If Err.Number > 0 Then
    'MsgBox Err.description
    Err.Clear
End If
Resume Next
End Sub

Sub doCOMBO(Index, datax As ADODB.Recordset, list, totalwidth)
Dim rec, i, extraW
Dim t As String
    Err.Clear
    With frmWarehouse.combo(Index)
        Do While Not datax.EOF
            rec = ""
            For i = 0 To frmWarehouse.matrix.TextMatrix(1, Index) - 1
                If list(i) = "error" Then
                    MsgBox "Definition error, please contact IMS"
                    Exit Sub
                Else
                    t = IIf(IsNull(datax(list(i))), "", datax(list(i)))
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
            .Height = 2340
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

Function InvtReceipt_Insert2a(NameSpace As String, PONumb As String, TranType As String, Companycode As String, Warehouse As String, user As String, cn As ADODB.Connection, Optional ManufacturerNumb As String, Optional TranFrom As String, Optional TransNum As String) As Integer

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
        InvtReceipt_Insert2a = .parameters("RV") = 0
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
On errror GoTo errorHandler
Dim n, rec, i, qty2Value
    With datax
        n = 0
        'Juan 2010-5-21
        'frmWarehouse.STOCKlist.Rows = .RecordCount + 1
        frmWarehouse.STOCKlist.Rows = 2
        frmWarehouse.STOCKlist.row = 1
        frmWarehouse.STOCKlist.col = 0
        frmWarehouse.STOCKlist.CellFontName = "MS Sans Serif"
        Do While Not .EOF
            Select Case frmWarehouse.tag
                'ReturnFromRepair, AdjustmentEntry,WellToWell,InternalTransfer,
                'AdjustmentIssue,WarehouseToWarehouse,Sales
                Case "02040400", "02050200", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                    n = n + 1
                    rec = Format(n) + vbTab
                    rec = rec + Trim(!StockNumber) + vbTab
                    rec = rec + IIf(IsNull(!unitPRICE), "0.00", Format(!unitPRICE, "0.00")) + vbTab
                    rec = rec + IIf(IsNull(!description), "", !description) + vbTab
                    rec = rec + IIf(IsNull(!unit), "", !unit) + vbTab
                    'Juan 2010-6-5
                    'rec = rec + Format(!qty) + vbTab
                    rec = rec + Format(!qty, "0.00") + vbTab
                    '------------------------
                    rec = rec + IIf(IsNull(!unit), "", !unit)
                'WarehouseIssue Juan 2012-3-23 to add serial
                Case "02040200"
                    n = n + 1
                    rec = Format(n) + vbTab
                    rec = rec + Trim(!StockNumber) + vbTab
                    rec = rec + Trim(!serialNumber) + vbTab
                    rec = rec + IIf(IsNull(!unitPRICE), "0.00", Format(!unitPRICE, "0.00")) + vbTab
                    rec = rec + IIf(IsNull(!description), "", !description) + vbTab
                    rec = rec + IIf(IsNull(!unit), "", !unit) + vbTab
                    'Juan 2010-6-5
                    'rec = rec + Format(!qty) + vbTab
                    rec = rec + Format(!qty, "0.00") + vbTab
                    '------------------------
                    rec = rec + IIf(IsNull(!unit), "", !unit)
                Case "02040100" 'WarehouseReceipt
                    frmWarehouse.STOCKlist.ColAlignment(7) = 0
                    rec = Format(!poItem) + vbTab
                    rec = rec + Trim(!StockNumber) + vbTab
                    rec = rec + IIf(IsNull(!QTYpo), "0.00", Format(!QTYpo, "0.00")) + vbTab
                    'Juan 2010-9-19
                    ' rec = rec + Format(!qty1, "0.00") + vbTab
                    Dim toBeReceived, toBeReceived2 As Double
                    If Null = !QTY1_invoice Then
                        toBeReceived = !qty1
                    Else
                        If !QTY1_invoice > 0 Then
                            toBeReceived = !QTY1_invoice
                        Else
                            toBeReceived = !qty1
                        End If
                    End If
                    rec = rec + Format(toBeReceived, "0.00") + vbTab
                    rec = rec + IIf(IsNull(!unit), "", !unit) + vbTab
                    
                    'Dim qty2
                    ' qty2 = Format(!qty2, "0.00")
                    ' rec = rec + qty2 + vbTab
                    If Null = !QTY2_invoice Then
                        toBeReceived2 = !qty2
                    Else
                        If !QTY2_invoice > 0 Then
                            toBeReceived2 = !QTY2_invoice
                        Else
                            toBeReceived2 = !qty2
                        End If
                    End If
                    rec = rec + Format(toBeReceived2, "0.00") + vbTab
                    rec = rec + IIf(IsNull(!unit2), "", !unit2) + vbTab
                    '-----------------------
                    rec = rec + IIf(IsNull(!description), "", !description) + vbTab
                    poItem = Format(!poItem)
                    rec = rec + Format(!poItem) + vbTab
                    rec = rec + Format(toBeReceived, "0.00") + vbTab
                    rec = rec + Format(toBeReceived2, "0.00")
            End Select
            frmWarehouse.STOCKlist.addITEM rec
            If n = 20 Then
                DoEvents
                frmWarehouse.STOCKlist.Refresh
            End If
            .MoveNext
        Loop
        If frmWarehouse.STOCKlist.Rows > 2 Then frmWarehouse.STOCKlist.RemoveItem (1)
        frmWarehouse.STOCKlist.RowHeightMin = 240
        frmWarehouse.STOCKlist.row = 0
    End With
    
errorHandler:
If Err.Number > 0 Then
    'MsgBox "fillSTOCKlist " + Err.description
    Err.Clear
    Resume Next
End If
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
    inProgress = False 'Juan 2010-7-22
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
Sub markROW(grid As MSHFlexGrid, Optional editing As Boolean)
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
        
        'Juan 2010-7-23
        'If IsNumeric(.text) Then
        'Juan 2010-8-14
        'If IsNumeric(.text) Or .text = "È" Then
        '--------------------
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
            If .text = "È" Then
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
                .CellFontName = "Wingdings 3"
                .CellFontSize = 10
                .text = "Æ"
            End If

            If .name = frmWarehouse.STOCKlist.name Then
                Call PREdetails
            Else
                'AdjustmentIssue juan 2012--3-24 to add serial
                If frmWarehouse.tag = "02040200" Then
                   Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 4), .TextMatrix(.row, 5), .row)
                Else
                    Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 5), .TextMatrix(.row, 6), .TextMatrix(.row, 2), .row)
                End If
            End If
        'Else
        '    unmarkRow (stock)
        'End If
    End With
    For i = 0 To 2
        frmWarehouse.grid(i).Visible = False
    Next
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
                    .cell(5).Enabled = False
                End If
                
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
        Call workBOXESlist("FIX")
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
Sub putBOX(box As textBOX, Left, Top, width, backcolor)
    With box
        .Left = Left
        .width = width
        .Top = Top
        .Height = 180
        ' Juan 2011-5-8 commented out because we found some situations with many line items
        ' TODO, it needs a mechanism when is bigger
        'If (frmWarehouse.Tree.Nodes.Count > 15 And .Index < 16) Or frmWarehouse.Tree.Nodes.Count < 16 Then
            .ZOrder
            .Visible = True
        ' End If
        .backcolor = backcolor
        Select Case frmWarehouse.tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
                'If .name = "quantity" Then .Visible = False
                If .name = "balanceBOX" Then .Visible = False
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
    End With
End Sub

Function topNODE(Index) As Integer
Dim heightFactor, spaceFactor As Integer
    spaceFactor = 45
    heightFactor = 265
    Select Case frmWarehouse.tag
        Case "02040400" 'ReturnFromRepair
            heightFactor = 325
            spaceFactor = 80
        Case "02050200" 'AdjustmentEntry
            heightFactor = 325
            spaceFactor = 80
        Case "02040200" 'WarehouseIssue
            heightFactor = 325
            spaceFactor = 40
        Case "02040500" 'WellToWell
            heightFactor = 325
            spaceFactor = 80
        Case "02040700" 'InternalTransfer
            heightFactor = 325
            spaceFactor = 80
        Case "02050300" 'AdjustmentIssue
            heightFactor = 325
            spaceFactor = 80
        Case "02040600" 'WarehouseToWarehouse
            heightFactor = 325
            spaceFactor = 80
        Case "02040100" 'WarehouseReceipt
        Case "02050400" 'Sales
            heightFactor = 325
            spaceFactor = 80
        Case "02040300" 'Return from Well
            heightFactor = 325
            spaceFactor = 80
    End Select
    topNODE = frmWarehouse.Tree.Top + spaceFactor + (heightFactor * (Index - nodeONtop))
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

Sub unmarkAllRows(grid As MSHFlexGrid)
Dim i  As Integer
Dim stock
Dim imsLock As imsLock.Lock

Screen.MousePointer = 11
frmWarehouse.Refresh
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

Sub unmarkRow(stock, Optional unmarkIt As Boolean)
    'Juan 2010-7-4
    Dim imsLock As imsLock.Lock
    If IsMissing(unmarkIt) Then unmarkIt = True
    If unmarkIt Then
        With frmWarehouse.STOCKlist
            If .text = "Æ" Then
                .col = 0
                .CellFontName = "MS Sans Serif"
                .CellFontSize = 8.5
                .text = .row
            End If
        End With
    End If
    '------------------
    
    Call fillDETAILlist("", "", "")
    
    'Unlock
    Set imsLock = New imsLock.Lock
    Call imsLock.Unlock_Row(STOCKlocked, cn, CurrentUser, rowguid, True, "STOCKMASTER", stock, False)
    Set imsLock = Nothing
    '------
End Sub

Sub updateStockListBalance() 'Juan 2010-9-19 to re-load the proper values of the stocklist, specially after have removed a row
    Dim i, ii As Integer
    Dim StockNumber As String
    Dim balance, qtySummaryList, qtyStockLIst As Double
    With frmWarehouse
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

Sub validateQTY(box As textBOX, Index)
Dim n
Dim d As Integer
    noRETURN = True
    With box
        If Index <> totalNode Then
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

Function summaryQTYshort(StockNumber) As Integer
summaryQTYshort = 0
    With frmWarehouse.SUMMARYlist
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 1)) = Trim(StockNumber) Then
                summaryQTYshort = 1
                Exit Function
            End If
        Next
    End With
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
Dim i, size, point, balanceCol
On Error Resume Next
'On Error GoTo errHandler
    With frmWarehouse
        size = .Tree.Nodes.Count
        'topvalue is for the total
        'topvalue2 is for the lines
        topvalue = 160
        topvalue2 = 0
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
                topvalue = 0
            Case "02040200" 'WarehouseIssue
                topvalue = topvalue - 90
            Case "02040500" 'WellToWell
                topvalue = topvalue - 160
            Case "02040700" 'InternalTransfer
                Select Case size
                    Case Is > 15
                        topvalue = topvalue - 70
                        topvalue2 = 100
                    Case Is > 4
                        topvalue = topvalue - 100
                        topvalue2 = 40
                    Case Else
                        topvalue = topvalue - 160
                        topvalue2 = -20
                End Select
            Case "02050300" 'AdjustmentIssue
                topvalue = topvalue - 120
            Case "02040600" 'WarehouseToWarehouse
                topvalue = topvalue - 160
            Case "02040100" 'WarehouseReceipt
                topvalue2 = 90
            Case "02050400" 'Sales
                topvalue = topvalue - 120
            Case "02040300" 'Return from Repair
                topvalue = topvalue - 160
        End Select
        If size > 0 Then
            For i = 1 To size
                Err.Clear
                If .quantity(i) <> "" Then
                        If Err.Number = 0 Then
                            Select Case UCase(work)
                                Case "CLEAN"
                                    If i > 0 Then
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
                                        'Juan 2010-6-6
                                        Select Case frmWarehouse.tag
                                            Case "02040100" 'WarehouseReceipt
                                                Unload .quantity2BOX(i)
                                                Unload .unit2BOX(i)
                                        End Select
                                        '----------------------
                                    End If
                                Case "FIX"
                                    balanceCol = 5
                                    Select Case frmWarehouse.tag
                                        Case "02040400" 'ReturnFromRepair
                                            point = 2
                                        Case "02050200" 'AdjustmentEntry
                                            point = 0
                                        Case "02040300" 'ReturnFromWell
                                            point = 1
                                        Case "02040100" 'WarehouseReceipt
                                            balanceCol = 8
                                        Case Else
                                            point = 0
                                    End Select
                                    If i = size Then
                                        If Not .newBUTTON.Enabled Then Call putBOX(.quantity(totalNode), .linesV(1).Left + 20, topNODE(size) + topvalue, .detailHEADER.ColWidth(4 + point) - 50, &HC0C0C0)
                                        Call putBOX(.quantityBOX(totalNode), .linesV(5 + point).Left + 30, topNODE(size) + topvalue, .detailHEADER.ColWidth(4 + point) - 50, &HC0C0C0)
                                        If Not .newBUTTON.Enabled Then Call putBOX(.balanceBOX(totalNode), .linesV(balanceCol + point).Left + 30, topNODE(size) + topvalue, .detailHEADER.ColWidth(balanceCol + point) - 50, &HC0C0C0)
                                    Else
                                        If Not .newBUTTON.Enabled Then Call putBOX(.quantity(i), .linesV(1).Left + 40, topNODE(i) + topvalue2, .detailHEADER.ColWidth(1) - 80, vbWhite)
                                        Call putBOX(.logicBOX(i), .linesV(2).Left + 55, topNODE(i) + topvalue2, .detailHEADER.ColWidth(2) - 80, vbWhite)
                                        Call putBOX(.sublocaBOX(i), .linesV(3).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(3) - 50, vbWhite)
                                        Call putBOX(.quantityBOX(i), .linesV(4 + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(4 + point) - 50, vbWhite)
                                        If Not .newBUTTON.Enabled Then
                                            Call putBOX(.balanceBOX(i), .linesV(balanceCol + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(balanceCol + point) - 50, vbWhite)
                                        End If
                                        Select Case .tag
                                            Case "02040400", "02040300" 'ReturnFromRepair, ReturnFromWell
                                                Call putBOX(.NEWconditionBOX(i), .linesV(4).Left + 30, topNODE(i), .detailHEADER.ColWidth(4) - 50, vbWhite)
                                                If frmWarehouse.tag = "02040400" Then
                                                    Call putBOX(.repairBOX(i), .linesV(5).Left + 30, topNODE(i), .detailHEADER.ColWidth(5) - 50, vbWhite)
                                                End If
                                                .NEWconditionBOX(i).ZOrder
                                            Case "02050200" 'AdjustmentEntry
                                                Call putBOX(.logicBOX(i), .linesV(1).Left + 55, topNODE(i) + topvalue, .detailHEADER.ColWidth(1) - 80, vbWhite)
                                                Call putBOX(.sublocaBOX(i), .linesV(2).Left + 30, topNODE(i) + topvalue, .detailHEADER.ColWidth(2) - 50, vbWhite)
                                                Call putBOX(.priceBOX(i), .linesV(3).Left + 30, topNODE(i) + topvalue, .detailHEADER.ColWidth(3) - 50, vbWhite)
                                                Call putBOX(.NEWconditionBOX(i), .linesV(4).Left + 30, topNODE(i) + topvalue, .detailHEADER.ColWidth(4) - 50, vbWhite)
                                                Call putBOX(.quantityBOX(i), .linesV(5).Left + 30, topNODE(i) + topvalue, .detailHEADER.ColWidth(5) - 50, vbWhite)
                                            Case "02040100" 'WarehouseReceipt
                                                'Juan 2010-5-10
                                                Call putBOX(.unitBOX(i), .linesV(4 + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(4 + point) - 50, vbWhite)
                                                Call putBOX(.quantityBOX(i), .linesV(5 + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(5 + point) - 50, vbWhite)
                                                Call putBOX(.unit2BOX(i), .linesV(6 + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(6 + point) - 50, vbWhite)
                                                Call putBOX(.quantity2BOX(i), .linesV(7 + point).Left + 30, topNODE(i) + topvalue2, .detailHEADER.ColWidth(7 + point) - 50, vbWhite)
                                                '---------------------
                                        End Select
                                    End If
                            End Select
                        Else
                            Err.Clear
                        End If
                End If
            Next
        End If
    End With
'errHandler:
'    If Err.Number > 0 Then
'        MsgBox "workBOXESlist error: " + Err.description
'        Resume Next
'    End If
    Err.Clear
End Sub

Sub selectROW(grid As MSHFlexGrid, Optional clean As Boolean)
On Error GoTo getOUT
Dim changeCOLORS As Boolean
Dim i, currentCOL, currentROW As Integer
    Screen.MousePointer = 11
    With frmWarehouse.grid(0)
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
                Call getCOLORSrow(grid, .cols - 1)
                changeCOLORS = True
            End If
        Else
            .tag = .row
            Call getCOLORSrow(grid, .cols - 1)
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

Sub calculations(updateStockList As Boolean, Optional isDynamic As Boolean)
Dim this, r, summary, balance, balance2, balanceTotal, col
Dim i As Integer
Dim once As Boolean
Dim originalQty As Double
once = True
balanceTotal = 0

'This applies to reciept only
'balance and balance2 are meant to store the difference between the PO qty to be recieved and what is being recieved.
'that is for primary and secondary qty's
'originalQTY is meant to store the qty to be recieved as originally got from the PO and thus is never updated

'On Error GoTo errorHandler
On Error Resume Next
    With frmWarehouse
        'Global declarations
        Dim colRef, colRef2, colTot As Integer
        Dim fromStockList As Boolean
        isDynamic = True
        fromStockList = False
        colRef = 2
        colRef2 = 7
        colTot = 5
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
                colTot = 6
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                colRef = 9
                colTot = 3
                isDynamic = False
                fromStockList = True
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        'Restore initial values
'        For i = 0 To .STOCKlist.Rows - 1
'            .STOCKlist.TextMatrix(i, colTot) = .STOCKlist.TextMatrix(i, colRef)
'            .STOCKlist.TextMatrix(i, colTot + 2) = .STOCKlist.TextMatrix(i, colRef + 1)
'        Next
        'WarehouseReceipt
        If .tag = "02040100" Then
            r = findSTUFF(.commodityLABEL, .STOCKlist, 1, .poItemLabel, 8)
        Else
            r = findSTUFF(.commodityLABEL, .STOCKlist, 1)
        End If
        If r > 0 Then
            'When isDynamic variable is false means we are taking the values from the stockList
            
            If isDynamic Then
            Else
                If IsNumeric(.STOCKlist.TextMatrix(r, colRef)) Then
                    originalQty = CDbl(.STOCKlist.TextMatrix(r, colRef))
                    If originalQty > 0 Then
                        this = 0
                        balance = originalQty
                    Else
                        Exit Sub
                    End If
                    .STOCKlist.row = r
                    Call selectROW(.STOCKlist)
                End If
            End If
        End If
        
        'Main cycle to scan the active tree nodes-------------
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
                        Next
                        balance = balance - subTot
                    End If '-----------------------------------
                    'Step to update cells on screen-----------------
                    If qBoxExists Then
                        balance = balance - .quantityBOX(i)
                        this = this + CDbl(.quantityBOX(i))
                        .balanceBOX(i) = Format(balance, "0.00")
                        balanceTotal = balanceTotal + balance
                    End If
                Else
                    Err.Clear
                End If
            End If
        Next
        'Final step to totalize--------------------
        .quantityBOX(totalNode) = Format(this, "0.00")
        If isDynamic Then
            .balanceBOX(totalNode) = Format(balanceTotal, "0.00")
            balance = balanceTotal
        Else
            .balanceBOX(totalNode) = Format(balance, "0.00")
        End If
        If ratioValue > 1 Then
            balance2 = balance * ratioValue
        Else
            balance2 = balance
        End If
        If updateStockList Then .STOCKlist.TextMatrix(r, colTot) = Format(balance, "0.00")
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                If updateStockList Then .STOCKlist.TextMatrix(r, colTot + 2) = Format(balance2, "0.00")
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
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

Sub calculations2(summaryListRow, nodeSubLoca, activeTreeNode, Optional isHiding As Boolean)
'Juan 2010-10-9
'This is an alternative procedure  to calculate ater ther first submit is done (isFirstSubmit=false)
'It is based on the summaryValues grid rather than directly from the tree
Dim this, stockListRow, originalSubLocation, col
Dim i As Integer
Dim originalQty, originalTotalQty, lineQty, lineTotalQty, balanceQty, balanceTotalQty As Double
On Error GoTo errorHandler
    With frmWarehouse
        'Global declarations
        Dim colRef, colRef2, colTot As Integer
        colRef = 2
        colRef2 = 7
        colTot = 5
        Select Case .tag
            Case "02040400" 'ReturnFromRepair
            Case "02050200" 'AdjustmentEntry
            Case "02040200" 'WarehouseIssue
            Case "02040500" 'WellToWell
            Case "02040700" 'InternalTransfer
            Case "02050300" 'AdjustmentIssue
            Case "02040600" 'WarehouseToWarehouse
            Case "02040100" 'WarehouseReceipt
                colRef = 9
                colTot = 3
                fromStockList = True
            Case "02050400" 'Sales
            Case "02040300" 'Return from Well
        End Select
        'This is the starting point reference, the stock number itself + original qtyArray
        'WarehouseReceipt
        If .tag = "02040100" Then
            stockListRow = findSTUFF(.commodityLABEL, .STOCKlist, 1, .poItemLabel, 8)
        Else
            stockListRow = findSTUFF(.commodityLABEL, .STOCKlist, 1)
        End If
        originalSubLocation = .SUMMARYlist.TextMatrix(summaryListRow, 10)
        originalTotalQty = 0
        lineTotalQty = 0
        balanceTotalQty = 0
        For i = 0 To UBound(qtyArray)
            originalQty = qtyArray(i) 'Taken from the array first time submitted
            lineQty = 0
            'Validating the from sub location from the original tree (stored into subLocationArray()) with the current editing tree which might be
            'only part of a bigger record
            Dim doProcess As Boolean
            doProcess = False
            If InStr(UCase(nodeSubLoca), "SUBLOCATION:") > 0 Then doProcess = True
            If InStr(UCase(nodeSubLoca), "POOL") > 0 Then doProcess = True
            If InStr(UCase(nodeSubLoca), "SERIAL") > 0 Then doProcess = True
            If doProcess Then
                Dim originallSubLoca
                originallSubLoca = LTrim(Mid(nodeSubLoca, 13))
                If originallSubLoca = subLocationArray(i) Then
                    lineQty = frmWarehouse.quantityBOX(activeTreeNode)
                Else
                    For j = 1 To .SUMMARYlist.Rows - 1
                        'The reason for this select case is to manage if there is difrerences between the previous and current values
                        If .SUMMARYlist.TextMatrix(j, 1) = .commodityLABEL.Caption Then
                        Else
                            If .SUMMARYlist.TextMatrix(j, 10) = originalsubloca Then
                                lineQty = CDbl(.SUMMARYlist.TextMatrix(j, colRef2))
                            End If
                        End If
                    Next
                End If
                Select Case .tag
                    Case "02040100" 'WarehouseReceipt
                        balanceQty = lineQty
                    Case Else
                        balanceQty = originalQty - lineQty
                End Select
                If originallSubLoca = subLocationArray(i) Then
                    .balanceBOX(activeTreeNode) = Format(balanceQty, "0.00")
                End If
            End If
            originalTotalQty = originalTotalQty + originalQty
            lineTotalQty = lineTotalQty + lineQty
            balanceTotalQty = balanceTotalQty + balanceQty
        Next
        'Final step to totalize--------------------
        .quantityBOX(totalNode) = Format(lineTotalQty, "0.00")
        Select Case .tag
            Case "02040100", "02050200" 'WarehouseReceipt AdjustmentEntry
                .balanceBOX(totalNode) = Format(balanceQty, "0.00")
            Case Else
                .balanceBOX(totalNode + 1) = Format(balanceQty, "0.00")
        End Select
        If isHiding Then
            'It does not affect stocklist
        Else
            .STOCKlist.TextMatrix(stockListRow, colTot) = Format(balanceTotalQty, "0.00")
        End If
        
    End With
    Exit Sub
errorHandler:
    If Err.Number = 340 Then
    Else
        'MsgBox Err.description
        Err.Clear
    End If
    Resume Next
End Sub


Function findSTUFF(toFIND, grid As MSHFlexGrid, col, Optional toFIND2, Optional col2 As Integer) As Integer
Dim i
Dim findIT As Boolean
    findSTUFF = 0
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
                    If IsMissing(toFIND2) Or IsMissing(col2) Then
                        findSTUFF = i
                        Exit For
                    Else
                         If UCase(Trim(.TextMatrix(i, col2))) = UCase(Trim(toFIND2)) Then
                            findSTUFF = i
                            Exit For
                         Else
                         End If
                    End If
                End If
            Next
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
            .ReportFileName = App.path + "CRreports\transactionGlobal.rpt"
            .ParameterFields(1) = "ponumb;" + frmWarehouse.cell(0) + ";TRUE"
            'call translator.Translate_Reports("transactionGlobal.rpt")
        Else
            '*******************
            'CHECK THIS PATH
            .ReportFileName = App.path + "CRreports\transaction.rpt"
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
If skipAlphaSearch = True Then
    skipAlphaSearch = False
    Exit Sub
End If
    If cellACTIVE <> "" Then
        With gridACTIVE
            If Not .Visible Then .Visible = True
            If .Rows < Val(.tag) Then .tag = 1
            If IsNumeric(.tag) Then
                .col = column
                Call gridCOLORnormal(gridACTIVE, Val(.tag))
            End If
            If .cols <= column Then Exit Sub
            .col = column
            .tag = ""
            found = False
            
            For i = 1 To .Rows - 1
                word = Trim(UCase(.TextMatrix(i, column)))
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    Call gridCOLORdark(gridACTIVE, i)
                    .tag = .row
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
                    .TopRow = Val(.tag)
                End If
            End If
        End With
    End If
End Sub
'Addec by Muzammil 03/04/04
'Check if the sublocation exist
Function DoesItemExist(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column) As Boolean
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    
    If cellACTIVE <> "" Then
        With gridACTIVE
            
            For i = 1 To .Rows - 1
                If .cols = 1 And column >= 1 Then Exit Function 'Juan 2013-12-28 added to avoid exception
                word = Trim(UCase(.TextMatrix(i, column)))
                
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    DoesItemExist = True
                    Exit Function
                End If
            
            Next
            cellACTIVE = ""
        End With
    ElseIf cellACTIVE = "" Then
    
        found = True
    
    End If
    DoesItemExist = found
End Function
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

Public Function getDATA(Access, parameters) As ADODB.Recordset
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = Access
        Set getDATA = .Execute(, parameters)
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
            'WarehouseToWarehouse,Sales
            Case "02040400", "02040500", "02040700", "02050300", "02040600", "02050400", "02040300"
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 3), .TextMatrix(.row, 4), .row)
            'AdjustmentIssue juan 2012--3-24 to add serial
            Case "02040200"
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 4), .TextMatrix(.row, 5), .row)
            Case "02040100" 'WarehouseReceipt
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 7), .TextMatrix(.row, 4), .TextMatrix(.row, 3), .row)
            Case "02050200" 'AdjustmentEntry
                Call fillDETAILlist(.TextMatrix(.row, 1), .TextMatrix(.row, 2), .TextMatrix(.row, 3), -1, .row)
        End Select
    End With
Screen.MousePointer = 0
End Sub


Public Function loadFQA(Companycode As String, Optional LocationCode As String) As Boolean

On Error GoTo ErrHand
loadFQA = False
Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

RsCompany.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Level ='C'"

RsCompany.Open , cn

Do While RsCompany.EOF

    SSOleDBFQA.addITEM RsCompany("FQA")
    RsCompany.MoveNext
    
Loop


'Get Location FQA

RsLocation.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='L'"

RsLocation.Open , cn


Do While RsLocation.EOF

    SSOleDBFQA.addITEM RsCompany("FQA")
    RsLocation.MoveNext
    
Loop


'Get US Chart FQA

RsUC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='UC'"

RsUC.Open , cn


Do While RsUC.EOF

    SSOleDBFQA.addITEM RsUC("FQA")
    RsUC.MoveNext
    
Loop

'Get Cam Chart FQA

RsCC.source = "select FQA from FQA where Namespace ='" & nameSP & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='CC'"

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






