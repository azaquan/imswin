VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Warehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum kindTRANSAC
    ReturnFromRepair
    AdjustmentEntry
    WarehouseIssue
    WellToWell
    InternalTransfer
    AdjustmentIssue
    WarehouseToWarehouse
    WarehouseReceipt
    Sales
    ReturnFromWell
    GlobalTransfer
    Fabrication
End Enum
Public Enum EmailClients
   Unknown = 0
   ATT
   Outlook
End Enum

Public frm As Object

Public Property Get Connection() As ADODB.Connection
    Set Connection = cn
End Property
Public Property Set Connection(conn As ADODB.Connection)
    Set cn = conn
End Property
Public Property Let cUid(ByRef u As String)
    uid = u
End Property
Public Property Let cPwd(ByRef p As String)
    pwd = p
End Property
Function getName(var) As String
    Select Case var
        Case Fabrication
            getName = "Fabrication"
        Case ReturnFromRepair
            getName = "ReturnFromRepair"
        Case AdjustmentEntry
            getName = "AdjustmentEntry"
        Case WarehouseIssue
            getName = "WarehouseIssue"
        Case WellToWell
            getName = "WellToWell"
        Case InternalTransfer
            getName = "InternalTransfer"
        Case AdjustmentIssue
            getName = "AdjustmentIssue"
        Case WarehouseToWarehouse
            getName = "WarehouseToWarehouse"
        Case WarehouseReceipt
            getName = "WarehouseReceipt"
        Case Sales
            getName = "Sales"
        Case ReturnFromWell
            getName = "ReturnFromWell"
        Case GlobalTransfer
            getName = "GlobalTransfer"
    End Select
End Function

Public Property Let initialCatalog(ByRef i As String)
    InitCatalog = i
End Property
Public Property Let cEmailOutFolder(ByRef e As String)
    emailOutFolder = e
End Property
Public Property Let NameSpace(ByRef NS As String)
    nameSP = NS
End Property
Public Property Let NameSpaceName(ByRef NSname As String)
    nameSPname = NSname
End Property
Public Property Let currUSER(ByRef user As String)
    CurrentUser = user
End Property
Public Property Let reportPATH(ByRef repoP As String)
    repoPATH = repoP
End Property
Public Property Let dsnFILE(dsn As String)
    dsnF = dsn
End Property
Public Property Let language(langSEL As String)
    If UCase(langSEL) = "US" Then
        TR_LANGUAGE = "US"
    Else
        TR_LANGUAGE = langSEL
    End If
End Property

Public Property Let ExtendedCurrency(ByRef CurrencyCode As String) 'M
    GExtendedCurrency = CurrencyCode
End Property

Public Sub Loading(transaction As kindTRANSAC)
Dim title, tag
Screen.MousePointer = 11

    Load frmWarehouse
    With frmWarehouse
        .Visible = False
        'cell 0 (common)
        .matrix.TextMatrix(0, 0) = "GETISSUENUMBER"
        .matrix.TextMatrix(1, 0) = 1
        .matrix.TextMatrix(2, 0) = "ii_trannumb"
        .matrix.TextMatrix(3, 0) = "ii_trannumb"
        .matrix.TextMatrix(4, 0) = "ii_trannumb"
        .matrix.TextMatrix(5, 0) = "0"
        .matrix.TextMatrix(7, 0) = ""
        '.matrix.TextMatrix(8, 0) = "Transaction #"
        .matrix.TextMatrix(8, 0) = getName("cel0")
        .matrix.TextMatrix(9, 0) = "1200"
        .matrix.TextMatrix(10, 0) = 0
        
        'cell 1 (common)
        .matrix.TextMatrix(0, 1) = "GETCOMPANY"
        .matrix.TextMatrix(1, 1) = 2
        .matrix.TextMatrix(2, 1) = "com_name, com_compcode"
        .matrix.TextMatrix(3, 1) = "com_name"
        .matrix.TextMatrix(4, 1) = "com_compcode"
        .matrix.TextMatrix(5, 1) = "0, 0"
        .matrix.TextMatrix(6, 1) = nameSP
        .matrix.TextMatrix(7, 1) = ""
        '.matrix.TextMatrix(8, 1) = "Company, Code"
        .matrix.TextMatrix(8, 1) = translator.getIt("cel1")
        .matrix.TextMatrix(9, 1) = "3000, 1000"
        .matrix.TextMatrix(10, 1) = 1
        
        'cell 2(common)
        .matrix.TextMatrix(8, 2) = translator.getIt("cel2")
        
        'cell 3(common)
        .matrix.TextMatrix(8, 3) = translator.getIt("cel3")
        
        Select Case transaction
            Case Fabrication
                title = "Fabrication"
                tag = "02040800"
                
                'cell 0
                frmFabrication.matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                frmFabrication.matrix.TextMatrix(1, 0) = 1
                frmFabrication.matrix.TextMatrix(2, 0) = "ir_trannumb"
                frmFabrication.matrix.TextMatrix(3, 0) = "ir_trannumb"
                frmFabrication.matrix.TextMatrix(4, 0) = "ir_trannumb"
                frmFabrication.matrix.TextMatrix(5, 0) = "0"
                frmFabrication.matrix.TextMatrix(6, 0) = nameSP + ", @1, F"
                frmFabrication.matrix.TextMatrix(7, 0) = ""
                frmFabrication.matrix.TextMatrix(9, 0) = "1200"
                frmFabrication.matrix.TextMatrix(10, 0) = 0
                                
                'cell 2
                frmFabrication.matrix.TextMatrix(0, 2) = "Get_Location2"
                frmFabrication.matrix.TextMatrix(1, 2) = 2
                frmFabrication.matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                frmFabrication.matrix.TextMatrix(3, 2) = "loc_name"
                frmFabrication.matrix.TextMatrix(4, 2) = "loc_locacode"
                frmFabrication.matrix.TextMatrix(5, 2) = "0, 0"
                frmFabrication.matrix.TextMatrix(6, 2) = "FABRICA, TRUE"
                frmFabrication.matrix.TextMatrix(7, 2) = ""
                frmFabrication.matrix.TextMatrix(9, 2) = "4000, 1000"
                frmFabrication.matrix.TextMatrix(10, 2) = 1
                
                '.Label(2).Caption = "From"
                frmFabrication.label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")

                
                'cell 3
                frmFabrication.matrix.TextMatrix(0, 3) = "Get_Location2"
                frmFabrication.matrix.TextMatrix(1, 3) = 2
                frmFabrication.matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                frmFabrication.matrix.TextMatrix(3, 3) = "loc_name"
                frmFabrication.matrix.TextMatrix(4, 3) = "loc_locacode"
                frmFabrication.matrix.TextMatrix(5, 3) = "0, 0"
                frmFabrication.matrix.TextMatrix(6, 3) = "BASE, TRUE"
                frmFabrication.matrix.TextMatrix(7, 3) = ""
                frmFabrication.matrix.TextMatrix(9, 3) = "4000, 1000"
                frmFabrication.matrix.TextMatrix(10, 3) = 1
                
                '.Label(3).Caption = "To Warehouse"
                frmFabrication.label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                frmFabrication.label(3).Visible = True
                frmFabrication.label(3).ZOrder
                frmFabrication.cell(3).Visible = True
                frmFabrication.cell(3).ZOrder
                
                'cell 4
                frmFabrication.label(4).Visible = False
                frmFabrication.cell(4).Visible = False
                
            Case ReturnFromRepair
                title = "Return From Repair"
                tag = "02040400"
                
                'cell 0
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, RR"
                .matrix.TextMatrix(7, 0) = ""
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                                
                'cell 2
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "REPAIR, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                
                '.Label(2).Caption = "From"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                'cell 3
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1
                
                '.Label(3).Caption = "To Warehouse"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                .label(3).Visible = True
                .cell(3).Visible = True
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case AdjustmentEntry
                title = "Adjustment Entry"
                tag = "02050200"
                
                'cell 0
                .cell(2).Top = .cell(1).Top
                .label(2).Top = .label(1).Top
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, AE"
                .matrix.TextMatrix(7, 0) = ""
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                
                'cell 2
                .label(2).Caption = "Warehouse"
                .matrix.TextMatrix(0, 2) = "AdjustmentEntry"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = ""
                .matrix.TextMatrix(7, 2) = ""
                '.matrix.TextMatrix(8, 2) = "Location, Code"
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                .label(3).Visible = False
                .cell(3).Visible = False
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case WarehouseIssue
                title = "Warehouse Issue"
                tag = "02040200"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, I"
                .matrix.TextMatrix(8, 0) = ""
                
                'cell 2
                '.Label(2).Caption = "From Warehouse"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "BASE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                '.matrix.TextMatrix(8, 2) = "Location, Code"
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                '.Label(3).Caption = "Issue To"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                .label(3).Visible = True
                .cell(3).Visible = True
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "SITE, REPAIR, FABRICA, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                '.matrix.TextMatrix(8, 3) = "Location, Code"
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1

                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case WellToWell
                title = "Well to Well"
                tag = "02040500"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, TI"
                .matrix.TextMatrix(8, 0) = ""
                
                'cell 2
                '.Label(2).Caption = "From Well"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "SITE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                '.matrix.TextMatrix(8, 2) = "Location, Code"
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                '.Label(3).Caption = "To Well"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                
                .label(3).Visible = True
                .cell(3).Visible = True
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "SITE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                '.matrix.TextMatrix(8, 3) = "Location, Code"
                .matrix.TextMatrix(9, 3) = "4000,1000"
                .matrix.TextMatrix(10, 3) = 1
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case InternalTransfer
                title = "Internal Transfer"
                tag = "02040700"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, IT"
                .matrix.TextMatrix(8, 0) = ""
                
                'cell 2
                .cell(2).Top = .cell(1).Top
                .label(2).Top = .label(1).Top
                '.Label(2).Caption = "Warehouse"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "BASE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                .label(3).Visible = False
                .cell(3).Visible = False
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case AdjustmentIssue
                title = "Adjustment Issue"
                tag = "02050300"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, AI"
                .matrix.TextMatrix(8, 0) = ""
                
                'cell 2
                '.Label(2).Caption = "Warehouse"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                .cell(2).Top = .cell(1).Top
                .label(2).Top = .label(1).Top
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "BASE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                .label(3).Visible = False
                .cell(3).Visible = False
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case WarehouseToWarehouse
                title = "Warehouse To Warehouse"
                tag = "02040600"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, TI"
                .matrix.TextMatrix(8, 0) = ""
                                
                'cell 2
                .label(2).Caption = "From Warehouse"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "BASE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                '.Label(3).Caption = "To Warehouse"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                
                .label(3).Visible = True
                .cell(3).Visible = True
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                '.matrix.TextMatrix(8, 3) = "Location, Code"
                .matrix.TextMatrix(9, 3) = "4000,1000"
                .matrix.TextMatrix(10, 3) = 1

                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
            Case WarehouseReceipt
                title = "Warehouse Receipt"
                tag = "02040100"
                        
                'cell 0
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, R"
                .matrix.TextMatrix(7, 0) = ""
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                
                'cell 2
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "OTHER, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                '.Label(2).Caption = "Receipt From"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                'cell 3
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1
                '.Label(3).Caption = "To Warehouse"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                
                .label(3).Visible = True
                .cell(3).Visible = True
                
                'cell 4
                '.Label(4).Caption = "Purchase Order"
                .label(4).Caption = translator.getIt(getName(transaction) + ".Label(4)")
                
                .label(4).Visible = True
                .cell(4).Visible = True
                .matrix.TextMatrix(0, 4) = "query"
                .matrix.TextMatrix(1, 4) = 7
                .matrix.TextMatrix(2, 4) = "po_ponumb, PO_Date, po_buyr, po_sendby, po_apprby, po_stas, po_freigforwr"
                .matrix.TextMatrix(3, 4) = "po_ponumb"
                .matrix.TextMatrix(4, 4) = "po_ponumb"
                .matrix.TextMatrix(5, 4) = "0, 3, 0, 0, 0, 3, 0"
                .matrix.TextMatrix(6, 4) = ""
                .matrix.TextMatrix(7, 4) = ""
                '.matrix.TextMatrix(8, 4) = "PO#, Date, Buyer, Send By, Approve By, Status, Forwarded"
                .matrix.TextMatrix(8, 4) = translator.getIt(getName(transaction) + ".cel4")
                .matrix.TextMatrix(9, 4) = "1300, 900, 1400, 1400, 1400, 600, 0"
                .matrix.TextMatrix(10, 4) = 0
                
            Case Sales
                title = "Sales"
                tag = "02050400"
                
                'cell 0
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, SL"
                .matrix.TextMatrix(8, 0) = ""
                
                'cell 2
                '.Label(2).Caption = "From Warehouse"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "BASE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000,1000"
                .matrix.TextMatrix(10, 2) = 1
                
                'cell 3
                .label(3).Caption = "Sold to"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                
                .label(3).Visible = True
                .cell(3).Visible = True
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "OTHER, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1
                
                'cell 4
                '.Label(4).Caption = "Sold To"
                .label(4).Caption = translator.getIt(getName(transaction) + ".Label(3)") 'same as label(4)
                
                .label(4).Visible = True
                .cell(4).Visible = True
                .matrix.TextMatrix(0, 4) = "suppliers"
                .matrix.TextMatrix(1, 4) = 2
                .matrix.TextMatrix(2, 4) = "sup_name, sup_code"
                .matrix.TextMatrix(3, 4) = "sup_name"
                .matrix.TextMatrix(4, 4) = "sup_code"
                .matrix.TextMatrix(5, 4) = "0, 0"
                .matrix.TextMatrix(6, 4) = ""
                .matrix.TextMatrix(7, 4) = ""
                .matrix.TextMatrix(8, 4) = "Supplier, Code"
                .matrix.TextMatrix(9, 4) = "4000, 1600"
                .matrix.TextMatrix(10, 4) = 1
                
            Case ReturnFromWell
                title = "Return from Well"
                tag = "02040300"
                
                'cell 0
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, RT"
                .matrix.TextMatrix(7, 0) = ""
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                                
                'cell 2
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "SITE, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                
                '.label(2).Caption = "From Site"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")
                
                'cell 3
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1
                .label(3).Caption = "To Warehouse"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                
                .label(3).Visible = True
                .cell(3).Visible = True
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
            Case Else
                Unload frmWarehouse
                Exit Sub
        End Select
        'Juan (2015-02-16) for Multilingual
        msg1 = translator.getIt(getName(transaction) + ".title")
        title = IIf(msg1 = "", title, msg1)
        '---------------------------------------------
        
        .Caption = title + " " + tag
        .tag = tag
        Set frm = frmWarehouse
        .Show 1
        
    End With
Screen.MousePointer = 0
End Sub
Public Sub Loading3(transaction As kindTRANSAC)
Dim title, tag
Screen.MousePointer = 11

    Load frmFabrication
    With frmFabrication
        .Visible = False
        'cell 0 (common)
        .matrix.TextMatrix(0, 0) = "GETISSUENUMBER"
        .matrix.TextMatrix(1, 0) = 1
        .matrix.TextMatrix(2, 0) = "ii_trannumb"
        .matrix.TextMatrix(3, 0) = "ii_trannumb"
        .matrix.TextMatrix(4, 0) = "ii_trannumb"
        .matrix.TextMatrix(5, 0) = "0"
        .matrix.TextMatrix(7, 0) = ""
        '.matrix.TextMatrix(8, 0) = "Transaction #"
        .matrix.TextMatrix(8, 0) = getName("cel0")
        .matrix.TextMatrix(9, 0) = "1200"
        .matrix.TextMatrix(10, 0) = 0
        
        'cell 1 (common)
        .matrix.TextMatrix(0, 1) = "GETCOMPANY"
        .matrix.TextMatrix(1, 1) = 2
        .matrix.TextMatrix(2, 1) = "com_name, com_compcode"
        .matrix.TextMatrix(3, 1) = "com_name"
        .matrix.TextMatrix(4, 1) = "com_compcode"
        .matrix.TextMatrix(5, 1) = "0, 0"
        .matrix.TextMatrix(6, 1) = nameSP
        .matrix.TextMatrix(7, 1) = ""
        '.matrix.TextMatrix(8, 1) = "Company, Code"
        .matrix.TextMatrix(8, 1) = translator.getIt("cel1")
        .matrix.TextMatrix(9, 1) = "3000, 1000"
        .matrix.TextMatrix(10, 1) = 1
        
        'cell 2(common)
        .matrix.TextMatrix(8, 2) = translator.getIt("cel2")
        
        'cell 3(common)
        .matrix.TextMatrix(8, 3) = translator.getIt("cel3")
        
                title = "Fabrication"
                tag = "02040800"
                
                'cell 0
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, F"
                .matrix.TextMatrix(7, 0) = ""
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                                
                'cell 2
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "FABRICA, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                
                '.Label(2).Caption = "From"
                .label(2).Caption = translator.getIt(getName(transaction) + ".Label(2)")

                
                'cell 3
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1
                
                '.Label(3).Caption = "To Warehouse"
                .label(3).Caption = translator.getIt(getName(transaction) + ".Label(3)")
                .label(3).Visible = True
                .label(3).ZOrder
                .cell(3).Visible = True
                .cell(3).ZOrder
                
                'cell 4
                .label(4).Visible = False
                .cell(4).Visible = False
                
                
        'Juan (2015-02-16) for Multilingual
        msg1 = translator.getIt(getName(transaction) + ".title")
        title = IIf(msg1 = "", title, msg1)
        '---------------------------------------------
        
        .Caption = title + " " + tag
        .tag = tag
        Set frm = frmWarehouse
        .Show 1
        
    End With
Screen.MousePointer = 0
End Sub


'
'
Public Sub Loading2(transaction As kindTRANSAC)
Dim title, tag
Screen.MousePointer = 11

    Load frmGlobalWH
    With frmGlobalWH
        .Visible = False
        'cell 0 (common)
        .matrix.TextMatrix(0, 0) = "GETISSUENUMBER"
        .matrix.TextMatrix(1, 0) = 1
        .matrix.TextMatrix(2, 0) = "ii_trannumb"
        .matrix.TextMatrix(3, 0) = "ii_trannumb"
        .matrix.TextMatrix(4, 0) = "ii_trannumb"
        .matrix.TextMatrix(5, 0) = "0"
        .matrix.TextMatrix(7, 0) = ""
        '.matrix.TextMatrix(8, 0) = "Transaction #"
        .matrix.TextMatrix(8, 0) = translator.getIt(getName("cel0"))
        .matrix.TextMatrix(9, 0) = "1200"
        .matrix.TextMatrix(10, 0) = 0
        
        'cell 1 (common)
        .matrix.TextMatrix(0, 1) = "GETCOMPANY"
        .matrix.TextMatrix(1, 1) = 2
        .matrix.TextMatrix(2, 1) = "com_name, com_compcode"
        .matrix.TextMatrix(3, 1) = "com_name"
        .matrix.TextMatrix(4, 1) = "com_compcode"
        .matrix.TextMatrix(5, 1) = "0, 0"
        .matrix.TextMatrix(6, 1) = nameSP
        .matrix.TextMatrix(7, 1) = ""
        '.matrix.TextMatrix(8, 1) = "Company, Code"
        .matrix.TextMatrix(8, 1) = translator.getIt("cel1")
        .matrix.TextMatrix(9, 1) = "3000, 1000"
        .matrix.TextMatrix(10, 1) = 1
        
        Select Case transaction
            Case GlobalTransfer
                title = "Global Transfer"
                'Juan (2015-02-27) for Multilingual
                msg1 = translator.getIt(getName(transaction) + ".title")
                title = IIf(msg1 = "", title, msg1)
                '---------------------------------------------
                        
                tag = "02050801"
                
                'cell 0
                .matrix.TextMatrix(0, 0) = "GETRECEPTIONNUMBER"
                .matrix.TextMatrix(1, 0) = 1
                .matrix.TextMatrix(2, 0) = "ir_trannumb"
                .matrix.TextMatrix(3, 0) = "ir_trannumb"
                .matrix.TextMatrix(4, 0) = "ir_trannumb"
                .matrix.TextMatrix(5, 0) = "0"
                .matrix.TextMatrix(6, 0) = nameSP + ", @1, RR"
                .matrix.TextMatrix(7, 0) = ""
                '.matrix.TextMatrix(8, 0) = "Transaction #"
                .matrix.TextMatrix(8, 0) = translator.getIt(getName(transaction) + ".cel0")
                .matrix.TextMatrix(9, 0) = "1200"
                .matrix.TextMatrix(10, 0) = 0
                                
                'cell 2
                .matrix.TextMatrix(0, 2) = "Get_Location2"
                .matrix.TextMatrix(1, 2) = 2
                .matrix.TextMatrix(2, 2) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 2) = "loc_name"
                .matrix.TextMatrix(4, 2) = "loc_locacode"
                .matrix.TextMatrix(5, 2) = "0, 0"
                .matrix.TextMatrix(6, 2) = "REPAIR, TRUE"
                .matrix.TextMatrix(7, 2) = ""
                '.matrix.TextMatrix(8, 2) = "Location, Code"
                .matrix.TextMatrix(8, 2) = translator.getIt(getName(transaction) + ".cel3")
                .matrix.TextMatrix(9, 2) = "4000, 1000"
                .matrix.TextMatrix(10, 2) = 1
                
                .label(2).Caption = "From Company"
                
                'cell 3
                .matrix.TextMatrix(0, 3) = "Get_Location2"
                .matrix.TextMatrix(1, 3) = 2
                .matrix.TextMatrix(2, 3) = "loc_name, loc_locacode"
                .matrix.TextMatrix(3, 3) = "loc_name"
                .matrix.TextMatrix(4, 3) = "loc_locacode"
                .matrix.TextMatrix(5, 3) = "0, 0"
                .matrix.TextMatrix(6, 3) = "BASE, TRUE"
                .matrix.TextMatrix(7, 3) = ""
                '.matrix.TextMatrix(8, 3) = "Location, Code"
                .matrix.TextMatrix(8, 3) = translator.getIt(getName(transaction) + ".cel3")
                .matrix.TextMatrix(9, 3) = "4000, 1000"
                .matrix.TextMatrix(10, 3) = 1


            Case Else
                Unload frmGlobalWH
                Exit Sub
        End Select
        
        .Caption = title + " " + tag
        .tag = tag
        Set frm = frmGlobalWH
        .Show 1
        
    End With
Screen.MousePointer = 0
End Sub

