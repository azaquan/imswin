VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDI_IMS 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "IMS for Windows?"
   ClientHeight    =   8190
   ClientLeft      =   1140
   ClientTop       =   -450
   ClientWidth     =   11880
   Icon            =   "MDI_IMS.frx":0000
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrStateMonitor 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   4200
   End
   Begin VB.Timer tmrPeriod 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog cmdDialog 
      Left            =   4560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6324
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      UserName        =   "sa"
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Menu mnu_MasterFile 
      Caption         =   "&Master Files"
      Tag             =   "01000000"
      Begin VB.Menu itm_Purchasing 
         Caption         =   "Purchasing"
         Tag             =   "01010000"
         Begin VB.Menu itm_Supplier 
            Caption         =   "Supplier Utility"
            Tag             =   "01010100"
            Begin VB.Menu itm_InterSupply 
               Caption         =   "Supplier Utility"
               Tag             =   "01010101"
            End
            Begin VB.Menu itm_PrintIntSupply 
               Caption         =   "Print International Supplier Records"
               Tag             =   "01010102"
            End
            Begin VB.Menu itm_PrintLocalSupply 
               Caption         =   "Print Local Supplier Records"
               Tag             =   "01010103"
            End
            Begin VB.Menu itm_ListSupCodeused 
               Caption         =   "List of Supplier Code used in Transaction Order"
               Tag             =   "01010104"
            End
            Begin VB.Menu itm_ListSupCodeNotUsed 
               Caption         =   "Local Supplier"
               Tag             =   "01010105"
            End
         End
         Begin VB.Menu itm_Shipper 
            Caption         =   "Shipper Utility"
            Tag             =   "01010200"
         End
         Begin VB.Menu itm_ShipTerms 
            Caption         =   "Shipment Terms && Conditions"
            Tag             =   "01010300"
         End
         Begin VB.Menu itm_Originator 
            Caption         =   "Originator"
            Tag             =   "01010400"
         End
         Begin VB.Menu itm_CurrencyTable 
            Caption         =   "Currency"
            Tag             =   "01010500"
         End
         Begin VB.Menu itm_ListPriority 
            Caption         =   "Shipping Mode"
            Tag             =   "01010600"
         End
         Begin VB.Menu itm_DocumentType 
            Caption         =   "Document Type"
            Tag             =   "01010700"
         End
         Begin VB.Menu itm_ServiceUtility 
            Caption         =   "Service Utility"
            Tag             =   "01010800"
         End
         Begin VB.Menu itm_CustomCategory 
            Caption         =   "Custom Category"
            Tag             =   "01010900"
         End
         Begin VB.Menu itm_termofdelivery 
            Caption         =   "Terms of Delivery"
            Tag             =   "01011000"
         End
         Begin VB.Menu itm_termandcondition 
            Caption         =   "Terms && Conditions"
            Tag             =   "01011100"
         End
         Begin VB.Menu itm_forwarder 
            Caption         =   "Forwarder"
            Tag             =   "01011200"
         End
         Begin VB.Menu itm_ToBe 
            Caption         =   "To Be Used For Utility"
            Tag             =   "01011300"
         End
         Begin VB.Menu itm_phone 
            Caption         =   "Phone Directory"
            Tag             =   "01011400"
         End
         Begin VB.Menu itm_unit 
            Caption         =   "Unit"
            Tag             =   "01011500"
         End
         Begin VB.Menu itm_category 
            Caption         =   "Category"
            Tag             =   "01011600"
         End
         Begin VB.Menu itm_manufacturer 
            Caption         =   "Manufacturer"
            Tag             =   "01011700"
         End
         Begin VB.Menu itm_groupe 
            Caption         =   "Group"
            Tag             =   "01011800"
         End
         Begin VB.Menu itm_servicecode 
            Caption         =   "Service Code Category"
            Tag             =   "01011900"
         End
         Begin VB.Menu itm_charge 
            Caption         =   "Charge Account"
            Tag             =   "01012000"
         End
         Begin VB.Menu itm_stocktype 
            Caption         =   "Stock Type"
            Enabled         =   0   'False
            Tag             =   "01012100"
         End
      End
      Begin VB.Menu itm_Logistics 
         Caption         =   "Logistics"
         Tag             =   "01020000"
         Begin VB.Menu itm_BillTo 
            Caption         =   "Bill To"
            Tag             =   "01020100"
         End
         Begin VB.Menu itm_ShipTo 
            Caption         =   "Ship To"
            Tag             =   "01020200"
         End
         Begin VB.Menu itm_SoldTo 
            Caption         =   "Sold To"
            Tag             =   "01020300"
         End
         Begin VB.Menu itm_Destination 
            Caption         =   "Destination"
            Tag             =   "01020400"
         End
      End
      Begin VB.Menu itm_Warehouse 
         Caption         =   "Inventory Management"
         Tag             =   "01030000"
         Begin VB.Menu itm_TransactionType 
            Caption         =   "Transaction type"
            Tag             =   "01030100"
         End
         Begin VB.Menu itm_PhoneDirectory2 
            Caption         =   "Phone Directory"
            Tag             =   "01030200"
         End
         Begin VB.Menu itm_Country 
            Caption         =   "Country"
            Tag             =   "01030300"
         End
         Begin VB.Menu itm_Location 
            Caption         =   "Location"
            Tag             =   "01030400"
         End
         Begin VB.Menu itm_LocationSite 
            Caption         =   "Location / Site"
         End
         Begin VB.Menu itm_Logicals 
            Caption         =   "Logical Warehouse"
            Tag             =   "01030500"
         End
         Begin VB.Menu itm_SubLocation 
            Caption         =   "Sub-Location"
            Tag             =   "01030600"
         End
         Begin VB.Menu itm_Condition 
            Caption         =   "Condition"
            Tag             =   "01030700"
         End
         Begin VB.Menu itm_Company 
            Caption         =   "Company"
            Tag             =   "01030800"
         End
      End
      Begin VB.Menu itm_Other 
         Caption         =   "System"
         Tag             =   "01040000"
         Begin VB.Menu itm_status 
            Caption         =   "Status"
            Tag             =   "01040100"
         End
         Begin VB.Menu itm_Site 
            Caption         =   "Site"
            Tag             =   "01040200"
         End
         Begin VB.Menu itm_siteconsolidation 
            Caption         =   "Site Consolidation"
            Tag             =   "01040300"
         End
         Begin VB.Menu itm_KeysTable 
            Caption         =   "Autonumbering"
            Tag             =   "01040400"
         End
         Begin VB.Menu itm_electronicdistributionsystem 
            Caption         =   "Electronic Distribution (System)"
            Tag             =   "01040500"
         End
         Begin VB.Menu itm_electronicdistributionuser 
            Caption         =   "Electronic Distribution (User)"
            Tag             =   "01040600"
         End
         Begin VB.Menu itm_systemfile 
            Caption         =   "System File"
            Tag             =   "01040700"
         End
      End
      Begin VB.Menu itm_Exit 
         Caption         =   "E&xit"
         Tag             =   "01050000"
      End
   End
   Begin VB.Menu mnu_Activities 
      Caption         =   "&Activities"
      Tag             =   "02000000"
      Begin VB.Menu itm_Catalog 
         Caption         =   "Cataloging"
         Tag             =   "02010000"
         Begin VB.Menu itm_Modify 
            Caption         =   "Create/Modify Stock Record"
            Tag             =   "02010100"
         End
         Begin VB.Menu itm_Search 
            Caption         =   "Search on Stock Records"
            Tag             =   "02010200"
         End
      End
      Begin VB.Menu itm_Purchasing2 
         Caption         =   "Purchasing"
         Tag             =   "02020000"
         Begin VB.Menu itm_TransOrder 
            Caption         =   "Create/Revise Order"
            Tag             =   "02020100"
         End
         Begin VB.Menu itm_TransOrderMsg 
            Caption         =   "Create Order Tracking Message"
            Tag             =   "02020200"
         End
         Begin VB.Menu itm_TransOrderClose 
            Caption         =   "Close/Cancel Order"
            Tag             =   "02020300"
         End
         Begin VB.Menu itm_PrintOrder 
            Caption         =   "Print Order"
            Tag             =   "02020400"
         End
         Begin VB.Menu itm_TransOrderTBA 
            Caption         =   "Approve && Send Order"
            Tag             =   "02020500"
         End
         Begin VB.Menu itm_generalstatusreporttransaction 
            Caption         =   "General Status Report (by transaction)"
            Tag             =   "02020600"
         End
      End
      Begin VB.Menu itm_LogisticsA 
         Caption         =   "Logistics"
         Tag             =   "02030000"
         Begin VB.Menu itm_FreightReception 
            Caption         =   "Receive Freight"
            Tag             =   "02030100"
         End
         Begin VB.Menu itm_PackingManage 
            Caption         =   "Create Shipping Manifest"
            Tag             =   "02030200"
         End
         Begin VB.Menu itm_PackingTracking 
            Caption         =   "Create Shipping Manifest Tracking Message"
            Tag             =   "02030300"
         End
      End
      Begin VB.Menu itm_WarehouseA 
         Caption         =   "Inventory Management"
         Tag             =   "02040000"
         Begin VB.Menu itm_WReceipt 
            Caption         =   "Order Receipt"
            Tag             =   "02040100"
         End
         Begin VB.Menu itm_WIssue 
            Caption         =   "Issue"
            Tag             =   "02040200"
         End
         Begin VB.Menu itm_WReturnWell 
            Caption         =   "Return from Well Site"
            Tag             =   "02040300"
         End
         Begin VB.Menu itm_WReturnRepair 
            Caption         =   "Return from Repair"
            Tag             =   "02040400"
         End
         Begin VB.Menu itm_Well_Well 
            Caption         =   "Well to Well Transfer"
            Tag             =   "02040500"
         End
         Begin VB.Menu itm_WWarehouse_Warehouse 
            Caption         =   "Warehouse to Warehouse Transfer "
            Tag             =   "02040600"
         End
         Begin VB.Menu itm_WInternalTransfer 
            Caption         =   "Logical Warehouse-Sub Location Movement"
            Tag             =   "02040700"
         End
         Begin VB.Menu itm_WGlobalTransfer 
            Caption         =   "Global Transfer"
         End
      End
      Begin VB.Menu itm_accounting 
         Caption         =   "Financial Management"
         Tag             =   "02050000"
         Begin VB.Menu ita_WInitialLoad 
            Caption         =   "Inventory Initial Load"
            Tag             =   "02050100"
         End
         Begin VB.Menu itm_WQuantityAdjOn 
            Caption         =   "Inventory Write On"
            Tag             =   "02050300"
         End
         Begin VB.Menu itm_WQuantityAdjustmentOff 
            Caption         =   "Inventory Write Off"
            Tag             =   "02050200"
         End
         Begin VB.Menu itm_WSale 
            Caption         =   "Sale"
            Tag             =   "02050400"
         End
         Begin VB.Menu ita_SAPinquiry 
            Caption         =   "SAP inquiry"
            Tag             =   "02050500"
         End
         Begin VB.Menu itm_CondCodeValuation 
            Caption         =   "Condition Code Valuation"
            Tag             =   "02050600"
         End
         Begin VB.Menu itm_supinv 
            Caption         =   "Supplier Invoice Input"
            Tag             =   "02050700"
         End
         Begin VB.Menu itm_TranValuationRep 
            Caption         =   "Transaction Valuation Report"
            Tag             =   "02050800"
         End
         Begin VB.Menu itm_monthend 
            Caption         =   "SAPAdjustment"
            Tag             =   "02050900"
         End
         Begin VB.Menu itm_SAPAnalysisrep 
            Caption         =   "SAP Analysis Report"
            Tag             =   "02051000"
         End
         Begin VB.Menu itm_AuditSAPValuation 
            Caption         =   "Audit SAP Valuation"
            Tag             =   "02051100"
         End
      End
   End
   Begin VB.Menu mnu_report 
      Caption         =   "&Reports"
      Tag             =   "03000000"
      Begin VB.Menu itmr_catalog 
         Caption         =   "Cataloging"
         Tag             =   "03010000"
         Begin VB.Menu itmr_stockmaster 
            Caption         =   "Stock Master"
            Tag             =   "03010100"
         End
         Begin VB.Menu itmr_manustockXref 
            Caption         =   "Manufacturer/Stock Number X-Reference"
            Tag             =   "03010200"
         End
         Begin VB.Menu itmr_stockmasterExcel 
            Caption         =   "Export Stock Master to Excel"
            Tag             =   "03010300"
         End
      End
      Begin VB.Menu itmr_pruchasing 
         Caption         =   "Purchasing"
         Tag             =   "03020000"
         Begin VB.Menu itmr_TranOrderrep 
            Caption         =   "Print Order"
            Tag             =   "03020100"
         End
         Begin VB.Menu itm_stockhistory 
            Caption         =   "Stock Number History"
            Tag             =   "03020200"
         End
         Begin VB.Menu itm_OpenOrder 
            Caption         =   "Open Order"
            Tag             =   "03020300"
         End
         Begin VB.Menu itm_Orderactivity 
            Caption         =   "Order Activity Report"
            Tag             =   "03020400"
         End
         Begin VB.Menu itmr_OrderTracking 
            Caption         =   "Order Tracking Record"
            Tag             =   "03020500"
         End
         Begin VB.Menu itmr_Orderdeliveryschedule 
            Caption         =   "Order Delivery Schedule"
            Tag             =   "03020600"
         End
         Begin VB.Menu itmr_latedelivery 
            Caption         =   "Late Delivery Report"
            Tag             =   "03020700"
         End
         Begin VB.Menu itmr_lateshipping 
            Caption         =   "Late Shipping Report"
            Tag             =   "03020800"
         End
         Begin VB.Menu itmr_OrderAuitLog 
            Caption         =   "Order Audit Log"
            Tag             =   "03020900"
         End
         Begin VB.Menu itm_generalstatusreportreq 
            Caption         =   "General Status Report (by req)"
            Tag             =   "03021000"
         End
      End
      Begin VB.Menu itmr_InventoryManagement 
         Caption         =   "Inventory Management"
         Tag             =   "03030000"
         Begin VB.Menu itmr_OrderToBeReceive 
            Caption         =   "Order to be Received"
            Tag             =   "03030100"
         End
         Begin VB.Menu itmr_InventoryperStocknumber 
            Caption         =   "Inventory per Stock Number"
            Tag             =   "03030200"
         End
         Begin VB.Menu itmr_TranPerdateRange 
            Caption         =   "Transactions per Date Range"
            Tag             =   "03030300"
         End
         Begin VB.Menu itm_StockonHand 
            Caption         =   "Stock On-Hand"
            Tag             =   "03030400"
         End
         Begin VB.Menu itmr_PhysicalInventory 
            Caption         =   "Physical Inventory"
            Tag             =   "03030500"
         End
         Begin VB.Menu itmr_slowmovinginventory 
            Caption         =   "Slow Moving Inventory"
            Tag             =   "03030600"
         End
         Begin VB.Menu itmr_historicalStockMovement 
            Caption         =   "Historical Stock Movement"
            Tag             =   "03030700"
         End
         Begin VB.Menu itmr_SOHAccrosslocation 
            Caption         =   "Stock On-Hand Across All Locations"
            Tag             =   "03030800"
         End
         Begin VB.Menu itmr_SOHperSTcokNmber 
            Caption         =   "Stock On-Hand per Stock Number"
            Tag             =   "03030900"
         End
      End
      Begin VB.Menu itmr_financialManagement 
         Caption         =   "Financial Management"
         Tag             =   "03040000"
         Begin VB.Menu itmr_SAPvaluationInquiry 
            Caption         =   "SAP Valuation Inquiry"
            Tag             =   "03040100"
         End
         Begin VB.Menu itmr_TranvaluationReport 
            Caption         =   "Transaction Valuation Report"
            Tag             =   "03040200"
         End
         Begin VB.Menu itmr_SAPAnalysisReport 
            Caption         =   "SAP Analysis Report"
            Tag             =   "03040300"
         End
      End
      Begin VB.Menu mnu_reportsecurity 
         Caption         =   "Security"
         Tag             =   "03050000"
         Begin VB.Menu itm_accesslevel 
            Caption         =   "Access Level"
            Tag             =   "03050100"
         End
         Begin VB.Menu itm_applicationuserstatus 
            Caption         =   "Application User Status"
            Tag             =   "03050200"
         End
         Begin VB.Menu itm_loginlogoff 
            Caption         =   "Login/Logoff"
            Tag             =   "03050300"
         End
         Begin VB.Menu itm_securitychangelog 
            Caption         =   "Security Changes Log"
            Tag             =   "03050400"
         End
         Begin VB.Menu itm_individualuserprofile 
            Caption         =   "Individual User Profile"
            Tag             =   "03050500"
         End
         Begin VB.Menu itm_accesslevelbuyeruser 
            Caption         =   "Access Level + Buyer + User"
            Tag             =   "03050600"
         End
      End
   End
   Begin VB.Menu mnu_System 
      Caption         =   "&System"
      Tag             =   "04000000"
      Begin VB.Menu itm_security 
         Caption         =   "Security Utility"
         Tag             =   "04010000"
         Begin VB.Menu itm_BuyersTable 
            Caption         =   "Buyers table utility/User application rights"
            Tag             =   "04010100"
         End
         Begin VB.Menu itm_userprofile 
            Caption         =   "User Profile"
            Tag             =   "04010200"
         End
         Begin VB.Menu itm_InitialUser 
            Caption         =   "Initial user password settings"
            Tag             =   "04010300"
         End
         Begin VB.Menu itm_ChangePassword 
            Caption         =   "Change personal password"
            Tag             =   "04010400"
         End
         Begin VB.Menu itm_TemporaryPassword 
            Caption         =   "Temporary Password"
            Tag             =   "04010500"
         End
         Begin VB.Menu itm_menuoption 
            Caption         =   "Menu Option"
            Tag             =   "04010600"
         End
         Begin VB.Menu itm_menulevel 
            Caption         =   "Menu Level"
            Tag             =   "04010700"
         End
         Begin VB.Menu itm_MenuTemplate 
            Caption         =   "Menu Template"
            Tag             =   "04010800"
         End
         Begin VB.Menu itm_useraccesslevel 
            Caption         =   "User Access Level"
            Tag             =   "04010900"
         End
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Tag             =   "05000000"
      Begin VB.Menu itm_colors 
         Caption         =   "Colors"
         Tag             =   "05010000"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Tag             =   "06000000"
      Begin VB.Menu itm_Help 
         Caption         =   "Help"
         Tag             =   "06010000"
      End
      Begin VB.Menu itm_About 
         Caption         =   "About"
         Tag             =   "06020000"
      End
   End
End
Attribute VB_Name = "MDI_IMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''' Idle time API used for unlocking and closing Application / Form added by jawdat 1.31.02
Private Type POINTAPI
    x As Long
    y As Long
    End Type
   
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'''Private Const INTERVAL As Long = 10 'where "10" = length of time when pc becommes idle
Dim IsIdle As Boolean, IsIdle2 As Boolean 'True when idling or While in idle-state
Dim MousePos As POINTAPI 'holds mouse position
Dim startOfIdle As Long, startofidle2 As Long
Dim CountForTimer As Integer, j As Integer, i As Integer, idleStateEngagedFlag As Boolean
''''''''''''End idle time API calls


'Declare a Constant DoubleQuote which is
'the Ascii value 34, or " (double quote)
Public Property Get DoubleQuote() As String
    DoubleQuote = Chr(34)
End Property

'set crystal report parameters and load report

Private Sub itm_accesslevelbuyeruser_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\menubig.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00199") 'J added
        .WindowTitle = IIf(msg1 = "", "Access Level-user-buyer", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'set crystal report parameters and load report

Private Sub itm_applicationuserstatus_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\userstatus.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00383") 'J added
        .WindowTitle = IIf(msg1 = "", "User Status", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

Private Sub itm_globalTransfer_Click()

End Sub

'load form and show it

Private Sub itm_individualuserprofile_Click()
    Load frm_individualuserprofile
    frm_individualuserprofile.Show
End Sub

'load form and show it

Private Sub itm_InterSupply_Click()
On Error Resume Next

'    Load the International Supplier form
    Screen.MousePointer = vbHourglass

    Load frm_IntSupe
    frm_IntSupe.ZOrder 0

    Call frm_IntSupe.Move(0, 0)

    frm_IntSupe.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub itm_LocationSite_Click()
    Load frm_LocationSITE
    frm_LocationSITE.ZOrder
    frm_LocationSITE.Show
End Sub

'load login or off form and show it

Private Sub itm_loginlogoff_Click()
    Load frm_loginlogoff
    frm_loginlogoff.Show
End Sub

'set crystal report parameters and load report

Private Sub itm_PrintIntSupply_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Intsupp.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00384") 'J added
        .WindowTitle = IIf(msg1 = "", "Internal Supplier", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'set crystal report parameters and load report

Private Sub itm_PrintLocalSupply_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Locsupp.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00385") 'J added
        .WindowTitle = IIf(msg1 = "", "Local Supplier", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'set crystal report parameters and load report

Private Sub itm_ListSupCodeused_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Usedsupp.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00386") 'J added
        .WindowTitle = IIf(msg1 = "", "Used Supplier", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'set crystal report parameters and load report

Private Sub itm_ListSupCodeNotUsed_Click()
'On Error GoTo ErrHandler
'    'Report call
'    With MDI_IMS.CrystalReport1
'        .Reset
'        .ReportFileName = FixDir(App.Path) + "CRreports\Notusedsupp.rpt"
'        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'
'        'Modified by Juan (9/27/2000) for Multilingual
'        msg1 = translator.Trans("M00387") 'J added
'        .WindowTitle = IIf(msg1 = "", "Not Used Supplier", msg1) 'J modified
'        '---------------------------------------------
'
'        .Action = 1: .Reset
'    End With

'    Load frm_LocationSITE   changed, 2.13.02 M
   Screen.MousePointer = vbHourglass

    Load frm_LocSupe
    frm_LocSupe.ZOrder 0

    Call frm_LocSupe.Move(0, 0)

    frm_LocSupe.Show
    Screen.MousePointer = vbDefault
End Sub
'set crystal report parameters and load report

Private Sub itm_securitychangelog_Click()
    Load frmSecurityChangeLog
    frmSecurityChangeLog.Show
'On Error GoTo ErrHandler
'    'Report call
'    With MDI_IMS.CrystalReport1
'        .Reset
'        .ReportFileName = FixDir(App.Path) + "CRreports\securitchange.rpt"
'        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'
'        'Modified by Juan (9/27/2000) for Multilingual
'        msg1 = translator.Trans("M00388") 'J added
'        .WindowTitle = IIf(msg1 = "", "Security change", msg1) 'J modified
'        '---------------------------------------------
'
'        .Action = 1: .Reset
'    End With
'        Exit Sub
'
'ErrHandler:
'    If Err Then
'        MsgBox Err.Description
'        Err.Clear
'    End If
End Sub

'load form

Private Sub itm_Shipper_Click()
On Error Resume Next
Dim ctl As Control

    'Load the Shipper form
    Load frm_Shipper
    frm_Shipper.ZOrder 0
    'Set Caption for the Shipper
    'form window
'    frm_Shipper.Caption = "Shipper"
    frm_Shipper.Visible = True
End Sub

'set crystal report parameters and load report

Private Sub itm_ShipTerms_Click()
    Load frm_ShiptermsEdit
    frm_ShiptermsEdit.ZOrder 0
'    frm_ShiptermsEdit.Caption = "Ship Terms & Condition"
    frm_ShiptermsEdit.Visible = True
End Sub

'load form

Private Sub itm_Originator_Click()
    Load frm_Originator
    frm_Originator.ZOrder 0
'    frm_Originator.Caption = "Originator"
    frm_Originator.Visible = True
End Sub

'load form and show it

Private Sub itm_CurrencyTable_Click()
    'Load the Currency form
    Call LockWindowUpdate(HWND)
    
    Load frmCurrency
    frmCurrency.Show
    Call LockWindowUpdateOff
End Sub

'load form and set to visible

Private Sub itm_ListPriority_Click()
    Load frm_Priority
    frm_Priority.ZOrder 0
    frm_Priority.Visible = True
End Sub

'load form and show it

Private Sub itm_DocumentType_Click()
    'Load the Document Type form
    Load frm_Document
    frm_Document.ZOrder 0
'    frm_Document.Caption = "Document Type"
    frm_Document.Show
End Sub

'load form and show it

Private Sub itm_ServiceUtility_Click()
    Load frm_ServiceCode
    frm_ServiceCode.ZOrder 0
'    frm_ServiceCode.Caption = "Service Codes"
    frm_ServiceCode.Show
End Sub

'load form and show it

Private Sub itm_CustomCategory_Click()
    'Load the Custom form
    Load frm_Custom
    frm_Custom.ZOrder 0
   ' Call Move(0, 0, width, height)
    'Set Caption for the Custom
    'form window
'    frm_Custom.Caption = "Custom Category"
    frm_Custom.Show
End Sub

'load form and show it

Private Sub itm_termofdelivery_Click()
    Load frmTermDelivery
    frmTermDelivery.ZOrder 0
'    frmTermDelivery.Caption = "Term of Delivery"
    frmTermDelivery.Show

End Sub

'load form and show it

Private Sub itm_termandcondition_Click()
    Load frmTermCondition
    frmTermCondition.ZOrder 0
'    frmTermCondition.Caption = "Terms of Condition"
    frmTermCondition.Show
End Sub

'load form and show it

Private Sub itm_forwarder_Click()
    frmForwarder.Show
End Sub

'load form and show it

Private Sub itm_ToBe_Click()
    Load frm_ToBe
    frm_ToBe.ZOrder 0
'    frm_ToBe.Caption = "To Be"
    frm_ToBe.Show
End Sub

'load form and show it

Private Sub itm_phone_Click()
    Load frm_Phone
    frm_Phone.ZOrder 0
'    frm_Phone.Caption = "Phone Directory"
    frm_Phone.Show
End Sub

'load form and show it

Private Sub itm_Unit_Click()
    'Load the Unit form
    Load frm_Unit
    frm_Unit.ZOrder 0
'    frm_Unit.Caption = "Unit"
    frm_Unit.Show
End Sub

'load form and show it

Private Sub itm_Category_Click()
    'Load the Category form
    Load frm_Category
    frm_Category.ZOrder 0
    'Set Caption for the Category form
'    frm_Category.Caption = "Category"
    frm_Category.Show
End Sub

'load form and show it

Private Sub itm_Manufacturer_Click()
'    'Load the manufacturer form
    Load frm_Manufacturer
    frm_Manufacturer.ZOrder 0
    frm_Manufacturer.Show
    'Set Caption for the Manufacturer
'    'form window.show
End Sub

'load form and show it

Private Sub itm_groupe_Click()
    'Load the Group Code form
    Load frm_GroupCode
    frm_GroupCode.ZOrder 0
    'Set Caption for the Group Code
    'form window
'    frm_GroupCode.Caption = "Group Code"
    frm_GroupCode.Show
End Sub

'load form and show it

Private Sub itm_servicecode_Click()
    Load frmServiceCate
    frmServiceCate.ZOrder 0
'    frmServiceCate.Caption = "Service Code Category"
    frmServiceCate.Show
End Sub

'load form and show it

Private Sub itm_charge_Click()
    'Load the Charge form
    Load frm_Charge
    frm_Charge.ZOrder 0
    'Set Caption for the Charge form
'    frm_Charge.Caption = "Charge"
    frm_Charge.Show
End Sub

Private Sub itm_StockType_Click()
    'Load frm_StockType
    'frm_StockType.ZOrder 0
    'frm_StockType.Caption = "Stock Type"
    'frm_StockType.Show
End Sub

'load form and show it

Private Sub itm_BillTo_Click()
    'Load the Bill to: form
    Load frm_Billto
    frm_Billto.ZOrder
    frm_Billto.Show
End Sub

'load form and show it

Private Sub itm_ShipTo_Click()
    'Load the Ship To form
    Load frm_ShipTo
    frm_ShipTo.ZOrder 0
'    frm_ShipTo.Caption = "Ship To:"
    frm_ShipTo.Show
End Sub

'load form and show it

Private Sub itm_SoldTo_Click()
    Load frm_SoldTo
    frm_SoldTo.ZOrder 0
'    frm_SoldTo.Caption = "Sold To"
    frm_SoldTo.Show
End Sub

'load form and show it

Private Sub itm_Destination_Click()
    'Load the Destination form
    Load frm_Destination
    frm_Destination.ZOrder 0
'    frm_Destination.Caption = "Destination"
    frm_Destination.Show
End Sub

'load form and show it

Private Sub itm_TransactionType_Click()
'    Load frm_Transaction
    frmTrantype.ZOrder 0
'    frmTrantype.Caption = "Transaction Type"
    frmTrantype.Show
End Sub

'load form and show it

Private Sub itm_PhoneDirectory2_Click()
    Load frm_Phone
    frm_Phone.ZOrder 0
'    frm_Phone.Caption = "Phone Directory"
    frm_Phone.Show
End Sub

'load form and show it

Private Sub itm_Country_Click()
    'Load the Country form
    Load frm_Country
    frm_Country.ZOrder 0
    'Set Caption for the Country
    'form window
'    frm_Country.Caption = "Country"
    frm_Country.Show
End Sub

'load form and show it

Private Sub itm_Location_Click()
    'Load the manufacturer form
    Load frm_Location
    frm_Location.ZOrder 0
    'Set Caption for the Manufacturer
    'form window
'    frm_Location.Caption = "Location"
    frm_Location.Show
End Sub

'load form and show it

Private Sub itm_Logicals_Click()
    'Load the Logicals form
    Load frm_logical
    frm_logical.ZOrder 0
    'Set Caption for the Logicals
    'form window
'    frm_Logicals.Caption = "Logicals Warehouse"
    frm_logical.Show
End Sub

'load form and show it

Private Sub itm_SubLocation_Click()
    Load frm_SubLocation
    frm_SubLocation.ZOrder 0
'    frm_SubLocation.Caption = "Sub-Location"
    frm_SubLocation.Show
End Sub

'load form and show it

Private Sub itm_Condition_Click()
    'Load the Condition form
    Load frm_Condition
    frm_Condition.ZOrder 0
    frm_Condition.Show
End Sub

'load form and show it

Private Sub itm_Company_Click()
    'Load the Company form
    Load frm_Company
    frm_Company.ZOrder 0
    'Set Caption for the Company
    'form window
'    frm_Company.Caption = "Company"
    frm_Company.Show
End Sub

'load form and show it

Private Sub itm_status_Click()
    Load frmStatus
    frmStatus.ZOrder 0
'    frmStatus.Caption = "Status"
    frmStatus.Show
End Sub

'load form and show it

Private Sub itm_Site_Click()
    'Load the SiteDescript form
    Load frm_SiteDescript

    frm_SiteDescript.ZOrder 0
'    frm_SiteDescript.Caption = "Site Description"
    frm_SiteDescript.Show
End Sub

'load form and show it

Private Sub itm_siteconsolidation_Click()
    Load frmSiteConsolidation
    frmSiteConsolidation.ZOrder 0
'    frmSiteConsolidation.Caption = "Site Consolidation"
    frmSiteConsolidation.Show
End Sub

'load form and show it

Private Sub itm_KeysTable_Click()
    Load frmChrono
    frmChrono.Show
End Sub

'load form and show it

Private Sub itm_electronicdistributionsystem_Click()
    Load frmElecDistribution
    frmElecDistribution.Show
End Sub

'load form and show it

Private Sub itm_electronicdistributionuser_Click()
    Load frmEUserDistribution
    frmEUserDistribution.Show
End Sub

'load form and show it

Private Sub itm_systemfile_Click()
    Load frm_systemfile
    frm_systemfile.Show
End Sub

'close form free memory

Private Sub itm_Exit_Click()
    Unload Me
    Set MDI_IMS = Nothing
End Sub

'load form and show it

Private Sub itm_Modify_Click()
    Load Frm_StockMaster
    Call Frm_StockMaster.Move(0, 0)
    Frm_StockMaster.Show
    mDidUserOpenStkMasterForm = True
End Sub

'load form and show it

Private Sub itm_Search_Click()
    Load frm_StockSearch
    frm_StockSearch.ZOrder 0
    Call frm_StockSearch.Move(0, 0)
    frm_StockSearch.Show
End Sub

'load form and show it

Private Sub itm_TransOrder_Click()
    Screen.MousePointer = vbHourglass

    With frm_NewPurchase
        Call .Move(0, 0)
        .Visible = True
    End With
    Screen.MousePointer = vbNormal
End Sub

'load form and show it

Private Sub itm_TransOrderMsg_Click()
    Load Frm_TrackingPONew
    Frm_TrackingPONew.Show
End Sub

'load form and show it

Private Sub itm_TransOrderClose_Click()
    Load frmClose
    frmClose.Show
End Sub

'load form and show it

Private Sub itm_PrintOrder_Click()
    Load frm_transact_order
    frm_transact_order.Show
End Sub

'load form and show it

Private Sub itm_TransOrderTBA_Click()
    Call frmPOApproval.Show(vbModeless)
End Sub

'load form and show it

Private Sub itm_generalstatusreporttransaction_Click()
    Load frm_gnrlstatustransac
    frm_gnrlstatustransac.Show
End Sub

'load form and show it

Private Sub itm_FreightReception_Click()
    frmReception.Show
End Sub

'load form and show it

Private Sub itm_PackingManage_Click()
    Load frmPackingList
    frmPackingList.Show

End Sub

'load form and show it

Private Sub itm_PackingTracking_Click()
    Load frmTrackManifest
    frmTrackManifest.Show
End Sub

'get crystal report parameters and load report

Private Sub itm_useraccesslevel_Click()
Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.ShowMenuOptions(mfUser, deIms.NameSpace, deIms.cnIms)
    If Err Then Call MsgBox(Err.Description, vbCritical)
    
    Err.Clear
    Set SC = Nothing
End Sub

Private Sub itm_WGlobalTransfer_Click()
    Call frmNavigator.lblSubWharehousing_Click(0)
End Sub

'load form and show it

Private Sub itm_WReceipt_Click()
    Call frmNavigator.lblSubWharehousing_Click(0)
End Sub

'load form and show it

Private Sub itm_WIssue_Click()
    Call frmNavigator.lblSubWharehousing_Click(2)
End Sub

'load form and show it

Private Sub itm_WReturnWell_Click()
    Call frmNavigator.lblSubWharehousing_Click(3)
End Sub

'load form and show it

Private Sub itm_WReturnRepair_Click()
    Call frmNavigator.lblSubWharehousing_Click(4)
End Sub

'load form and show it

Private Sub itm_Well_Well_Click()
    Call frmNavigator.lblSubWharehousing_Click(5)
End Sub

'load form and show it

Private Sub itm_WWarehouse_Warehouse_Click()
    Call frmNavigator.lblSubWharehousing_Click(6)
End Sub

'load form and show it

Private Sub itm_WInternalTransfer_Click()
    Call frmNavigator.lblSubWharehousing_Click(7)
End Sub

'load form and show it

Private Sub ita_WInitialLoad_Click()
    Load frmWHInitialAdjustment
    Call frmWHInitialAdjustment.Move(0, 0)
    frmWHInitialAdjustment.Visible = True
'    frmWHInitialAdjustment.Caption = "Inventory Initial Load"
End Sub

'load form and show it

Private Sub itm_WQuantityAdjOn_Click()
    Call frmNavigator.lblSubAccounting_Click(2)
End Sub

'load form and show it

Private Sub itm_WQuantityAdjustmentOff_Click()
    Call frmNavigator.lblSubAccounting_Click(1)
End Sub

'load form and show it

Private Sub itm_WSale_Click()
    Call frmNavigator.lblSubAccounting_Click(3)
End Sub

'load form and show it

Private Sub ita_SAPinquiry_Click()
    Load frm_sap_inquiry
    frm_sap_inquiry.Show
End Sub

'load form and show it

Private Sub itm_CondCodeValuation_Click()
    'Load the Condition form
    Load frm_Condition
    frm_Condition.ZOrder 0
    frm_Condition.Show
End Sub

'load form and show it

Private Sub itm_supinv_Click()
    Load frmInvoice
    Call frmInvoice.Move(0, 0)
    frmInvoice.Show
End Sub

'load form and show it

Private Sub itm_TranValuationRep_Click()
   Load frm_tranvaluationreport
   frm_tranvaluationreport.Show
   frm_tranvaluationreport.ZOrder
End Sub

'load form and show it

Private Sub itm_monthend_Click()
    frmSapAdjustment.Show
End Sub

'load form and show it

Private Sub itm_SAPAnalysisrep_Click()
    frm_sap_analysis.Show
End Sub

Private Sub itm_AuditSAPValuation_Click()

    'Modified by Juan (9/27/2000) for Multilingual
    msg1 = translator.Trans("M00088") 'J added
    MsgBox IIf(msg1 = "", ("does not exist yet"), msg1) 'J modified
    '---------------------------------------------
    
End Sub

'load form and show it

Private Sub itmr_SAPAnalysisReport_Click()
    frm_sap_analysis.Show
End Sub

'load form and show it

Private Sub itmr_SAPvaluationInquiry_Click()
    frm_sap_inquiry.Show
End Sub

'get crystal report parameters and load form

Private Sub itmr_stockmaster_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\stckmaster.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00166") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock Master", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'get crystal report parameters and load form

Private Sub itmr_manustockXref_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Xcrossmanu.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00389") 'J added
        .WindowTitle = IIf(msg1 = "", "X reference Manufacturer", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'get crystal report parameters and load form

Private Sub itmr_stockmasterExcel_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\stckmasterX.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00390") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock Master to Excel", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form and show it

Private Sub itmr_TranOrderrep_Click()
    Load frm_transact_order
    frm_transact_order.Show
End Sub

'load form and show it

Private Sub itm_stockhistory_Click()
  Load frm_stockhistory
  frm_stockhistory.Show
End Sub

'get crystal report parameters and load form

Private Sub itm_OpenOrder_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\orderopen.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00172") 'J added
        .WindowTitle = IIf(msg1 = "", "Open Orders", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form and show it

Private Sub itm_Orderactivity_Click()
    Load frm_order_activity
    frm_order_activity.Show
End Sub

'load form and show it

Private Sub itmr_OrderTracking_Click()
    Load frm_ordertracking
    frm_ordertracking.Show
End Sub

'load form and show it

Private Sub itmr_Orderdeliveryschedule_Click()
Load frm_orderdelivery
frm_orderdelivery.Show
End Sub

'load form and show it

Private Sub itmr_latedelivery_Click()
  Load frm_latedelivery
frm_latedelivery.Show
End Sub

'load form and show it

Private Sub itmr_lateshipping_Click()
   Load frm_lateshipping
   frm_lateshipping.Show
End Sub

Private Sub itmr_OrderAuitLog_Click()

    'Modified by Juan (9/27/2000) for Multilingual
    msg1 = translator.Trans("M00088") 'J added
    MsgBox IIf(msg1 = "", ("does not exist yet"), msg1) 'J modified
    '---------------------------------------------

End Sub

Private Sub itm_generalstatusreportreq_Click()

    'Modified by Juan (9/27/2000) for Multilingual
    msg1 = translator.Trans("M00088") 'J added
    MsgBox IIf(msg1 = "", ("does not exist yet"), msg1) 'J modified
    '---------------------------------------------

End Sub

'get crystal report parameters and load form

Private Sub itmr_OrderToBeReceive_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\ordertoberevcd.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00181") 'J added
        .WindowTitle = IIf(msg1 = "", "Order to be received", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form and show it

Private Sub itmr_InventoryperStocknumber_Click()
'    MsgBox "Feature does not exist as yet"
   Load frm_inventoryperstocknu
   frm_inventoryperstocknu.Show
End Sub

'load form and show it

Private Sub itmr_TranPerdateRange_Click()
 Load frm_tranperdaterange
 frm_tranperdaterange.Show
End Sub

'load form and show it

Private Sub itm_StockonHand_Click()
    Load StockOnHand
    StockOnHand.ZOrder 0
    StockOnHand.Visible = True
    StockOnHand.Show
End Sub

'load form and show it

Private Sub itmr_PhysicalInventory_Click()
'Load frm_physicalinventory
' frm_physicalinventory.Show
    Load StockOnHandNew
    StockOnHandNew.Show
End Sub

'load form and show it

Private Sub itmr_slowmovinginventory_Click()
   Load frm_slowmoving
    frm_slowmoving.Show
End Sub

'load form and show it

Private Sub itmr_historicalStockMovement_Click()
  Load frm_historicalstock
  frm_historicalstock.Show
End Sub

'load form and show it

Private Sub itmr_SOHAccrosslocation_Click()
 Load frm_sohaccrosslocation
 frm_sohaccrosslocation.Show
End Sub

'load form and show it

Private Sub itmr_SOHperSTcokNmber_Click()
    Load frmStockOnHandStock
    frmStockOnHandStock.ZOrder 0
    frmStockOnHandStock.Visible = True
    frmStockOnHandStock.Show
End Sub

'load form and show it

Private Sub itm_About_Click()
Dim About As ImsCmDlg
Dim str As String

    Set About = New ImsCmDlg
        
    str = "IMS Inventory Tracking " & App.Major & "." & App.Minor & App.Revision
    Call About.ShowAbout(Icon, str, "", "Copyright ? 1999 - 2005 IMS")
    Set About = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_BuyersTable_Click()
On Error Resume Next

Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.ShowBuyers(deIms.NameSpace, deIms.cnIms)
    If Err Then Call MsgBox(Err.Description, vbCritical)
    
    Err.Clear
    Set SC = Nothing

End Sub

'load form and show it

Private Sub itm_colors_Click()
    frm_Color.Show
    frm_Color.ZOrder 0
'    frm_Color.Caption = "Color Scheme"
End Sub


'call function get help file

Private Sub itm_Help_Click()
    Call ShowHelpContents(HWND, App.HelpFile, 0)
End Sub

'get crystal report parameters and load form

Private Sub itm_TemporaryPassword_Click()
Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.NameSpace = deIms.NameSpace
    Set SC.Connection = deIms.cnIms
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.AssignTempOwnerPassWord(CurrentUser, False)
    
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_InitialUser_Click()
Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.NameSpace = deIms.NameSpace
    Set SC.Connection = deIms.cnIms
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.AssignTempOwnerPassWord(CurrentUser, True)
    
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_UserProfile_Click()
Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.NameSpace = deIms.NameSpace
    Set SC.Connection = deIms.cnIms
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.AddUser(CurrentUser)
    
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_ChangePassword_Click()
Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.NameSpace = deIms.NameSpace
    Set SC.Connection = deIms.cnIms
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    If SC.CanChangePassword(deIms.NameSpace, CurrentUser, deIms.cnIms) Then
        SC.ChangePassword
    Else
    
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00382") 'J added
        MsgBox IIf(msg1 = "", "Your Password is not old enough", msg1) 'J modified
        '---------------------------------------------
    

    End If
    
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_MenuOption_Click()
On Error Resume Next

Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.ShowMenuOptions(mfOption, deIms.NameSpace, deIms.cnIms)
    If Err Then Call MsgBox(Err.Description, vbCritical)
    
    Err.Clear
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_MenuLevel_Click()
On Error Resume Next

Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.ShowMenuOptions(mfLevel, deIms.NameSpace, deIms.cnIms)
    If Err Then Call MsgBox(Err.Description, vbCritical)
    
    Err.Clear
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_MenuTemplate_Click()
On Error Resume Next

Dim SC As imsSecMod

    Set SC = New imsSecMod
    SC.UserName = CurrentUser
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Call SC.ShowMenuOptions(mfTemplate, deIms.NameSpace, deIms.cnIms)
    If Err Then Call MsgBox(Err.Description, vbCritical)
    
    Err.Clear
    Set SC = Nothing
End Sub

'get crystal report parameters and load form

Private Sub itm_Accesslevel_Click()
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\accesslevel.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00194") 'J added
        .WindowTitle = IIf(msg1 = "", "Access Level", msg1) 'J modified
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form and show it

Private Sub itmr_TranvaluationReport_Click()
    Load frm_tranvaluationreport
    frm_tranvaluationreport.Show
End Sub

'load form set back ground color and get user level
'set user menu

Private Sub dllform_Click()
 
End Sub

Private Sub MDIForm_Activate()
    tmrPeriod.Enabled = True
    tmrStateMonitor = True
End Sub

Private Sub MDIForm_Load()
Dim str As String
Dim ctl As Control
Dim rs As New ADODB.Recordset

Dim i As IMSFile
Dim Dsnname, LANG As String
On Error Resume Next

m_OutlookLocation = "C:\OutLook\"

    'Added by Juan Gonzalez (8/29/2000) for Multilingual
    translator.Translate_Forms ("MDI_IMS")
    '------------------------------------------------------
    
    Call Move(0, 0)
    Call LockWindowUpdate(HWND)

    Call LogExec("Loading Background")
        
    frm_Color.ChangeFormsColor
            
    frmNavigator.Show
    
    'Hide
    DoEvents: DoEvents
    Call StayOnTop(frm_Load.HWND, True)
    
    Set i = New IMSFile
    str = FixDir(App.Path) & "RDCPrint.exe"
    
    If i.FileExists(str) Then _
        Call LaunchApp(str & " /RegServer", "", SW_NORMAL, True)
        
    str = ""
    Set i = Nothing
    
    'Changed by Juan Gonzalez for translation requires
    'rs.Source = "select mu_meopid,mu_accsflag,menuOPTION.mo_meopname from menuuser,menuOPTION where mu_npecode ='" & deIms.NameSpace & "'and mo_npecode= mu_npecode and mu_userid ='" & CurrentUser & "'and menuoption.mo_meopid = menuuser.mu_meopid"
    If translator.TR_LANGUAGE = "*" Then
        LANG = "US"
    Else
        LANG = translator.TR_LANGUAGE
    End If
    
    rs.Source = "SELECT MENUUSER.mu_meopid, MENUUSER.mu_accsflag, TRMENU.msg_lang, " _
        & "TRMENU.msg_text AS mo_meopname, TRMENU.trs_obj FROM MENUUSER INNER JOIN " _
        & "TRMENU ON MENUUSER.mu_meopid = TRMENU.trs_obj WHERE (MENUUSER.mu_npecode = " _
        & "'" & deIms.NameSpace & "') AND (MENUUSER.mu_userid = '" & CurrentUser & "') AND (TRMENU.msg_lang = '" & LANG & "') "
    rs.ActiveConnection = deIms.cnIms
    rs.Open
    msg1 = translator.Trans("M00220")
    Call frm_Load.ShowMessage(IIf(msg1 = "", "Processing Navigator Options", msg1))
    msg1 = translator.Trans("M00221")
    msg2 = translator.Trans("M00222")
    Call LogExec(IIf(msg1 = "", "Retrieving Menu Options", msg1) & vbCrLf & IIf(msg2 = "", "Processing Form navigator", msg2))
    '-------------------------------------------------
    
  'Enabling and Disabling the Labels in frmNavigator form
    
  For Each ctl In frmNavigator.Controls
  
  
  If TypeOf ctl Is Label Then
  rs.MoveFirst
  
    DoEvents: DoEvents
  
    'Modified by Juan Gonzalez (8/29/200) for Multilingual and performance
    'Do While Not (rs.EOF)
        'DoEvents
        'If ctl.Tag = rs!mu_meopid And ctl.ForeColor = &HC00000 Then
            'ctl.Caption = rs!mo_meopname
            'ctl.Enabled = rs!mu_accsflag
            'If rs!mu_accsflag = 0 Then
                'ctl.MousePointer = 1
                'ctl.FontUnderline = False
            'Else
                'ctl.MousePointer = 99
                'ctl.MouseIcon = LoadResPicture(101, vbResCursor)
            'End If
        'End If
        'rs.MoveNext
    'Loop
    rs.MoveFirst
    rs.Find "mu_meopid = '" + ctl.Tag + "'"
    If rs.EOF Then
        ctl.Enabled = False
    Else
        If ctl.Tag = rs!mu_meopid And ctl.ForeColor = &HC00000 Then
            ctl.Caption = rs!mo_meopname
            ctl.Enabled = rs!mu_accsflag
            If rs!mu_accsflag = 0 Then
                ctl.MousePointer = 1
                ctl.FontUnderline = False
            Else
                ctl.MousePointer = 99
                ctl.MouseIcon = LoadResPicture(101, vbResCursor)
            End If
        End If
    End If
    '----------------------------------------
    
  End If
    Next
    
  
  
   'Enabling and Disabling the Labels in Mdi form depending upon
   'the specified user privilages
   
    'Modified by Juan Gonzalez (8/29/2000) for Translations fix
    msg1 = translator.Trans("M00217")
    Call LogExec(IIf(msg1 = "", "Processing Menu Items", msg1))
    Call frm_Load.ShowMessage(IIf(msg1 = "", "Processing Menu Items", msg1))
    '----------------------------------------------------------
    
    'Modified by Juan Gonzalez (8/29/2000) to improve performance
    For Each ctl In Me.Controls
        If TypeOf ctl Is Menu Then
            rs.MoveFirst
            DoEvents
            rs.MoveFirst
            rs.Find "mu_meopid = '" + ctl.Tag + "'"
            If Not rs.EOF Then
            'Do While Not (rs.EOF)
                'DoEvents
                
                'Added by Juan Gonzalez (8/29/2000) for Multilingual
                'Link from the principal navigator labels to menu
                If frmNavigator.lrhActivities.Tag = rs!mu_meopid Then frmNavigator.lrhActivities.Caption = rs!mo_meopname
                If frmNavigator.lrhReports.Tag = rs!mu_meopid Then frmNavigator.lrhReports.Caption = rs!mo_meopname
                If frmNavigator.lrhTables.Tag = rs!mu_meopid Then frmNavigator.lrhTables.Caption = rs!mo_meopname
                If frmNavigator.lrhSystem.Tag = rs!mu_meopid Then frmNavigator.lrhSystem.Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(0).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(0).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(1).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(1).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(2).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(2).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(3).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(3).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(4).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(4).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(5).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(5).Caption = rs!mo_meopname
                If frmNavigator.lblSubActivities(6).Tag = rs!mu_meopid Then frmNavigator.lblSubActivities(6).Caption = rs!mo_meopname
                If frmNavigator.lblTblSubCat(0).Tag = rs!mu_meopid Then frmNavigator.lblTblSubCat(0).Caption = rs!mo_meopname
                If frmNavigator.lblTblSubCat(1).Tag = rs!mu_meopid Then frmNavigator.lblTblSubCat(1).Caption = rs!mo_meopname
                If frmNavigator.lblTblSubCat(2).Tag = rs!mu_meopid Then frmNavigator.lblTblSubCat(2).Caption = rs!mo_meopname
                If frmNavigator.lblTblSubCat(5).Tag = rs!mu_meopid Then frmNavigator.lblTblSubCat(5).Caption = rs!mo_meopname
                If frmNavigator.lblTblSubCat(6).Tag = rs!mu_meopid Then frmNavigator.lblTblSubCat(6).Caption = rs!mo_meopname
                If frmNavigator.lblReportMenu(0).Tag = rs!mu_meopid Then frmNavigator.lblReportMenu(0).Caption = rs!mo_meopname
                If frmNavigator.lblReportMenu(1).Tag = rs!mu_meopid Then frmNavigator.lblReportMenu(1).Caption = rs!mo_meopname
                If frmNavigator.lblReportMenu(2).Tag = rs!mu_meopid Then frmNavigator.lblReportMenu(2).Caption = rs!mo_meopname
                If frmNavigator.lblReportMenu(3).Tag = rs!mu_meopid Then frmNavigator.lblReportMenu(3).Caption = rs!mo_meopname
                If frmNavigator.lblReportMenu(5).Tag = rs!mu_meopid Then frmNavigator.lblReportMenu(5).Caption = rs!mo_meopname
                '------------------------------------------------------
                
                'If ctl.Tag = rs!mu_meopid Then
                    ctl.Caption = rs!mo_meopname
                    ctl.Enabled = rs!mu_accsflag
                    'Exit Do
                'End If
                DoEvents
                'rs.MoveNext
            'Loop
            End If
        End If
    Next
    '-----------------------------------------------------------------------------
    
    
    frmNavigator.Show
    If Err Then Err.Clear
    Call LockWindowUpdate(0)
    CrystalReport1.WindowParentHandle = HWND
    MDI_IMS.CrystalReport1.DialogParentHandle = HWND
    
    DoEvents: DoEvents
    StatusBar1.Panels(4).Text = CurrentUser
    
    'Modified by Juan Gonzalez(8/29/200) for Multilingual
    msg1 = translator.Trans("M00218")
    Call LogExec(IIf(msg1 = "", "Attempting to logon on to crysatl Reports Sql Server DLL", msg1))
    Call frm_Load.ShowMessage(IIf(msg1 = "", "Attempting to logon on to crysatl Reports Sql Server DLL", msg1))
    '-------------------------------------------------------
    
    'Tag = MDI_IMS.CrystalReport1.LogOnServer("p2ssql.dll", "ims", InitialCatalog, UserID, DBPassword)
    'Tag = Tag & ";" & MDI_IMS.CrystalReport1.LogOnServer("p2sodbc.dll", "imsO", InitialCatalog, UserID, DBPassword)
    ' Dsnname = GetDSNNameFromFile   'M
     
    'Tag = MDI_IMS.CrystalReport1.LogOnServer("p2ssql.dll", ConnInfo.DSource, ConnInfo.InitCatalog, ConnInfo.UId, ConnInfo.Pwd)  'M
    Tag = Tag & ";" & MDI_IMS.CrystalReport1.LogOnServer("p2sodbc.dll", ConnInfo.Dsnname, ConnInfo.InitCatalog, ConnInfo.UId, ConnInfo.Pwd)     'M
    
  '  DSN = "dsn=" & Dsnname & ";UID=" & "" & UserID & ";PWD=" & DBPassword & ";DBQ=<CRWDC>DBQ=" & Datasource
    
    'CrystalReport1.Connect = "UID=;PWD=;DBQ=<CRWDC>DBQ="
    If Err Then Call LogErr(Name & "Form_load", Err.Description, Err)
    
    'Juan 2010-9-25 added to center the form on the screen
    With MDI_IMS
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
    '----------------------
End Sub

'Resize Form

Private Sub MDIForm_Resize()
On Error Resume Next

    If WindowState = vbMinimized Then Exit Sub
    
    'Modified by Juan Gonzalez (8/29/2000) for Multilinguales
    'If IsLoaded("frm_bkgnd") Then frm_bkgnd.Height = Height - 1800
    If IsLoaded("frmNavigator") Then
        'frmNavigator.Height = Height
        frmNavigator.Top = Int((MDI_IMS.Height - frmNavigator.Height) / 2) - 500
        frmNavigator.Left = Int((MDI_IMS.Width - frmNavigator.Width) / 2)
    End If
    'If IsLoaded("frmNavigator") And IsLoaded("frm_bkgnd") Then frmNavigator.Left = frm_bkgnd.Width
    '-----------------------------------------------------------
    
End Sub
   
'unload form close all form and free memory

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Dim id
Dim i As Integer, l As Integer

'If MsgBox("Are you sure you want to Exit?", vbCritical + vbYesNo, "Imswin") = vbYes Then
    
            id = Split(Tag, ";")
            
            If id(0) = "" Then id(0) = Tag
            Hide
                
            Call CloseAllChild
            Call CloseAllForms
            
        '    Call ReleaseMutex(MutexHandle)
        '    Call CloseHandle(MutexHandle)
            If deIms.cnIms.State And adStateOpen = adStateOpen Then
                Call InsertIntoXLogin(deIms.NameSpace, "LOGOFF", CurrentUser, deIms.cnIms)
            End If
            
            If Err Then Err.Clear
            
            l = UBound(id)
            If Err Then Err.Clear
            
            For i = 0 To l
                Call MDI_IMS.CrystalReport1.LogOffServer(id(i), True)
            Next
            
            Unload deIms
            Set deIms = Nothing
            Set MDI_IMS = Nothing
            
            'Call PostQuitMessage(0)
            If Err Then Err.Clear
'Else
   
 '      Cancel = True
  
'End If
End Sub

'set browsing status

Public Sub WriteBrowsingStatus(mode As FormMode)
Dim str As String

    Select Case mode
    
        Case 0
            str = "Browsing"
            
        Case mdCreation
            str = "Creation"
        
        Case mdModification
            str = "Modification"
            
        Case mdvisualization
            str = "Visualization"
            
        End Select
 
   StatusBar1.Panels(2).Text = str
End Sub

'function close forms and set memory free

Private Sub CloseAllChild()

    Do While Not ActiveForm Is Nothing
        Unload ActiveForm
        Set ActiveForm = Nothing
    Loop
End Sub

'close form and free memory

Public Function CloseAllForms() As Boolean
Dim frm As Form

    For Each frm In Forms
        If Not frm Is Me Then
            Unload frm
            Set frm = Nothing
        End If

    Next frm

    CloseAllForms = Forms.Count <= 1
End Function

'get crystal report parameters and load form

Public Sub SaveReport(Filename As String, Optional FileType As PrintFileTypeConstants = crptHTML32Ext)
On Error Resume Next

    With CrystalReport1
        .ProgressDialog = False
        .Destination = crptToFile
        .PrintFileType = FileType
        .PrintFileName = Filename
        
        .Action = 1
        .Destination = crptToWindow
        If .LastErrorNumber Then MsgBox .LastErrorString
    End With
    
    If Err Then Err.Clear
End Sub
Public Sub PrintDirectReport(Filename As String)
On Error Resume Next

    With CrystalReport1
        .ProgressDialog = False
        .Destination = crptToPrinter
        .PrintFileName = Filename
        .Action = 1
        If .LastErrorNumber Then MsgBox .LastErrorString
    End With
    
    If Err Then Err.Clear
End Sub

'set status bar

Public Sub WriteStatus(Status As String, Panel As Integer)
On Error Resume Next
    StatusBar1.Panels(Panel).Text = Status
End Sub
'


Private Sub tmrStateMonitor_Timer()
''    Dim State As Integer
''    Dim tmpPos As POINTAPI
''    Dim ret As Long
''    Dim IdleFound As Boolean
''    Dim i As Integer
''    IdleFound = False
''
''    For i = 1 To 256
''
''        State = GetAsyncKeyState(i)
''
''
''        If State = -32767 Then
''
''            IdleFound = True
''            IsIdle = False
''            IsIdle2 = False
''        End If
''    Next
''
''    ret = GetCursorPos(tmpPos)
''
''    If tmpPos.X <> MousePos.X Or tmpPos.Y <> MousePos.Y Then
''        IsIdle = False
''        IsIdle2 = False
''        IdleFound = True 'values
''         MousePos.X = tmpPos.X
''        MousePos.Y = tmpPos.Y
''    End If
''    If Not IdleFound Then
''         If Not IsIdle Then
''
''            IsIdle = True
''            startOfIdle = Timer
''
''        End If
''         If Not IsIdle2 Then
''
''            IsIdle2 = True
''            startofidle2 = Timer
''
''        End If
''    End If
End Sub



Private Sub tmrPeriod_Timer()

Dim INTERVAL As Long, INTERVAL2 As Long
Dim TimeOut_Length As New ADODB.Recordset
Dim SQL_Form_Timeout As String, Form_TimeOut, App_TimeOut
Dim psys_npecode
psys_npecode = deIms.NameSpace


If CurrentUser <> "" Then
SQL_Form_Timeout = "SELECT psys_FormTimeout, psys_Apptimeout FROM PESYS WHERE psys_npecode = '" & psys_npecode & "'"

    Set TimeOut_Length = New ADODB.Recordset
    TimeOut_Length.Open SQL_Form_Timeout, deIms.cnIms


            Form_TimeOut = TimeOut_Length("psys_FormTimeout")
            App_TimeOut = TimeOut_Length("psys_Apptimeout")


INTERVAL = Form_TimeOut * 60
INTERVAL2 = App_TimeOut * 60

Dim timer2, startofidle2
timer2 = Timer
startofidle2 = startOfIdle

    If IsIdle Then



       If Timer - startOfIdle >= INTERVAL Then
       Call IdleStateEngaged(Timer)

       IsIdle = True
       End If
   Else
 

 
On Error Resume Next

   End If



        If IsIdle2 Then

            If timer2 - startofidle2 >= INTERVAL2 Then
            Call IdleStateEngaged2(timer2)
            startofidle2 = timer2
            IsIdle2 = True
            End If
        End If
End If

End Sub

Public Sub IdleStateEngaged(ByVal IdleStartTime As Long)
idleStateEngagedFlag = True

      ''' close form and unlock, no need to call unlock functions because all of the forms unlock when they are closed

j = Forms.Count



If IsNothing(frmNavigator.SC.frm) = False Then
 Unload frmNavigator.SC.frm
 End If
If IsNothing(frmNavigator.WH.frm) = False Then
Unload frmNavigator.WH.frm
  ' Call IdleStateDisengaged(Timer)
 End If
 
If ((j > 3) And (i <> 1000)) Then
On Error Resume Next
Unload (VB.Forms(3))
End If


If ((Forms.Count = 3) And (i <> 1000)) Or idleStateEngagedFlag = True Then
Load Timeout
'Timeout.Show vbModal
i = 1000

End If



End Sub
Public Sub IdleStateEngaged2(ByVal IdleStartTime As Long)
 End
 ''' close app and no need to unlock because the user hasn't touched the machine since either SQL AGENT or close form ran
End Sub


Public Sub IdleStateDisengaged(ByVal IdleStopTime As Long)

i = 0
idleStateEngagedFlag = False
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock

Call imsLock.Update_time(deIms.cnIms, CurrentUser)


End Sub

