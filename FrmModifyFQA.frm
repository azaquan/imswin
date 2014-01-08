VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form FrmModifyFQA 
   Caption         =   "Modify FQA"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   10320
   Tag             =   "02051200"
   Begin LRNavigators.LROleDBNavBar LROleDBNavBar1 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      FirstVisible    =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NewVisible      =   0   'False
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintVisible    =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin VB.Frame FrmHeader 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   10095
      Begin VB.OptionButton OptInvoice 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier Invoice"
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Top             =   150
         Width           =   1815
      End
      Begin VB.OptionButton OptWarehouse 
         Alignment       =   1  'Right Justify
         Caption         =   "Warehouse Transaction"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   150
         Width           =   2415
      End
      Begin VB.OptionButton OptPO 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Order"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   150
         Width           =   2055
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSDBPO 
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   573
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSDbWarehouse 
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   2
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   573
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSDbInvoice 
      Height          =   330
      Left            =   5160
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   573
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBCamChart 
      Height          =   735
      Left            =   8160
      TabIndex        =   14
      Top             =   7080
      Width           =   1455
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBStockType 
      Height          =   735
      Left            =   8160
      TabIndex        =   15
      Top             =   6960
      Width           =   975
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1720
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBUsChart 
      Height          =   735
      Left            =   8160
      TabIndex        =   16
      Top             =   6600
      Width           =   1455
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBLocation 
      Height          =   735
      Left            =   8160
      TabIndex        =   17
      Top             =   6360
      Width           =   975
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1720
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleCompany 
      Height          =   735
      Left            =   8160
      TabIndex        =   18
      Top             =   6120
      Width           =   855
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   1508
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGridFQA 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   1950
      Width           =   10065
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   529
      Columns(0).Caption=   "#"
      Columns(0).Name =   "Line#"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1720
      Columns(1).Caption=   "Stockno"
      Columns(1).Name =   "Stockno"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1931
      Columns(2).Caption=   "Quantity"
      Columns(2).Name =   "Quantity"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2275
      Columns(3).Caption=   "UnitPrice"
      Columns(3).Name =   "UnitPrice"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2566
      Columns(4).Caption=   "ExtendedUnitPrice"
      Columns(4).Name =   "ExtendedUnitPrice"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1296
      Columns(5).Caption=   "Company"
      Columns(5).Name =   "Tocompany"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1429
      Columns(6).Caption=   "Location"
      Columns(6).Name =   "ToLocation"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2170
      Columns(7).Caption=   "UsChart#"
      Columns(7).Name =   "ToUsChart#"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1244
      Columns(8).Caption=   "St Type"
      Columns(8).Name =   "ToStockType"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1984
      Columns(9).Caption=   "CamChart#"
      Columns(9).Name =   "ToCamChart#"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   17754
      _ExtentY        =   9340
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   50
      TabIndex        =   2
      Top             =   1800
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Modify FQA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label LblCount 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   2415
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
      Left            =   6960
      TabIndex        =   13
      Top             =   7320
      Width           =   3060
   End
   Begin VB.Label LblInvoice 
      Caption         =   "Invoice #"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label LblWare 
      Caption         =   "W/H Transaction"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblPO 
      Caption         =   "Transaction Order"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "FrmModifyFQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 
 Enum EnmOptions
    mdpo = 1
    mdWarehouse
    mdInvoice
    
End Enum

'''Type PO
'''
'''    liitnumb As String
'''    StockNo As String
'''    Quantity As String
'''    UnitDesc As String
'''    Currdesc As String
'''    UnitPrice As String
'''    ExtUnitPrice As String
'''    Company As String
'''    Location As String
'''    ToUsChar As String
'''    ToStockType As String
'''    ToCamChar As String
'''
'''End Type
'''
'''Type Warehouse
'''
'''    LineNo
'''    StockNo
'''    qty
'''    Currency
'''    UnitPrice
'''    ExtendedCurrency
'''    EXTENDEDUNITPRICE
'''    Company
'''    Location
'''    UsChar ToStockType & vbTab & Rs!ToCamChar
'''
'''End Type
'''
'''Type Invoice
'''
'''End Type

Dim FInvoicePopulated As Boolean
Dim FPOPopulated As Boolean
Dim FWarePopulated As Boolean
Dim FPopulateCombosWithFQA As Boolean
Dim FCurrentOption As EnmOptions
Dim FChangeflag As Boolean
Dim FFormMode As FormMode
Public Function ConverFormToOption(Options As EnmOptions)

Select Case Options

Case mdpo

Case mdWarehouse

Case mdInvoice
     
End Select

End Function

Public Function GetTransaction(Options As EnmOptions)

Dim Rs As New ADODB.Recordset


Select Case Options

 Case mdpo
        
    OptWarehouse.FontBold = False
    OptPO.FontBold = True
    OptInvoice.FontBold = False
    
    lblPO.Visible = True
    SSDBPO.Visible = True
    If SSDBPO.Visible = True Then SSDBPO.SetFocus
    LblInvoice.Visible = False
    SSDbInvoice.Visible = False
    
    LblWare.Visible = False
    SSDbWarehouse.Visible = False
            
    SSGridFQA.RemoveAll
    SSGridFQA.Caption = " Transaction Order "
    SSGridFQA.Columns("stockno").Visible = True
    SSGridFQA.Columns("Quantity").Caption = "Quantity"
    SSGridFQA.Columns("UnitPrice").Caption = "Unit Price"
    SSGridFQA.Columns("extendedunitprice").Visible = False
    SSDBPO.text = ""
    LblCount.Caption = ""

    If FPOPopulated = False Then
        
        SSDBPO.RemoveAll
        Rs.Source = "select Po_ponumb from po where po_npecode ='" & deIms.NameSpace & "' and po_stas not in ('oh','ca')"
        Rs.ActiveConnection = deIms.cnIms
        Rs.Open
        
        Do While Not Rs.EOF
            
            SSDBPO.AddItem Rs!po_ponumb
            
            Rs.MoveNext
        
        Loop
        
        FPOPopulated = True
    
    End If
    
    FCurrentOption = mdpo
    
 Case mdWarehouse
    
    
    OptWarehouse.FontBold = True
    OptPO.FontBold = False
    OptInvoice.FontBold = False
    
    lblPO.Visible = False
    SSDBPO.Visible = False
    
    LblInvoice.Visible = False
    SSDbInvoice.Visible = False
    
    LblWare.Visible = True
    SSDbWarehouse.Visible = True
    SSDbWarehouse.SetFocus
    SSGridFQA.Columns("stockno").Visible = True
    SSGridFQA.Columns("extendedunitprice").Visible = True
    SSGridFQA.Columns("Quantity").Caption = "Quantity"
    SSGridFQA.Columns("UnitPrice").Caption = "Unit Price"
    SSGridFQA.Columns("extendedunitprice").Visible = True
    'To fill the tranasction no combo box
    SSDbWarehouse.text = ""
    LblCount.Caption = ""
    If FInvoicePopulated = False Then

        Rs.Source = "Select ii_trannumb 'Transactionno',ii_trantype 'Transactionttype' FROM INVTISSUE where ii_npecode='" & deIms.NameSpace & "' union"
        Rs.Source = Rs.Source & " SELECT ir_trannumb 'Transactionno',ir_trantype 'Transactionttype' FROM INVTRECEIPT where ir_npecode ='" & deIms.NameSpace & "'"
        Rs.Source = Rs.Source & " Order By 'Transactionno'"
        Rs.Open , deIms.cnIms
        SSDbWarehouse.Columns(1).Visible = False
        SSDbWarehouse.RemoveAll
        
        Do While Not Rs.EOF
        
            SSDbWarehouse.AddItem Rs!TransactionNo & vbTab & Rs!Transactionttype
            Rs.MoveNext
            
        Loop
        FInvoicePopulated = True
    End If

    'To Get the Transaction with the FQa CODE

    SSGridFQA.RemoveAll
    SSGridFQA.Caption = " Warehouse Transactions "
    
    FCurrentOption = mdWarehouse
    
 Case mdInvoice
    
    
     OptWarehouse.FontBold = False
     OptPO.FontBold = False
     OptInvoice.FontBold = True
        
     lblPO.Visible = True
     SSDBPO.Visible = True
    SSDBPO.SetFocus
     LblInvoice.Visible = True
     SSDbInvoice.Visible = True
    
     LblWare.Visible = False
     SSDbWarehouse.Visible = False
    
     SSGridFQA.RemoveAll
     SSGridFQA.Caption = " Supplier Invoice "
     SSGridFQA.Columns("Quantity").Caption = "Description"
     SSGridFQA.Columns("UnitPrice").Caption = "Charges"
     SSDBPO.text = ""
     SSDbInvoice.text = ""
     LblCount.Caption = ""
     SSGridFQA.Columns("stockno").Visible = False
     SSGridFQA.Columns("extendedunitprice").Visible = False
     FCurrentOption = mdInvoice
     
 End Select

    settheGrid

End Function

Private Sub Form_Load()
Me.Height = 8340
Me.Width = 10440
 

SSGridFQA.StyleSets.Add ("CellBeingModified")
SSGridFQA.StyleSets("CellBeingModified").BackColor = vbYellow '&H80C0FF
SSGridFQA.ActiveCell.StyleSet = "CellBeingModified"
LROleDBNavBar1.LastPrintSepVisible = False
LROleDBNavBar1.EditVisible = True
LROleDBNavBar1.EditEnabled = True
Call PopulateCombosWithFQA
FFormMode = ChangeModeOfForm(lblStatus, mdVisualization)
LROleDBNavBar1.PrintVisible = False
LROleDBNavBar1.PrintSaveSepVisible = False
OptPO.Value = True
Call GetTransaction(mdpo)
Call DisableButtons(Me, LROleDBNavBar1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
FInvoicePopulated = False
FPOPopulated = False
FWarePopulated = False
FPopulateCombosWithFQA = False
FCurrentOption = mdpo
FChangeflag = True
End Sub

Private Sub LROleDBNavBar1_BeforeCancelClick()
Dim output As VbMsgBoxResult
Dim x As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer
output = MsgBox("You will loose any changes you have made. Are you sure you want to go ahead?", vbInformation + vbYesNo, "Ims")

If output = vbYes Then


Select Case FCurrentOption

Case mdpo

        Call SSDBPO_Click
        
Case mdInvoice

        Call SSDbInvoice_Click
        
Case mdWarehouse

        Call SSDbWarehouse_Click
        
End Select

Call EnableDisableNavbar(False)

''    FChangeflag = False
''    LROleDBNavBar1.EditEnabled = True
''    LROleDBNavBar1.SaveEnabled = False
''    FFormMode = ChangeModeOfForm(lblStatus, mdVisualization)
''    FrmHeader.Enabled = True
''    SSDbInvoice.Enabled = True
''    SSDBPO.Enabled = True
''    SSDbWarehouse.Enabled = True
    
End If

End Sub

Private Sub LROleDBNavBar1_BeforeSaveClick()
Dim Rs As ADODB.Recordset
Dim RowCount As Integer
Dim i As Integer
On Error GoTo ErrHand

RowCount = SSGridFQA.Rows
Set Rs = New ADODB.Recordset
Select Case FCurrentOption

Case mdpo

    Rs.Source = "select ItemNo, ToCompany ,ToLocation, ToUsChar,  ToStockType, ToCamChar,ModiUser,ModiDate from pofqa where ponumb ='" & Trim(SSDBPO.text) & "' and npce_code = '" & deIms.NameSpace & "'"
    Rs.Open , deIms.cnIms, adOpenKeyset, adLockBatchOptimistic
    
    
Case mdInvoice

    Rs.Source = "select ""Lineno"" Itemno, ""desc"", currencycode, amount, ToCompanyFqa TOCOMPANY, ToLocationFqa TOLOCATION, ToUsChart TOUSCHAR, ToStocktype, ToCamchar, modiuser,modidate from invoicefqa where Ponumb='" & Trim(SSDBPO.text) & "' and namespace= '" & deIms.NameSpace & "'"
    Rs.Open , deIms.cnIms, adOpenKeyset, adLockBatchOptimistic
    
   
    
Case mdWarehouse

    Rs.Source = "select  ItemNo,ToCompany ,ToLocation, ToUsChar, ToStockType, ToCamChar,ModiUser,ModiDate  from inventoryfqa where  TransactionNo='" & Trim(SSDbWarehouse.text) & "' and Npce_code ='" & deIms.NameSpace & "'"
    Rs.Open , deIms.cnIms, adOpenKeyset, adLockBatchOptimistic
    
End Select

   For i = 0 To RowCount - 1
        
        SSGridFQA.row = i
        Rs.MoveFirst
        Rs.Find ("itemno ='" & Trim(SSGridFQA.Columns("line#").Value) & "'")
        
        If Rs.AbsolutePosition <> adPosEOF Then
            
            Rs!ToCompany = SSGridFQA.Columns("tocompany").Value
            Rs!Tolocation = SSGridFQA.Columns("Tolocation").Value
            Rs!ToUsChar = SSGridFQA.Columns("ToUsChart#").Value
            Rs!ToStockType = SSGridFQA.Columns("ToStockType").Value
            Rs!ToCamChar = SSGridFQA.Columns("ToCamChart#").Value
            Rs!modiuser = CurrentUser
            Rs!modidate = Now
        
        End If
    
    Next i
    
    
      Rs.UpdateBatch
      
      FChangeflag = True
      
      Call EnableDisableNavbar(False)
      
Exit Sub
ErrHand:

MsgBox "Errors Occured while trying to save. Please try again." & Err.Description, vbCritical, "Ims"

Err.Clear

LROleDBNavBar1.SaveEnabled = True

End Sub

Private Sub LROleDBNavBar1_OnCloseClick()
Unload Me
End Sub

Private Sub LROleDBNavBar1_OnEditClick()
'''
'''FrmHeader.Enabled = False
'''LROleDBNavBar1.SaveEnabled = True
'''LROleDBNavBar1.EditEnabled = False
'''FFormMode = ChangeModeOfForm(lblStatus, mdModification)
'''SSDBPO.Enabled = False
'''SSDbInvoice.Enabled = False
'''SSDbWarehouse.Enabled = False
Select Case FCurrentOption
Case mdpo

    If Len(Trim(SSDBPO.text)) = 0 Then
        MsgBox "Please make sure that a Transaction Order is selected.", vbInformation, "Ims"
        Exit Sub
    End If
        

Case mdWarehouse

    
    If Len(Trim(SSDbWarehouse.text)) = 0 Then
        MsgBox "Please make sure that a Warehouse transaction is selected.", vbInformation, "Ims"
        Exit Sub
    End If

Case mdInvoice

    
    If Len(Trim(SSDbInvoice.text)) = 0 Then
        MsgBox "Please make sure that an Invoice is selected.", vbInformation, "Ims"
        Exit Sub
    End If

End Select
Call EnableDisableNavbar(True)

End Sub

Private Sub LROleDBNavBar1_OnSaveClick()
LROleDBNavBar1.SaveEnabled = True
End Sub

Private Sub OptInvoice_Click()

''If ChangeOption = True Then
''    Call GetTransaction(mdInvoice)
''ElseIf FCurrentOption <> mdInvoice Then
''    MsgBox "Please save or cancel the changes before you switch to any other option.", vbInformation, "Ims"
''End If



    Call GetTransaction(mdInvoice)


End Sub

Private Sub OptPO_Click()

''If ChangeOption = True Then
''    Call GetTransaction(mdpo)
''ElseIf FCurrentOption <> mdpo Then
''    MsgBox "Please save or cancel the changes before you switch to any other option.", vbInformation, "Ims"
''End If


    Call GetTransaction(mdpo)



End Sub

Private Sub OptWarehouse_Click()

''If ChangeOption = True Then
''
''    Call GetTransaction(mdWarehouse)
''
''ElseIf FCurrentOption <> mdWarehouse Then
''
''    MsgBox "Please save or cancel the changes before you switch to any other option.", vbInformation, "Ims"
''
''End If
    
    Call GetTransaction(mdWarehouse)

End Sub

Private Sub SSDbInvoice_Click()
Dim Rs As New ADODB.Recordset
On Error GoTo ErrHand

     Rs.Source = " select ""Lineno"",""desc"",amount,curr_code, curr_desc, ToCompanyFqa, ToLocationFqa, ToUsChart, ToStocktype,"
     Rs.Source = Rs.Source & " ToCamchar  from invoicefqa "
     Rs.Source = Rs.Source & " inner join currency on curr_code =currencycode and curr_npecode=namespace "
     Rs.Source = Rs.Source & " where namespace ='" & deIms.NameSpace & "' and Ponumb ='" & Trim(SSDBPO.text) & "' and Invoiceno ='" & Trim(SSDbInvoice.text) & "' order by ""Lineno"""

     Rs.Open , deIms.cnIms
     
     SSGridFQA.RemoveAll
     
     Do While Not Rs.EOF
    
         SSGridFQA.AddItem Rs!LineNo & vbTab & vbTab & Rs!desc & vbTab & IIf(Trim(UCase(Rs!curr_code)) = "USD", "USD", Trim(Rs!Curr_desc)) & " " & Rs!amount & vbTab & vbTab & Rs!ToCompanyFqa & vbTab & Rs!ToLocationFqa & vbTab & Rs!ToUSChart & vbTab & Rs!ToStockType & vbTab & Rs!ToCamChar
         
         Rs.MoveNext
    
     Loop
    
    LblCount.Caption = Rs.RecordCount & " Records Found"
    
Exit Sub
ErrHand:

MsgBox "Errors occurred while trying to populate the grid with the  Invoice details." & Err.Description, vbInformation, "Ims"
Err.Clear

End Sub

Private Sub SSDbInvoice_GotFocus()
SSDbInvoice.SelLength = 0
SSDbInvoice.SelStart = 0
Call HighlightBackground(SSDbInvoice)
End Sub

Private Sub SSDbInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSDbInvoice.DroppedDown Then SSDbInvoice.DroppedDown = True
End Sub

Private Sub SSDbInvoice_LostFocus()
Call NormalBackground(SSDbInvoice)
End Sub

Private Sub SSDbInvoice_Validate(Cancel As Boolean)
If Len(Trim(SSDbInvoice.text)) = 0 Then Exit Sub
If SSDbInvoice.IsItemInList = False Then
    MsgBox "Please select a valid invoice from the list.", vbInformation, "Ims"
    Cancel = True
    SSDbInvoice.SetFocus
End If
End Sub

Private Sub SSDBPO_Click()

Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim ConversionRate As Double
On Error GoTo ErrHand

If OptPO.Value = True Then

    
    Rs.Source = "  select po_ponumb,poi_comm,poi_liitnumb ,curr_code,curr_desc, poi_primreqdqty,uni_desc,poi_unitprice, ToCompany, ToLocation, ToUsChar, ToStockType,ToCamChar from po"
    Rs.Source = Rs.Source & "  inner join poitem on poi_ponumb=po_ponumb and poi_npecode=po_npecode"
    Rs.Source = Rs.Source & "  inner join currency on curr_code =po_currcode and curr_npecode =po_npecode"
    Rs.Source = Rs.Source & "  inner join unit on uni_npecode = po_npecode and uni_code= poi_primuom"
    Rs.Source = Rs.Source & "  inner join pofqa on Ponumb = poi_ponumb and Npce_code=poi_npecode and ItemNo = poi_liitnumb"
    Rs.Source = Rs.Source & "  where po_ponumb='" & Trim(SSDBPO.text) & "' and po_npecode='" & deIms.NameSpace & "' order by cast(poi_liitnumb as int) "
    
    Rs.Open , deIms.cnIms
    
    SSGridFQA.RemoveAll

    If Rs.RecordCount > 0 Then
        
'    ConversionRate = GetConversionRate

    Do While Not Rs.EOF

        SSGridFQA.AddItem Rs!poi_liitnumb & vbTab & Trim(Rs!poi_comm) & vbTab & Rs!poi_primreqdqty & " " & Trim(Rs!uni_desc) & vbTab & IIf(Trim(UCase(Rs!curr_code)) = "USD", "USD", Trim(Rs!Curr_desc)) & " " & Round(Rs!poi_unitprice, 2) & vbTab & "" & vbTab & Rs!ToCompany & vbTab & Rs!Tolocation & vbTab & Rs!ToUsChar & vbTab & Rs!ToStockType & vbTab & Rs!ToCamChar
        
        SSGridFQA.MoveLast
        
        For i = 0 To SSGridFQA.Cols - 1
        
            SSGridFQA.Columns(i).TagVariant = SSGridFQA.Columns(i).Value
            
        Next i
            
        Rs.MoveNext
    
    Loop
    End If
    LblCount.Caption = Rs.RecordCount & " Records Found"
    
ElseIf OptInvoice.Value = True Then
    SSDbInvoice.text = ""
    SSDbInvoice.RemoveAll

    Rs.Source = "select inv_invcnumb from invoice where inv_ponumb ='" & Trim(SSDBPO.text) & "' and inv_npecode='" & deIms.NameSpace & "'"
    
    Rs.Open , deIms.cnIms

    Do While Not Rs.EOF
    
        SSDbInvoice.AddItem Rs!inv_invcnumb
    
        Rs.MoveNext
    
    Loop
    
End If
    
Exit Sub

ErrHand:

MsgBox "Errors occurred while trying to show the details for the Trnasaction Order. " & Err.Description, vbInformation, "Ims"
Err.Clear

End Sub

Private Sub SSDBPO_GotFocus()
SSDBPO.SelLength = 0
SSDBPO.SelStart = 0
Call HighlightBackground(SSDBPO)
End Sub

Private Sub SSDBPO_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSDBPO.DroppedDown Then SSDBPO.DroppedDown = True
End Sub

Private Sub SSDBPO_LostFocus()
Call NormalBackground(SSDBPO)
End Sub

Private Sub SSDBPO_Validate(Cancel As Boolean)
If Len(Trim(SSDBPO.text)) = 0 Then Exit Sub
If SSDBPO.IsItemInList = False Then
    MsgBox "Please select a valid transaction order from the list.", vbInformation, "Ims"
    Cancel = True
    SSDBPO.SetFocus
End If
End Sub

Private Sub SSDbWarehouse_Click()

Dim Rs As New ADODB.Recordset
Dim Cmd As New ADODB.Command
Dim Sql As String
Dim rsCURRENCY As New ADODB.Recordset
Dim ConversionRate As Double
On Error GoTo ErrHand

    SSDbWarehouse.Tag = Trim(SSDbWarehouse.Columns(1).text)
    With Cmd
    .CommandType = adCmdStoredProc
    .CommandText = "GetFQAForInventory"
    .ActiveConnection = deIms.cnIms
    .Parameters.Append .CreateParameter("@TRANTYPE", adChar, adParamInput, 3, Trim(SSDbWarehouse.Tag))
    .Parameters.Append .CreateParameter("@TRANNO", adChar, adParamInput, 15, Trim(SSDbWarehouse.text))
    .Parameters.Append .CreateParameter("@NAMESPACE", adChar, adParamInput, 5, Trim(deIms.NameSpace))
    Set Rs = .Execute
    End With
    SSGridFQA.RemoveAll
    
    If Rs.RecordCount > 0 Then
        
       ConversionRate = GetConversionRate(Rs!EXTENDEDCURRENCYCODE)
    
        Do While Not Rs.EOF
        
            SSGridFQA.AddItem Rs!LineNo & vbTab & Rs!StockNo & vbTab & Rs!qty & vbTab & IIf(Trim(UCase(Rs!CurrencyCode)) = "USD", "USD", Trim(Rs!Currency)) & " " & Rs!UnitPrice & vbTab & Rs!ExtendedCurrency & " " & Rs!UnitPrice * ConversionRate & vbTab & Rs!ToCompany & vbTab & Rs!Tolocation & vbTab & Rs!ToUsChar & vbTab & Rs!ToStockType & vbTab & Rs!ToCamChar
    
            Rs.MoveNext
    
        Loop
    
    End If
    
    LblCount.Caption = Rs.RecordCount & " Records Found"
    
Exit Sub
ErrHand:

MsgBox "Errors occurred while trying to show the details for the Trnasaction Order. " & Err.Description, vbInformation, "Ims"
Err.Clear

End Sub


Public Function settheGrid() As Boolean

End Function


Public Function PopulateCombosWithFQA() As Boolean

On Error GoTo ErrHand
FPopulateCombosWithFQA = False
Dim rsCOMPANY As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUc As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

'Get Company FQA

'LocationCode = Trim(LocationCode)

rsCOMPANY.Source = "select FQA from FQA where Namespace ='" & deIms.NameSpace & "' and Level ='C' order by FQA"

rsCOMPANY.Open , deIms.cnIms

Do While Not rsCOMPANY.EOF

    SSOleCompany.AddItem rsCOMPANY("FQA")
    rsCOMPANY.MoveNext
    
Loop

RsLocation.Source = "select distinct(FQA) from FQA where Namespace ='" & deIms.NameSpace & "' and Level ='LB' OR LEVEL ='LS' order by FQA"

RsLocation.Open , deIms.cnIms

'If RsLocation.RecordCount = 0 Then SSOleDBLocation.AddItem LocationCode
Do While Not RsLocation.EOF

    SSOleDBLocation.AddItem RsLocation("FQA")
    RsLocation.MoveNext
    
Loop


'Get US Chart FQA

RsUc.Source = "select distinct(FQA) from  FQA where Namespace ='" & deIms.NameSpace & "'  and Level ='UC'  order by FQA"

RsUc.Open , deIms.cnIms


Do While Not RsUc.EOF

    SSOleDBUsChart.AddItem RsUc("FQA")
    RsUc.MoveNext
    
Loop

'Get Cam Chart FQA

RsCC.Source = "select  distinct(FQA) from FQA where Namespace ='" & deIms.NameSpace & "'  and Level ='CC'  order by FQA"

RsCC.Open , deIms.cnIms


Do While Not RsCC.EOF

    SSOleDBCamChart.AddItem RsCC("FQA")
    RsCC.MoveNext
    
Loop

Set rsCOMPANY = Nothing
Set RsLocation = Nothing
Set RsUc = Nothing
Set RsCC = Nothing

FPopulateCombosWithFQA = True

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to fill the combo boxes." & Err.Description, vbCritical, "Ims"

Err.Clear

End Function

Private Sub SSDbWarehouse_GotFocus()
SSDbWarehouse.SelLength = 0
SSDbWarehouse.SelStart = 0
Call HighlightBackground(SSDbWarehouse)
End Sub

Private Sub SSDbWarehouse_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSDbWarehouse.DroppedDown Then SSDbWarehouse.DroppedDown = True
End Sub

Private Sub SSDbWarehouse_LostFocus()
Call NormalBackground(SSDbWarehouse)
End Sub

Private Sub SSDbWarehouse_Validate(Cancel As Boolean)
If Len(Trim(SSDbWarehouse.text)) = 0 Then Exit Sub
If SSDbWarehouse.IsItemInList = False Then
    MsgBox "Please select a valid warehouse transaction from the list.", vbInformation, "Ims"
    Cancel = True
    SSDbWarehouse.SetFocus
End If
End Sub

Private Sub SSGridFQA_BeforeRowColChange(Cancel As Integer)

If FFormMode <> mdModification Then Exit Sub

Select Case SSGridFQA.Col

Case 7

    If CheckifFqaExist(SSGridFQA.Columns(7).text, "uc") = False Then Cancel = True: MsgBox " Please enter a valid USChart# .", vbInformation, "Ims"

Case 9

    If CheckifFqaExist(SSGridFQA.Columns(9).text, "cc") = False Then Cancel = True: MsgBox " Please enter a valid CamChart#.", vbInformation, "Ims"

End Select

End Sub

Private Sub SSGridFQA_Change()
Dim GOldValue As String

GOldValue = SSGridFQA.Columns(SSGridFQA.Col).CellText(SSGridFQA.Bookmark)

If Trim(UCase(SSGridFQA.Columns(SSGridFQA.Col).TagVariant)) <> Trim(UCase(SSGridFQA.Columns(SSGridFQA.Col).Value)) Then
    
    FChangeflag = True

Else
    
    FChangeflag = False

End If

End Sub

Private Sub SSGridFQA_InitColumnProps()
SSGridFQA.Columns("company").DropDownHwnd = SSOleCompany.HWND
SSGridFQA.Columns("location").DropDownHwnd = SSOleDBLocation.HWND
SSGridFQA.Columns("uschart#").DropDownHwnd = SSOleDBUsChart.HWND
'SSOleDBFQA.columns("stocktype").DropDownHwnd = SSOleDBStockType.hWnd
SSGridFQA.Columns("camchart#").DropDownHwnd = SSOleDBCamChart.HWND
End Sub

Public Function ChangeOption() As Boolean

If FChangeflag = True Then
  
  If FCurrentOption = mdpo Then OptPO.Value = True
  If FCurrentOption = mdWarehouse Then OptWarehouse.Value = True
  If FCurrentOption = mdInvoice Then OptInvoice.Value = True
  
  ChangeOption = False
  
Else
    
    ChangeOption = True

End If


End Function

Private Sub SSGridFQA_KeyPress(KeyAscii As Integer)
If FFormMode = mdVisualization Then KeyAscii = 0: Exit Sub

Dim column As Integer
Dim row As Integer

Select Case SSGridFQA.Col

    Case 0
    
        KeyAscii = 0
        
    Case 1
    
        KeyAscii = 0
            
    Case 2
    
    KeyAscii = 0
            
    
    Case 3
    
    KeyAscii = 0
    
    Case 4
    
    KeyAscii = 0
    
    Case 5
    
    KeyAscii = 0
    
    Case 6
    
    KeyAscii = 0
    
    Case 7
            
    Case 8
            
       If Len((SSGridFQA.Columns(8).text) & Chr(KeyAscii)) > 4 Then
        
            MsgBox "Please make sure that the Stock type is not more than 4 digits.", vbInformation, "Ims"
            SSGridFQA.Columns(8).text = Mid(SSGridFQA.Columns(8).text, 1, 4)
            KeyAscii = 0
            
       End If
         
            
End Select

End Sub

Private Sub SSOleCompany_DropDown()
If FFormMode = mdVisualization Then SSOleCompany.DroppedDown = False
End Sub

Private Sub SSOleDBCamChart_DropDown()
If FFormMode = mdVisualization Then SSOleDBCamChart.DroppedDown = False
End Sub

Private Sub SSOleDBLocation_DropDown()
If FFormMode = mdVisualization Then SSOleDBLocation.DroppedDown = False
End Sub

Private Sub SSOleDBUsChart_DropDown()
If FFormMode = mdVisualization Then SSOleDBUsChart.DroppedDown = False
End Sub


Public Function GetConversionRate(Optional ExtCurrencycode As String) As Double

Dim rsCURRENCY As New ADODB.Recordset
Dim Sql As String
On Error GoTo ErrHand

If Len(Trim(ExtCurrencycode)) = 0 Then

    Sql = " select top 1 curd_value from currencydetl where curd_code=(select psys_extendedcurcode from pesys where psys_npecode =curd_npecode) and getdate() between curd_from and curd_to and curd_npecode='" & deIms.NameSpace & "' order by curd_id desc"

Else

    Sql = " select top 1 curd_value from currencydetl where curd_code='" & ExtCurrencycode & "' and getdate() between curd_from and curd_to and curd_npecode='" & deIms.NameSpace & "' order by curd_id desc"
       
End If
     
rsCURRENCY.Source = Sql
        
rsCURRENCY.Open , deIms.cnIms

If rsCURRENCY.RecordCount = 0 Then

    GetConversionRate = 0

Else

    GetConversionRate = IIf(IsNull(rsCURRENCY!curd_value), 0, rsCURRENCY!curd_value)

End If

Exit Function

ErrHand:

MsgBox "Errors occurred while trying to get the conversion rate. Extended Currency prices will be displayed 0." & Err.Description, vbCritical, "Ims"

Err.Clear

End Function

Public Function EnableDisableNavbar(enable As Boolean)

FrmHeader.Enabled = Not enable

LROleDBNavBar1.SaveEnabled = enable
LROleDBNavBar1.CancelEnabled = enable
LROleDBNavBar1.EditEnabled = Not enable

If enable = False Then
    FFormMode = ChangeModeOfForm(lblStatus, mdVisualization)
Else
    FFormMode = ChangeModeOfForm(lblStatus, mdModification)
End If

SSDBPO.Enabled = Not enable

SSDbInvoice.Enabled = Not enable

SSDbWarehouse.Enabled = Not enable
'SSGridFQA.Columns("uschart#").StyleSet =
End Function
