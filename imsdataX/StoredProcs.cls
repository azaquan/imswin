VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "imsspInt"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A77025C"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'Private User As String
'##ModelId=38955A7703A7
Public Enum imsAddressType
    '##ModelId=38955A780069
    atFax
    '##ModelId=38955A78007D
    atEmail
End Enum

'##ModelId=38955A7903E6
Public Event OnError(ByVal Description As String, ByVal ErrNumber As Long)

'##ModelId=38955A780091

'assign store procedure parameters and call it to get recordset

Public Function GetCompany(NameSpace As String, Usercode As String, cn As ADODB.Connection, Optional Code As String, Optional Active As Boolean) As String
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "Get_Wharehouse_Company"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@USERCODE", adChar, adParamInput, 10, Usercode)
        '.Parameters.Append .CreateParameter("@Active", adBoolean, adParamInput, , Active)
        .Parameters.Append .CreateParameter("@COMPANY", adVarChar, adParamOutput, 40)
        .Parameters.Append .CreateParameter("@CODE", adVarChar, adParamOutput, 10)
        Call .Execute(Options:=adExecuteNoRecords)
        
        Code = .Parameters("@CODE") & ""
        GetCompany = .Parameters("@COMPANY") & ""
        
    End With
    
    MTSCommit
    Set cmd = Nothing
End Function

'##ModelId=38955A7800AF

'assign store procedure parameters and call it to get location recordset

Public Function GetLocation(NameSpace As String, LocationType As String, CompanyName As String, cn As ADODB.Connection, Optional EqualTo As Boolean = True) As Variant
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "Get_Location"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@LOCATIONTYPE", adVarChar, adParamInput, 10, LocationType)
        
        .Parameters.Append .CreateParameter("@NAMESPACE", adChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@COMPANYCODE", adChar, adParamInput, 10, CompanyName)
        
        .Parameters.Append .CreateParameter("EqualsTo", adBoolean, adParamInput, , EqualTo)
        Set GetLocation = .Execute
    End With
    
    Set cmd = Nothing
    
End Function

'##ModelId=38955A7800CD

'SQL statement to get company code

Public Function GetCompanyCode(NameSpace As String, CompanyName As String, cn As ADODB.Connection) As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    Set cmd = New ADODB.Command
    
    With cmd
        
        Set .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "SELECT com_compcode FROM COMPANY"
        .CommandText = .CommandText & " WHERE com_name = '" & CompanyName & "'"
        .CommandText = .CommandText & " AND com_npecode = '" & NameSpace & "'"
        
        Set rs = .Execute
        
        If rs.RecordCount = 0 Then Exit Function
        GetCompanyCode = rs!com_compcode & ""
    End With
    
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
End Function

'##ModelId=38955A7800EB

'assign store procedure parameters and call it to get warehouse reciptient number

Public Function GetWareHouseReceiptPoNumb(NameSpace As String, Location As String, cn As ADODB.Connection) As Variant
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "GET_PoNum_ForWareHouse"
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@LOCATION", adChar, adParamInput, 10, Location)
        Set GetWareHouseReceiptPoNumb = .Execute
        
    End With
    
    Set cmd = Nothing
End Function

'##ModelId=38955A780109

'assign store procedure parameters and call it to get poitem reciptient
'The last parameter FNRR added by Muzammil 18/01/01
Public Function GetPoitemForReceipt(NameSpace As String, PONumb As String, cn As ADODB.Connection, Optional FNRR As Integer) As Variant
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "Get_POItems_For_WareHouse_Receipt"
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, PONumb)
        .Parameters.Append .CreateParameter("@FNRR", adInteger, adParamInput, 1, FNRR)
        
        Set GetPoitemForReceipt = .Execute
        
    End With
    
    Set cmd = Nothing
End Function

'##ModelId=38955A780127

'assign store procedure parameters and call it to get inventory reciptient

Public Function GetInvtReceipt(NameSpace As String, PONumb As String, Transnumb As String, cn As ADODB.Connection) As Variant
    
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "Get_InvtReceiptDetl"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("np", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("tr", adInteger, adParamInput, , Transnumb)
        .Parameters.Append .CreateParameter("PONUMB", adVarChar, adParamInput, 15, PONumb)
        
        Set GetInvtReceipt = .Execute
    End With
    
    Set cmd = Nothing
End Function

'##ModelId=38955A780145

'SQL statement to get currency values for po number

Public Function GetCurrencyValueFromPo(NameSpace As String, PONumb As String, cn As ADODB.Connection) As String
On Error Resume Next
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    Set cmd = MakeCommand(cn, adCmdText)
    With cmd
        .CommandText = "SELECT po_currcode FROM PO where PO_PONUMB = '" & PONumb & "'"
        .CommandText = .CommandText & " AND po_npecode = '" & NameSpace & "'"
        
        Set rs = .Execute
        
        .CommandText = "SELECT curd_code, curd_value FROM CURRENCYDETL" & vbCrLf
        .CommandText = .CommandText & "WHERE curd_npecode = '" & NameSpace & "'" & vbCrLf
        .CommandText = .CommandText & "AND curd_code = '" & rs!po_currcode & "'" & vbCrLf
        .CommandText = .CommandText & "AND curd_to > '" & Format(Date, "mm/dd/yyyy") & "'" & vbCrLf
        .CommandText = .CommandText & "AND curd_from < '" & Format(Date, "mm/dd/yyyy") & "'" & vbCrLf
        
        Set rs = Nothing
        Set rs = .Execute
        
        On Error GoTo 0
        If rs.EOF And rs.BOF Then
            Err.Raise 1000, "Get Currency Value From Po", "No valid entry for currency"
        Else
            GetCurrencyValueFromPo = rs!curd_value & ""
        End If
        
        
        
        Set rs = Nothing
        Set cmd = Nothing
    End With
End Function

'##ModelId=38955A780163

'make command conncetion

Private Function MakeCommand(cn As ADODB.Connection, CommandType As ADODB.CommandTypeEnum, Optional Text As String) As ADODB.Command
 'ADDED ERROR HANDLING 08/03/00 / MUZAMMIL
 On Error Resume Next
    
    Set MakeCommand = Nothing
    
    cn.Errors.Clear
    Set MakeCommand = New ADODB.Command
    Set MakeCommand.ActiveConnection = cn
    MakeCommand.CommandType = CommandType
    MakeCommand.CommandText = Text
End Function

'##ModelId=38955A780181

'SQL statement to get currency description

Public Function GetCurrencyDescriptionForPo(NameSpace As String, PONumb As String, Code As String, cn As ADODB.Connection) As String
On Error Resume Next
Dim rs As ADODB.Recordset

    With MakeCommand(cn, adCmdText)
        .CommandText = "SELECT curr_desc, curr_code FROM CURRENCY, PO where po_ponumb = '" & PONumb & "'" & vbCrLf
        .CommandText = .CommandText & " AND po_npecode = '" & NameSpace & "'" & vbCrLf
        .CommandText = .CommandText & " AND curr_code = po_currcode"
        
        Set rs = .Execute
        Code = rs!curr_code & ""
        GetCurrencyDescriptionForPo = rs!curr_desc & ""
    End With
    Set rs = Nothing
End Function

'##ModelId=38955A78019F

'SQL statement to get stock type description

Public Function GetStockType(NameSpace As String, cn As ADODB.Connection, Optional Code As String) As Variant
Dim str As String

    str = "SELECT sty_stcktype Code, sty_desc Description FROM STOCKTYPE WHERE sty_npecode = '" & NameSpace & "'"
    
    If Not ((IsMissing(Code)) Or Len(Trim$(Code))) = 0 Then _
        str = str & " AND sty_stcktype = '" & Code & "'"
        
        With MakeCommand(cn, adCmdText)
            .CommandText = str
            Set GetStockType = .Execute
            
        End With
End Function

'##ModelId=38955A7801C8

'assign store procedure parameters and call it to get inventory transaction number

Public Function GetTransNumb(NameSpace As String, cn As ADODB.Connection) As Long

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_InvtNumb"
        
        .Parameters.Append .CreateParameter("@RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NameSpace", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@numb", adInteger, adParamOutput, 4, Null)
        
        Call .Execute(Options:=adExecuteNoRecords)
        GetTransNumb = .Parameters("@numb").Value
    End With
    
    If GetTransNumb Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A7801E6

'insert record to inventory recipient

Public Function InvtReceipt_Insert(NameSpace As String, PONumb As String, TranType As String, CompanyCode As String, Warehouse As String, User As String, cn As ADODB.Connection, Optional ManufacturerNumb As String, Optional TranFrom As String, Optional TransNum As String) As Boolean

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
        
        .Parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, User)
        
        Call .Execute(Options:=adExecuteNoRecords)
        InvtReceipt_Insert = .Parameters("RV") = 0
    End With
        
    If InvtReceipt_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function

'##ModelId=38955A780218
Public Function InvtReceiptDetl_Insert() As Boolean

End Function

'##ModelId=38955A78022C

'insert a record to inventory reciptient remark

Public Function InvtReceiptRem_Insert(NameSpace As String, CompanyCode As String, Warehouse As String, TansNum As String, Remarks As String, User As String, cn As ADODB.Connection) As Boolean
                                      
                                      

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "InvtReceiptRem_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WH", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("TN", adVarChar, adParamInput, 15, TansNum)
        
        .Parameters.Append .CreateParameter("LN", adInteger, adParamInput, , 1)
        .Parameters.Append .CreateParameter("Rem", adVarChar, adParamInput, 7000, Remarks)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        InvtReceiptRem_Insert = .Parameters("RV") = 0
        
    End With
        
    If InvtReceiptRem_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A780254

'assign store procedure parameters and  call it to get warehouse price

Public Function Get_Warehouse_Price(NameSpace As String, PONumb As String, LineItem As String, cn As ADODB.Connection) As Double
        
        With MakeCommand(cn, adCmdStoredProc)
            .CommandText = "Get_Warehouse_Price"
            
            .Parameters.Append .CreateParameter("np", adVarChar, adParamInput, 5, NameSpace)
            .Parameters.Append .CreateParameter("PONUMB", adVarChar, adParamInput, 15, PONumb)
            .Parameters.Append .CreateParameter("Litem", adVarChar, adParamInput, 6, LineItem)
            .Parameters.Append .CreateParameter("RETVAL", adDouble, adParamOutput)
            Call .Execute(Options:=adExecuteNoRecords)
            
            Get_Warehouse_Price = .Parameters("RETVAL").Value
        End With
        
        MTSCommit
End Function


'##ModelId=38955A780272

'insert record to quantity one table

Public Function Quantity_In_stock1_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, Description As String, User As String, cn As ADODB.Connection) As Boolean
                                          

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock1_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, RTrim$(CompanyCode))
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, RTrim$(StockNumber))
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("Description", adVarChar, adParamInput, 1500, Description)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        
        

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock1_Insert = .Parameters("RV").Value = 0
    End With
    
    If Quantity_In_stock1_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function


'##ModelId=38955A7802A4

'insert record to quantity two table

Public Function Quantity_In_stock2_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, LogicalWareHouse As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock2_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock2_Insert = .Parameters("RV").Value = 0
        
    End With

    If Quantity_In_stock2_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A7802D6

'insert record to quantity three table

Public Function Quantity_In_stock3_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, LogicalWareHouse As String, SubLocation As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock3_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SubLoc", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock3_Insert = .Parameters("RV").Value = 0
        
    End With

    If Quantity_In_stock3_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A780312

'insert record to quantity four table

Public Function Quantity_In_stock4_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, LogicalWareHouse As String, SubLocation As String, ConditionCode As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock4_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SubLoc", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock4_Insert = .Parameters("RV").Value = 0
        
    End With
    
    If Quantity_In_stock4_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A78034E

'inset record to quantity five table

Public Function Quantity_In_stock5_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, LogicalWareHouse As String, SubLocation As String, ConditionCode As String, Transnumb As String, transerl As Long, TransWare As String, TransType As String, TransComp As String, FromTo As String, TTransnumb As String, TTransComp As String, TransLineNumb As Long, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock5_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SubLoc", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("transnumb", adVarChar, adParamInput, 15, Transnumb)
        
        .Parameters.Append .CreateParameter("transerl", adInteger, adParamInput, , transerl)
        .Parameters.Append .CreateParameter("transware", adVarChar, adParamInput, 10, TransWare)
        .Parameters.Append .CreateParameter("transtype", adVarChar, adParamInput, 2, TransType)
        
        .Parameters.Append .CreateParameter("tranComp", adVarChar, adParamInput, 10, TransComp)
        .Parameters.Append .CreateParameter("Ttransnumb", adVarChar, adParamInput, 15, TTransnumb)
        .Parameters.Append .CreateParameter("TTranComp", adVarChar, adParamInput, 10, TTransComp)
        .Parameters.Append .CreateParameter("tranLineNumb", adInteger, adParamInput, , TransLineNumb)
        .Parameters.Append .CreateParameter("FromTo", adVarChar, adParamInput, 10, FromTo)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock5_Insert = .Parameters("RV").Value = 0
        
    End With

    If Quantity_In_stock5_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A7803BC

'insert record to quantity six table

Public Function Quantity_In_stock6_Insert(NameSpace As String, CompanyCode As String, _
                                          StockNumber As String, Warehouse As String, _
                                          PrimaryQnty As Double, SecondaryQnty As Double, _
                                          LogicalWareHouse As String, SubLocation As String, _
                                          ConditionCode As String, Serial As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock6_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SubLoc", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("Serial", adVarChar, adParamInput, 30, Serial)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock6_Insert = .Parameters("RV").Value = 0
        
        
    End With
    
    If Quantity_In_stock6_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If
End Function

'##ModelId=38955A790024

'insert record to quantity seven table

Public Function Quantity_In_stock7_Insert(NameSpace As String, CompanyCode As String, StockNumber As String, _
                                          Warehouse As String, PrimaryQnty As Double, SecondaryQnty As Double, _
                                          LogicalWareHouse As String, SubLocation As String, ConditionCode As String, _
                                          Transnumb As String, FromTo As String, transerl As Long, TransWare As String, _
                                          TransType As String, TransComp As String, TTransnumb As String, TTransComp As String, _
                                          TransLineNumb As Long, Serial As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Quantity_In_stock7_Insert"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("PRMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("SECQNTY", adDouble, adParamInput, , SecondaryQnty)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SubLoc", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("transnumb", adVarChar, adParamInput, 15, Transnumb)
        
        .Parameters.Append .CreateParameter("transerl", adInteger, adParamInput, , transerl)
        .Parameters.Append .CreateParameter("transware", adVarChar, adParamInput, 10, TransWare)
        .Parameters.Append .CreateParameter("transtype", adVarChar, adParamInput, 2, TransType)
        
        .Parameters.Append .CreateParameter("tranComp", adVarChar, adParamInput, 10, TransComp)
        .Parameters.Append .CreateParameter("Ttransnumb", adVarChar, adParamInput, 15, TTransnumb)
        .Parameters.Append .CreateParameter("TTranComp", adVarChar, adParamInput, 10, TTransComp)
        .Parameters.Append .CreateParameter("tranLineNumb", adInteger, adParamInput, , TransLineNumb)
        .Parameters.Append .CreateParameter("FromTo", adVarChar, adParamInput, 10, FromTo)
        .Parameters.Append .CreateParameter("Serial", adVarChar, adParamInput, 30, Serial)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)

        Call .Execute(Options:=adExecuteNoRecords)
        Quantity_In_stock7_Insert = .Parameters("RV").Value = 0
        
    End With

    If Quantity_In_stock7_Insert Then
        MTSCommit
    Else
        MTSRollback
    End If

End Function

'##ModelId=38955A7900F7

'assign store procedure parameters and call it to update sap value

Public Function Update_Sap(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, CurrencyValue As Double, UnitPrice As Double, ConditionCode As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "UPDATE_SAP"
        
        Dim v As Variant
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("UP", adDouble, adParamInput, 10, UnitPrice)
        
        v = ConditionCode
        
        If Len(Trim$(v)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, v)
        .Parameters.Append .CreateParameter("PRIMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("CurrValue", adDouble, adParamInput, , CurrencyValue)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Update_Sap = .Parameters("RV").Value = 0
        
        If Update_Sap Then
            MTSCommit
        Else
            MTSRollback
        End If
        
    End With

End Function


'##ModelId=38955A79013D

'assign store procedure parameters and call it to get stock information with condition

Public Function GetStockInfo(NameSpace As String, Warehouse As String, Company As String, cn As ADODB.Connection) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "GetStockInfoFromQTYST4"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("WH", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, Company)
        
        Set GetStockInfo = .Execute
        
    End With
End Function

'assign store procedure parameters and call it to get stock information

Public Function GetStockInformation(NameSpace As String, Warehouse As String, Company As String, cn As ADODB.Connection) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "GetStockInfoFromQTYST1"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("WH", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, Company)
        
        Set GetStockInformation = .Execute
        
    End With
End Function

'##ModelId=38955A79015B

'assign store procedure parameters and call it to get logical warehouse information
'for quantity tables

Public Function Get_LogicalWarehouse_FromQTYST(NameSpace As String, StockNumber As String, Warehouse As String, Company As String, cn As ADODB.Connection, Optional Condition As String) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_LogicalWarehouse_FromQTYST"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("Company", adVarChar, adParamInput, 10, Company)
        .Parameters.Append .CreateParameter("WareHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("CONDITION", adVarChar, adParamInput, 10, Condition)
        
        
        Set Get_LogicalWarehouse_FromQTYST = .Execute
    End With
    
    MTSCommit
End Function



'##ModelId=38955A790183

'assign store procedure parameters and call it to get sub location information
'for quantity tables

Public Function Get_SubLocation_FromQTYST(NameSpace As String, StockNumber As String, Warehouse As String, Company As String, LogicalWareHouse As String, cn As ADODB.Connection) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_SubLocation_FromQTYST"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("Company", adVarChar, adParamInput, 10, Company)
        .Parameters.Append .CreateParameter("WareHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
'        .Parameters.Append .CreateParameter("Condition", adVarChar, adParamInput, 2, Condition)
        
        Set Get_SubLocation_FromQTYST = .Execute
    End With
    
    
    MTSCommit
End Function

Public Function Get_SubLocation_FromQTYSTnew(NameSpace As String, StockNumber As String, Warehouse As String, Company As String, LogicalWareHouse As String, cn As ADODB.Connection, Optional Condition As String) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_SubLocation_FromQTYSTnew"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("Company", adVarChar, adParamInput, 10, Company)
        .Parameters.Append .CreateParameter("WareHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("Condition", adVarChar, adParamInput, 2, Condition)
        
        Set Get_SubLocation_FromQTYSTnew = .Execute
    End With
    
    
    MTSCommit
End Function


'##ModelId=38955A7901AB

'assign store procedure parameters and call it to get condition information

Public Function Get_Condition_FromQTYST(NameSpace As String, StockNumber As String, Warehouse As String, Company As String, LogicalWareHouse As String, SubLocation As String, cn As ADODB.Connection) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_Condition_FromQTYST"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("Company", adVarChar, adParamInput, 10, Company)
        .Parameters.Append .CreateParameter("WareHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SB", adVarChar, adParamInput, 10, SubLocation)
        
        Set Get_Condition_FromQTYST = .Execute
    End With
    
    MTSCommit
End Function


'##ModelId=38955A7901D3

'assign store procedure parameters and call it to get sap values

Public Function GetSapValue(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, ConditionCode As String, cn As ADODB.Connection) As Double

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "GetSapValue"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("Value", adDouble, adParamOutput)
        
        Call .Execute(Options:=adExecuteNoRecords)
        GetSapValue = CDbl(.Parameters("Value"))
        
    End With

End Function

'##ModelId=38955A790205

'assign store procedure parameters and call it to get serial number for quantity tables

Public Function Get_SerialNumberFromStockNumber(NameSpace As String, StockNumber As String, Warehouse As String, Company As String, LogicalWareHouse As String, SubLocation As String, ConditionCode As String, cn As ADODB.Connection) As Variant
                                                
    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Get_SerialNumberFromStockNumber"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("Company", adVarChar, adParamInput, 10, Company)
        .Parameters.Append .CreateParameter("WareHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("LW", adVarChar, adParamInput, 10, LogicalWareHouse)
        .Parameters.Append .CreateParameter("SB", adVarChar, adParamInput, 10, SubLocation)
        .Parameters.Append .CreateParameter("Cond", adVarChar, adParamInput, 2, ConditionCode)
        
        Set Get_SerialNumberFromStockNumber = .Execute
        
    End With
    
    MTSCommit
End Function

'##ModelId=38955A790237

'assign store procedure parameters and call it to get condition information

Public Function GetConditions(NameSpace As String, ConditionCode As String, Higher As Boolean, cn As ADODB.Connection) As Variant

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "GetConditions"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("cc", adVarChar, adParamInput, 2, ConditionCode)
        .Parameters.Append .CreateParameter("GR", adBoolean, adParamInput, , Higher)
        
        Set GetConditions = .Execute
    End With
    
End Function

'##ModelId=38955A790255

'SQL statement to get stock master information

Public Function GetSpecificStockInfo(sStockNumber As String, NameSpace As String, cn As ADODB.Connection, Optional Exact As Boolean = True) As Variant
Dim cmd As ADODB.Command

    'Set GetSpecificStockInfo = Nothing
    Set cmd = New ADODB.Command
    With cmd
        
        .CommandType = adCmdText
        Set .ActiveConnection = cn
        
        .CommandText = "Select stk_stcknumb, stk_stcktype, stk_catecode, stk_poolspec, stk_desc"
        .CommandText = .CommandText & " From StockMaster Where stk_npecode = '" & NameSpace & "'"
        
        If Len(Trim$(sStockNumber)) Then
            .CommandText = .CommandText & " AND stk_stcknumb "
            .CommandText = .CommandText & IIf(Exact, "= '" & sStockNumber & "'", "like '" & sStockNumber & "%'")
        End If
        
        Set GetSpecificStockInfo = New ADODB.Recordset
        GetSpecificStockInfo.LockType = adLockReadOnly
        GetSpecificStockInfo.CursorType = adOpenForwardOnly
        GetSpecificStockInfo.CursorLocation = adUseServer
        
        Set GetSpecificStockInfo = .Execute
        
    End With
    
    Set cmd = Nothing
End Function

'##ModelId=38955A790273

'assign store procedure parameters and call it to get sap repair cost

Public Function Update_Sap_With_repair_Cost(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, PrimaryQnty As Double, CurrencyValue As Double, RepairCost As Double, UnitPrice As Double, ConditionCode As String, User As String, cn As ADODB.Connection) As Boolean

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "Update_Sap_With_repair_Cost"
        
        Dim v As Variant
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        .Parameters.Append .CreateParameter("UP", adSingle, adParamInput, 10, UnitPrice)
        
        .Parameters.Append .CreateParameter("REPCOST", adDouble, adParamInput, , RepairCost)
        v = ConditionCode
        
        If Len(Trim$(v)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, v)
        .Parameters.Append .CreateParameter("PRIMQNTY", adDouble, adParamInput, , PrimaryQnty)
        .Parameters.Append .CreateParameter("CurrValue", adDouble, adParamInput, , CurrencyValue)
        .Parameters.Append .CreateParameter("User", adVarChar, adParamInput, 20, User)
        

        Call .Execute(Options:=adExecuteNoRecords)
        Update_Sap_With_repair_Cost = .Parameters("RV").Value = 0
        
        If Update_Sap_With_repair_Cost Then
            MTSCommit
        Else
            MTSRollback
        End If
        
    End With

End Function

'##ModelId=38955A7902AF

'assign store procedure parameters and call it to get quantity on hand information

Public Function QuantityOnHand(NameSpace As String, CompanyCode As String, StockNumber As String, Warehouse As String, ConditionCode As String, cn As ADODB.Connection) As Double

    With MakeCommand(cn, adCmdStoredProc)
        .CommandText = "QuantityOnHand"
        
        Dim v As Variant
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        .Parameters.Append .CreateParameter("WHouse", adVarChar, adParamInput, 10, Warehouse)
        v = ConditionCode
        
        If Len(Trim$(v)) = 0 Then v = Null
        .Parameters.Append .CreateParameter("Condition", adChar, adParamInput, 2, v)
        
        .Parameters.Append .CreateParameter("RETVAL", adDouble, adParamOutput)

        Call .Execute(Options:=adExecuteNoRecords)
        QuantityOnHand = .Parameters("RETVAL").Value & ""
        
    End With

End Function

'##ModelId=38955A7902D7

'assign store procedure paramters and call it to get sap stock values

Public Function Get_Sap_Stock_Values(NameSpace As String, CompanyCode As String, Location As String, cn As ADODB.Connection, Optional StockNumber As Variant, Optional ConditionCode, Optional Value As Double = -1) As Variant
                                     
                                     
Dim cmd As ADODB.Command


    Set cmd = MakeCommand(cn, adCmdStoredProc)
    
    With cmd
        .CommandText = "Get_Sap_Stock_Values"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("CC", adVarChar, adParamInput, 10, CompanyCode)
        .Parameters.Append .CreateParameter("Location", adVarChar, adParamInput, 10, Location)
        
        If Not IsNull(StockNumber) Then _
            If Len(Trim$(StockNumber)) = 0 Then StockNumber = Null

        .Parameters.Append .CreateParameter("StockNumb", adVarChar, adParamInput, 20, StockNumber)
        

        If Not IsNull(ConditionCode) Then _
            If Len(Trim$(ConditionCode)) = 0 Then ConditionCode = Null
            
        .Parameters.Append .CreateParameter("CondCode", adVarChar, adParamInput, 10, ConditionCode)
        
        .Parameters.Append .CreateParameter("Value", adDouble, adParamOutput)

        Set Get_Sap_Stock_Values = .Execute
        Value = .Parameters("Value")
        
    End With
    
    MTSCommit
    Set cmd = Nothing
End Function

'##ModelId=38955A790309

'insert record to xlogin table

Public Sub InsertIntoXLogin(NameSpace As String, Activity As String, User As String, cn As ADODB.Connection)
On Error GoTo Cancelled
Dim cmd As ADODB.Command

    
    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        .CommandText = "INSERT INTO XLOGIN"
        .CommandText = .CommandText & " (log_npecode,log_userid, log_date, log_actn)"
        .CommandText = .CommandText & " VALUES (?,?,?,?)"
    
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, User)
        
        .Parameters.Append .CreateParameter(, adDBTimeStamp, adParamInput)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, Activity)
        
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    
    MTSCommit
    Set cmd = Nothing
    
    Exit Sub
Cancelled:
    If Err Then MTSRollback
End Sub

'##ModelId=38955A790327

'assign store procedure parameters and call it to insert record to xevent table

Public Function XEVENT_INSERT(NameSpace As String, UserId As String, From As Variant, sEvent As String, Status As Variant, cn As ADODB.Connection) As Boolean
                              
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "XEVENT_INSERT"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("UID", adVarChar, adParamInput, 15, UserId)
        
        If Len(Trim$(From & "")) = 0 Then From = Null
        .Parameters.Append .CreateParameter("From", adVarChar, adParamInput, 15, From)
        .Parameters.Append .CreateParameter("Event", adVarChar, adParamInput, 40, sEvent)
        
        If Len(Trim$(Status & "")) = 0 Then Status = Null
        .Parameters.Append .CreateParameter("Status", adVarChar, adParamInput, 10, Status)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        XEVENT_INSERT = .Parameters(0).Value = 0
        
    End With
    
    If XEVENT_INSERT Then
        MTSCommit
    Else
        MTSRollback
    End If
    
    Set cmd = Nothing
End Function


'##ModelId=38955A790350

'call store procedure to get name space

Public Function GetNameSpaces(cn As ADODB.Connection, Optional RecordsReturned As Long) As Variant
On Error Resume Next
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset


    
    Set cmd = New ADODB.Command
      
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandText = "GETNAMESPACES"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("Retval", adInteger, adParamReturnValue)
        
        Set rs = .Execute
        
        If Err Then MsgBox Err.Description
        If rs.EOF And rs.BOF Then RecordsReturned = 0
        RecordsReturned = IIf(IsNull(rs!RecordsReturned), 0, rs!RecordsReturned)
            
        Set GetNameSpaces = rs
    End With
    
    MTSCommit
    Set cmd = Nothing
    
End Function

'##ModelId=38955A790364

'assign store procedure parameters and call it to get buyer approval information

Public Function GetPOsForApproval(NameSpace As String, UserName As String, cn As ADODB.Connection) As Variant
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdStoredProc)
        
    With cmd
        .CommandText = "BUYER_APP"
        .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, NameSpace)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, UserName)
        
        Set GetPOsForApproval = .Execute
    End With
    
    Set cmd = Nothing
End Function

'##ModelId=38955A790382

'user microsofe transaction server command

Public Sub MTSCommit()
On Error Resume Next
    If GetObjectContext Is Nothing Then Exit Sub
    If GetObjectContext.IsInTransaction Then GetObjectContext.SetComplete
End Sub

'##ModelId=38955A79038C

'user microsofe transaction server command

Public Sub MTSRollback()
    If GetObjectContext Is Nothing Then Exit Sub
    If GetObjectContext.IsInTransaction Then GetObjectContext.SetAbort
End Sub

'##ModelId=38955A7903A0

'SQL statemnet to get address from phone directory

Public Function GetAddresses(ByVal NameSpace As String, ByVal ConnectionString As String, LockType As LockTypeEnum, AddressType As imsAddressType) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command
Dim cnn As ADODB.Connection
On Error GoTo ErrHandler

    Set cnn = New ADODB.Connection
    cnn.ConnectionString = ConnectionString
    
    cnn.Open
    Set GetAddresses = Nothing
    Set cmd = MakeCommand(cnn, adCmdText)
    str = IIf(AddressType = atEmail, "phd_mail", "phd_faxnumb")
    
    With cmd
    
        .Prepared = True
        .CommandText = "SELECT phd_name, phd_code, " & str
       
        
        
        .CommandText = .CommandText & " FROM PHONEDIR where " & str & " is not null"
        .CommandText = .CommandText & " AND phd_npecode = ? and len(" & str & ") > 3"
        .CommandText = .CommandText & " ORDER BY phd_name"
        Set GetAddresses = .Execute(0, Array(NameSpace))
    End With
    
    GetAddresses.Close
    Call GetAddresses.Open(, , adOpenStatic, LockType)
    
    MTSCommit
    Set cnn = Nothing
    Exit Function
ErrHandler:

    MTSRollback
    If Err Then RaiseEvent OnError(Err.Description, Err.Number)
End Function


'##ModelId=38955A7903BE

'assign store procedure parameters and call it to get menuuser information

Public Function Getmenuuser(NameSpace As String, UserId As String, optionid As String, cn As ADODB.Connection) As Boolean
   Dim cmd As ADODB.Command
   Dim str As String
                           
   Set cmd = New ADODB.Command
                                            
  With cmd
         Set .ActiveConnection = cn
         .CommandType = adCmdStoredProc
         .CommandText = "getmenuaccess"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@optionid", adVarChar, adParamInput, 10, optionid)
        .Parameters.Append .CreateParameter("@userid", adVarChar, adParamInput, 15, UserId)
        
        Call .Execute(Options:=adExecuteNoRecords)
        Getmenuuser = .Parameters("RT")
    End With
    
    
    Set cmd = Nothing
End Function

'assign store procedure parameters and call it to update po total cost

Public Sub UpdatePoTotalCost(PONumber As String, NameSpace As String, cn As ADODB.Connection)
Dim cmd As ADODB.Command
On Error GoTo Cancelled

    Set cmd = MakeCommand(cn, adCmdStoredProc, "UPDATETOTALPOCOST")
    
    With cmd
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("PO", adVarChar, adParamInput, 15, PONumber)
        
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    
    MTSCommit
    Set cmd = Nothing
    
    Exit Sub
Cancelled:
    If Err Then MTSRollback: Err.Clear
End Sub

'08/02/00 CHANGED BY MUZAMMIL
'REDUCED A PARAMETER(GENDER) FROM THIS FUNCTION
'Public Function GetPORecipients(NameSpace As String, DocType As String, Gender As String, cn As ADODB.Connection, Optional RecordsReturned As Long) As ADODB.Recordset


Public Function GetPORecipients(NameSpace As String, DocType As String, cn As ADODB.Connection, Optional RecordsReturned As Long) As ADODB.Recordset
On Error Resume Next

Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdStoredProc, "GetEmailAndFaxForRecipients")
    
    With cmd
    
        If .Parameters.count = 0 Then
            .Parameters.Append .CreateParameter(, adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5)
        End If
       ' 08/02/00 CHANGE BY MUZAMMIL
       ' MADE CHANGES TO THE STORED PROCEDURE AS WELL
       '.Parameters(3) = Gender
        .Parameters(2) = DocType
        .Parameters(1) = NameSpace
        Set GetPORecipients = .Execute
        
           
        RecordsReturned = IIf(IsNull(.Parameters(0)), 0, .Parameters(0))
    End With
    
    Set cmd = Nothing
    If Err Then Err.Clear
End Function

'SQL statement to get supplier email and fax information

Public Function GetSupplierEmailForPO(NameSpace As String, SupplierCode As String, cn As ADODB.Connection) As String
On Error Resume Next
Dim cmd As ADODB.Command

    
    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        
        cmd.CommandText = "select ? = sup_faxnumb, ? = sup_mail from SUPPLIER where sup_code = '"
        cmd.CommandText = cmd.CommandText & SupplierCode & "' AND sup_npecode = '" & NameSpace & "'"
        
        cmd.Parameters.Append .CreateParameter("mail", adVarChar, adParamOutput, 60)
        cmd.Parameters.Append .CreateParameter("faxNo", adVarChar, adParamOutput, 60)
        
        Call .Execute(Options:=adExecuteNoRecords)

        If Len(Trim$(cmd.Parameters(0) & "")) Then
            GetSupplierEmailForPO = cmd.Parameters(0) & ""
            
        
        Else
             GetSupplierEmailForPO = cmd.Parameters(1) & ""
        End If
    End With

End Function

'SQL statement to get computing factor information

Public Function ComputingFactor(NameSpace As String, StockNumber As String, cn As ADODB.Connection) As Double
#If DBUG = 0 Then
    On Error Resume Next
#End If

Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        .CommandText = "SELECT ? = stk_compfctr FROM STOCKMASTER"
        .CommandText = .CommandText & " WHERE stk_stcknumb = '" & StockNumber & "'"
        .CommandText = .CommandText & " AND stk_npecode = '" & NameSpace & "'"
        
        .Parameters.Append .CreateParameter(, adDouble, adParamOutput)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        ComputingFactor = IIf(IsNull(.Parameters(0)), 0, .Parameters(0))
    End With
    
    MTSCommit
    Set cmd = Nothing
End Function

'assign store procedure parameters and call it to get stock master unit

Public Sub GetStockUnit(NameSpace As String, StockNumber As String, PrimaryUnit As String, SecondaryUnit As String, cn As ADODB.Connection)
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdStoredProc, "GetStockUnit")
    
    With cmd
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("SN", adVarChar, adParamInput, 20, StockNumber)
        
        .Parameters.Append .CreateParameter("PU", adVarChar, adParamOutput, 4)
        .Parameters.Append .CreateParameter("SU", adVarChar, adParamOutput, 4)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        PrimaryUnit = .Parameters(2) & ""
        SecondaryUnit = .Parameters(3) & ""
    End With
        
        
End Sub

'SQL statement to get company information

Public Function GetCompanies(NameSpace As String, cn As ADODB.Connection)
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        .CommandText = "select com_compcode, com_comname "
        .CommandText = .CommandText & " from company"
        .CommandText = .CommandText & "where "
    End With
    
    
End Function

'SQL statement to check stock serial number exist or not

Public Function IsStockNumberSerial(NameSpace As String, StockNumber As String, cn As ADODB.Connection) As Boolean
On Error Resume Next

Dim cmd As ADODB.Command


    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        .CommandText = "SELECT ? = stk_poolspec From STOCKMASTER"
        .CommandText = .CommandText & " Where (stk_npecode = '" & NameSpace & "')"
        .CommandText = .CommandText & " and (stk_stcknumb = '" & StockNumber & "')"
        
        
        .Parameters.Append .CreateParameter("PoolSpecific", adBoolean, adParamOutput)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        IsStockNumberSerial = .Parameters(0).Value = False
    End With


    If Err Then Err.Clear
End Function


