VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LookUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim miutility As Iutility

Public Function GetStockNUmbers() As ADODB.Recordset
On Error GoTo handler
  Set GetStockNUmbers = New ADODB.Recordset
  Set GetStockNUmbers.ActiveConnection = miutility.Connection
  GetStockNUmbers.CursorType = adOpenKeyset
  GetStockNUmbers.CacheSize = 200
  GetStockNUmbers.Source = "SELECT stk_stcknumb, stk_desc FROM STOCKMASTER WHERE stk_npecode ='" & miutility.Namespace & "' and  stk_flag =1"
  GetStockNUmbers.Open

  Exit Function
handler:
  Err.Clear
  
End Function

Public Sub Configure(Iutility As Iutility)
Set miutility = Iutility
End Sub

Public Function GetManuFActurer(StockNumber As String) As ADODB.Recordset
Dim query As String

On Error GoTo handler
  
  Set GetManuFActurer = New ADODB.Recordset
  GetManuFActurer.ActiveConnection = miutility.Connection
  
  query = "SELECT stm_manucode, stm_partnumb, stm_estmpric From STOCKMANUFAC "
  query = query & " WHERE stm_npecode = '" & miutility.Namespace & "' AND stm_stcknumb = '" & StockNumber & "' and stm_flag<>0 "
  
  GetManuFActurer.Source = query
  GetManuFActurer.Open
  

  Exit Function
  
handler:
  Err.Clear
  
End Function


Public Function GetAllUnits() As ADODB.Recordset
On Error GoTo handler
 Dim query  As String
 Set GetAllUnits = New ADODB.Recordset
 GetAllUnits.ActiveConnection = miutility.Connection
 query = "SELECT uni_code, uni_desc FROM UNIT WHERE "
  query = query & " uni_actvflag = 1 AND uni_npecode ='" & miutility.Namespace & "' ORDER BY uni_desc, uni_code "

 GetAllUnits.Source = query
  GetAllUnits.CursorType = adOpenDynamic
 GetAllUnits.Open

 
Exit Function
handler:
  
Err.Clear
End Function

Public Function GetUnitForTheStckNo(StockNumber As String) As ADODB.Recordset
Dim RsUnit As ADODB.Recordset
Dim query As String
 On Error GoTo handler
   
   StockNumber = Trim$(StockNumber)
   
    query = "SELECT stk_primuon, stk_secouom, stk_desc,stk_compfctr,stk_estmprice From STOCKMASTER "
    query = query & " WHERE stk_stcknumb = '" & StockNumber & "' AND stk_npecode = '" & miutility.Namespace & "'"
    Set RsUnit = New ADODB.Recordset
    RsUnit.ActiveConnection = miutility.Connection
    RsUnit.Source = query
    RsUnit.CursorType = adOpenDynamic
    RsUnit.Open
    Set GetUnitForTheStckNo = RsUnit
    Exit Function
handler:
    
    Err.Clear
End Function

Public Function GetRequisitions(Namespace As String) As ADODB.Recordset
On Error Resume Next
    
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
    Set Cmd.ActiveConnection = miutility.Connection
       
    With Cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GET_BRQ"
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, miutility.Namespace)
        
        Set GetRequisitions = .Execute
    End With
    
    Set Cmd = Nothing
    If Err Then Err.Clear
    
End Function
Public Function GetReqisitionLineItem(ReqLineitem As String, ReqPonumb As String) As ADODB.Recordset
  Dim text As String
  
    text = "SELECT poitem.*, e.eccn_no, p.source FROM POITEM "
    text = text & " left outer join eccn e on e.eccnid = poitem.poi_eccnid "
    text = text & " left outer join picklist p on p.sourceid = poitem.poi_sourceid "
    text = text & " WHERE poi_npecode = '" & miutility.Namespace & "'"
    text = text & " AND poi_liitnumb = '" & ReqLineitem & "'"
    text = text & " and poi_ponumb = '" & ReqPonumb & "'"
    
    Set GetReqisitionLineItem = New ADODB.Recordset
    GetReqisitionLineItem.Source = text
    GetReqisitionLineItem.ActiveConnection = miutility.Connection
    GetReqisitionLineItem.CursorType = adOpenKeyset
    GetReqisitionLineItem.Open
    

End Function
Public Function GetAddresses(AddressType As String) As ADODB.Recordset
Dim str As String
Dim CmdPo_RecepientList As ADODB.Command
Dim cnn As ADODB.Connection
On Error GoTo Errhandler

    AddressType = Trim(UCase(AddressType))

    Set CmdPo_RecepientList = New ADODB.Command

    With CmdPo_RecepientList

    .CommandType = adCmdStoredProc
    .CommandText = "PO_RECEPIENTLIST"
    .ActiveConnection = miutility.Connection
    .Parameters.Append .CreateParameter("@NPECODE", adBSTR, adParamInput, 10, miutility.Namespace)
    .Parameters.Append .CreateParameter("@EMAIL_FAX", adBSTR, adParamInput, 7, AddressType)
    
    Set GetAddresses = .Execute
    
    End With
    
    With GetAddresses
        .Close
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With
    Exit Function
Errhandler:

    Err.Raise Err.Number, , Err.Description
    Err.Clear
End Function

Public Function GetDoctypeForUser(UserName As String) As ADODB.Recordset

On Error GoTo handler

  Set GetDoctypeForUser = New ADODB.Recordset
    
    GetDoctypeForUser.Source = "select buyr_docutype from BUYER_RIGHT where buyr_username='" & UserName & "' and buyr_npecode='" & miutility.Namespace & "' AND buyr_docutype in (select doc_code from doctype where doc_npecode=buyr_npecode and doc_actvflag<>0) "
    GetDoctypeForUser.CursorType = adOpenStatic
    GetDoctypeForUser.Open , miutility.Connection
    
  Exit Function
    
    
handler:

  Err.Raise Err.Number, , Err.Description
  Err.Clear
End Function

Public Function GetMYSite() As String

Dim Rs As New ADODB.Recordset
On Error GoTo handler

Rs.Source = "select psys_site from pesys where psys_npecode='" & miutility.Namespace & "'"
Rs.CursorType = adOpenStatic
Rs.Open , miutility.Connection
GetMYSite = Trim$(Rs!psys_site)
Exit Function
handler:
  Err.Raise Err.Number, , Err.Description
  Err.Clear
End Function

Public Function GetLocalSuppliers() As ADODB.Recordset
Dim query As String

On Error GoTo handler

 Set GetLocalSuppliers = New ADODB.Recordset

 query = "SELECT sup_code, sup_name, sup_city, sup_phonnumb,sup_contaname,sup_contaph,sup_contaFax ,sup_mail,sup_faxnumb,sup_actvflag "
 query = query & " FROM SUPPLIER WHERE sup_npecode ='" & miutility.Namespace & "' and sup_flag=0  ORDER BY sup_name"
 
 GetLocalSuppliers.Source = query
 GetLocalSuppliers.ActiveConnection = miutility.Connection
 GetLocalSuppliers.CursorType = adOpenKeyset
 GetLocalSuppliers.Open
 
 Exit Function
handler:
  Err.Clear
End Function

Public Function GetUserMenuLevel(UserId As String) As Integer
Dim Rs As New ADODB.Recordset
Rs.Source = "select usr_menuleve from xuserprofile where usr_npecode='" & miutility.Namespace & "' and usr_userid ='" & Trim$(UserId) & "'"
Rs.ActiveConnection = miutility.Connection
Rs.Open

GetUserMenuLevel = Rs!usr_menuleve
End Function

Public Function CurrencyDetlExist(Currcode As String) As Boolean
Dim Rscurrdetl As New ADODB.Recordset
Dim query As String
On Error GoTo handler
Currcode = Trim$(Currcode)

query = "SELECT count(*) EXIST FROM CURRENCYDETL WHERE CURD_CODE ='" & Currcode & "'"
query = query & " AND getdate() >=CURD_FROM AND getdate() <=CURD_TO AND CURD_NPECODE='" & miutility.Namespace & "'"
Rscurrdetl.Source = query
Rscurrdetl.ActiveConnection = miutility.Connection
Rscurrdetl.Open , , adOpenKeyset
CurrencyDetlExist = Rscurrdetl!EXIST

Exit Function
handler:
  Err.Raise Err.Number, , Err.Description
  Err.Clear
End Function

Public Function GetCustom() As ADODB.Recordset
Dim query As String
On Error GoTo handler
Set GetCustom = New ADODB.Recordset
query = "SELECT cust_cate FROM CUSTOM WHERE (cust_npecode = '" & miutility.Namespace & "' ) and cust_actvflag <>0 "
query = query & " order by cust_cate"
GetCustom.Source = query
GetCustom.Open , miutility.Connection, adOpenKeyset
Exit Function
handler:
 Err.Raise Err.Number, , Err.Description
 Err.Clear
 
End Function

Public Function CanUserDeleteLineItem(Po_number As String, po_revnumber As Integer, LineNumber As Integer) As Integer
Dim query As String
Dim Rs As ADODB.Recordset
On Error GoTo handler
CanUserDeleteLineItem = 1
Set Rs = New ADODB.Recordset
Rs.Source = "select count(*) muzammil from poitemrevision where poir_ponumb='" & Po_number & "' and poir_npecode='" & miutility.Namespace & "'  and poir_revinumb='" & po_revnumber & "' And poir_liitnumb = '" & LineNumber & "'"
Rs.ActiveConnection = miutility.Connection
Rs.Open
 CanUserDeleteLineItem = Rs!muzammil
Exit Function
handler:
  MsgBox "Error occurred while verifying whether the line item could be deleted.Err.desc " & Err.Description
  Err.Clear
End Function

Public Function CanUserDeleteRecepient(Po_number As String, po_revnumber As Integer, RecepientName As String) As Integer
Dim query As String
Dim Rs As ADODB.Recordset
On Error GoTo handler
CanUserDeleteRecepient = 1
Set Rs = New ADODB.Recordset
RecepientName = UCase(Trim$(RecepientName))
Rs.Source = "select count(*) muzammil from POrecrevision where porrc_ponumb='" & Po_number & "' and porrc_npecode='" & miutility.Namespace & "'  and porrc_revinumb='" & po_revnumber & "' And porrc_rec = '" & RecepientName & "'"
Rs.ActiveConnection = miutility.Connection
Rs.Open
 CanUserDeleteRecepient = Rs!muzammil
Exit Function
handler:
  MsgBox "Error occurred while verifying whether the Recepient could be deleted.Err.desc " & Err.Description
  Err.Clear
End Function
Public Function CanUserDeleteRemark(Po_number As String, po_revnumber As Integer, LineNumber As Integer) As Integer
Dim query As String
Dim Rs As ADODB.Recordset
On Error GoTo handler
CanUserDeleteRemark = 1
Set Rs = New ADODB.Recordset
'RecepientName = UCase(Trim$(RecepientName))
Rs.Source = "select count(*) muzammil from poremrevision where porr_ponumb='" & Po_number & "' and porr_npecode='" & miutility.Namespace & "'  and porr_revinumb='" & po_revnumber & "' And porr_linenumb = '" & LineNumber & "'"
Rs.ActiveConnection = miutility.Connection
Rs.Open
 CanUserDeleteRemark = Rs!muzammil
Exit Function
handler:
  MsgBox "Error occurred while verifying whether the Recepient could be deleted.Err.desc " & Err.Description
  Err.Clear
End Function

Public Function CanUserDeleteClause(Po_number As String, po_revnumber As Integer, LineNumber As Integer) As Integer
Dim query As String
Dim Rs As ADODB.Recordset
On Error GoTo handler
CanUserDeleteClause = 1
Set Rs = New ADODB.Recordset
'RecepientName = UCase(Trim$(RecepientName))
Rs.Source = "select count(*) muzammil from poclauserevision where porcl_ponumb='" & Po_number & "' and porcl_npecode='" & miutility.Namespace & "'  and porcl_revinumb='" & po_revnumber & "' And porcl_linenumb = '" & LineNumber & "'"
Rs.ActiveConnection = miutility.Connection
Rs.Open
 CanUserDeleteClause = Rs!muzammil
Exit Function
handler:
  MsgBox "Error occurred while verifying whether the Recepient could be deleted.Err.desc " & Err.Description
  Err.Clear
End Function

Public Function GetDefaultRecForDoctype(Doctype As String, Recepients() As String) As Integer

Dim Rs As New ADODB.Recordset
Dim Rec() As String
Dim i As Integer

On Error GoTo Errhandler

GetDefaultRecForDoctype = 1

'JCG 2008-12-15
'Rs.Source = "select dis_mail from DISTRIBUTION where dis_npecode='" & miutility.Namespace & "' and dis_gender ='USER' AND DIS_ID ='" & Doctype & "'"
Rs.Source = "select dis_mail from DISTRIBUTION where dis_npecode='" & miutility.Namespace & "' and dis_gender ='USER' AND DIS_ID ='" & Doctype & "' and dis_active = 1"
'----------------------

Rs.ActiveConnection = miutility.Connection

Rs.Open

Do While Not Rs.Eof

    ReDim Preserve Rec(i)

    Rec(i) = Rs("dis_mail")
    
    Rs.MoveNext
    
    i = i + 1
    
Loop

Recepients = Rec

GetDefaultRecForDoctype = 0

Exit Function

Errhandler:

Err.Clear

End Function

Public Function CanDocTypeAutoDist(DocCode As String, True_False As Boolean) As Integer

Dim Rs As New ADODB.Recordset

On Error GoTo Errhandler

CanDocTypeAutoDist = 1

True_False = False

Rs.Source = "select doc_autodist from doctype where doc_code='" & DocCode & "' and doc_npecode='" & miutility.Namespace & "'"

Rs.ActiveConnection = miutility.Connection

Rs.Open

True_False = Rs("doc_autodist")

CanDocTypeAutoDist = 0

Exit Function

Errhandler:

Err.Clear

End Function

Public Function LoadFQAFromLocation(Companycode As String, LocationCode As String, ByRef CompanyFQA As String, ByRef LocationFQA As String, ByRef UC As String, ByRef CC As String, ByRef StockType As String)

Dim RsCompany As New ADODB.Recordset
Dim RsLocation As New ADODB.Recordset
Dim RsUC As New ADODB.Recordset
Dim RsCC As New ADODB.Recordset

Dim RsToFQA As ADODB.Recordset

On Error GoTo ErrHand

'Get Company FQA

RsCompany.Source = "select FQA from FQA where Namespace ='" & miutility.Namespace & "' and Companycode ='" & Companycode & "' and Level ='C'"

RsCompany.Open , miutility.Connection


'Get Location FQA

RsLocation.Source = "select FQA from FQA where Namespace ='" & miutility.Namespace & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='LB' or  Level ='LS'"

RsLocation.Open , miutility.Connection



'Get US Chart FQA

RsUC.Source = "select FQA from FQA where Namespace ='" & miutility.Namespace & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='UC'"

RsUC.Open , miutility.Connection


'Get Cam Chart FQA

RsCC.Source = "select FQA from FQA where Namespace ='" & miutility.Namespace & "' and Companycode ='" & Companycode & "' and Locationcode='" & LocationCode & "' and Level ='CC'"

RsCC.Open , miutility.Connection
            
            
            If RsCompany.Eof = False Then
                    CompanyFQA = RsCompany("FQA")
            Else
                    CompanyFQA = ""
            End If
            
            If RsLocation.Eof = False Then
                    LocationFQA = RsLocation("FQA")
            Else
                    LocationFQA = ""
            End If
            
            If RsUC.Eof = False Then
                    UC = RsUC("FQA")
            Else
                    UC = ""
            End If
            
            StockType = "0000" 'rsFrom("FromStockType")
            
            If RsCC.Eof = False Then
            
                    CC = RsCC("FQA")
            Else
                    CC = ""
            End If
            







Exit Function
ErrHand:


MsgBox "Errors occurred while trying to fill the combo boxes." & Err.Description, vbCritical, "Ims"
End Function
    

Public Function LoadFQAFromStock(Stockno As String) As Boolean

Dim Rs

End Function

Public Function GetAllUSAccounts() As ADODB.Recordset

Dim Rs As New ADODB.Recordset

On Error GoTo ErrHand

Rs.Source = "select DISTINCT(fqa) from fqa where level ='UC' and namespace ='" & miutility.Namespace & "' order by FQA"
Rs.Open , miutility.Connection
Set GetAllUSAccounts = Rs
Exit Function
ErrHand:
Err.Clear
End Function

Public Function GetAllCCAccounts() As ADODB.Recordset

Dim Rs As New ADODB.Recordset

On Error GoTo ErrHand

Rs.Source = "select DISTINCT(fqa) from fqa where level ='CC' and namespace ='" & miutility.Namespace & "' order by FQA"
Rs.Open , miutility.Connection
Set GetAllCCAccounts = Rs
Exit Function
ErrHand:
Err.Clear
End Function


Public Function GetAllLocations() As ADODB.Recordset
Dim Rs As New ADODB.Recordset
On Error GoTo ErrHand

Rs.Source = "select DISTINCT(fqa) from fqa where namespace ='" & miutility.Namespace & "' and  level ='LB' or level ='LC' order by FQA"
Rs.Open , miutility.Connection

Set GetAllLocations = Rs

Exit Function
ErrHand:
Err.Clear


End Function

Public Function GetListofEccns(AllORActive As Integer) As ADODB.Recordset

On Error GoTo Errhandler

Dim RsEccn As New ADODB.Recordset

RsEccn.Source = "SELECT eccnid, eccn_no, eccn_desc FROM eccn WHERE (eccn_npecode = '" & miutility.Namespace & "')" ' and eccn_active =1"
If AllORActive = 1 Then RsEccn.Source = RsEccn.Source & " and eccn_active =1"
RsEccn.Source = RsEccn.Source & " order by eccn_no "
RsEccn.Open , miutility.Connection, adOpenKeyset, adLockPessimistic

Set GetListofEccns = RsEccn

Exit Function

Errhandler:

Err.Clear

End Function
Public Function GetListofEccnSource(AllORActive As Integer) As ADODB.Recordset

On Error GoTo Errhandler

Dim RsEccn As New ADODB.Recordset

RsEccn.Source = "SELECT sourceid, source FROM picklist WHERE (npecode = '" & miutility.Namespace & "')"
If AllORActive = 1 Then RsEccn.Source = RsEccn.Source & " and source_active =1"
RsEccn.Open , miutility.Connection, adOpenKeyset, adLockPessimistic

Set GetListofEccnSource = RsEccn

Exit Function

Errhandler:

Err.Clear

End Function
