VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsGetSupplierInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A700325"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable(s) to hold property value(s)
'##ModelId=38955A7203B4
Private fPO_Number As String
Attribute fPO_Number.VB_VarDescription = "local variable(s) to hold property value(s)"
'##ModelId=38955A730026
Private fSupplierName As String
'##ModelId=38955A73003A
Private fAddress1 As String
'##ModelId=38955A730044
Private fAddress2 As String
'##ModelId=38955A730058
Private fCity As String
'##ModelId=38955A730062
Private fState As String
'##ModelId=38955A730076
Private fZipCode As String
'##ModelId=38955A730080
Private fCountry As String
'##ModelId=38955A730094
Private fTelephone As String
'##ModelId=38955A73009E
Private fFAX As String
'##ModelId=38955A7300B2
Private fEMail As String
'##ModelId=38955A7300BC
Private FNameSpace As String
'##ModelId=38955A730120
Private con As ADODB.Connection
'##ModelId=38955A730333
Public Event FindComplete(ByVal Found As Boolean)


'##ModelId=38955A73018E
Public Property Let NameSpace(ByVal vData As String)
    FNameSpace = vData
End Property

'##ModelId=38955A730210
Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property

'##ModelId=38955A730224

'check po number in class if it not exist, show error message

Public Function Find(Optional ByVal PONumb As String) As Boolean
    
    If IsMissing(PONumb) Then
    
        If Len(Trim$(fPO_Number)) = 0 Then _
            Err.Raise 512, "Supplier info", "Nothing to find"
        
        PONumb = fPO_Number
        
    ElseIf Len(Trim$(PONumb)) = 0 Then _
        Err.Raise 512, "Supplier info", "Cannot on spaces"
        
    End If
    
    fPO_Number = PONumb
    Find = FindSupplier(PONumb)
    
    RaiseEvent FindComplete(Find)
End Function

'##ModelId=38955A730239

'assign store procedure parameters and call it to get supplier
'recordset

Private Function FindSupplier(PONumb As String) As Boolean
On Error GoTo ErrHandler

Dim str As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = con
    cmd.CommandText = "GETSUPPLIERINFOFROMPONUMBER_SP"
    
    cmd.Parameters.Append cmd.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@PONUMB", adVarChar, adParamInput, 15, PONumb)
    cmd.Parameters.Append cmd.CreateParameter("NAMESPACE", adVarChar, adParamInput, 5, FNameSpace)
    
    
    Set rs = cmd.Execute
    If cmd.Parameters("RETURN_VALUE") = 0 Then GoTo CleanUp
    
    fCity = rs("sup_city") & ""
    fState = rs("sup_stat") & ""
    fEMail = rs("sup_mail") & ""
    fFAX = rs("sup_faxnumb") & ""
    fZipCode = rs("sup_zipc") & ""
    fCountry = rs("sup_ctry") & ""
    fAddress1 = rs("sup_adr1") & ""
    fAddress2 = rs("sup_adr2") & ""
    fSupplierName = rs("sup_name") & ""
    fTelephone = rs("sup_phonnumb") & ""
    
    FindSupplier = True
    
CleanUp:

    Set cmd = Nothing
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Public Property Let EMail(ByVal vData As String)
'    fEMail = vData
'End Property

'##ModelId=38955A73024D
Public Property Get EMail() As String
    EMail = fEMail
End Property

'Public Property Let FAX(ByVal vData As String)
'    fFAX = vData
'End Property

'##ModelId=38955A730261
Public Property Get FAX() As String
    FAX = fFAX
End Property

'Public Property Let Telephone(ByVal vData As String)
'    fTelephone = vData
'End Property

'##ModelId=38955A73026B
Public Property Get Telephone() As String
    Telephone = fTelephone
End Property

'Public Property Let Country(ByVal vData As String)
'    fCountry = vData
'End Property

'##ModelId=38955A73027F
Public Property Get Country() As String
    Country = fCountry
End Property

'Public Property Let ZipCode(ByVal vData As String)
'    fZipCode = vData
'End Property

'##ModelId=38955A730289
Public Property Get ZipCode() As String
    ZipCode = fZipCode
End Property

'Public Property Let State(ByVal vData As String)
'    fState = vData
'End Property

'##ModelId=38955A73029D
Public Property Get State() As String
    State = fState
End Property

'Public Property Let City(ByVal vData As String)
'    fCity = vData
'End Property

'##ModelId=38955A7302A7
Public Property Get City() As String
    City = fCity
End Property

'Public Property Let Address2(ByVal vData As String)
'    fAddress2 = vData
'End Property

'##ModelId=38955A7302BB
Public Property Get Address2() As String
    Address2 = fAddress2
End Property

'Public Property Let Address1(ByVal vData As String)
'    fAddress1 = vData
'End Property

'##ModelId=38955A7302C5
Public Property Get Address1() As String
    Address1 = fAddress1
End Property

'Public Property Let SupplierName(ByVal vData As String)
'    fSupplierName = vData
'End Property

'##ModelId=38955A7302D9
Public Property Get SupplierName() As String
Attribute SupplierName.VB_UserMemId = 0
Attribute SupplierName.VB_MemberFlags = "200"
    SupplierName = fSupplierName
End Property

'##ModelId=38955A7302E3
Public Property Let PO_Number(ByVal vData As String)
    fPO_Number = vData
End Property

'##ModelId=38955A730301
Public Property Get PO_Number() As String
    PO_Number = fPO_Number
End Property

'##ModelId=38955A730329
Public Property Get Connection() As ADODB.Connection
    Set Connection = con
End Property

'##ModelId=38955A73030B
Public Property Set Connection(vNewValue As ADODB.Connection)
    Set con = vNewValue
End Property
