VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsGetPOInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A760305"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable(s) to hold property value(s)
'##ModelId=38955A7603B9
Private fPO_Date As Date
Attribute fPO_Date.VB_VarDescription = "local variable(s) to hold property value(s)"
'##ModelId=38955A7603CD
Private FNameSpace As String
'##ModelId=38955A7603D7
Private fPOCurrency As String
'##ModelId=38955A770003
Private fBuyer As String
'##ModelId=38955A77000D
Private fBuyerTelephone As String
'##ModelId=38955A770021
Private fPO_Number As String
'##ModelId=38955A770038
Private fSupplier As imsGetSupplierInfo
'##ModelId=38955A77005E
Private fDateRequested As Date

'##ModelId=38955A77007C
Dim con As ADODB.Connection
'##ModelId=38955A770180
Public Event FindComplete(ByVal Found As Boolean)

'##ModelId=38955A770090
Public Property Get DateRequested() As Date
    DateRequested = fDateRequested
End Property

'##ModelId=38955A77009A
Public Property Let NameSpace(ByVal vData As String)
    FNameSpace = vData
    fSupplier.NameSpace = vData
End Property


'##ModelId=38955A7700B8
Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property

'##ModelId=38955A7700CC
Public Property Get PO_Date() As Date
    PO_Date = fPO_Date
End Property


'##ModelId=38955A7700D6
Public Property Get PO_Number() As String
Attribute PO_Number.VB_UserMemId = 0
Attribute PO_Number.VB_MemberFlags = "200"
    PO_Number = fPO_Number
End Property

'##ModelId=38955A7700EA

'check po number in class if it not exist show message

Public Function GetValues(Optional ByVal PONumb As String) As Boolean
    
    If IsMissing(PONumb) Then
    
        If Len(Trim$(fPO_Number)) = 0 Then _
            Err.Raise 512, "Supplier info", "Nothing to find"
        
        PONumb = fPO_Number
        
    ElseIf Len(Trim$(PONumb)) = 0 Then _
        Err.Raise 512, "Supplier info", "Cannot on spaces"
        
    End If
    
    fPO_Number = PONumb
    GetValues = GetPoInfoValues
    
    RaiseEvent FindComplete(GetValues)
    Exit Function
    
Handler:

    If Err Then
        GetValues = False
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

'##ModelId=38955A7700FE
Public Property Get Supplier() As imsGetSupplierInfo
    
    If fSupplier.PO_Number <> fPO_Number Then _
        fSupplier.Find (fPO_Number)
        
    Set Supplier = fSupplier
End Property

'##ModelId=38955A770112
Public Property Get BuyerTelephone() As String
    BuyerTelephone = fBuyerTelephone
End Property

'##ModelId=38955A77011C
Public Property Get Buyer() As String
    Buyer = fBuyer
End Property

'##ModelId=38955A770130
Public Property Get POCurrency() As String
    POCurrency = fPOCurrency
End Property

'##ModelId=38955A770144

'set parameters and call store procedure to get po information

Private Function GetPoInfoValues() As Boolean
On Error GoTo Handler

Dim lng As Long
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    Set cmd = New ADODB.Command
    
    cmd.CommandText = "GETPOINFO"
    cmd.CommandType = adCmdStoredProc
    Set cmd.ActiveConnection = con
    
    cmd.Parameters.Append cmd.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@PONUMB", adVarChar, adParamInput, 15, fPO_Number)
    cmd.Parameters.Append cmd.CreateParameter("NAMESPACE", adVarChar, adParamInput, 5, FNameSpace)

    Set rs = cmd.Execute
    
    lng = cmd.Parameters("RETURN_VALUE")
    If lng = 0 Then GoTo CleanUp
    
    On Error Resume Next
    fPO_Date = rs("po_date")
    fBuyer = rs("buy_username") & ""
    fPOCurrency = rs("curr_desc") & ""
    fDateRequested = rs("po_reqddelvdate")
    fBuyerTelephone = rs("buy_phonnumb") & ""
    
    
CleanUp:
    cmd.Parameters.Delete (0)
    cmd.Parameters.Delete (0)
    cmd.Parameters.Delete (0)
    
    Set cmd = Nothing
    Set rs = Nothing
    GetPoInfoValues = lng
    Exit Function

Handler:
    If Err Then
        GetPoInfoValues = False
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'##ModelId=38955A77016C
Public Property Get Connection() As ADODB.Connection
    Set Connection = con
End Property

'##ModelId=38955A77014E
Public Property Set Connection(vNewValue As ADODB.Connection)
    Set con = vNewValue
    Set fSupplier.Connection = con
End Property

'##ModelId=38955A770176
Private Sub Class_Initialize()
    Set fSupplier = New imsGetSupplierInfo
End Sub
