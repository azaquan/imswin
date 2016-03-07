VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)

Dim Changed As Boolean
Dim FLastCode As String
Dim cn As ADODB.Connection
Private FCompanyCode As String
Private FLocationCode As String
Private FLocationName As String
Private FAddress1 As String
Private FAddress2 As String
Private FCity As String
Private FState As String
Private FCountry As String
Private FPhoneNumber As String
Private FFaxNumber As String
Private FEmail As String
Private FTelexNumber As String
Private FContact As String
Private FGender As String
Private FZipCode As String
Private FNameSpace As String
Private FUser As String
Private FMode As Boolean
Private FFlag As Boolean


Public Property Let User(ByVal vData As String)
    vData = Trim$(vData)
    If FUser <> vData Then FUser = vData: Changed = True
End Property

Public Property Get User() As String
    User = FUser
End Property

Public Property Get Flag() As Boolean
    Flag = FFlag
End Property

Public Property Let Flag(ByVal vData As Boolean)
    vData = Trim$(vData)
    If FFlag <> vData Then FFlag = vData: Changed = True
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = Changed
End Property

Public Property Let ZipCode(ByVal vData As String)
    
    vData = Trim$(vData)
    If FZipCode <> vData Then
        Changed = True
        FZipCode = vData
    End If
End Property

Public Property Get ZipCode() As String
    ZipCode = FZipCode
End Property

Public Property Let Gender(ByVal vData As String)
    vData = Trim$(vData)
    If FGender <> vData Then FGender = vData: Changed = True
End Property

Public Property Get Gender() As String
    Gender = FGender
End Property

Public Property Let Contact(ByVal vData As String)
    vData = Trim$(vData)
    If FContact <> vData Then FContact = vData: Changed = True
End Property

Public Property Get Contact() As String
    Contact = FContact
End Property

Public Property Let TelexNumber(ByVal vData As String)
    If FTelexNumber <> vData Then FTelexNumber = vData: Changed = True
End Property

Public Property Get TelexNumber() As String
    TelexNumber = FTelexNumber
End Property

Public Property Let Email(ByVal vData As String)
    vData = Trim$(vData)
    If FEmail <> vData Then FEmail = vData: Changed = True
End Property

Public Property Get Email() As String
    Email = FEmail
End Property

Public Property Let Faxnumber(ByVal vData As String)
    vData = Trim$(vData)
    If FFaxNumber <> vData Then FFaxNumber = vData: Changed = True
End Property

Public Property Get Faxnumber() As String
    Faxnumber = FFaxNumber
End Property

Public Property Let PhoneNumber(ByVal vData As String)
    vData = Trim$(vData)
    If FPhoneNumber <> vData Then FPhoneNumber = vData: Changed = True
End Property

Public Property Get PhoneNumber() As String
    PhoneNumber = FPhoneNumber
End Property

Public Property Let Country(ByVal vData As String)
    vData = Trim$(vData)
    If FCountry <> vData Then FCountry = vData: Changed = True
End Property

Public Property Get Country() As String
   Country = FCountry
End Property

Public Property Let State(ByVal vData As String)
    vData = Trim$(vData)
    If FState <> vData Then FState = vData: Changed = True
End Property

Public Property Get State() As String
    State = FState
End Property

Public Property Let City(ByVal vData As String)
    vData = Trim$(vData)
    If FCity <> vData Then FCity = vData: Changed = True
End Property

Public Property Get City() As String
    City = FCity
End Property

Public Property Let address2(ByVal vData As String)
    vData = Trim$(vData)
    If FAddress2 <> vData Then FAddress2 = vData: Changed = True
End Property

Public Property Get address2() As String
    address2 = FAddress2
End Property

Public Property Let address1(ByVal vData As String)
    vData = Trim$(vData)
    If FAddress1 <> vData Then FAddress1 = vData: Changed = True
End Property

Public Property Get address1() As String
    address1 = FAddress1
End Property

Public Property Let LocationName(ByVal vData As String)
    vData = Trim$(vData)
    If FLocationName <> vData Then FLocationName = vData: Changed = True
End Property

Public Property Get LocationName() As String
    LocationName = FLocationName
End Property

Public Property Let LocationCode(ByVal vData As String)
    vData = Trim$(vData)
'    If FMode = False Then Exit Property
    If FLocationCode <> vData Then FLocationCode = vData: Changed = True
End Property

Public Property Get LocationCode() As String
     LocationCode = FLocationCode
End Property

Public Property Let CompanyCode(ByVal vData As String)
    vData = Trim$(vData)
'    If FMode = False Then Exit Property
    If FCompanyCode <> vData Then FCompanyCode = vData: Changed = True
End Property

Public Property Get CompanyCode() As String
    CompanyCode = FCompanyCode
End Property

Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
    vNewValue = Trim$(vNewValue)
'    If FMode = False Then Exit Property
    If FNameSpace <> vNewValue Then FNameSpace = vNewValue: Changed = True
End Property

Public Property Get Connection() As ADODB.Connection

End Property

Public Property Let Connection(ByVal vNewValue As ADODB.Connection)

End Property

'set gender llist details

Public Function GenderList() As String()
Dim str(5) As String

    str(0) = "PO"
    str(1) = "BASE"
    str(2) = "OTHER"
    str(3) = "SITE"
    str(4) = "REPAIR"
    str(5) = "FABRICA"
    GenderList = str
    
End Function

'SQL statement get company list

Public Function CompanyList(NameSpace As String, con As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command
    
    
    If con Is Nothing Then Set con = cn
    If con Is Nothing Then Exit Function
    If Not con.State And adStateOpen Then con.Open
    
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        
        .Prepared = True
        .CommandText = "SELECT com_compcode Comp "
        
        .CommandText = .CommandText & "FROM COMPANY WHERE"
        .CommandText = .CommandText & " (com_npecode = ?) and "
        .CommandText = .CommandText & " com_actvflag = 1 ORDER BY com_npecode"
        
        
        Set CompanyList = .Execute(0, Array(NameSpace))
    End With
End Function

'SQL statement get location values

Public Function Find(Optional SNamespace As String, Optional sCode As String, Optional CompCode As String, Optional con As ADODB.Connection) As imsLocation
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
    
    
    FLastCode = FLocationCode
    If Len(Trim$(sCode)) = 0 Then sCode = FLocationCode
    If Len(Trim$(CompCode)) = 0 Then CompCode = FCompanyCode
    If Len(Trim$(SNamespace)) = 0 Then SNamespace = FNameSpace
   
    
    'If Len(Trim$(sCode)) = 0 Then Exit Function
    'If Len(Trim$(Compcode)) = 0 Then Exit Function
    'If Len(Trim$(sNameSpace)) = 0 Then Exit Function
    
    If con Is Nothing Then Set con = cn
    If con Is Nothing Then Exit Function
    If Not con.State And adStateOpen Then con.Open
    
    Set cmd = New ADODB.Command
    
    With cmd
        .ActiveConnection = con
        .CommandText = " SELECT loc_locacode, loc_npecode, loc_compcode, loc_name, loc_npecode,"
        .CommandText = .CommandText & "  loc_adr1, loc_adr2, loc_city, loc_stat, loc_zipc,"
        
        .CommandText = .CommandText & "  loc_ctry, loc_phonnumb, loc_faxnumb,"
        .CommandText = .CommandText & " loc_telxnumb, loc_mail, loc_cont, loc_gender, loc_actvflag"
        
        .CommandText = .CommandText & " FROM LOCATION WHERE"
        .CommandText = .CommandText & " (loc_npecode = ?) AND (loc_locacode = ?) and (loc_compcode = ?)"
        'Modified By Muzammil 10/1/2000
        'Reason - This form would dispaly only Locations which are not SITE
        '.CommandText = .CommandText & " and loc_gender <>'SITE'"
        .CommandText = .CommandText & " ORDER BY loc_locacode"
        
        If Len(sCode) = 0 Then sCode = " "
        If Len(CompanyCode) = 0 Then CompanyCode = " "
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 5, SNamespace)
        
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 10, sCode)
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 10, CompCode)
        
        Set rs = .Execute
    End With
    
    Set cmd = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        With rs
            FCity = Trim$(!loc_city & "")
            FContact = Trim$(!loc_cont & "")
            FAddress1 = Trim$(!loc_adr1 & "")
            FAddress2 = Trim$(!loc_adr2 & "")
            FCompanyCode = Trim$(!loc_compcode & "")
            FCountry = Trim$(!loc_ctry & "")
            FState = Trim$(!loc_stat & "")
            FZipCode = Trim$(!loc_zipc & "")
            FEmail = Trim$(!loc_mail & "")
            FFaxNumber = Trim$(!loc_faxnumb & "")
            FGender = Trim$(!loc_gender & "")
            FLocationCode = Trim$(!loc_locacode & "")
            FLocationName = Trim$(!loc_name & "")
            FPhoneNumber = Trim$(!loc_phonnumb & "")
            FNameSpace = Trim$(!loc_npecode & "")
            FFlag = IIf(!loc_actvflag, 1, 0)
        End With
    Else
        Call AddNew
    End If
    
    Set Find = Me
    FMode = False
    Changed = False
End Function

'SQl statement get location code

Public Function LocationCodes(Optional SNamespace As String, Optional Comp As String, Optional con As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command
On Error GoTo Handled

     
    If Len(Trim$(Comp)) = 0 Then Comp = FCompanyCode
    If Len(Trim$(SNamespace)) = 0 Then SNamespace = FNameSpace
   
    If con Is Nothing Then Set con = cn
    
    If con Is Nothing Then Exit Function
    If Not con.State And adStateOpen Then con.Open
    
    If Len(Trim$(Comp)) = 0 Then Exit Function
    If Len(Trim$(SNamespace)) = 0 Then Exit Function
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        
        .Prepared = True
        .CommandText = "SELECT loc_locacode Code"
        .CommandText = .CommandText & " FROM LOCATION WHERE"
        .CommandText = .CommandText & " (loc_npecode = ?) and (loc_compcode = ?)"
        .CommandText = .CommandText & " and loc_gender <> 'SITE'"
        
        
        .CommandText = .CommandText & " ORDER BY loc_locacode"

        Set LocationCodes = .Execute(0, Array(SNamespace, Comp))
    End With
    
    FMode = False
    Set cmd = Nothing
    Exit Function
Handled:
    If Err Then Err.Clear: MsgBox "Unkown error"
End Function

Public Sub Revert()
    If Len(Trim$(FCompanyCode)) Then Exit Sub
    If Len(Trim$(FLocationCode)) Then Exit Sub
    
End Sub

Public Sub Refresh()
    Call Revert
End Sub

'call stored procedure to insert a location record

Public Function Update(Optional con As ADODB.Connection) As Boolean
On Error GoTo Handler

Dim va As Variant
Dim cmd As ADODB.Command

    If con Is Nothing Then Set con = cn
    
    If con Is Nothing Then Exit Function
    If Not con.State And adStateOpen Then con.Open
    
    Set cmd = New ADODB.Command
    
    If Not Validate Then Exit Function
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "LOCATIONINSERT"
        
        .parameters.Append .CreateParameter("rt", adInteger, adParamReturnValue)
        .parameters.Append .CreateParameter("lc", adVarChar, adParamInput, 10, FLocationCode)
        
        .parameters.Append .CreateParameter("np", adVarChar, adParamInput, 5, FNameSpace)
        .parameters.Append .CreateParameter("cc", adVarChar, adParamInput, 10, FCompanyCode)
        .parameters.Append .CreateParameter("nm", adVarChar, adParamInput, 35, FLocationName)
        
        va = FAddress1
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("ad1", adVarChar, adParamInput, 25, va)
        
        va = FAddress2
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("ad2", adVarChar, adParamInput, 25, va)
        
        va = FCity
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("city", adVarChar, adParamInput, 25, va)

        
        va = FState
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("st", adVarChar, adParamInput, 2, va)
        
        va = FZipCode
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("zp", adVarChar, adParamInput, 11, va)
        
        
        va = FCountry
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("ctry", adVarChar, adParamInput, 25, FCountry)
        
        
        va = FPhoneNumber
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("pn", adVarChar, adParamInput, 25, va)
        
        va = FFaxNumber
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("fax", adVarChar, adParamInput, 50, va)
        
        
        va = FTelexNumber
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("tx", adVarChar, adParamInput, 25, va)
        
        va = FEmail
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("Email", adVarChar, adParamInput, 59, va)
        
        va = FContact
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("cont", adVarChar, adParamInput, 25, va)
        
        va = FGender
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("gender", adVarChar, adParamInput, 10, va)
        
        .parameters.Append .CreateParameter("flag", adBoolean, adParamInput, , FFlag)
        
        va = FUser
        If Len(va) = 0 Then va = Null
        .parameters.Append .CreateParameter("user", adVarChar, adParamInput, 20, FUser)
        
        Call .Execute(Options:=adExecuteNoRecords)
        
        Update = .parameters(0).value = 0
    End With

    Set cmd = Nothing
    FMode = Update = False
    MsgBox "Insert into Location was completed successfully"
    Exit Function
    
Handler:
    Set cmd = Nothing
    If Err Then MsgBox Err.Description: Err.Clear
End Function

'clear location class data fields

Public Function AddNew() As imsLocation
    
    FLastCode = FLocationCode
    
    FCity = ""
    FEmail = ""
    FState = ""
    FGender = ""
    FContact = ""
    FZipCode = ""
    FCountry = ""
    FAddress1 = ""
    FAddress2 = ""
    FCompanyCode = ""
    FFaxNumber = ""
    FLocationCode = ""
    FLocationName = ""
    FPhoneNumber = ""
    FFlag = vbChecked
    
    FMode = True
    Changed = False
    Set AddNew = Me
End Function

Public Function InsertMode() As Boolean
    InsertMode = FMode
End Function

'validate data format

Public Function Validate() As Boolean
Dim msg As String


    If Len(FLocationCode) = 0 Then
        msg = "Location code cannot be empty"
    ElseIf Len(FCompanyCode) = 0 Then
        msg = "Company Code cannot be left empty"
        
    ElseIf Len(FLocationName) = 0 Then
        msg = "Location name cannot be left empty"
        
    ElseIf Len(FGender) = 0 Then
        msg = "Location Type cannot ne left empty"
        
    End If
    
    Validate = Len(msg) = 0
    If Len(msg) Then MsgBox msg
End Function

Public Property Get LastCode() As String
    LastCode = FLastCode
End Property



Public Function LocationCodesOfSites(Optional SNamespace As String, Optional Comp As String, Optional con As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command
On Error GoTo Handled

     
    If Len(Trim$(Comp)) = 0 Then Comp = FCompanyCode
    If Len(Trim$(SNamespace)) = 0 Then SNamespace = FNameSpace
   
    If con Is Nothing Then Set con = cn
    
    If con Is Nothing Then Exit Function
    If Not con.State And adStateOpen Then con.Open
    
    If Len(Trim$(Comp)) = 0 Then Exit Function
    If Len(Trim$(SNamespace)) = 0 Then Exit Function
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        
        .Prepared = True
        .CommandText = "SELECT loc_locacode Code"
        .CommandText = .CommandText & " FROM LOCATION WHERE"
        .CommandText = .CommandText & " (loc_npecode = ?) and (loc_compcode = ?)"
        
        'Modified By Muzammil 10/1/00
        'Reason - This form is used only for Admin only
        .CommandText = .CommandText & " and loc_gender = 'SITE'"
        .CommandText = .CommandText & " ORDER BY loc_locacode"

        Set LocationCodesOfSites = .Execute(0, Array(SNamespace, Comp))
    End With
    
    FMode = False
    Set cmd = Nothing
    Exit Function
Handled:
    If Err Then Err.Clear: MsgBox "Unkown error"
End Function

