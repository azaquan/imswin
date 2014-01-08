VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private fcode As String 'local copy
Private fCompanyName As Variant 'local copy
Private FAddress1 As Variant 'local copy
Private FAddress2 As Variant 'local copy
Private FCity As Variant 'local copy
Private FState As Variant 'local copy
Private FZipCode As Variant 'local copy
Private FCountry As Variant 'local copy
Private fPhoneNumb As Variant 'local copy
Private fFaxnumb As Variant 'local copy
Private fTelexnumb As Variant 'local copy
Private FEmail As Variant 'local copy
Private FContact As Variant 'local copy
Private Factvflag As Boolean

Private FModifiedUSer As String
Private fCreateUser As String

Public Property Let Actvflag(ByVal vData As Boolean)
    Factvflag = vData
End Property

Public Property Get Actvflag() As Boolean

        Actvflag = Factvflag
   
End Property

Public Property Let Contact(ByVal vData As Variant)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FContact = vData
        Else
            vData = Null
        End If
    End If
End Property



Public Property Get Contact() As Variant

        Contact = FContact
   
End Property

Public Property Let Email(ByVal vData As Variant)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FEmail = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Email() As Variant

        Email = FEmail
    
End Property

Public Property Let Telexnumb(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fTelexnumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Telexnumb() As Variant

        Telexnumb = fTelexnumb
    
End Property

Public Property Let Faxnumb(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fFaxnumb = vData
        Else
            vData = Null
        End If
    End If
    
    
End Property

Public Property Get Faxnumb() As Variant

        Faxnumb = fFaxnumb
    
End Property

Public Property Let PhoneNumb(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fPhoneNumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get PhoneNumb() As Variant

        PhoneNumb = fPhoneNumb
    
End Property

Public Property Let Country(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FCountry = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Country() As Variant

        Country = FCountry
    
End Property

Public Property Let ZipCode(ByVal vData As Variant)
     If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FZipCode = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get ZipCode() As Variant

        ZipCode = FZipCode
    
End Property

Public Property Let State(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FState = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get State() As Variant

        State = FState
    
End Property

Public Property Let City(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FCity = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get City() As Variant

        City = FCity
    
End Property

Public Property Let address2(ByVal vData As Variant)
  If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FAddress2 = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get address2() As Variant

        address2 = FAddress2
   
End Property

Public Property Let address1(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FAddress1 = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get address1() As Variant

        address1 = FAddress1
    
End Property

Public Property Let CompanyName(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fCompanyName = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get CompanyName() As Variant

        CompanyName = fCompanyName
    
End Property

Public Property Let Code(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fcode = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get Code() As String

    Code = fcode
End Property

'SQL statement to get company name

Public Function GetCompanyCode(NameSpace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select com_compcode from COMPANY"
        .CommandText = .CommandText & " where com_npecode = '" & NameSpace & "'"
        .CommandText = .CommandText & " ORDER BY com_compcode"
        
         Set GetCompanyCode = .Execute
    End With
    
End Function

'SQL statement to get company list

Public Function GetCompanyList(Code As String, NameSpace As String, cn As ADODB.Connection) As imsCompany
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        Set .ActiveConnection = cn
        .CommandText = " SELECT com_name, com_adr1, "
        .CommandText = .CommandText & " com_adr2, com_city, "
        .CommandText = .CommandText & " com_stat, com_zipc, "
        .CommandText = .CommandText & " com_ctry, com_phonnumb, "
        .CommandText = .CommandText & " com_faxnumb, com_mail, "
        .CommandText = .CommandText & " com_telxnumb,com_cont,com_actvflag "
        .CommandText = .CommandText & " From COMPANY "
        .CommandText = .CommandText & " WHERE com_compcode =  '" & Code & "'"
        .CommandText = .CommandText & " AND com_npecode = '" & NameSpace & "'"

        Set rst = .Execute
    End With

    
    FAddress1 = rst!com_adr1 & ""
    FAddress2 = rst!com_adr2 & ""
    fFaxnumb = rst!com_faxnumb & ""
    FCity = rst!com_ctry & ""
    FContact = rst!com_cont & ""
    FCountry = rst!com_ctry & ""
    FEmail = rst!com_mail & ""
    fCompanyName = rst!com_name & ""
    FZipCode = rst!com_zipc & ""
    FState = rst!com_stat & ""
    fPhoneNumb = rst!com_phonnumb & ""
    fTelexnumb = rst!com_telxnumb & ""
    Factvflag = IIf(rst!com_actvflag, 1, 0)
    
    
    
    Set GetCompanyList = Me
End Function

Public Property Get CreatedBy() As String
    CreatedBy = fCreateUser
End Property

Public Property Let CreatedBy(ByVal vNewValue As String)
    fCreateUser = vNewValue
End Property

Public Property Get ModifiedBy() As String
    CreatedBy = FModifiedUSer
End Property

Public Property Let ModifiedBy(ByVal vNewValue As String)
    FModifiedUSer = vNewValue
End Property

