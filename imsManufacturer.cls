VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsManufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private fManuFcode As String 'local copy
Private fManufName As String 'local copy
Private FAddress1 As String 'local copy
Private FAddress2 As Variant 'local copy
Private FCity As String 'local copy
Private FZipCode As Variant 'local copy
Private FState As Variant 'local copy
Private fPhoneNumb As Variant 'local copy
Private fFaxnumb As Variant 'local copy
Private fTelexnumb As Variant 'local copy
Private FEmail As Variant 'local copy
Private FCountry As String 'local copy
Private FContact As Variant 'local copy
'local variable(s) to hold property value(s)
Private fnpecode As String 'local copy
Public Property Let npecode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fnpecode = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get npecode() As String

    npecode = fnpecode
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



Public Property Let Country(ByVal vData As String)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FCountry = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get Country() As String

    Country = FCountry
End Property



Public Property Let Email(ByVal vData As Variant)
 If Not IsNull(vData) Then
        If Len(Trim$(vData)) Then
            FEmail = vData
        Else
            FEmail = Null
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
            fTelexnumb = Null
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



Public Property Let City(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FCity = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get City() As String

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

Public Property Let address1(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FAddress1 = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get address1() As String

    address1 = FAddress1
End Property



Public Property Let ManufName(ByVal vData As String)
        If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fManufName = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get ManufName() As String

    ManufName = fManufName
End Property

Public Property Let ManuFcode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fManuFcode = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get ManuFcode() As String
    ManuFcode = fManuFcode
End Property

'SQL sttatment to get manufacturer code

Public Function GetManufacturerCode(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
    Dim cmd As ADODB.Command


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select man_code from MANUFACTURER"
        .CommandText = .CommandText & " where man_npecode = '" & Namespace & "'"
        
         Set GetManufacturerCode = .Execute
    End With
    
End Function

'SQL statement to get manufacurer list

Public Function GetManufacturerlist(Code As String, Namespace As String, cn As ADODB.Connection) As imsManufacturer
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    

        Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        Set .ActiveConnection = cn
        .CommandText = "SELECT man_code, man_name, man_adr1, man_adr2,"
        .CommandText = .CommandText & " man_city, man_stat, man_zipc, man_ctry, "
        .CommandText = .CommandText & " man_phonnumb, man_faxnumb, man_telxnumb, "
        .CommandText = .CommandText & " man_mail , man_cont "
        .CommandText = .CommandText & " From MANUFACTURER "
        .CommandText = .CommandText & " WHERE man_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " AND man_code ='" & Code & "'"
            
        Set rst = .Execute
        
   End With
   
        FAddress1 = rst!man_adr1 & ""
        FAddress2 = rst!man_adr2 & ""
        FCity = rst!man_city & ""
        FContact = rst!man_cont & ""
        FCountry = rst!man_ctry & ""
        FEmail = rst!man_mail & ""
        fFaxnumb = rst!man_faxnumb & ""
        FState = rst!man_stat & ""
        FZipCode = rst!man_zipc & ""
        fPhoneNumb = rst!man_phonnumb & ""
        fManuFcode = rst!man_code & ""
        fManufName = rst!man_name & ""
        fTelexnumb = rst!man_telxnumb & ""
        
    Set GetManufacturerlist = Me
End Function
