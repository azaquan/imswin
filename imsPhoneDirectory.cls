VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsPhoneDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private FNpecode As String 'local copy
Private fcode As String 'local copy
Private fphonename As Variant 'local copy
Private FAddress1 As Variant 'local copy
Private FAddress2 As Variant 'local copy
Private FCity As Variant 'local copy
Private FState As Variant 'local copy
Private FZipCode As Variant 'local copy
Private FCountry As Variant 'local copy
Private FPhoneNumber As Variant 'local copy
Private FFaxNumber As Variant 'local copy
Private FTelexNumber As Variant 'local copy
Private FEmail As Variant 'local copy
Private FContact As Variant 'local copy
'local variable(s) to hold property value(s)
Private fCellular As Variant 'local copy
Private fBeeper As Variant 'local copy
Private fHome As Variant 'local copy
Private fAlternative As Variant 'local copy
Public Property Let Alternative(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fAlternative = vData
        Else
            vData = Null
        End If
    End If
    
End Property



Public Property Get Alternative() As Variant

        Alternative = fAlternative
    
End Property

Public Property Let Home(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fHome = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get Home() As Variant

        Home = fHome
    
End Property

Public Property Let Beeper(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fBeeper = vData
        Else
            vData = Null
        End If
    End If
    
End Property



Public Property Get Beeper() As Variant

        Beeper = fBeeper
    
End Property

Public Property Let Cellular(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fCellular = vData
        Else
            vData = Null
        End If
    End If
    
End Property



Public Property Get Cellular() As Variant

        Cellular = fCellular
    
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
            FEmail = Null
        End If
    End If
    
End Property

Public Property Get Email() As Variant

        Email = FEmail
    
End Property



Public Property Let TelexNumber(ByVal vData As Variant)
    If Not IsNull(vData) Then
        If Len(Trim$(vData)) Then
             FTelexNumber = vData
        Else
            FTelexNumber = Null
        End If
    End If
    
End Property

Public Property Get TelexNumber() As Variant

        TelexNumber = FTelexNumber
   
End Property

Public Property Let Faxnumber(ByVal vData As Variant)
  If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FFaxNumber = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get Faxnumber() As Variant

        Faxnumber = FFaxNumber
    
End Property



Public Property Let PhoneNumber(ByVal vData As Variant)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FPhoneNumber = vData
        Else
            vData = Null
        End If
    End If
End Property



Public Property Get PhoneNumber() As Variant

        PhoneNumber = FPhoneNumber
    
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

Public Property Let phonename(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fphonename = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get phonename() As Variant

        phonename = fphonename
    
End Property



Public Property Let Code(ByVal vData As String)
        
    vData = Trim$(vData)
                    
    If Len(Trim$(vData)) Then
        fcode = vData
    Else
        vData = Null
    End If
    
End Property


Public Property Get Code() As String

    Code = fcode
End Property



Public Property Let Npecode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FNpecode = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Npecode() As String

    Npecode = FNpecode
End Property

'SQL statement to get phone directory list

Public Function GetPhoneDirectorylist(Code As String, Namespace As String, cn As ADODB.Connection) As imsPhoneDirectory
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = " SELECT  phd_name, phd_adr1, phd_adr2, "
        .CommandText = .CommandText & " phd_city, phd_stat, phd_zipc, phd_phonnumb, "
        .CommandText = .CommandText & " phd_faxnumb, phd_telxnumb, phd_cellnumb, "
        .CommandText = .CommandText & " phd_homenumb, phd_beepnumb, phd_altnnumb, "
        .CommandText = .CommandText & " phd_mail , phd_cont, phd_ctry"
        .CommandText = .CommandText & " From PHONEDIR "
        .CommandText = .CommandText & " WHERE phd_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " AND phd_code = '" & Code & "'"
             Set rst = .Execute
    End With
        
        fcode = Code
        FAddress1 = rst!phd_adr1 & ""
        FAddress2 = rst!phd_adr2 & ""
        FFaxNumber = rst!phd_faxnumb & ""
        FCity = rst!phd_city & ""
        FContact = rst!phd_cont & ""
        FCountry = rst!phd_ctry & ""
        FEmail = rst!phd_mail & ""
        fphonename = rst!phd_name & ""
        FZipCode = rst!phd_zipc & ""
        FState = rst!phd_stat & ""
        FPhoneNumber = rst!phd_phonnumb & ""
        FTelexNumber = rst!phd_telxnumb & ""
        fAlternative = rst!phd_altnnumb & ""
        fBeeper = rst!phd_beepnumb & ""
        fCellular = rst!phd_cellnumb & ""
        fHome = rst!phd_homenumb & ""
    Set GetPhoneDirectorylist = Me
End Function

'SQL statement to get phone directory code

Public Function GetPhonedirectoryCode(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
    Dim cmd As ADODB.Command


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select phd_code from PHONEDIR"
        .CommandText = .CommandText & " where phd_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " ORDER BY phd_code"
        
         Set GetPhonedirectoryCode = .Execute
    End With
    
End Function
