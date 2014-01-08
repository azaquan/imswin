VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private fforwnpecode As String 'local copy
Private fforwcode As String 'local copy
Private fforwname As String 'local copy
Private FAddress1 As Variant 'local copy
Private FAddress2 As Variant 'local copy
Private FCity As Variant 'local copy
Private FZipCode As Variant 'local copy
Private FState As Variant 'local copy
Private fcontury As Variant 'local copy
Private fphonnumb As Variant 'local copy
Private fFaxnumb As Variant 'local copy
Private fTelexnumb As Variant 'local copy
Private FEmail As Variant 'local copy
Private FContact As Variant 'local copy

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



Public Property Let phonnumb(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fphonnumb = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get phonnumb() As Variant

        phonnumb = fphonnumb
    
End Property



Public Property Let contury(ByVal vData As Variant)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            fcontury = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get contury() As Variant

        contury = fcontury
   
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



Public Property Let forwname(ByVal vData As String)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fforwname = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get forwname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.forwname
    forwname = fforwname
End Property



Public Property Let forwcode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fforwcode = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get forwcode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.forwcode
    forwcode = fforwcode
End Property



Public Property Let forwnpecode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             fforwnpecode = vData
        Else
            vData = Null
        End If
    End If
End Property


Public Property Get forwnpecode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.forwnpecode
    forwnpecode = fforwnpecode
End Property

Public Function GetForwarder(Code As String, Namespace As String, cn As ADODB.Connection) As imsForwarder
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)

    With cmd
        Set .ActiveConnection = cn
        .CommandText = " SELECT forw_name, forw_adr1, "
        .CommandText = .CommandText & " forw_adr2, forw_city, "
        .CommandText = .CommandText & " forw_stat, forw_zipc, "
        .CommandText = .CommandText & " forw_ctry, forw_phonnumb, "
        .CommandText = .CommandText & " forw_faxnumb, forw_mail, "
        .CommandText = .CommandText & " forw_telxnumb,forw_cont "
        .CommandText = .CommandText & " From FORWARDER "
        .CommandText = .CommandText & " WHERE forw_code =  '" & Code & "'"
        .CommandText = .CommandText & " AND forw_npecode = '" & Namespace & "'"

        Set rst = .Execute
    End With

    
    FAddress1 = rst!forw_adr1 & ""
    FAddress2 = rst!forw_adr2 & ""
    fFaxnumb = rst!forw_faxnumb & ""
    FCity = rst!forw_ctry & ""
    FContact = rst!forw_cont & ""
    fcontury = rst!forw_ctry & ""
    FEmail = rst!forw_mail & ""
    fforwname = rst!forw_name & ""
    FZipCode = rst!forw_zipc & ""
    FState = rst!forw_stat & ""
    fphonnumb = rst!forw_phonnumb & ""
    fTelexnumb = rst!forw_telxnumb & ""
    
    Set GetForwarder = Me
    
End Function


Public Function GetForwarderCode(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command


    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "Select forw_code from FORWARDER"
        .CommandText = .CommandText & " where forw_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " ORDER BY forw_code"
        
         Set GetForwarderCode = .Execute
    End With
    
        
    
End Function
