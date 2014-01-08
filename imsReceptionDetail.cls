VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsReceptionDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private FNpecode As String 'local copy
Private FRecpnumber As String 'local copy
Private FRecepSerl As Integer 'local copy
Private FRecplineitem As String 'local copy
Private FRecpriqtydelived As Double 'local copy
Private FRecpsecqtydelived As Double 'local copy
Private FRepdescription As Variant 'local copy
Private FRecpUintprice As Double 'local copy
Private FRecptotalprice As Double 'local copy
Private FRecpDatedelivd As Date 'local copy
Private FRecppartnumb As String 'local copy
Public Property Let Recppartnumb(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecppartnumb = vData
        Else
            FRecppartnumb = Null
        End If
    End If
    
End Property

Public Property Get Recppartnumb() As String

    Recppartnumb = FRecppartnumb
End Property

Public Property Let RecpDatedelivd(ByVal vData As Date)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpDatedelivd = vData
        Else
            FRecpDatedelivd = Null
        End If
    End If
    
End Property


Public Property Get RecpDatedelivd() As Date

    RecpDatedelivd = FRecpDatedelivd
End Property

Public Property Let Recptotalprice(ByVal vData As Double)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecptotalprice = vData
        Else
            FRecptotalprice = Null
        End If
    End If
    
End Property
Public Property Get Recptotalprice() As Double

    Recptotalprice = FRecptotalprice
End Property

Public Property Let RecpUintprice(ByVal vData As Double)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpUintprice = vData
        Else
            FRecpUintprice = Null
        End If
    End If
    
End Property

Public Property Get RecpUintprice() As Double

    RecpUintprice = FRecpUintprice
End Property

Public Property Let Repdescription(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRepdescription = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Repdescription() As Variant

        Repdescription = FRepdescription
    
End Property

Public Property Let Recpsecqtydelived(ByVal vData As Double)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpsecqtydelived = vData
        Else
            FRecpsecqtydelived = Null
        End If
    End If
    
End Property

Public Property Get Recpsecqtydelived() As Double

    Recpsecqtydelived = FRecpsecqtydelived
End Property

Public Property Let Recpriqtydelived(ByVal vData As Double)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpriqtydelived = vData
        Else
            FRecpriqtydelived = Null
        End If
    End If
    
End Property

Public Property Get Recpriqtydelived() As Double

    Recpriqtydelived = FRecpriqtydelived
End Property

Public Property Let Recplineitem(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecplineitem = vData
        Else
            FRecplineitem = Null
        End If
    End If
    
End Property

Public Property Get Recplineitem() As String

    Recplineitem = FRecplineitem
End Property

Public Property Let RecepSerl(ByVal vData As Integer)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecepSerl = vData
        Else
            FRecepSerl = Null
        End If
    End If
    
End Property

Public Property Get RecepSerl() As Integer

    RecepSerl = FRecepSerl
End Property

Public Property Let Recpnumber(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpnumber = vData
        Else
            FRecpnumber = Null
        End If
    End If
    
End Property

Public Property Get Recpnumber() As String
    
    Recpnumber = FRecpnumber

End Property
Public Property Let Npecode(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FNpecode = vData
        Else
            FNpecode = Null
        End If
    End If
    
End Property


Public Property Get Npecode() As String
    Npecode = FNpecode
End Property

'call store procedure and set parameters, insert data to database

Public Function InsertReceptiondelt(cn As ADODB.Connection) As Boolean
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New Command
    
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "RECEPTIONDETLINSERT_SP"
        
        If .Parameters.Count = 0 Then
            .Parameters.Append .CreateParameter("@recd_recpnumb", adVarChar, adParamInput, 15)
            .Parameters.Append .CreateParameter("@recd_npecode", adVarChar, adParamInput, 5)
            .Parameters.Append .CreateParameter("@recd_primqtydlvd", adDecimal, adParamInput, 22)
            .Parameters.Append .CreateParameter("@recd_desc", adVarChar, adParamInput, 1500)
            .Parameters.Append .CreateParameter("@recd_unitpric", adCurrency, adParamInput, 22)
            .Parameters.Append .CreateParameter("@recd_liitnumb", adVarChar, adParamInput, 6)
            .Parameters.Append .CreateParameter("@recd_partnumb", adVarChar, adParamInput, 20)
        End If
        
        .Parameters("@recd_recpnumb") = FRecpnumber
        .Parameters("@recd_npecode") = FNpecode
        .Parameters("@recd_primqtydlvd") = FRecpriqtydelived
        .Parameters("@recd_desc") = FRepdescription
        .Parameters("@recd_unitpric") = FRecpUintprice
        .Parameters("@recd_liitnumb") = FRecplineitem
        .Parameters("@recd_partnumb") = FRecppartnumb
        
        .Execute , , adExecuteNoRecords
    End With
    
     InsertReceptiondelt = True
     Set cmd = Nothing
     Exit Function
     
Noinsert:
    InsertReceptiondelt = False
    Err.Raise Err.number, Err.Description, Err.HelpFile, Err.HelpContext

    
End Function
