VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsTrackingMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private FNamespace As String 'local copy
Private FPonumb As String 'local copy
Private FMessageNumb As String 'local copy
Private FFlag As Boolean 'local copy
Private FMessageDate As Date 'local copy
Private FSubject As String 'local copy
Private FSupplierRe As String 'local copy
Private FForwarderRe As String 'local copy
Private FRecipient1 As Variant 'local copy
Private FRecipient2 As Variant 'local copy
Private FRecipient3 As Variant 'local copy
Private FRecipient4 As Variant 'local copy
Private FRecipient5 As Variant 'local copy
Private FNewDeliveryDate As Date 'local copy
Private FEstimateDate As Date 'local copy
Private FArrivalDate As Date 'local copy
Private FShipVia As String 'local copy
Private FIncludeMessage As Boolean 'local copy
Private FRemark As Variant 'local copy

Private FOperator As String 'local copy
'local variable(s) to hold property value(s)
Private FUser As String 'local copy
Public Property Let User(ByVal vData As String)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FUser = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get User() As String

    User = FUser
End Property



Public Property Let Operator(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FOperator = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get Operator() As String

    Operator = FOperator
End Property



Public Property Let Remark(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRemark = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Remark() As Variant

        Remark = FRemark
    
End Property

Public Property Let IncludeMessage(ByVal vData As Boolean)

    FIncludeMessage = vData
End Property

Public Property Get IncludeMessage() As Boolean

    IncludeMessage = FIncludeMessage
End Property

Public Property Let ShipVia(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FShipVia = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get ShipVia() As String

    ShipVia = FShipVia
End Property

Public Property Let ArrivalDate(ByVal vData As Date)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FArrivalDate = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get ArrivalDate() As Date

    ArrivalDate = FArrivalDate
End Property

Public Property Let EstimateDate(ByVal vData As Date)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FEstimateDate = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get EstimateDate() As Date

    EstimateDate = FEstimateDate
End Property

Public Property Let NewDeliveryDate(ByVal vData As Date)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FNewDeliveryDate = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get NewDeliveryDate() As Date

    NewDeliveryDate = FNewDeliveryDate
End Property

Public Property Let Recipient5(ByVal vData As Variant)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRecipient5 = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Recipient5() As Variant

        Recipient5 = FRecipient5
    
End Property

Public Property Let Recipient4(ByVal vData As Variant)
     If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRecipient4 = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Recipient4() As Variant

        Recipient4 = FRecipient4
    
End Property
Public Property Let Recipient3(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRecipient3 = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get Recipient3() As Variant

        Recipient3 = FRecipient3
    
End Property

Public Property Let Recipient2(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRecipient2 = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get Recipient2() As Variant

        Recipient2 = FRecipient2
    
End Property

Public Property Let Recipient1(ByVal vData As Variant)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FRecipient1 = vData
        Else
            vData = Null
        End If
    End If
End Property

Public Property Get Recipient1() As Variant

        Recipient1 = FRecipient1
    
End Property

Public Property Let ForwarderRe(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FForwarderRe = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get ForwarderRe() As String

    ForwarderRe = FForwarderRe
End Property

Public Property Let SupplierRe(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FSupplierRe = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get SupplierRe() As String

    SupplierRe = FSupplierRe
End Property

Public Property Let Subject(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FSubject = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Subject() As String

    Subject = FSubject
End Property

Public Property Let MessageDate(ByVal vData As Date)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FMessageDate = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get MessageDate() As Date

    MessageDate = FMessageDate
End Property

Public Property Let Flag(ByVal vData As Boolean)

    FFlag = vData
End Property

Public Property Get Flag() As Boolean

    Flag = FFlag
End Property

Public Property Let MessageNumb(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FMessageNumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get MessageNumb() As String

    MessageNumb = FMessageNumb
End Property

Public Property Let Ponumb(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FPonumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Ponumb() As String

    Ponumb = FPonumb
End Property

Public Property Let Namespace(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
             FNamespace = vData
        Else
            vData = Null
        End If
    End If
    
End Property

Public Property Get Namespace() As String

    Namespace = FNamespace
End Property

'SQL statement to get po number

Public Function GetOBSPonumber(Namespace As String, cn As ADODB.Connection) As Recordset
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "SELECT ob_ponumb FROM OBS "
        .CommandText = .CommandText & " where ob_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " ORDER BY ob_ponumb"
        
         Set GetOBSPonumber = .Execute
    End With

End Function

'SQL statement to get message number list

Public Function GetMessageNumberlist(Ponumb As String, Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim rst As Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "SELECT ob_mesgnumb FROM OBS "
        .CommandText = .CommandText & " where ob_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " and ob_ponumb = '" & Ponumb & "'"
        .CommandText = .CommandText & " ORDER BY ob_mesgnumb"
        
         Set GetMessageNumberlist = .Execute
    End With
    
'    If Rst.RecordCount = 0 Then GoTo CleanUp
'
'    Rst.MoveFirst
'
'    Do While ((Not Rst.EOF))
'        cmbMessage.AddItem Rst!ob_mesgnumb & ""
'
'        Rst.MoveNext
'    Loop
'
'CleanUp:
'    Rst.Close
'    Set cmd = Nothing
'    Set Rst = Nothing

End Function

'SQL statement to get OBS list

Public Function GetOBSList(Ponumb As String, Message As String, Namespace As String, cn As ADODB.Connection) As imsTrackingMessage
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "SELECT ob_ponumb, ob_mesgnumb,  ob_flag, ob_mesgdate, "
        .CommandText = .CommandText & " ob_subj, ob_suppreci, ob_forwreci, ob_rec1, "
        .CommandText = .CommandText & " ob_rec3, ob_rec2,ob_rec4, ob_rec5, ob_oper, "
        .CommandText = .CommandText & " ob_newdelvdate, ob_etd, ob_eta, ob_shipvia, "
        .CommandText = .CommandText & " ob_inclmesg, ob_remk FROM OBS "
        .CommandText = .CommandText & " where ob_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " and ob_ponumb = '" & Ponumb & "'"
        .CommandText = .CommandText & " and ob_mesgnumb = '" & Message & "'"
        .CommandText = .CommandText & " ORDER BY ob_mesgnumb"
        
         Set rst = .Execute
    End With
    
    FArrivalDate = rst!ob_eta & ""
    FEstimateDate = rst!ob_etd & ""
    FForwarderRe = rst!ob_forwreci & ""
    FFlag = IIf((rst!ob_flag), 1, 0)
    FPonumb = rst!ob_ponumb & ""
    FMessageNumb = rst!ob_mesgnumb & ""
    FMessageDate = rst!ob_mesgdate & ""
    FSubject = rst!ob_subj & ""
    FSupplierRe = rst!ob_suppreci & ""
    FRecipient1 = rst!ob_rec1 & ""
    FRecipient2 = rst!ob_rec2 & ""
    FRecipient3 = rst!ob_rec3 & ""
    FRecipient4 = rst!ob_rec4 & ""
    FRecipient5 = rst!ob_rec5 & ""
    FNewDeliveryDate = rst!ob_newdelvdate & ""
    FShipVia = rst!ob_shipvia & ""
    FRemark = rst!ob_remk & ""
    FIncludeMessage = IIf((rst!ob_inclmesg), 1, 0)
    FOperator = rst!ob_oper & ""
    
    Set GetOBSList = Me
    
End Function

'call store procedure and set parameters, insert record to database

Public Function InsertandUpdateTable(PO As String, Namespace As String, Message As String, _
                                     Subject As Variant, Flag As Boolean, Medate As Variant, _
                                     Suppreci As Variant, Forreci As Variant, Rec1 As Variant, _
                                     Rec2 As Variant, Rec3 As Variant, Rec4 As Variant, Rec5 As Variant, _
                                     Oper As Variant, NewDe As Variant, Etd As Variant, Atd As Variant, _
                                     Ship As Variant, Inflag As Boolean, Remark As Variant, User As String, _
                                     cn As ADODB.Connection) As Boolean
On Error GoTo Noinsert
Dim cmd As ADODB.Command
'Dim cn As ADODB.Connection
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATE_INSERT_OBS"
    
        .Parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, PO)
        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("@mesgnumb", adVarChar, adParamInput, 12, Message)
        .Parameters.Append .CreateParameter("@subj", adVarChar, adParamInput, 60, Subject)
        .Parameters.Append .CreateParameter("@flag", adBoolean, adParamInput, 1, Flag)
        .Parameters.Append .CreateParameter("@mesgdate", adVarChar, adParamInput, 15, Medate)
        .Parameters.Append .CreateParameter("@suppreci", adVarChar, adParamInput, 60, Suppreci)
        .Parameters.Append .CreateParameter("@forwreci", adVarChar, adParamInput, 60, Forreci)
        .Parameters.Append .CreateParameter("@rec1", adVarChar, adParamInput, 60, Rec1)
        .Parameters.Append .CreateParameter("@rec2", adVarChar, adParamInput, 60, Rec2)
        .Parameters.Append .CreateParameter("@rec3", adVarChar, adParamInput, 60, Rec3)
        .Parameters.Append .CreateParameter("@rec4", adVarChar, adParamInput, 60, Rec4)
        .Parameters.Append .CreateParameter("@rec5", adVarChar, adParamInput, 60, Rec5)
        .Parameters.Append .CreateParameter("@oper", adVarChar, adParamInput, 30, Oper)
        .Parameters.Append .CreateParameter("@newdelvdate", adVarChar, adParamInput, 12, NewDe)
        .Parameters.Append .CreateParameter("@etd", adVarChar, adParamInput, 12, Etd)
        .Parameters.Append .CreateParameter("@eta", adVarChar, adParamInput, 12, Atd)
        .Parameters.Append .CreateParameter("@shipvia", adVarChar, adParamInput, 2, Ship)
        .Parameters.Append .CreateParameter("@inclmesg", adBoolean, adParamInput, 1, Inflag)
        .Parameters.Append .CreateParameter("@REMARKS", adVarChar, adParamInput, 1000, Remark)
        .Parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, User)
        
        
        .Execute , , adExecuteNoRecords
    End With
    
    InsertandUpdateTable = True
    Set cmd = Nothing
    MsgBox "Record saved successfully."
    Exit Function
    
Noinsert:
    InsertandUpdateTable = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
