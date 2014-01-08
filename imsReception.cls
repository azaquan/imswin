VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"RecepCol"
Option Explicit

'local variable(s) to hold property value(s)
Private FNpecode As String 'local copy
Private FRecpNumb As String 'local copy
Private FRecDate As Date 'local copy
Private FRecponumb As String 'local copy
'local variable(s) to hold property value(s)
Private FLinenumb As Integer 'local copy

Private FRecepCol As RecepCol


Public Property Get RecepCol() As RecepCol
    If FRecepCol Is Nothing Then
        Set FRecepCol = New RecepCol
    End If


    Set RecepCol = FRecepCol
End Property


Public Property Set RecepCol(vData As RecepCol)
    Set FRecepCol = vData
End Property
Private Sub Class_Terminate()
    Set FRecepCol = Nothing
End Sub



Public Property Let Linenumb(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Linenumb = 5
    FLinenumb = vData
End Property


Public Property Get Linenumb() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Linenumb
    Linenumb = FLinenumb
End Property



Public Property Let Recponumb(ByVal vData As String)
   If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecponumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get Recponumb() As String

    Recponumb = FRecponumb
End Property

'Public Property Let RecDate(ByVal vData As Date)
'    If Not IsNull(vData) Then
'
'        If Len(Trim$(vData)) Then
'            FRecDate = vData
'        Else
'            vData = Null
'        End If
'    End If
'
'End Property


'Public Property Get RecDate() As Date
'
'    RecDate = FRecDate
'End Property

Public Property Let RecpNumb(ByVal vData As String)
    If Not IsNull(vData) Then
        
        If Len(Trim$(vData)) Then
            FRecpNumb = vData
        Else
            vData = Null
        End If
    End If
    
End Property


Public Property Get RecpNumb() As String

    RecpNumb = FRecpNumb
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
'Public Property Let User(ByVal vData As String)
'    If Not IsNull(vData) Then
'
'        If Len(Trim$(vData)) Then
'            FUser = vData
'        Else
'            vData = Null
'        End If
'    End If
'
'End Property
'
'Public Property Get User() As String
'
'    User = FUser
'End Property

'call function to insert record to database

Public Function Update(Optional RecpNumb As String) As Boolean
On Error GoTo ErrHandler

Dim cmd As ADODB.Command

    Update = Save("UPDATEPOITENTOBE", RecpNumb)
    'MsgBox "Insert into Packing List Detail is completed"
    Exit Function
    
    
ErrHandler:
    MsgBox "Insert into Reception Detail is failure"
    Update = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    MsgBox "Insert into Reception Detail is failure"
End Function

'get store procedure parameters

Private Function Save(ProcName As String, Optional RecpNumb As String, Optional FRecponumb As String) As Boolean
On Error GoTo ErrHandler

Dim cmd As ADODB.Command

    Save = False
    Set cmd = New ADODB.Command
    
    
    With cmd
        .CommandText = ProcName
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = deIms.cnIms
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, FNpecode)
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, FRecpNumb)
        .Parameters.Append .CreateParameter("@RECPNUMB", adVarChar, adParamInput, 15, FRecponumb)
        .Parameters.Append .CreateParameter("@POLIITEM", adVarChar, adParamInput, 6, FLinenumb)
       
        .Execute , , adExecuteNoRecords
    
        Save = .Parameters(0).Value = 0
        
    End With
    
    
    Set cmd = Nothing
    
    Exit Function
    
    
ErrHandler:
    Save = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function


