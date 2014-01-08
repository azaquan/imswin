VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsPackingListRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private FRecipient As ImsRecipients
Private FManifestNumber As String

Public Property Let ManifestNumber(ByVal vData As String)
    FManifestNumber = vData
End Property

Public Property Get ManifestNumber() As String
Attribute ManifestNumber.VB_UserMemId = 0
    ManifestNumber = FManifestNumber
End Property

Public Property Let RecpNumb(ByVal vData As Integer)
    FRecipient.RecpNumb = vData
End Property

Public Property Get RecpNumb() As Integer
    RecpNumb = FRecipient.RecpNumb
End Property

Public Property Let Receipient(ByVal vData As String)
    FRecipient.Receipient = vData
End Property

Public Property Get Receipient() As String
    Receipient = FRecipient.Receipient
End Property
Public Property Let User(ByVal vData As String)
    FRecipient.User = vData
End Property

Public Property Get User() As String
    User = FRecipient.User
End Property

Public Property Let NameSpace(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NameSpace = 5
    FRecipient.NameSpace = vData
End Property

Public Property Get NameSpace() As String
    NameSpace = FRecipient.NameSpace
End Property

Private Sub Class_Initialize()
    Set FRecipient = New ImsRecipients
End Sub

Private Sub Class_Terminate()
    Set FRecipient = Nothing
End Sub

'Call store procedure to insert a record to database

Public Function Update(cn As ADODB.Connection) As Boolean
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    Call Validate
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "Upd_Ins_PACKINGRECIP"
    
        .Parameters.Append .CreateParameter("@manfnumb", adVarChar, adParamInput, 15, FManifestNumber)
        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, FRecipient.NameSpace)
        .Parameters.Append .CreateParameter("@recpmumb", adInteger, adParamInput, 4, FRecipient.RecpNumb)
        .Parameters.Append .CreateParameter("@rec", adVarChar, adParamInput, 60, FRecipient.Receipient)
        .Parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, FRecipient.User)
        Call .Execute(Options:=adExecuteNoRecords)
    End With
    
    Update = True
    Set cmd = Nothing
   ' MsgBox "Insert into Packing List Receipients is completed"
    Exit Function
    
Noinsert:
    Update = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

'validate data format

Public Function Validate() As Boolean

    Validate = False
            
    If Len(FManifestNumber) = 0 Then _
        Err.Raise 60000, "Packing List", LoadResString(60000)

    If Len(FRecipient.NameSpace) = 0 Then _
        Err.Raise 60007, "Packing List", LoadResString(60007)
    
    If FRecipient.RecpNumb = 0 Then _
        Err.Raise 60008, "Packing List", LoadResString(60008)
    
    If Len(FRecipient.Receipient) = 0 Then _
        Err.Raise 60009, "Packing List", LoadResString(60009)
        
End Function


