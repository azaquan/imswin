VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsPackinListDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class wrapper for packingist detail table"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PackingListDelt"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private FMainfestNumber As String
Private FNamespace As String
Private FManiFestSerialNumb As Integer
Private FPoNumber As String
Private FLineNumber As String
Private FBoxNumber As String
Private FRequestedQty As Double
Private FUnitPrice As Currency
Private FTotalPrice As Double
Private FUser As String
Private FConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private FTobeship As Double 'local copy
Public Property Let User(ByVal vData As String)

    FUser = vData
End Property

Public Property Get User() As String

    User = FUser
End Property

Public Property Let Tobeship(ByVal vData As Double)

    FTobeship = vData
End Property


Public Property Get Tobeship() As Double

    Tobeship = FTobeship
End Property




Public Property Set Connection(ByRef vData As Connection)
    Set FConnection = vData
End Property

Public Property Let TotalPrice(ByVal vData As Double)
        FTotalPrice = vData
End Property

Public Property Get TotalPrice() As Double
    TotalPrice = FTotalPrice
End Property

Public Property Let UnitPrice(ByVal vData As Currency)
        FUnitPrice = vData
End Property

Public Property Get UnitPrice() As Currency
    UnitPrice = FUnitPrice
End Property

Public Property Let RequestedQty(ByVal vData As Double)

        FRequestedQty = vData
End Property

Public Property Get RequestedQty() As Double
    RequestedQty = FRequestedQty
End Property

Public Property Let BoxNumber(ByVal vData As String)
        FBoxNumber = vData
End Property

Public Property Get BoxNumber() As String
    BoxNumber = FBoxNumber
End Property

Public Property Let LineNumber(ByVal vData As String)
        FLineNumber = vData
End Property

Public Property Get LineNumber() As String
    LineNumber = FLineNumber
End Property

Public Property Let PoNumber(ByVal vData As String)
        FPoNumber = vData
End Property

Public Property Get PoNumber() As String
        PoNumber = FPoNumber
End Property

Public Property Let ManiFestSerialNumb(ByVal vData As Integer)
        FManiFestSerialNumb = vData
End Property

Public Property Get ManiFestSerialNumb() As Integer
   
        ManiFestSerialNumb = FManiFestSerialNumb
End Property

Public Property Let NameSpace(ByVal vData As String)
        FNamespace = vData
End Property

Public Property Get NameSpace() As String
    NameSpace = FNamespace
End Property

Public Property Let MainfestNumber(ByVal vData As String)
    FMainfestNumber = vData
End Property

Public Property Get MainfestNumber() As String
    MainfestNumber = FMainfestNumber
End Property

'call function to insert data to database

Public Function Update(Optional sMainfestNumber As String) As Boolean
On Error GoTo ErrHandler

Dim cmd As ADODB.Command

    Update = Save("Upd_Ins_PACKINGDETAIL", sMainfestNumber)
    'MsgBox "Insert into Packing List Detail is completed"
    Exit Function
    
    
ErrHandler:
    MsgBox "Insert into PackingList Detail is failure"
    Update = False
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    'MsgBox "Insert into PackingList Detail is failure"
End Function

'set store procedure parameters and validate data format

Private Function Save(ProcName As String, Optional sMainfestNumber As String, Optional sRequestedQty As String) As Boolean
On Error GoTo ErrHandler

Dim cmd As ADODB.Command

    Save = False
    Set cmd = New ADODB.Command
    
   
   Call Validate
    
    If sMainfestNumber <> "" Then _
        FMainfestNumber = sMainfestNumber
'    If FRequestedQty <> "" Then _
'        FRequestedQty = sRequestedQty
        
    If ((FConnection.State And adStateOpen) = adStateClosed) Then FConnection.Open
    With cmd
        .CommandText = ProcName
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = FConnection
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@manfnumb", adVarChar, adParamInput, 15, FMainfestNumber)
        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, FNamespace)
        .Parameters.Append .CreateParameter("@manfsrl", adInteger, adParamInput, 4, FManiFestSerialNumb)
        .Parameters.Append .CreateParameter("@ponum", adVarChar, adParamInput, 15, FPoNumber)
        .Parameters.Append .CreateParameter("@liitnumb", adVarChar, adParamInput, 6, FLineNumber)
        .Parameters.Append .CreateParameter("@boxnumb", adInteger, adParamInput, 4, FBoxNumber)
        .Parameters.Append .CreateParameter("@reqdqty", adDouble, adParamInput, 10, FTobeship)
        .Parameters.Append .CreateParameter("@unitpric", adDouble, adParamInput, 10, FUnitPrice)
        .Parameters.Append .CreateParameter("@User", adVarChar, adParamInput, 20, FUser)
        
        Call .Execute(Options:=adExecuteNoRecords)
    
        Save = .Parameters(0).Value = 0
        
    End With
    
    
    Set cmd = Nothing
    
    Exit Function
    
    
ErrHandler:
    Save = False
    MsgBox "Insert into Packing list details was failure"
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

'call store procedure to update poitem status

Public Function GetPoitemUpdate(NameSpace As String, Ponumb As String, Linenumb As String, Manfnumb As String, _
                            Manser As Integer, cn As ADODB.Connection) As Boolean

On Error Resume Next


Dim cmd As ADODB.Command

Set cmd = New ADODB.Command

    cn.Errors.Clear
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = adCmdStoredProc
        .CommandText = "UPDATE_POITEM_PACKING"
        
        .Parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@Namespace", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, Ponumb)
        .Parameters.Append .CreateParameter("@LINENUMB", adVarChar, adParamInput, 6, Linenumb)
        .Parameters.Append .CreateParameter("@MANFNUMB", adVarChar, adParamInput, 15, Manfnumb)
        .Parameters.Append .CreateParameter("@MANFSRL", adInteger, adParamInput, 10, Manser)
        
        Call .Execute(Options:=adExecuteNoRecords)
        GetPoitemUpdate = .Parameters("RT") = 0
    End With
    
    Set cmd = Nothing
    
    If cn.Errors.Count Then
        Debug.Print "GetPoitemUpdate Error count = " & cn.Errors.Count
        Debug.Print "GetPoitemUpdate Error  = " & cn.Errors(0).Description
    End If
    
    If Err Then Err.Clear
End Function

'validate data format and show message

Public Function Validate() As Boolean
Dim Msg, Style, Title
Dim Str As Integer

'Msg = " Lineitem# " & FLineNumber & " is being over shipped, Doyou want to continue ?"
'Style = vbYesNo + vbCritical + vbDefaultButton2
'Title = "Imswin"

    Validate = False

    
    If Len(FMainfestNumber) = 0 Then _
        Err.Raise 60000, "Packing List", LoadResString(60000)

    If Len(FNamespace) = 0 Then _
        Err.Raise 60007, "Packing List", LoadResString(60007)
   
    If Len(FManiFestSerialNumb) = 0 Then _
        Err.Raise 60001, "Packing List", LoadResString(60001)
    
    If Len(FPoNumber) = 0 Then _
        Err.Raise 60002, "Packing List", LoadResString(60002)
        
    If Len(FLineNumber) = 0 Then _
        Err.Raise 60003, "Packing List", LoadResString(60003)
    
   ' If Len(FBoxNumber) Then _
    '    Err.Raise
    'End If
    
    If Len(FRequestedQty) = 0 Then _
        Err.Raise 60004, "packing List", LoadResString(60004)
    
    If Len(FUnitPrice) = 0 Then _
        Err.Raise 60006, "Packing List", LoadResString(60006)
   
    If Len(FTotalPrice) = 0 Then _
        Err.Raise 60005, "Packing List ", LoadResString(60005)
        
'    If Not Len(FTobeship) = 0 Then
'        If FRequestedQty > FTobeship Then
'            If MsgBox(Msg, Style, Title) = vbNo Then
'                Exit Function
'            End If
'        Else
'        End If
'    End If
        
    
End Function

