VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsCurrencyDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"imsCurrency"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements imsNav

Private FID As Integer
Private FValue As String
Private FEndDate As String
Private FStartdate As String
Private uMode As TupdateMode
Private FimsCurrency As imsCurrency

Dim DataChanged As Boolean
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Public Event OnMoveComplete()
Public Event OnError(Description As String)

'check data format then get values

Public Property Let Value(ByVal vData As String)
On Error Resume Next

    If rs.RecordCount = 0 Then Exit Property
    If Not IsNumeric(vData) Then Exit Property
    If ((FValue <> vData) And (CDbl(vData)) > 0) Then
        FValue = vData
        DataChanged = True
        rs!CURD_VALUE = vData
    End If
    
    If Err Then
        RaiseEvent OnError(Err.Description)
        Err.Clear
    End If
    
End Property

Public Property Get Value() As String
    Value = FValue
End Property

'check data format then get values

Public Property Let EndDate(ByVal vData As String)
On Error Resume Next

    If rs.RecordCount = 0 Then Exit Property
    If Not IsDate(vData) Then
        RaiseEvent OnError("End Date is not a valid date")
        Exit Property
    End If
    
    If ((FEndDate <> vData)) Then
        FEndDate = vData
        rs!curd_to = vData
        DataChanged = True
    End If
End Property

Public Property Get EndDate() As String
    EndDate = FEndDate
End Property

'check data format then get values

Public Property Let Startdate(ByVal vData As String)
On Error Resume Next

    If rs.RecordCount = 0 Then Exit Property
    If Not IsDate(vData) Then
        RaiseEvent OnError("Start Date is not a valid date")
        Exit Property
    End If
    
    If ((FStartdate <> vData)) Then
        DataChanged = True
        FStartdate = vData
        rs!curd_from = vData
    End If
    
End Property

' get data values

Public Property Get Startdate() As String
    Startdate = FStartdate
End Property

' get data values

Private Sub Class_Initialize()
    Set FimsCurrency = New imsCurrency
End Sub

' get data values

Public Property Get imsCurrency() As imsCurrency
    Set imsCurrency = FimsCurrency
End Property

' get data values

Public Property Set imsCurrency(vData As imsCurrency)
    Set FimsCurrency = vData
End Property

'free memory

Private Sub Class_Terminate()
    Set FimsCurrency = Nothing
End Sub

Private Property Get imsNav_AbsolutePosition() As Long
    imsNav_AbsolutePosition = rs.AbsolutePosition
End Property

'check error cause give message

Private Sub imsNav_AddNew()
On Error Resume Next
    
    rs.AddNew
    If Err Then RaiseEvent OnError(Err.Description)
    
    rs!curd_id = IIf(rs.RecordCount, rs.RecordCount, 1)
    If Err Then Err.Clear
End Sub

Private Property Get imsNav_Bor() As Boolean
On Error Resume Next

    imsNav_Bor = rs.AbsolutePosition = rs.RecordCount
    
    If Err Then Err.Clear
End Property

'check error cause give message

Private Sub imsNav_CancelUpdate()
On Error Resume Next

    rs.CancelUpdate
    If Err Then RaiseEvent OnError(Err.Description)
    If rs.editmode = adEditAdd Then rs.MoveFirst
    
    If Err Then Err.Clear
End Sub

'set edit values

Private Property Get imsNav_Editting() As Boolean
On Error Resume Next
Dim i As Long
    
    i = rs.editmode
    imsNav_Editting = i <> adEditNone
    If Err Then Err.Clear
End Property

Private Property Get imsNav_Eor() As Boolean
    imsNav_Eor = rs.AbsolutePosition = 1
End Property

Private Function imsNav_Find(Criteria As String) As Boolean
    imsNav_Find = False
End Function

'move record to first position

Private Sub imsNav_MoveFirst()
On Error Resume Next
    
    
    rs.MoveFirst
    If Err Then Err.Clear
End Sub

'move record to last position

Private Sub imsNav_MoveLast()
On Error Resume Next

    rs.MoveLast
    If Err Then Err.Clear
End Sub

'move record to next position

Private Sub imsNav_MoveNext()
On Error Resume Next

   
    If Not rs.EOF Then rs.MoveNext
    
    If rs.EOF And rs.RecordCount <> 0 Then _
        If rs.EOF Then rs.MoveLast
        
    If Err Then Err.Clear
End Sub

'move record to previous position

Private Sub imsNav_MovePrevious()
On Error Resume Next

    If Not rs.BOF Then rs.MovePrevious
    
    If rs.BOF And rs.RecordCount <> 0 Then _
        If rs.BOF Then rs.MoveFirst
        
    If Err Then Err.Clear
End Sub

Public Property Get Navigator() As imsNav
    Set Navigator = Me
End Property

'if error cause give description

Private Sub imsNav_Requery()
On Error Resume Next

    rs.Requery
    
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

'set record set update, if error cause give error description

Private Sub imsNav_Update()
On Error Resume Next

    'If Navigator.Validate Then _
        If uMode = umImediate Then rs.Update
    
    rs.Update
    rs.UpdateBatch
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

'update all record sets

Private Sub imsNav_UpdateBatch()
On Error Resume Next

    Call rs.Update
    FimsCurrency.Navigator.Update
    
    Call rs.UpdateBatch
    
    Call rs.Move(0)
    DoEvents: DoEvents: DoEvents: DoEvents
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

Private Property Get imsNav_UpdateMode() As TupdateMode
    imsNav_UpdateMode = Abs(rs.LockType = adLockBatchOptimistic)
End Property

'validate data format, if error cause give messsage

Private Function imsNav_Validate() As Boolean


    imsNav_Validate = False
    'Verify the field is not null.
    
    If Not FimsCurrency Is Nothing Then
      
        
        rs!curd_code = FimsCurrency.Code
        rs!CURD_NPECODE = FimsCurrency.namespace
    End If
        
    rs!curd_to = FEndDate
    rs!curd_from = FStartdate
    
    If Err Then Err.Clear
    
    If IsNull(rs("curd_npecode")) Then
        RaiseEvent OnError("The field NameSpace cannot be null.")
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("curd_npecode")) Then
        If Len(Trim(rs("curd_npecode"))) = 0 Then
            RaiseEvent OnError("The field NameSpace does not contain valid text.")
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curd_code")) Then
        RaiseEvent OnError("The field Code cannot be null.")
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("curd_code")) Then
        If Len(Trim(rs("curd_code"))) = 0 Then
            RaiseEvent OnError("The field Code does not contain valid text.")
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curd_from")) Then
        RaiseEvent OnError("The field Start Date cannot be null.")
        Exit Function
    End If

    'Verify the date field contains a valid date.
    If Not IsNull(rs("curd_from")) Then
        If Not IsDate(rs("curd_from")) Then
            RaiseEvent OnError("The field Start Date does not contain a valid date.")
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curd_to")) Then
        RaiseEvent OnError("The field End Date cannot be null.")
        Exit Function
    End If

    'Verify the date field contains a valid date.
    If Not IsNull(rs("curd_to")) Then
        If Not IsDate(rs("curd_to")) Then
            RaiseEvent OnError("The field End Date does not contain a valid date.")
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curd_value")) Then
        RaiseEvent OnError("The field ' Value ' cannot be null.")
        Exit Function
    End If


    imsNav_Validate = True

End Function

'SQL statement get currency information

Public Sub GetValues(cn As ADODB.Connection, Optional SNamespace As String, Optional sCode As String)
On Error Resume Next

    If ((Not (rs Is Nothing)) And (rs.State And adStateOpen)) Then _
        rs.Update: rs.UpdateBatch
        
    rs.CancelBatch
    If Err Then Err.Clear
    
    Set rs = Nothing
    
    
    Set rs = New ADODB.Recordset
    
    Set rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    
    rs.Source = "SELECT curd_id, curd_npecode, "
    rs.Source = rs.Source & "curd_code, curd_from,"
    rs.Source = rs.Source & " curd_to, curd_value FROM "
    rs.Source = rs.Source & " CURRENCYDETL where ( "
    rs.Source = rs.Source & " ( curd_code = '" & sCode & "' ) "
    rs.Source = rs.Source & " AND ( curd_npecode = '" & SNamespace & "' ) )"
    rs.Source = rs.Source & " ORDER BY curd_from"
    
    rs.Open
    LoadValues
End Sub

'clear fields

Private Sub ClearField()
    FID = 0
    FValue = 0
    FEndDate = ""
    FStartdate = ""
End Sub

'load data from currency recordset

Private Sub LoadValues()
On Error Resume Next
Dim i As EditModeEnum

    i = rs.editmode
    If Err Then Err.Clear
    
    If rs.EOF Or rs.BOF Then
        Call ClearField: GoTo RaiseEvents
        
    ElseIf rs.RecordCount = 0 Then
        Call ClearField
        
    Else
        FID = Trim$(rs!curd_id & "")
        FValue = Trim$(rs!CURD_VALUE & "")
        FEndDate = IIf(i = adEditAdd, Date, Trim$(rs!curd_to & ""))
        FStartdate = IIf(i = adEditAdd, Date, Trim$(rs!curd_from & ""))
    End If
    
    If Err Then Err.Clear
    
RaiseEvents:
    DataChanged = False
    RaiseEvent OnMoveComplete
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    LoadValues
End Sub

Public Property Get Inserting()
    Inserting = rs.editmode = adEditAdd
End Property

