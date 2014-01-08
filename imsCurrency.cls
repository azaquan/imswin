VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Implements imsNav
Private fcode As String
Private uMode As TupdateMode
Private FNamespace As String
Private FInserting As Boolean
Private DataChanged As Boolean
Private FDescription As String
Private FConversionRatio As Variant
Dim WithEvents rs  As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Public Event OnMoveComplete()
Public Event OnError(Description As String)

'set conversion ratio data

Public Property Let ConversionRatio(ByVal NewValue As Variant)
    
    If Len(Trim$(NewValue & "")) = 0 Then NewValue = Null
    
    If FConversionRatio <> NewValue Then
        DataChanged = True
        FConversionRatio = NewValue
        rs!curr_convration = NewValue
    End If
End Property

Public Property Get ConversionRatio() As Variant
    ConversionRatio = FConversionRatio
End Property

Public Property Let namespace(ByVal NewValue As String)
    
    If NewValue <> FNamespace Then
        DataChanged = True
        rs!curr_npecode = FNamespace
        FNamespace = Trim$(NewValue)
    End If
    
End Property

Public Property Get namespace() As String
    namespace = FNamespace
End Property

'set description data

Public Property Let Description(ByVal NewValue As String)
    
    NewValue = Trim$(NewValue)
    If NewValue <> FDescription Then
        DataChanged = True
        rs!curr_desc = NewValue
        FDescription = Trim$(NewValue)
    End If
End Property

Public Property Get Description() As String
    Description = FDescription
End Property


Public Property Let Code(ByVal NewValue As String)
    NewValue = Trim$(NewValue)
    
    If NewValue <> fcode Then
        DataChanged = True
        fcode = Trim$(NewValue)
        rs!curr_code = NewValue
    End If
    
End Property

Public Property Get Code() As String
    Code = fcode
End Property

Private Property Get imsNav_AbsolutePosition() As Long
    imsNav_AbsolutePosition = rs.AbsolutePosition
End Property

Private Sub imsNav_AddNew()
On Error Resume Next
    
    rs.AddNew
    
    rs!curr_npecode = FNamespace
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

Private Property Get imsNav_Bor() As Boolean
On Error Resume Next

    imsNav_Bor = rs.AbsolutePosition = rs.RecordCount
    
    If Err Then Err.Clear
End Property

Private Sub imsNav_CancelUpdate()
On Error Resume Next

    rs.CancelUpdate
    If rs.editmode = adEditAdd Then rs.MoveFirst
    
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

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

'function get string data

Private Function imsNav_Find(Criteria As String) As Boolean
Dim BK As Variant
On Error Resume Next

    imsNav_Find = False
    
    BK = rs.Bookmark
    
    rs.CancelUpdate
    rs.CancelUpdate
    
    Criteria = Replace(Criteria, "Code", "curr_code", , , vbTextCompare)
    Call rs.Find(Criteria, 0, adSearchForward, adBookmarkFirst)
    Call rs.Find(Criteria, 0, adSearchForward, adBookmarkFirst)
    
    If rs.EOF Then
        rs.Bookmark = BK
        imsNav_Find = False
    Else
        imsNav_Find = True
    End If
End Function

'move record set first position

Private Sub imsNav_MoveFirst()
On Error Resume Next
    
    
    rs.MoveFirst
    If Err Then Err.Clear
End Sub

'move record set last position

Private Sub imsNav_MoveLast()
On Error Resume Next

    
    rs.MoveLast
    If Err Then Err.Clear
End Sub

'move record set to next position

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

Private Sub imsNav_Requery()
On Error Resume Next

    rs.Requery
    
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

'before save record set modify user to current user

Private Sub imsNav_Update()
On Error Resume Next
Dim df As Object
        
        
    rs!cur_modiuser = CurrentUser
    If Navigator.Validate Then rs.Update
    
    DoEvents: DoEvents: DoEvents
    If Err Then RaiseEvent OnError(Err.Description)
    If Err Then Err.Clear
End Sub

'set data to be save

Private Sub imsNav_UpdateBatch()
On Error Resume Next


    If False Then
        Dim ds As RDS.DataSpace, obj As Object
        
        Set ds = New RDS.DataSpace
        
        Set obj = ds.CreateObject("Server.Class1", "")
        
        Call obj.Update(rs)
        
        Set ds = Nothing
        Set obj = Nothing
    Else
    
        Call rs.UpdateBatch
        
        If Err Then RaiseEvent OnError(Err.Description)
        If Err Then Err.Clear
    End If
End Sub

Private Property Get imsNav_UpdateMode() As TupdateMode
    imsNav_UpdateMode = Abs(rs.LockType = adLockBatchOptimistic)
End Property

'=============================================================
'Name: Validate
'
'Date: 01/11/2000 11:39
'
'Description: This Function validates the RecordSet class data.
'
'Comment:
'=============================================================

'validate data format

Private Function imsNav_Validate() As Boolean
On Error GoTo Handler


    imsNav_Validate = False


    'If DataChanged = False Then rs.CancelUpdate
    If Len(Trim$(FConversionRatio & "")) = 0 Then FConversionRatio = Null
    
    If Len(Trim$(fcode)) = 0 Then
        Err.Raise 1000, "imsCurrency", "The field ' Code ' cannot be left empty"
        
    ElseIf Len(Trim$(FDescription)) = 0 Then
        Err.Raise 1000, "imsCurrency", "The field ' Description ' cannot be left empty"
    ElseIf Len(Trim$(fcode)) > 3 Then
        Err.Raise 1000, "imscurrency", "The field 'Code' cannot have more than 3 characters"
    
    End If

    rs!curr_code = fcode
    rs!curr_desc = FDescription
    rs!curr_npecode = FNamespace
    rs!curr_convration = FConversionRatio
    
    'Verify the integer field contains a valid value.
    If Not IsNull(rs("curr_convration")) Then
        If Not IsNumeric(rs("curr_convration")) _
            And InStr(rs("curr_convration"), ".") = 0 Then
            Err.Raise 1000, "imsCurrency", "The field ' ConversionRatio ' does not contain a valid number."
        Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curr_desc")) Then
        Err.Raise 1000, "imsCurrency", "The field ' Description ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("curr_desc")) Then
        If Len(Trim(rs("curr_desc"))) = 0 Then
            Err.Raise 1000, "imsCurrency", "The field ' Description ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curr_npecode")) Then
        Err.Raise 1000, "imsCurrency", "The field ' NameSpace ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("curr_npecode")) Then
        If Len(Trim(rs("curr_npecode"))) = 0 Then
            Err.Raise 1000, "imsCurrency", "The field ' NameSpace ' does not contain valid text."
            Exit Function
        End If
    End If

    'Verify the field is not null.
    If IsNull(rs("curr_code")) Then
        Err.Raise 1000, "imsCurrency", "The field ' Code ' cannot be null."
        Exit Function
    End If

    'Verify the text field contains text.
    If Not IsNull(rs("curr_code")) Then
        If Len(Trim(rs("curr_code"))) = 0 Then
            Err.Raise 1000, "imsCurrency", "The field ' Code ' does not contain valid text."
            Exit Function
        End If
    End If


    imsNav_Validate = True
    
Handler:

    If Err Then
    
        Dim str As String
        str = Err.Description
        
        Err.Clear
        RaiseEvent OnError(str)
    End If
End Function

'SQL statement get currency data

Public Sub GetValues(cn As ADODB.Connection, Optional SNamespace As String)
Dim SQL As String

    
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    SNamespace = IIf(SNamespace = "", FNamespace, SNamespace)
    
    SQL = "SELECT curr_code, curr_npecode,"
    SQL = SQL & " curr_desc, curr_convration,"
    SQL = SQL & " curr_modiuser , curr_creauser"
    SQL = SQL & " FROM CURRENCY WHERE curr_npecode = '" & SNamespace & "'"
    
    If True Then
        With rs
        
            Set .ActiveConnection = cn
            
            .Source = SQL
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            
            .Open
        End With
    Else
        Dim ds As RDS.DataSpace, obj As Object
        
        Set ds = New RDS.DataSpace
        
        Set obj = ds.CreateObject("RDSServer.DataFactory", "")
        Set rs = obj.Query("DSN=SAKHALIN;USER ID = sa", SQL)

'        dc.SQL = SQL
'        dc.Server = ""
'        dc.ExecuteOptions = adcExecSync
'        dc.FetchOptions = adcFetchBackground
'        dc.Connect = "DSN=SAKHALIN;USER ID = sa"
'
'        dc.Refresh
'        Do While dc.ReadyState = adcReadyStateLoaded
'            DoEvents
'        Loop
'
'
'        Set rs = dc.Recordset
'        Set rs.ActiveConnection = Nothing

        Set ds = Nothing
        Set obj = Nothing
    End If
 End Sub

'load data to form

Private Sub LoadValues()
On Error Resume Next

    If rs.EOF Or rs.BOF Then
        Call ClearField: Exit Sub
        
    ElseIf rs.RecordCount = 0 Then
        Call ClearField
        
    Else
        fcode = Trim$(rs!curr_code & "")
        FDescription = Trim$(rs!curr_desc & "")
        FNamespace = Trim$(rs!curr_npecode & "")
        FConversionRatio = Trim$(rs!curr_convration & "")
    End If
    
    DataChanged = False
    If Err Then Err.Clear
    RaiseEvent OnMoveComplete
End Sub

'cleart fields

Private Sub ClearField()
    fcode = ""
    FDescription = ""
    FConversionRatio = Null
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Call LoadValues
End Sub


Public Property Get Inserting()
    Inserting = rs.editmode = adEditAdd
End Property

'SQL statement get currency code list

Public Function CodeList(cn As ADODB.Connection, Optional SNamespace As String) As ADODB.Recordset
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdText
        Set .ActiveConnection = cn
        
        .CommandText = "SELECT curr_code code FROM CURRENCY"
        .CommandText = .CommandText & " WHERE ( curr_npecode = ? )"
        
        .Parameters.Append .CreateParameter("", adVarChar, adParamInput, 5, SNamespace)
        
        Set CodeList = .Execute
        
    End With
End Function

'Function Created by Muzammil.It is to know if the Code the User is entering is repeated or not.
Public Function IsCodeRepeated(namespace As String) As Boolean
Dim RsCodeCount As New ADODB.Recordset
RsCodeCount.Source = "select curr_code from CURRENCY where curr_code='" & fcode & "' and curr_npecode='sakha'"
RsCodeCount.ActiveConnection = deIms.cnIms
RsCodeCount.CursorType = adOpenForwardOnly
RsCodeCount.Open

If rs.editmode = adEditAdd Then
        If RsCodeCount.RecordCount > 0 Then
           RsCodeCount.Filter = adFilterNone
           IsCodeRepeated = True
        End If
 Else
       IsCodeRepeated = False
End If
End Function
