VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIsItemLoaded As Boolean

Private mNamespace As String
Private mStocknumber As String
Private mRsManufacturer As ADODB.Recordset

Private mConnection As ADODB.Connection

Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteALLPOITEM(cancel As Boolean)
Public Event AfterDeletePOITEMS()

Public Function AddNew() As Boolean
Dim cancel As Boolean

AddNew = False
  
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
  
  If cancel = False Then
           
           mRsManufacturer.AddNew
           
           mRsManufacturer!stm_npecode = mNamespace
           
   End If
   
  cancel = True
  
  AddNew = True
  
  Exit Function

handler:
     

  Err.Clear
End Function

Public Function DeleteCurrentLI() As Boolean
''Dim varBookmark  As Variant
''On Error GoTo handler
''DeleteCurrentLI = False
''
''
''If Not mRsManufacturer.AbsolutePosition = adPosEOF Then
''
''     mRsManufacturer.Delete adAffectCurrent
''
''  If mRsManufacturer.RecordCount > 0 Then
''            mRsManufacturer.MoveNext
''
''            If mRsManufacturer.AbsolutePosition = adPosEOF Then
''               mRsManufacturer.MoveFirst
''            End If
''
''            varBookmark = mRsManufacturer.Bookmark
''            mRsManufacturer.MoveFirst
''
''            Do While Not mRsManufacturer.Eof
''               mRsManufacturer!poi_liitnumb = mRsManufacturer.AbsolutePosition
''               mRsManufacturer.MoveNext
''            Loop
''            mRsManufacturer.Bookmark = varBookmark
''
''
''  End If
''
''End If
''DeleteCurrentLI = True
''Exit Function
''handler:
''   MsgBox "Error while deleting the line item." & Err.Description
''   Err.Clear
   
End Function



Public Function Update() As String

 On Error GoTo handler
 
 Update = "UnIndetified Errors."
      
   If mRsManufacturer Is Nothing Then Update = "": Exit Function
   
   If mRsManufacturer.RecordCount = 0 Then Update = "": Exit Function
   
   mRsManufacturer.UpdateBatch
     
   If mRsManufacturer.Eof = False Then
     
   End If
 
 Update = ""
     
 Exit Function
     
handler:
    
 Update = Err.Description
    
 Err.Clear

End Function

Public Function MoveToStocknumber(StockNumber As String) As Integer

'Returns 0 for Success
'Returns 1 for Unidentified Error
'Returns 2 for No Records

On Error GoTo handler

MoveToStocknumber = 1

'If Trim(UCase(mStocknumber)) <> Trim(UCase(StockNumber)) Then

    mStocknumber = Trim(UCase(StockNumber))

    If LoadCollection = False Then Exit Function

    If mRsManufacturer.state = 0 Then MoveToStocknumber = 2: Exit Function
    
    'If mRsManufacturer.State = 1 And mRsManufacturer.RecordCount = 0 Then MoveToStocknumber = 2: Exit Function

    If mRsManufacturer.RecordCount > 0 Then mRsManufacturer.MoveFirst
            
'End If
   
   MoveToStocknumber = 0
   Exit Function
   
handler:

   Err.Clear
End Function

Public Property Get Count() As Integer

On Error GoTo handler
Count = 0
Count = mRsManufacturer.RecordCount

Exit Property
handler:

Err.Clear

End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler
  If Not mRsManufacturer.Bof Then
   
     mRsManufacturer.MoveFirst
     If mRsManufacturer.Bof = True Then
       MoveFirst = False
     Else
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsManufacturer.Eof Then
   
     mRsManufacturer.MoveNext
     
     If mRsManufacturer.AbsolutePosition = adPosEOF Then mRsManufacturer.MoveLast
     If Not mRsManufacturer.Eof Then
        MoveNext = True
     End If
  End If
  Exit Function
handler:
  
End Function

Public Function MovePrevious() As Boolean
MovePrevious = False
On Error GoTo handler
  If Not mRsManufacturer.Bof Then
   
     mRsManufacturer.MovePrevious
     If mRsManufacturer.AbsolutePosition = adPosBOF And mRsManufacturer.Eof <> True And mRsManufacturer.Bof = True Then mRsManufacturer.MoveFirst
     If mRsManufacturer.Bof Then
      MovePrevious = False
     Else
      MovePrevious = True
     End If
  End If
 Exit Function
handler:
  
End Function

Public Function MoveLast() As Boolean

MoveLast = False
On Error GoTo handler
  If Not mRsManufacturer.Eof Then
   
     mRsManufacturer.MoveLast
     If mRsManufacturer.Eof = True Then
       MoveLast = False
     Else
       MoveLast = True
     End If
  End If
  Exit Function
handler:
  Err.Clear
End Function



Public Property Get Eof() As Boolean

  Eof = mRsManufacturer.Eof

End Property

Public Property Get Bof() As Boolean

Bof = mRsManufacturer.Bof

End Property


Public Function CancelUpdate() As Boolean
On Error GoTo handler
If mRsManufacturer.AbsolutePosition = adPosBOF And mRsManufacturer.Eof <> True And mRsManufacturer.Bof = True Then mRsManufacturer.MoveFirst
If mRsManufacturer.AbsolutePosition = adPosEOF And mRsManufacturer.Bof <> True And mRsManufacturer.Eof = True Then mRsManufacturer.MoveLast
If mRsManufacturer.AbsolutePosition = adPosUnknown And mRsManufacturer.RecordCount > 0 Then mRsManufacturer.MoveLast
If mRsManufacturer.EditMode = adEditAdd Then
   mRsManufacturer.CancelUpdate
End If

If mRsManufacturer.AbsolutePosition = adPosBOF And mRsManufacturer.Eof <> True And mRsManufacturer.Bof = True Then mRsManufacturer.MoveFirst
If mRsManufacturer.AbsolutePosition = adPosEOF And mRsManufacturer.Bof <> True And mRsManufacturer.Eof = True Then mRsManufacturer.MoveLast
If mRsManufacturer.AbsolutePosition = adPosUnknown And mRsManufacturer.RecordCount > 0 Then mRsManufacturer.MoveLast

Exit Function
   
handler:
   Err.Clear
End Function




Public Property Get EditMode() As String

'If mRsManufacturer.State = 1 And mRsManufacturer.RecordCount > 0 Then

  EditMode = mRsManufacturer.EditMode
  
'Else

'   EditMode = "-1"
   
'End If

End Property


Public Function Requery()
On Error GoTo handler
   mRsManufacturer.Requery
   Exit Function
   
handler:

  MsgBox "Error When Poitem tried to Requery Database"
  Err.Clear

End Function

Public Function AnyChangesBeforeSaving() As Boolean
On Error GoTo handler

AnyChangesBeforeSaving = False

AnyChangesBeforeSaving = True
Exit Function
handler:

   Err.Clear
End Function

Public Property Get ManufactCode() As String
ManufactCode = mRsManufacturer!stm_manucode
End Property

Public Property Let ManufactCode(ByVal vNewValue As String)
mRsManufacturer!stm_manucode = vNewValue
End Property

Public Property Get NameSpace() As String
NameSpace = mNamespace
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
'mRsManufacturer!stm_npecode = vNewValue
mNamespace = vNewValue

End Property

Public Property Get StockNumber() As String
StockNumber = mStocknumber
End Property

Public Property Let StockNumber(ByVal vNewValue As String)
mRsManufacturer!stm_stcknumb = vNewValue
mStocknumber = mRsManufacturer!stm_stcknumb
End Property

Public Property Get PartNumb() As String
PartNumb = mRsManufacturer!stm_partnumb & ""
End Property

Public Property Let PartNumb(ByVal vNewValue As String)


If Len(Trim(vNewValue & "")) = 0 Then

    mRsManufacturer!stm_partnumb = Null

Else

    mRsManufacturer!stm_partnumb = (vNewValue)

End If


'mRsManufacturer!stm_partnumb = vNewValue
End Property

Public Property Get Flag() As Boolean
Flag = mRsManufacturer!stm_partnumb
End Property

Public Property Let Flag(ByVal vNewValue As Boolean)
mRsManufacturer!stm_partnumb = vNewValue
End Property

Public Property Get Techspec() As String
Techspec = mRsManufacturer!stm_techspec & ""
End Property

Public Property Let Techspec(ByVal vNewValue As String)
mRsManufacturer!stm_techspec = vNewValue
End Property

Public Property Get Estmpric() As String
Estmpric = mRsManufacturer!stm_estmpric & ""
End Property

Public Property Let Estmpric(ByVal vNewValue As String)

If Len(Trim(vNewValue & "")) = 0 Then

    mRsManufacturer!stm_estmpric = Null

Else

    mRsManufacturer!stm_estmpric = CDbl(vNewValue)

End If

'mRsManufacturer!stm_estmpric = IIf(Len(Trim(vNewValue)) = 0, Null, vNewValue)
End Property

Private Function LoadCollection() As Boolean
Dim CmdHead As ADODB.Command

LoadCollection = False
On Error GoTo handler

Set CmdHead = New ADODB.Command

Set mRsManufacturer = Nothing

   With CmdHead
   
       .CommandType = adCmdStoredProc
       .CommandText = "GetStockManufacturer"
    
       Set .ActiveConnection = mConnection
       
       .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, mNamespace)
       .Parameters.Append .CreateParameter("@STOCKNUMBER", adVarChar, adParamInput, 15, mStocknumber)
       
   End With
    
   Set mRsManufacturer = CmdHead.Execute
       
       mRsManufacturer.Close
       mRsManufacturer.Open , , adOpenKeyset, adLockBatchOptimistic
    
   LoadCollection = True
   
   Exit Function
   
handler:
      
   Err.Clear
   
End Function

Public Property Get Connection() As ADODB.Connection
Set Connection = mConnection
End Property

Public Property Let Connection(ByVal vNewValue As ADODB.Connection)
Set mConnection = vNewValue
End Property

Public Property Get AbsolutePosition() As Integer

If mRsManufacturer.state = 1 Then

    AbsolutePosition = mRsManufacturer.AbsolutePosition
    
Else

    AbsolutePosition = 0
    
End If

End Property

Public Property Get state() As Integer

state = mRsManufacturer.state

End Property
