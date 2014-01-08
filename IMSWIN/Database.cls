VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function RunQuery(Select_query As String, Optional Where As String, Optional Orderby As String) As ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim SQLstring As String
On Error GoTo LogErr

SQLstring = Select_query

Where = Replace(Where, "where", "")
Orderby = Replace(Orderby, "orderby", "")

If Len(Where) > 0 Then SQLstring = SQLstring & " where " & Where
If Len(Orderby) > 0 Then SQLstring = SQLstring & " order by " & Orderby

Rs.Open SQLstring, deIms.cnIms, adOpenKeyset

Set RunQuery = Rs

Set Rs.ActiveConnection = Nothing

Exit Function

LogErr:
   
   MsgBox Err.Description
   Call LogErr("Database::RunQuery", Err.Description, Err)
   Err.Clear
End Function
