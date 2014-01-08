VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Ponumb As String
'Private m_POlineitems As ImsPOLI
'Private m_PONotes As imsPONotes
'Private m_Remarks As imsPORemarks
Private m_namespace As String
Private m_DB As Database

Public Property Get Ponumb() As String
Ponumb = m_Ponumb
End Property

Public Property Let Ponumb(ByVal vNewValue As String)
 m_Ponumb = Ponumb
End Property

Public Function getLINEitems() As ADODB.Recordset
Dim Select_query As String
Dim WHERE As String
Dim Orderby As String
Set m_DB = New Database

Select_query = "SELECT poi_ponumb,poi_liitnumb,poi_comm, poi_desc,poi_remk From poitem"
WHERE = "where poi_npecode = '" & m_namespace & "' "
Orderby = " poi_ponumb, cast(poi_liitnumb as int)"

Set getLINEitems = m_DB.RunQuery(Select_query, WHERE, Orderby)


'Set Getlineitem =
End Function

Public Function GetClause() As ADODB.Recordset
Dim Select_query As String
Dim WHERE As String
Dim Orderby As String
Set m_DB = New Database

Select_query = "SELECT poc_ponumb, poc_linenumb, poc_clau From POCLAUSE"
WHERE = "poc_npecode = '" & m_namespace & "' AND len(ltrim(rtrim( poc_clau)))>0 "
Orderby = "poc_ponumb,cast( poc_linenumb as int) "

Set GetClause = m_DB.RunQuery(Select_query, WHERE, Orderby)

End Function

Public Function GetRemarks() As ADODB.Recordset

Dim Select_query As String
Dim WHERE As String
Dim Orderby As String
Set m_DB = New Database

Select_query = "SELECT por_ponumb, por_linenumb, por_remk From POREM"
WHERE = "por_npecode = '" & m_namespace & "' AND len(ltrim(rtrim( POR_REMK)))>0 "
Orderby = "por_ponumb, cast(por_linenumb as int)"

Set GetRemarks = m_DB.RunQuery(Select_query, WHERE, Orderby)
End Function

Public Property Get NameSpace() As String
NameSpace = m_namespace
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
m_namespace = vNewValue
End Property

Public Function GetAllPOnumb(Optional TableName As String) As ADODB.Recordset
Dim Select_query As String
Dim WHERE As String
Set m_DB = New Database
TableName = Trim$(UCase(TableName))
Select Case TableName
       Case "POCLAUSE"
                  Select_query = "SELECT DISTINCT poc_ponumb FROM POCLAUSE "
                  WHERE = "poc_npecode='" & m_namespace & "' AND len(ltrim(rtrim( poc_clau)))>0 ORDER BY POC_PONUMB"
       Case "POREM"
                  Select_query = "SELECT DISTINCT por_ponumb FROM POREM "
                  WHERE = "POR_NPECODE='" & m_namespace & "' AND len(ltrim(rtrim( POR_REMK)))>0 ORDER BY POR_PONUMB"
       Case Else
            Select_query = "SELECT po_ponumb From PO"
            WHERE = "po_npecode = '" & m_namespace & "'  ORDER BY PO_PONUMB"
       
End Select

'''Select_query = "SELECT po_ponumb From PO"
'''WHERE = "po_npecode = '" & m_namespace & "'  "
    Set GetAllPOnumb = m_DB.RunQuery(Select_query, WHERE)

End Function

Public Function GetStocKsFromStockmaster() As ADODB.Recordset
Dim Select_query As String
Dim WHERE As String
Set m_DB = New Database
 
 Select_query = "SELECT stk_stcknumb,stk_desc From  stockmaster"
 WHERE = "stk_npecode = '" & m_namespace & "'  "
 
 Set GetStocKsFromStockmaster = m_DB.RunQuery(Select_query, WHERE)
End Function
