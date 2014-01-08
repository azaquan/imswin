VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImsXevents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim RsEvents As ADODB.Recordset
Private m_Cn As ADODB.Connection
Private m_TranNo As String
Private mUSERID As String
Private mnamespace As String


Public Property Get NameSpace() As String
NameSpace = RsEvents!eve_npecode
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
RsEvents!eve_npecode = vNewValue
mnamespace = Trim$(vNewValue)
End Property

Public Property Get MyLoginId() As String
MyLoginId = RsEvents!eve_from
End Property

Public Property Let MyLoginId(ByVal vNewValue As String)

RsEvents!eve_from = vNewValue

End Property
Public Property Get HisLoginId() As String
HisLoginId = RsEvents!eve_userid
End Property

Public Property Let HisLoginId(ByVal vNewValue As String)
mUSERID = vNewValue
RsEvents!eve_userid = vNewValue
End Property


Public Property Get OldVAlue() As Variant
OldVAlue = RsEvents!eve_oldval
End Property

Public Property Let OldVAlue(ByVal vNewValue As Variant)
RsEvents!eve_oldval = vNewValue
End Property

Public Property Get NewVAlue() As Variant
NewVAlue = RsEvents!eve_newval
End Property

Public Property Let NewVAlue(ByVal vNewValue As Variant)
RsEvents!eve_newval = vNewValue
End Property
Public Property Get EventDetail() As String
EventDetail = RsEvents!eve_evnt
End Property

Public Property Let EventDetail(ByVal vNewValue As String)
RsEvents!eve_evnt = vNewValue
End Property
Public Property Get TranNo() As String
TranNo = m_TranNo
End Property


Public Property Get ConnectionObject() As ADODB.Connection
Set ConnectionObject = m_Cn
End Property

Public Property Let ConnectionObject(ByVal vNewValue As ADODB.Connection)
Set m_Cn = vNewValue
Set RsEvents = New ADODB.Recordset

RsEvents.ActiveConnection = m_Cn
RsEvents.Source = "select * from xevent where eve_npecode='" & "asdf" & "'"
RsEvents.Open , , adOpenKeyset, adLockBatchOptimistic
Set RsEvents.ActiveConnection = Nothing
 
'RsEvents.ActiveConnection = m_Cn

End Property

Public Property Get STAs() As String
STAs = RsEvents!eve_stas
End Property

Public Property Let STAs(ByVal vNewValue As String)
RsEvents!eve_stas = Trim$(vNewValue)
End Property
Public Function update() As Boolean
If Trim$(UCase$(mUSERID & "")) <> "IMSUSA" Then
  Dim rsXuserpro As New ADODB.Recordset
  rsXuserpro.Source = "select usr_autoinacflag from xuserprofile where usr_npecode='" & mnamespace & "' and usr_userid='" & mUSERID & "'"
  rsXuserpro.ActiveConnection = m_Cn
  rsXuserpro.Open
  If rsXuserpro.EOF = True Then GoTo updateIt
  If rsXuserpro!usr_autoinacflag Then
updateIt:
        RsEvents.ActiveConnection = m_Cn
        RsEvents.MoveFirst
        m_TranNo = transnumber
        
        
        
        Do While Not RsEvents.EOF
           RsEvents!eve_trannumb = Trim$(m_TranNo)
           RsEvents.MoveNext
         Loop
        RsEvents.UpdateBatch
  End If
    Set rsXuserpro = Nothing
End If
End Function

Public Function AddNew()
RsEvents.AddNew

End Function


Private Function transnumber() As String

'transnumber = Trim$(Replace(Replace(Replace(Now(), "/", ""), " ", ""), ":", ""))
transnumber = Trim$(Replace(Replace(Replace(Format(Date, "YYYY/MM/DD") & Format(Time, "HH/MM/SS"), "/", ""), " ", ""), ":", ""))
End Function

Public Property Get RowNumber() As Integer
RowNumber = RsEvents.AbsolutePosition
End Property



