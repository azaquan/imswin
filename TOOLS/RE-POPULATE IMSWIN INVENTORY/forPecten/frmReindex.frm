VERSION 5.00
Begin VB.Form Frmreindex 
   Caption         =   "Reindex"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Frmreindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub rePopulate()

End Sub

Private Sub Form_Load()
Dim Rs0 As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset
Dim Rs5 As New ADODB.Recordset
Dim Rs6 As New ADODB.Recordset
Dim Rs7 As New ADODB.Recordset
Dim Rs8 As New ADODB.Recordset
Dim Rs9 As New ADODB.Recordset
Dim Rs10 As New ADODB.Recordset
Dim Rs11 As New ADODB.Recordset
Dim Str As String
Dim cmd As New ADODB.Command
Dim Cn As New ADODB.Connection
Dim InTransaction As Boolean

On Error GoTo Errhandler


InTransaction = True
DataEnvironment1.Cn.Open


Set Cn = DataEnvironment1.Cn

'Cn.Open
'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=scms;Persist Security Info=True;User ID=sa;Initial Catalog=SAKHALIN;Data Source=IMSDEV003"
'cn.Open

Cn.BeginTrans


Str = " IF EXISTS (SELECT TABLE_NAME From INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME = 'QTY') DROP VIEW QTY"

Rs0.Source = Str
Rs0.ActiveConnection = Cn
Rs0.Open
If Rs0.State <> 0 Then Rs0.Close

Str = " CREATE VIEW QTY AS (SELECT  qs5_compcode company, qs5_npecode namespace, qs5_ware warehouse, qs5_stcknumb stock#, qs5_logiware logic, qs5_subloca subloca, qs5_cond condition, qs5_primqty qty, qs5_secoqty qty2 From qtyst5 Union All SELECT qs6_compcode company, qs6_npecode namespace, qs6_ware warehouse, qs6_stcknumb stock#, qs6_logiware logic, qs6_subloca subloca, qs6_cond condition, qs6_primqty qty, qs6_secoqty qty2 From qtyst6)"

Rs1.Source = Str
Rs1.ActiveConnection = Cn
Rs1.Open
If Rs1.State <> 0 Then Rs1.Close

Str = " UPDATE qtyst4 SET qs4_primqty = QTY, qs4_secoqty = qty2 From (SELECT company, namespace, warehouse, stock#, logic, subloca, condition, SUM(qty) qty, SUM(qty2) qty2 FROM QTY LEFT OUTER JOIN LOCATION ON loc_locacode =  warehouse AND loc_npecode = namespace AND loc_compcode = company WHERE loc_gender <> 'OTHER' Group By company, namespace, warehouse, stock#, logic, subloca, condition) QT INNER JOIN QTYST4 ON condition = qs4_cond AND subloca = qs4_subloca AND logic = qs4_logiware AND company = qs4_compcode AND namespace = qs4_npecode AND warehouse = qs4_ware AND stock# = qs4_stcknumb WHERE qs4_primqty <> qty DROP VIEW QTY"

Rs2.Source = Str
Rs2.ActiveConnection = Cn
Rs2.Open
Set Rs2 = Nothing

Str = " IF EXISTS (SELECT TABLE_NAME From INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME = 'QTY') DROP VIEW QTY"

Rs3.Source = Str
Rs3.ActiveConnection = Cn
Rs3.Open
Set Rs3 = Nothing

Str = " CREATE VIEW QTY AS (SELECT qs4_compcode company, qs4_npecode namespace, qs4_ware warehouse, qs4_stcknumb stock#, qs4_logiware logic, qs4_subloca subloca, qs4_primqty qty, qs4_secoqty qty2 From qtyst4)"

Rs4.Source = Str
Rs4.ActiveConnection = Cn
Rs4.Open
Set Rs4 = Nothing

Str = " UPDATE qtyst3 SET qs3_primqty = QTY, qs3_secoqty = qty2 From (SELECT company, namespace, warehouse, stock#, logic, subloca, SUM(qty) qty, SUM(qty2) qty2 FROM QTY LEFT OUTER JOIN LOCATION ON loc_locacode =  warehouse AND loc_npecode = namespace AND loc_compcode = company WHERE loc_gender <> 'OTHER' Group By company, namespace, warehouse, stock#, logic, subloca) QT INNER JOIN QTYST3 ON company = qs3_compcode AND namespace = qs3_npecode AND warehouse = qs3_ware AND subloca = qs3_subloca AND logic = qs3_logiware AND stock# = qs3_stcknumb WHERE qs3_primqty <> qty DROP VIEW QTY"

Rs5.Source = Str
Rs5.ActiveConnection = Cn
Rs5.Open
Set Rs5 = Nothing

Str = " IF EXISTS (SELECT TABLE_NAME From INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME = 'QTY') DROP VIEW QTY"

Rs6.Source = Str
Rs6.ActiveConnection = Cn
Rs6.Open
Set Rs6 = Nothing

Str = " CREATE VIEW QTY AS (SELECT qs3_compcode company, qs3_npecode namespace, qs3_ware warehouse, qs3_stcknumb stock#, qs3_logiware logic, qs3_primqty qty, qs3_secoqty qty2 From qtyst3)"

Rs7.Source = Str
Rs7.ActiveConnection = Cn
Rs7.Open
Set Rs7 = Nothing

Str = " UPDATE qtyst2 SET qs2_primqty = QTY, qs2_secoqty = qty2, qs2_tbs = 1 From (SELECT company, namespace, warehouse, stock#, logic, SUM(qty) qty, SUM(qty2) qty2 FROM QTY LEFT OUTER JOIN LOCATION ON loc_locacode =  warehouse AND loc_npecode = namespace AND loc_compcode = company WHERE loc_gender <> 'OTHER' Group By company, namespace, warehouse, stock#, logic) QT INNER JOIN QTYST2 ON company = qs2_compcode AND namespace = qs2_npecode AND warehouse = qs2_ware AND logic = qs2_logiware AND stock# = qs2_stcknumb WHERE qs2_primqty <> qty DROP VIEW QTY"

Rs8.Source = Str
Rs8.ActiveConnection = Cn
Rs8.Open
Set Rs8 = Nothing

Str = " CREATE VIEW QTY AS (SELECT qs2_compcode company, qs2_npecode namespace, qs2_ware warehouse, qs2_stcknumb stock#, qs2_primqty qty, qs2_secoqty qty2 From qtyst2)"

Rs9.Source = Str
Rs9.ActiveConnection = Cn
Rs9.Open
Set Rs9 = Nothing

Str = " UPDATE qtyst1 SET qs1_primqty = QTY, qs1_secoqty = qty2, qs1_tbs = 1 From (SELECT company, namespace, warehouse, stock#, SUM(qty) qty, SUM(qty2) qty2 FROM QTY LEFT OUTER JOIN LOCATION ON loc_locacode =  warehouse AND loc_npecode = namespace AND loc_compcode = company WHERE loc_gender <> 'OTHER' Group By company, namespace, warehouse, stock#) QT INNER JOIN QTYST1 ON company = qs1_compcode AND namespace = qs1_npecode AND warehouse = qs1_ware AND stock# = qs1_stcknumb WHERE qs1_primqty <> qty DROP VIEW QTY"

Rs10.Source = Str
Rs10.ActiveConnection = Cn
Rs10.Open
Set Rs10 = Nothing




''cmd.CommandType = adCmdText
''cmd.CommandText = Str
''cmd.ActiveConnection = cn
''
''cmd.Execute


Cn.CommitTrans

MsgBox "Ims reindex ran Successfully."

Unload Me

Exit Sub
Errhandler:

If InTransaction = True Then Cn.RollbackTrans
    
MsgBox "Errors Occurred while trying to re-populate the tables. Err Description : " & Err.Description

Err.Clear

Unload Me

End Sub

Public Function SendEmail(Message As String, Subject As String)

    Dim Attmail As New imsutils.imsmisc
    Dim Address() As String
    Dim attachments() As String
    Dim Emails As String
    On Error GoTo Errhandler
    
    Emails = GetIniInformation
    
     Emails = Trim(Emails)
    
    If Len(Emails) = 0 Then
        
        ReDim Address(0)
        
        Address(0) = "muzammil@ims-sys.com"
        
    Else
    
    Address = Split(Emails, ";")
    
    End If
    
    Call Attmail.SendAttMail(Message, Subject, Address(), attachments())
    
    Exit Function
    
Errhandler:
    
Err.Clear

End Function
Public Function GetIniInformation() As String

Dim sa As Scripting.FileSystemObject

Dim t As Scripting.TextStream

Dim line As String

Dim location As Integer


Dim Filepath As String

Set sa = New Scripting.FileSystemObject

Filepath = App.Path & "\Reindex.txt"

If sa.FileExists(Filepath) = False Then

    MsgBox "There is no INI file associated with this program. Please create one."

    Exit Function
    
End If
 
   Set t = sa.OpenTextFile(Filepath)

   line = t.ReadLine
    
   location = InStr(line, "=")
    
   GetIniInformation = Mid(line, location + 1, Len(line))

   t.Close
   
   Set t = Nothing


End Function

