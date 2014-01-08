VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim Rowguid, lock_date As Date, locked As Boolean
    Dim VBPrimaryKey1, VBPrimaryKey2, VBPrimaryKey3
    Dim VBPrimaryKey4, VBPrimaryKey5, VBPrimaryKey6
    Dim VBPrimaryKey7, CurrentFormName, DBTableName
    Dim DBPrimaryKey1, DBPrimaryKey2, DBPrimaryKey3
    Dim DBPrimaryKey4, DBPrimaryKey5, DBPrimaryKey6
    Dim DBPrimaryKey7
    

Public Function Lock_Row()
   
    Dim Lock_cmd As New ADODB.Recordset, Locked_data As New ADODB.Recordset
    Dim sqlstring3 As String
''' insert current user information into the Lock table

    lock_date = Now()

Dim VBPrimaryKey
Dim j As Integer
Dim w(5)

j = 0

w(0) = VBPrimaryKey3
w(1) = VBPrimaryKey4
w(2) = VBPrimaryKey5
w(3) = VBPrimaryKey6
w(4) = VBPrimaryKey7

For j = 0 To 4
If w(j) = "" Then

'w(j) =
VBPrimaryKey(j) = ""

j = j + 1
End If
Next
    


        sqlstring3 = "INSERT INTO LOCK ([USER], DateOpened, TableName, PrimaryKey1, PrimaryKey2, PrimaryKey3,"
        sqlstring3 = sqlstring3 & " PrimaryKey4, PrimaryKey5, PrimaryKey6, PrimaryKey7) VALUES  ("
        sqlstring3 = sqlstring3 & "'" & CurrentUser & "','" & lock_date & "','" & DBTableName & "','" & VBPrimaryKey1 & "','" & VBPrimaryKey2 & ""
        sqlstring3 = sqlstring3 & "','" & VBPrimaryKey3 & "','" & VBPrimaryKey4 & "','" & VBPrimaryKey5 & "','"
        sqlstring3 = sqlstring3 & VBPrimaryKey6 & "','" & VBPrimaryKey7 & "')"


    Set Lock_cmd = New ADODB.Recordset
    Lock_cmd.Open sqlstring3, deIms.cnIms

Dim SQLstring4 As String
        
        SQLstring4 = "select rowguid from LOCK where DateOpened = '" & lock_date & "'"

    Set Locked_data = New ADODB.Recordset
    Locked_data.Open SQLstring4, deIms.cnIms
    
    Rowguid = Locked_data("rowguid")

locked = True

    End Function
Public Function Check_Lock(locked)

Dim sql_check

   
    Dim SQLstring2 As String, SQLstring1 As String, sqlstring3 As String
    Dim get_table_info As New ADODB.Recordset, get_keys As New ADODB.Recordset
    Dim Lock_Check As New ADODB.Recordset
    
    CurrentFormName = Forms(3).Name

    ''' collect the information about the current recordset, where it came from whats the primary key, etc.

SQLstring1 = "SELECT * FROM [SCHEMA] WHERE vbform = '" & CurrentFormName & "'"


    Set get_table_info = New ADODB.Recordset
    get_table_info.Open SQLstring1, deIms.cnIms

            DBTableName = get_table_info("DBTableName")
            DBPrimaryKey1 = get_table_info("DBPrimaryKey1")
            DBPrimaryKey2 = get_table_info("DBPrimaryKey2")
            DBPrimaryKey3 = get_table_info("DBPrimaryKey3")
            DBPrimaryKey4 = get_table_info("DBPrimaryKey4")
            DBPrimaryKey5 = get_table_info("DBPrimaryKey5")
            DBPrimaryKey6 = get_table_info("DBPrimaryKey6")
            DBPrimaryKey7 = get_table_info("DBPrimaryKey7")

      Set VBPrimaryKey1 = get_table_info("VBPrimaryKey1")
      Set VBPrimaryKey2 = get_table_info("VBPrimaryKey2")
      Set VBPrimaryKey3 = get_table_info("VBPrimaryKey3")
      Set VBPrimaryKey4 = get_table_info("VBPrimaryKey4")
      Set VBPrimaryKey5 = get_table_info("VBPrimaryKey5")
      Set VBPrimaryKey6 = get_table_info("VBPrimaryKey6")
      Set VBPrimaryKey7 = get_table_info("VBPrimaryKey7")

'''''' collect the actual primary keys from the main table

VBPrimaryKey1 = VB.Forms(3).Controls(VBPrimaryKey1.Value) 'VB.Forms(3).Controls("TxtSuppCode")
VBPrimaryKey2 = deIms.NameSpace


sql_check = "SELECT * FROM LOCK WHERE TableName = '" & DBTableName & "' and PrimaryKey1 = '" & VBPrimaryKey1 & "' and PrimaryKey2 = '" & VBPrimaryKey2 & "'"



Dim j As Integer
Dim w(5), n(5)

j = 0

w(0) = DBPrimaryKey3
w(1) = DBPrimaryKey4
w(2) = DBPrimaryKey5
w(3) = DBPrimaryKey6
w(4) = DBPrimaryKey7

For j = 0 To 4
If w(j) = "" Then

n(0) = VB.Forms(3).Controls(VBPrimaryKey3.Value)    'VBPrimaryKey3
n(1) = VB.Forms(3).Controls(VBPrimaryKey4.Value)    'VBPrimaryKey4
n(2) = VB.Forms(3).Controls(VBPrimaryKey5.Value)    'VBPrimaryKey5
n(3) = VB.Forms(3).Controls(VBPrimaryKey6.Value)    'VBPrimaryKey6
n(4) = VB.Forms(3).Controls(VBPrimaryKey7.Value)    'VBPrimaryKey7

   sql_check = sql_check & "AND " & w(j) & " = '" & n(j).Value & "' "

j = j + 1
End If
Next


    Set Lock_Check = New ADODB.Recordset
    Lock_Check.Open sql_check, deIms.cnIms


If Lock_Check.EOF Then
locked = False
Call Lock_Row
Else
''send a message box with the info
'' this is where the message box is sent to the current user

      Dim LockingUser
      
      Set LockingUser = Lock_Check("User")
      

    If CurrentUser <> LockingUser Then  '' jawdat currently will pop up msg box
 
 locked = True
  
  
  MsgBox ("This record is currently in use and being modified by " & LockingUser & ".  Close the form and try again.")

    End If
End If
End Function
Public Function Unlock_Row()
Dim Form As Boolean
Rowguid = "'" & Rowguid & "'"
Dim UnLock_data As New ADODB.Recordset, SQL_Unlock As String

''''''''''''' delete lock
If Form = True Then

SQL_Unlock = "DELETE FROM LOCK WHERE [User] = '" & CurrentUser & "'"

Else

SQL_Unlock = "DELETE FROM LOCK WHERE ROWGUID = " & Rowguid

End If

    Set UnLock_data = New ADODB.Recordset
   UnLock_data.Open SQL_Unlock, deIms.cnIms
    
locked = False

End Function

Public Function Update_time()

'''' update locked row with current timestamp so that SQL Agent will not delete

Dim SQL_Update_time As String, Update_LockTime As New ADODB.Recordset

lock_date = Now()

SQL_Update_time = "UPDATE LOCK SET DateOpened = '" & lock_date & "' WHERE [User] = '" & CurrentUser & "'"

    Set Update_LockTime = New ADODB.Recordset
    Update_LockTime.Open SQL_Update_time, deIms.cnIms

locked = True

End Function




