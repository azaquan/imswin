VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim rowguid, lock_date As Date, locked As Boolean, TableLocked As Boolean
    Dim VBPrimaryKey1, VBPrimaryKey2, VBPrimaryKey3, LockingUser
    Dim VBPrimaryKey4, VBPrimaryKey5, VBPrimaryKey6
    Dim VBPrimaryKey7, currentformname, dbtablename
    Dim DBPrimaryKey1, DBPrimaryKey2, DBPrimaryKey3
    Dim DBPrimaryKey4, DBPrimaryKey5, DBPrimaryKey6
    Dim DBPrimaryKey7, CurrentUser, connection As ADODB.connection
    

Public Function Lock_Row(currentformname, CurrentUser, connection, rowguid, Optional dbtablename)
   
    Dim Lock_cmd As New ADODB.Recordset, Locked_data As New ADODB.Recordset
    Dim sqlstring3 As String
''' insert current user information into the Lock table

    lock_date = Now()

Dim VBPrimaryKey
Dim j As Integer
Dim w(7)

j = 0



'w(0) = currentformname(1)
'w(1) = currentformname(2)
'w(2) = currentformname(3)
'w(3) = currentformname(4)
'w(4) = currentformname(5)
'w(5) = currentformname(6)
'w(6) = currentformname(7)

'For j = 0 To 4
'If w(j) = "" Then
'
''w(j) =
'currentformname(j) = ""
'
'j = j + 1
'End If
'Next
Dim comma_pos
If InStr(1, dbtablename, ",") Then
dbtablename = Mid(dbtablename, 1, InStr(1, dbtablename, ",") - 1)
End If


        sqlstring3 = "INSERT INTO LOCK ([USER], DateOpened, TableName, PrimaryKey1, PrimaryKey2, PrimaryKey3,"
        sqlstring3 = sqlstring3 & " PrimaryKey4, PrimaryKey5, PrimaryKey6, PrimaryKey7) VALUES  ("
        sqlstring3 = sqlstring3 & "'" & CurrentUser & "','" & lock_date & "','" & dbtablename & "','" & currentformname(1) & "','" & currentformname(2) & ""
        sqlstring3 = sqlstring3 & "','" & currentformname(3) & "','" & currentformname(4) & "','" & currentformname(5) & "','"
        sqlstring3 = sqlstring3 & currentformname(6) & "','" & currentformname(7) & "')"


    Set Lock_cmd = New ADODB.Recordset
    Lock_cmd.Open sqlstring3, connection


Dim SQLstring4 As String
        
        SQLstring4 = "select rowguid from LOCK where DateOpened = '" & lock_date & "'"

    Set Locked_data = New ADODB.Recordset
    Locked_data.Open SQLstring4, connection
    If Locked_data.EOF = False Then
    rowguid = Locked_data("rowguid")
End If
locked = True

    End Function
Public Function Check_Lock(locked, connection As ADODB.connection, CurrentUser, currentformname, currentformname1, rowguid, Optional dbtablename)

Dim sql_check, PrimaryKeyInDB1, PrimaryKeyInDB2, TableInDB1

   
    Dim SQLstring2 As String, SQLstring1 As String, sqlstring3 As String
    Dim get_table_info As New ADODB.Recordset, get_keys As New ADODB.Recordset
    Dim Lock_Check As New ADODB.Recordset
    


    ''' collect the information about the current recordset, where it came from whats the primary key, etc.

SQLstring1 = "SELECT * FROM [SCHEMA] WHERE vbform = '" & currentformname1 & "'"


    Set get_table_info = New ADODB.Recordset
    get_table_info.Open SQLstring1, connection

            dbtablename = get_table_info("DBTableName")
            DBPrimaryKey1 = get_table_info("DBPrimaryKey1")
            DBPrimaryKey2 = get_table_info("DBPrimaryKey2")
            DBPrimaryKey3 = get_table_info("DBPrimaryKey3")
            DBPrimaryKey4 = get_table_info("DBPrimaryKey4")
            DBPrimaryKey5 = get_table_info("DBPrimaryKey5")
            DBPrimaryKey6 = get_table_info("DBPrimaryKey6")
            DBPrimaryKey7 = get_table_info("DBPrimaryKey7")
'
'      Set VBPrimaryKey1 = get_table_info("VBPrimaryKey1")
'      Set VBPrimaryKey2 = get_table_info("VBPrimaryKey2")
    Set VBPrimaryKey3 = get_table_info("VBPrimaryKey3")
     Set VBPrimaryKey4 = get_table_info("VBPrimaryKey4")
      Set VBPrimaryKey5 = get_table_info("VBPrimaryKey5")
'      Set VBPrimaryKey6 = get_table_info("VBPrimaryKey6")
'      Set VBPrimaryKey7 = get_table_info("VBPrimaryKey7")

'''''' collect the actual primary keys from the main table

'VBPrimaryKey1 = VB.Forms(3).Controls(VBPrimaryKey1.Value) 'VB.Forms(3).Controls("TxtSuppCode")
'VBPrimaryKey2 = connection.NameSpace


sql_check = "SELECT * FROM LOCK WHERE TableName = '" & dbtablename & "' and PrimaryKey1 = '" & currentformname(1) & "' and PrimaryKey2 = '" & currentformname(2) & "'"



Dim j As Integer
Dim w(5), n(5)

w(3) = "PrimaryKey3"
w(4) = "PrimaryKey4"
w(5) = "PrimaryKey5"

j = 3
For j = 3 To 7

If currentformname(j) <> "" Then

n(j) = currentformname(j)

   sql_check = sql_check & "AND " & w(j) & " = '" & n(j) & "' "

j = j + 1
End If
Next


    Set Lock_Check = New ADODB.Recordset
    Lock_Check.Open sql_check, connection


If Lock_Check.EOF Then
locked = False
Call Lock_Row(currentformname, CurrentUser, connection, rowguid, dbtablename)
Else
''send a message box with the info
'' this is where the message box is sent to the current user

      Dim LockingUser
      
      
      Set LockingUser = Lock_Check("User")
      Set PrimaryKeyInDB1 = Lock_Check("PrimaryKey1")
      Set PrimaryKeyInDB2 = Lock_Check("PrimaryKey2")
      Set TableInDB1 = Lock_Check("TableName")
      
'  If ((CurrentUser = LockingUser) And (PrimaryKeyInDB1 = currentformname(1)) And (PrimaryKeyInDB2 = currentformname(2)) And (TableInDB1 = dbtablename)) Then
'
'   MsgBox ("You currently have this record open in another form, you must save your work and try again.")
' locked = True
'  End If
      
    If CurrentUser <> LockingUser Then  '' jawdat currently will pop up msg box

           
 locked = True
    
  MsgBox "This record is currently in use and being modified by " & LockingUser & ".  To modify data, click 'OK' and try again later.", vbCritical

    End If
End If
End Function



Public Function Unlock_Row(locked, connection As ADODB.connection, CurrentUser, rowguid, Optional grid1 As Boolean, Optional dbtablename As String, Optional POValue, Optional grid2 As Boolean)
Dim Form As Boolean ', POValue As String
If InStr(1, rowguid, "'") Then
'rowguid = "'" & rowguid & "'"
'Else
rowguid = Replace(rowguid, "'", "")
End If
Dim UnLock_data As New ADODB.Recordset, SQL_Unlock As String

''''''''''''' delete lock

If grid2 = True Then
SQL_Unlock = "DELETE FROM LOCK WHERE [User] = '" & CurrentUser & "' and tablename = '" & dbtablename & "'"

Else

If grid1 = True Then

SQL_Unlock = "DELETE FROM LOCK WHERE [User] = '" & CurrentUser & "' and primarykey1 = '" & POValue & "'"



'If Form <> "" Then
Else
If Form = True Then

SQL_Unlock = "DELETE FROM LOCK WHERE [User] = '" & CurrentUser & "'"

Else

SQL_Unlock = "DELETE FROM LOCK WHERE ROWGUID = '" & rowguid & "'"
End If
End If
End If


If rowguid <> "" Then
    Set UnLock_data = New ADODB.Recordset
   UnLock_data.Open SQL_Unlock, connection
End If
locked = False

End Function



Public Function Update_time(connection As ADODB.connection, CurrentUser)

'''' update locked row with current timestamp so that SQL Agent will not delete

Dim SQL_Update_time As String, Update_LockTime As New ADODB.Recordset

lock_date = Now()

SQL_Update_time = "UPDATE LOCK SET DateOpened = '" & lock_date & "' WHERE [User] = '" & CurrentUser & "'"

    Set Update_LockTime = New ADODB.Recordset
    Update_LockTime.Open SQL_Update_time, connection

locked = True

End Function

Public Function GetPrimaryControls(currentformname, connection As ADODB.connection) As String()

Dim SQLstring1 As String
Dim x(7) As String

Dim Rs As New ADODB.Recordset

SQLstring1 = "SELECT * FROM [SCHEMA] WHERE vbform = '" & currentformname & "'"


    Set Rs = New ADODB.Recordset
    Rs.Open SQLstring1, connection

       x(1) = Rs("VBPrimaryKey1") & ""
       x(2) = Rs("VBPrimaryKey2") & ""
       x(3) = Rs("VBPrimaryKey3") & ""
       x(4) = Rs("VBPrimaryKey4") & ""
       x(5) = Rs("VBPrimaryKey5") & ""
       x(6) = Rs("VBPrimaryKey6") & ""
       x(7) = Rs("VBPrimaryKey7") & ""
       
GetPrimaryControls = x


End Function



Public Function LOCK_TABLE(currentformname, connection As ADODB.connection, CurrentUser)

Dim LOCK_TABLE1 As New ADODB.Recordset, get_table_info2 As New ADODB.Recordset
Dim SQLLockTable, SQLstring1
Dim lock_date As Date

SQLstring1 = "SELECT * FROM [SCHEMA] WHERE vbform = '" & currentformname & "'"


    Set get_table_info2 = New ADODB.Recordset
    get_table_info2.Open SQLstring1, connection

            dbtablename = get_table_info2("DBTableName")
'            DBPrimaryKey1 = get_table_info2("DBPrimaryKey1")
'            DBPrimaryKey2 = get_table_info2("DBPrimaryKey2")
'            DBPrimaryKey3 = get_table_info2("DBPrimaryKey3")
'            DBPrimaryKey4 = get_table_info2("DBPrimaryKey4")
'            DBPrimaryKey5 = get_table_info2("DBPrimaryKey5")
'            DBPrimaryKey6 = get_table_info2("DBPrimaryKey6")
'            DBPrimaryKey7 = get_table_info2("DBPrimaryKey7")


lock_date = Now()

        SQLLockTable = "INSERT INTO LOCK ([USER], DateOpened, TableName) VALUES  ("
        SQLLockTable = SQLLockTable & "'" & CurrentUser & "','" & lock_date & "','" & dbtablename & "')"


    Set LOCK_TABLE1 = New ADODB.Recordset
    LOCK_TABLE1.Open SQLLockTable, connection
    
TableLocked = True
End Function
Public Function CHECK_TABLE(TableLocked, currentformname, connection As ADODB.connection, CurrentUser)

Dim get_table_check As New ADODB.Recordset, get_table_data As New ADODB.Recordset
Dim SQL_Table_Check As String, SQL_locked_Check

SQL_Table_Check = "SELECT * FROM [SCHEMA] WHERE vbform = '" & currentformname & "'"


    Set get_table_data = New ADODB.Recordset
    get_table_data.Open SQL_Table_Check, connection

    dbtablename = get_table_data("DBTableName")

SQL_locked_Check = "SELECT * FROM LOCK WHERE TableName = '" & dbtablename & "'"

    Set get_table_check = New ADODB.Recordset
    get_table_check.Open SQL_locked_Check, connection


If get_table_check.EOF Then

TableLocked = False

Call LOCK_TABLE(currentformname, connection, CurrentUser)

Else
''send a message box with the info
'' this is where the message box is sent to the current user

      Dim LockingUser
      
      Set LockingUser = get_table_check("User")
      

    If CurrentUser <> LockingUser Then  '' jawdat currently will pop up msg box
 
 TableLocked = True
  
  
  MsgBox "This table is currently in use and being modified by " & LockingUser & ".  To modify data, click 'Ok' and try again later.", vbCritical

    End If

End If

End Function

Public Function UnLock_table(TableLocked, currentformname, connection As ADODB.connection, CurrentUser)

Dim SQL_Unlock_Table As String, SQL_Table_Check As String
Dim get_table_check As New ADODB.Recordset

SQL_Table_Check = "SELECT * FROM [SCHEMA] WHERE vbform = '" & currentformname & "'"


    Set get_table_check = New ADODB.Recordset
    get_table_check.Open SQL_Table_Check, connection

    dbtablename = get_table_check("DBTableName")


SQL_Unlock_Table = "DELETE FROM LOCK WHERE [User] = '" & CurrentUser & "' AND TableName = '" & dbtablename & "'"



   Set UnLock_table = New ADODB.Recordset
   UnLock_table.Open SQL_Unlock_Table, connection
    
TableLocked = False

End Function

