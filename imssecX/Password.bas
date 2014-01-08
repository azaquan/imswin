Attribute VB_Name = "modPassword"
Option Explicit

Public Enum LoginStatus
    lsUnknown = -1 ' we dono know status
    lsUnknownUser = 1 ' Unknown user
    
    lsSuccess = 0 ' normal logon
    lsInActive = 2 ' inactive user
    
    lsSuspended = 3 'suspended user
    lsSuspendedWithPasssword = 4 'suspended user with password
    
    lsTempSuspended = 5 ' Temporary user
    lsTempSuspendedWithPasssword = 6 ' Temporary user with password
    
    lsFirstLogon = 7
    lsFirstLogonWithAdmin = 8
    lsFirstLogonWithOwner = 9
    lsFirstLogonWithOutOwnerOrAdmin = 10
End Enum

'set store procedure parameters and call it

Public Function PASSWORDAGE(Namespace As String, UserId As String, MiniMum As Boolean, cn As ADODB.Connection) As Long
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = cn
        .CommandText = "PASSWORDAGE"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UserID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("Minimum", adBoolean, adParamInput, , MiniMum)
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamOutput)
        
        
        .Execute , , adExecuteNoRecords
        
        PASSWORDAGE = .Parameters("RetVal")
        
    End With
    
    Set cmd = Nothing
End Function

'set store procedure parameters and call it

Public Function GetMinPasswordLength(Namespace As String, UserId As String, cn As ADODB.Connection) As Long
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "GetMinPasswordLength"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UserID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamOutput)
        
        
        .Execute , , adExecuteNoRecords
        
        GetMinPasswordLength = .Parameters("RetVal")
        
    End With
    
    Set cmd = Nothing
End Function

'set store procedure parameters and call it

Public Function PassWordExist(Namespace As String, UserId As String, Password As String, cn As ADODB.Connection) As Boolean
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "PSWDEXISTRET"
        
        .Parameters.Append .CreateParameter("RV", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UserID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("PassWrd", adVarChar, adParamInput, 15, Password)
        
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamOutput)
        
        
        cn.Errors.Clear
        .Execute , , adExecuteNoRecords
        
        PassWordExist = .Parameters("RetVal")
        
    End With
        
    Set cmd = Nothing
End Function

'set store procedure parameters and call it

Public Function UNKNOWPASS(Namespace As String, UserId As String, NUMBINV As Integer, FLAG As Integer, cn As ADODB.Connection) As Boolean
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "UNKNOWPASS"
        
        
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamReturnValue)
        
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UserID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("numbinv", adInteger, adParamInput, , NUMBINV)
        
        .Parameters.Append .CreateParameter("flag", adInteger, adParamInput, 4, FLAG)
        
        
        
        .Execute , , adExecuteNoRecords
        
         UNKNOWPASS = .Parameters("RetVal") = 0
        
    End With
        
End Function

'set store procedure parameters and call it

Public Function CheckUserStatus(Namespace As String, UserId As String, cn As ADODB.Connection, rs As ADODB.Recordset) As LoginStatus
Dim cmd As ADODB.Command
On Error GoTo Handled

    CheckUserStatus = lsUnknown
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "USERSTATUS"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UID", adVarChar, adParamInput, 15, UserId)
        
        If Err Then Err.Clear
        
        Set rs = .Execute
        CheckUserStatus = rs!Status
        
    End With
        

    Set cmd = Nothing
    Exit Function
    
Handled:
    Call Err.Raise(1000, "User Status", "Error verifying user status")
End Function

'set store procedure parameters and call it to update xuserprifile table

Public Sub UpdateUserPassword(Namespace As String, UserId As String, Password As String, cn As ADODB.Connection)
On Error GoTo Cancelled
Dim cmd As ADODB.Command
Dim NewCn As New ADODB.Connection

    NewCn.ConnectionString = cn.ConnectionString
    NewCn.Open
    
    
    
    Set cmd = New ADODB.Command
    
    With cmd
    
    '    Set .ActiveConnection = cn
         Set .ActiveConnection = NewCn
         
        .CommandType = adCmdStoredProc
        .CommandText = "udsp_XUserProfile_Upd"
        
        .Parameters.Append .CreateParameter("N", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("UID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("PWD", adVarChar, adParamInput, 15, Password)
        
        cn.Errors.Clear
        .Execute , , adExecuteNoRecords
        
    End With

    Set cmd = Nothing

    If Not GetObjectContext Is Nothing Then GetObjectContext.SetComplete
    Exit Sub
    
Cancelled:
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetAbort
End Sub

'SQL statement to update xuserprofile after login process

Public Sub UpdateLastLogon(Namespace As String, UserId As String, cn As ADODB.Connection)
On Error GoTo Cancelled
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdText
        
        .CommandText = "UPDATE XUSERPROFILE SET "
        .CommandText = .CommandText & " usr_datelastlogn = "
        .CommandText = .CommandText & " GETDATE(), "
        .CommandText = .CommandText & " usr_numbinvaatte = 0, "
        .CommandText = .CommandText & " usr_numbremaatte = usr_maxiatte"
        
        .CommandText = .CommandText & " WHERE usr_userid = ?"
        .CommandText = .CommandText & " AND usr_npecode = ?"

        .Execute , Array(UserId, Namespace), adExecuteNoRecords
        
    End With
    
    Set cmd = Nothing
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetComplete
    Exit Sub
    
Cancelled:
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetAbort
End Sub

'set store procedure parameters and call it

Public Sub Updateexp(UserId As String, Password As String, cn As ADODB.Connection)
On Error GoTo Cancelled
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "XUserProfile_exp"
        
        '.Parameters.Append .CreateParameter("N", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("UID", adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter("PWD", adVarChar, adParamInput, 15, Password)
        
        
        .Execute , , adExecuteNoRecords
        
    End With

    Set cmd = Nothing

    If Not GetObjectContext Is Nothing Then GetObjectContext.SetComplete
    Exit Sub
    
Cancelled:
    If Not GetObjectContext Is Nothing Then GetObjectContext.SetAbort
End Sub

'SQL statement update number remand attempt field
'and invalid attempt field

Public Sub AddInvalidAttempt(UserName As String, Namespace As String, cn As ADODB.Connection)
Dim cmd As ADODB.Command
Dim str As String


    Exit Sub
    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdText
        Set .ActiveConnection = cn
        
        .CommandText = "UPDATE XUSERPROFILE"
        .CommandText = .CommandText & " SET usr_numbremaatte = usr_numbremaatte - 1, "
        .CommandText = .CommandText & " usr_numbinvaatte = usr_numbinvaatte + 1"
        .CommandText = .CommandText & " WHERE (usr_npecode = ?)"
        .CommandText = .CommandText & " AND (usr_userid = ?)"
        
        .Parameters.Append .CreateParameter("", adVarChar, adParamInput, 5, Namespace)
        .Parameters.Append .CreateParameter("", adVarChar, adParamInput, 15, UserName)
        
        .Execute , , adExecuteNoRecords
    End With
    
    
    Set cmd = Nothing
    
End Sub

'SQL statement tp get user level for user

Public Function UserLevel(UserId As String, Namespace As String, cn As ADODB.Connection) As Integer
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(cn, adCmdText)
    
    With cmd
        .CommandText = "select ? = usr_leve from XUSERPROFILE"
        .CommandText = .CommandText & " where usr_userid = ? "
        .CommandText = .CommandText & " and usr_npecode = ? "
        

        Call .Parameters.Append(.CreateParameter("Userlevel", adInteger, adParamOutput))
        Call .Parameters.Append(.CreateParameter("UserID", adVarChar, adParamInput, 15))
        Call .Parameters.Append(.CreateParameter("NameSpace", adVarChar, adParamInput, 15))

        
        .Parameters("UserID") = UserId
        .Parameters("NameSpace") = Namespace
        Call .Execute(Options:=adExecuteNoRecords)
        
        UserLevel = .Parameters("Userlevel")
    End With
       
End Function

