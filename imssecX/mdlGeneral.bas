Attribute VB_Name = "mdlGeneral"
Option Explicit

Public ReportPath As String
Public CurrentUser As String
Public Const CryptKey = "table"
Public Declare Sub DebugBreak Lib "kernel32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Added by Juan Gonzalez (8/28/200) for Translation fixes
Public TR_MESSAGES As ADODB.Recordset
Public TR_CONTROLS As ADODB.Recordset
Global msg1, msg2 As String
Public TR_LANGUAGE As String

Public Type ValueChanged
    Changed As Boolean
    RowNumb As Integer
End Type

'Added Juan 11/03/00
Global m_dsnname As String


'set store procedure parameters and call it

Public Function PSWDEXITRET(NameSpace As String, UserName As String, Password As String, cn As ADODB.Connection) As Boolean

    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandText = "PSWDEXITRET"
        
        .Parameters.Append .CreateParameter("rt", adBoolean, adParamReturnValue)
        .Parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter("UN", adVarChar, adParamInput, 15, UserName)
        .Parameters.Append .CreateParameter("PW", adVarChar, adParamInput, 15, Password)
        .Parameters.Append .CreateParameter("RetVal", adBoolean, adParamOutput)
        
        .Execute
        
        PSWDEXITRET = .Parameters("Retval")
                
    End With
    
    Set cmd = Nothing
End Function

'set text boxse lenght

Public Sub SelectWhenFocus(txt As TextBox)
    With txt
        .SelStart = 0
        .SelLength = Len(txt)
    End With

End Sub

'SQL statement to check password exist or not

Public Function IsCurrentPassword(NameSpace As String, UserId As String, Password As String, cn As ADODB.Connection) As Boolean
Dim i As Integer
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        .CommandType = adCmdText
        Set .ActiveConnection = cn
        
        .CommandText = "SELECT ? = COUNT(*) FROM XUSERPROFILE"
        .CommandText = .CommandText & " WHERE usr_pswd  = ? "
        .CommandText = .CommandText & " AND usr_userid = ? "
        .CommandText = .CommandText & " AND usr_npecode = ?"
        
        .Parameters.Append .CreateParameter("Rt", adInteger, adParamOutput)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, Password)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, NameSpace)
        
        Call .Execute(, , adExecuteNoRecords)
        
        IsCurrentPassword = .Parameters("RT").Value
        
    End With
    
    Set cmd = Nothing
End Function

'populate name space data grid

Public Function AddNameSpaces(rs As ADODB.Recordset, cbo As ComboBox, Optional code As Boolean = False) As Collection
On Error Resume Next

    If rs Is Nothing Then Exit Function
    If rs.BOF And rs.EOF Then Exit Function
    'If rs.RecordCount = 0 Then Exit Function
    If rs.State And adStateOpen <> adStateOpen Then Exit Function
    
    cbo.Clear
    Set AddNameSpaces = New Collection
    
    Do While Not rs.EOF
    
        cbo.AddItem IIf(code, rs!npce_code & "", rs!npce_name & "")
        Call AddNameSpaces.Add(CStr(rs!npce_code & ""), CStr(rs!npce_name & ""))
        
        rs.MoveNext
        
    Loop
    
    rs.Close
    Set rs = Nothing
    cbo.ListIndex = 0
End Function

Public Sub translate_reports(currentFORM As String, reportNAME As String, mainREPORT As Boolean, cn As ADODB.Connection, crystalR As CrystalReport)

'Procedure for labels translations in every report
    Dim subreportQUERY As New ADODB.Recordset
    Dim n, i, ii, X, xx As Integer
    Dim DSNtext, tableNAME, Text, dsnPWD, dsnUID, dsnDSQ As String
    Dim cnSTRING() As String
    Dim sql, mainREP, subREP As String
    
    X = InStr(App.Path, ":") + 2
    DSNtext = UCase(Mid(App.Path, X, InStr(X, App.Path, "\") - X)) 'J added
    If TR_LANGUAGE = "" Then TR_LANGUAGE = "US"
    If TR_LANGUAGE = "US" Or TR_LANGUAGE = "*" Then
    Else
        If TR_CONTROLS Is Nothing Then
            Set TR_CONTROLS = New ADODB.Recordset
            sql = "SELECT TRMESSAGE.msg_lang, TRTRANSLATION.trs_enty, TRTRANSLATION.trs_obj, " _
                & "TRMESSAGE.msg_text FROM TRTRANSLATION LEFT OUTER JOIN TRMESSAGE ON TRTRANSLATION.trs_mesgnum = TRMESSAGE.msg_numb"
            TR_CONTROLS.Open sql, cn, adOpenStatic
        End If
        
        With TR_CONTROLS
            .Filter = ""
            .Filter = "trs_enty = '" + reportNAME + "' and msg_lang = '" + TR_LANGUAGE + "'"
            
            If .RecordCount > 0 Then
                n = 0
                For i = 0 To VB.Forms.count
                    If VB.Forms(i).Name = currentFORM Then
                        For ii = 0 To VB.Forms.count
                            If TypeName(crystalR) = "CrystalReport" Then
                                If (TR_LANGUAGE <> "*" And TR_LANGUAGE <> "") And TR_LANGUAGE <> "US" Then
                                    Do While Not .EOF
                                        crystalR.Formulas(n) = !trs_obj + " = '" + !msg_text + "'"
                                        n = n + 1
                                        .MoveNext
                                    Loop
                                End If
                                Exit For
                            End If
                        Next
                        Exit For
                    End If
                Next
            End If
        End With
        Err.Clear
    End If
    
    'This part is for multidatabase purposes
    If mainREPORT Then
        X = crystalR.RetrieveLogonInfo - 1
        X = crystalR.RetrieveDataFiles - 1
    Else
        mainREP = crystalR.ReportFileName
        mainREP = Mid(mainREP, InStrRev(mainREP, "\") + 1)
        subREP = reportNAME
        tableNAME = ""
        sql = "SELECT * FROM REPORTS WHERE report = '" + mainREP + "' and subreport = '" + subREP + "'"
        Set subreportQUERY = New ADODB.Recordset
        With subreportQUERY
            .Open sql, cn, adOpenForwardOnly
            X = .RecordCount
            ReDim alltables(X) As String
            X = 0
            If .RecordCount > 0 Then
                Do While Not .EOF
                    alltables(X) = !tableNAME
                    X = X + 1
                    .MoveNext
                Loop
                X = X - 1
            End If
            .Close
        End With
    End If
        
    cnSTRING = Split(cn.ConnectionString, ";")
    For Each Text In cnSTRING
        Select Case Left(UCase(Text), InStr(Text, "="))
            Case "PASSWORD="
                dsnPWD = Mid(Text, InStr(Text, "=") + 1)
            Case "USER ID="
                dsnUID = Mid(Text, InStr(Text, "=") + 1)
            Case "INITIAL CATALOG="
                dsnDSQ = Mid(Text, InStr(Text, "=") + 1)
        End Select
    Next
    
    
    For n = 0 To X
        DSNtext = cn.ConnectionString
        crystalR.LogonInfo(n) = "dsn=" + m_dsnname + ";dsq=" _
        & dsnDSQ + ";uid=" + dsnUID + ";pwd=" + dsnPWD
        
        tableNAME = crystalR.DataFiles(n)
        If tableNAME = "" Then
            If Not mainREPORT Then
                 If IsNull(alltables(n)) Or alltables(n) = "" Then
                Else
                    xx = InStr(alltables(n), ".")
                    tableNAME = Mid(alltables(n), IIf(xx = 0, 1, xx))
                End If
            End If
        Else
            xx = InStr(tableNAME, ".")
            tableNAME = Mid(tableNAME, IIf(xx = 0, 1, xx))
        End If
        crystalR.DataFiles(n) = dsnDSQ + tableNAME
    Next
End Sub

'set store procedure parameters and call it

Public Sub UPDATEXPASSWORD(NameSpace As String, UserId As String, _
                           Password As String, PasswordType As String, _
                           UserType As String, cn As ADODB.Connection)
                           
                           

Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
    
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATEXPASSWORD"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, NameSpace)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, UserId)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, Password)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, PasswordType)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, UserType)
        
        .Execute , , adExecuteNoRecords
        
    End With
    
    Set cmd = Nothing
End Sub

'call encrypt function

Public Function Encrypt(Text As String) As String
Dim enc As ImsSecX.imsCryptoClass
    
    If Len(Text) = 0 Then Exit Function
    Set enc = New ImsSecX.imsCryptoClass
    Encrypt = enc.Encrypttext(Text, CryptKey)
    
    Set enc = Nothing
End Function

'function disable controls and combo boxse
Public Sub DisableButtons(frmname As Form, oNavbar As Object, NameSpace As String, CurrentUser As String, cn As ADODB.Connection)
Dim bl As Boolean
Dim ctl As Control
On Error Resume Next

    bl = Getmenuuser(NameSpace, CurrentUser, frmname.Tag, cn)

    If bl = True Then Exit Sub
    If Not (TypeOf oNavbar Is LRNavigators.NavBar _
    Or TypeOf oNavbar Is LRNavigators.LROleDBNavBar) Then Exit Sub


    If Err Then Err.Clear
    For Each ctl In frmname.Controls


        If TypeOf ctl Is ComboBox Then
            ctl.Locked = True

        ElseIf TypeOf ctl Is SSOleDBCombo Then ctl.AllowInput = False
        
        ElseIf Not (TypeOf ctl Is Label) Then ctl.Enabled = False
        
        Debug.Print ctl.Name
            Debug.Print TypeName(ctl)
        End If
        
        If TypeOf ctl Is SSOleDBGrid Then ctl.Enabled = True:    ctl.AllowUpdate = False
        'IF TYPEOF CTL IS SSOleDBDropDown THEN CTL.
    If Err Then Err.Clear
    Next ctl
    
    oNavbar.Enabled = True
    oNavbar.NewEnabled = False
    oNavbar.SaveEnabled = False
    oNavbar.CancelEnabled = False

End Sub

'function error check and save file to log

Public Sub LogErr(RoutineName As String, ErrorDescription As String, ErrorNumber As Long, Optional Clear As Boolean = False)

Dim str As String
Dim i As IMSFile
Dim ms As imsmisc
Dim lrsys As lrSysInfo

Dim FileName As String
Dim Filenumb As Integer
On Error Resume Next

    Set i = New IMSFile
    Set ms = New imsmisc
    Set lrsys = New lrSysInfo
        
    str = ms.FixDir(StripNullChar(lrsys.SystemDirectory)) + "LogFiles\"
    
    
    If Not i.DirectoryExists(str) Then Call MkDir(str)
    
    str = FixDir(str) & "ims\"
    If Not i.DirectoryExists(str) Then Call MkDir(str)
    
    
    Filenumb = FreeFile
    FileName = str + i.ChangeFileExt(App.EXEName + Format$(Date, "ddmmyy"), "imserrlog")
    
    Open FileName For Append As Filenumb
    
        Print #Filenumb, "Module:             " & App.EXEName
        Print #Filenumb, "Routine:            " & RoutineName
        Print #Filenumb, "Error Number:       " & ErrorNumber
        Print #Filenumb, "Error Source:       " & Err.Source
        Print #Filenumb, "Error Description:  " & ErrorDescription
        Print #Filenumb, "Error Date:         " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        
        Print #Filenumb, "": Print #Filenumb, ""
    Close #Filenumb
    
    Set i = Nothing
    Set ms = Nothing
    Set lrsys = Nothing
End Sub

'make command conncetion

Public Function MakeCommand(cn As ADODB.Connection, CommandType As ADODB.CommandTypeEnum) As ADODB.Command
    Set MakeCommand = Nothing
    
    Set MakeCommand = New ADODB.Command
    Set MakeCommand.ActiveConnection = cn
    MakeCommand.CommandType = CommandType
End Function

'strip null character

Public Function StripNullChar(Source As String) As String
    StripNullChar = Left(Source, (InStr(1, Source, vbNullChar, vbTextCompare) - 1))
End Function

Function Trans(MessageCode) As String
'Function for retrieve direct texts for translation
    If (TR_LANGUAGE <> "*" And TR_LANGUAGE <> "") And TR_LANGUAGE <> "US" Then
        With TR_MESSAGES
            .Filter = "lan_code = '" + TR_LANGUAGE + "'"
            .MoveFirst
            .Find "msg_numb = '" + MessageCode + "'"
            If .EOF Then
                Trans = ""
            Else
                Trans = !msg_text
            End If
        End With
    End If
End Function

Sub Translate_Forms(Form_name As String)
'Procedure for captions translations in every form
    If (TR_LANGUAGE <> "*" And TR_LANGUAGE <> "") And TR_LANGUAGE <> "US" Then
        Dim i, j, k, indexARRAY, indexTAB, indexCONTROL, indexCOL As Integer
        Dim originalFILTER, nameCONTROLs, nameCONTROLs2 As String
        Dim withARRAY, withTAB As Boolean
        On Error Resume Next
        
        With TR_CONTROLS
            originalFILTER = ""
            For i = 0 To VB.Forms.count - 1
                If VB.Forms(i).Name = Form_name Then
                    .Filter = "msg_lang = '" + TR_LANGUAGE + "' and trs_enty = '" + Form_name + "'"
                    If .RecordCount > 0 Then
                        .Find "trs_obj = '" + Form_name + "'"
                        If Not .EOF Then
                            VB.Forms(i).Caption = !msg_text
                        End If
                        For j = 0 To VB.Forms(i).Controls.count - 1
                            nameCONTROLs = VB.Forms(i).Controls(j).Name
                            indexARRAY = -1
                            indexARRAY = VB.Forms(i).Controls(j).Index
                            If indexARRAY >= 0 Then
                                nameCONTROLs = VB.Forms(i).Controls(j).Name + "(" + Format(indexARRAY) + ")"
                            Else
                                indexTAB = -1
                                indexTAB = VB.Forms(i).Controls(j).Tabs
                                If indexTAB = "" Then indexTAB = -1
                                If indexTAB >= 0 Then
                                    For k = 0 To indexTAB - 1
                                        nameCONTROLs = VB.Forms(i).Controls(j).Name + ".tab(" + Format(k) + ")"
                                        .MoveFirst
                                        .Find "trs_obj = '" + nameCONTROLs + "'"
                                        If Not .EOF Then
                                            VB.Forms(i).Controls(j).TabCaption(k) = !msg_text
                                        End If
                                    Next
                                Else
                                    indexTAB = VB.Forms(i).Controls(j).Tabs.count
                                    If indexTAB > 0 Then
                                        For k = 1 To indexTAB
                                            nameCONTROLs = VB.Forms(i).Controls(j).Name + ".tabs(" + Format(k) + ")"
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs + "'"
                                            If Not .EOF Then
                                                VB.Forms(i).Controls(j).Tabs(k).Caption = !msg_text
                                            End If
                                        Next
                                    Else
                                        nameCONTROLs = VB.Forms(i).Controls(j).Name
                                    End If
                                End If
                            End If
                            If indexTAB < 0 Then
                                .MoveFirst
                                .Find "trs_obj like '" + nameCONTROLs + "%'"
                                If Not .EOF Then
                                    indexCOL = -1
                                    indexCOL = VB.Forms(i).Controls(j).Columns.count
                                    If indexCOL >= 0 Then
                                        For k = 0 To indexCOL - 1
                                            nameCONTROLs2 = nameCONTROLs + "." + VB.Forms(i).Controls(j).Columns(k).Caption
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                            If Not .EOF Then
                                                VB.Forms(i).Controls(j).Columns(k).Caption = !msg_text
                                            Else
                                                nameCONTROLs2 = nameCONTROLs + ".Columns(" + Format(k) + ")"
                                                .MoveFirst
                                                .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                                If Not .EOF Then VB.Forms(i).Controls(j).Columns(k).Caption = !msg_text
                                            End If
                                        Next
                                        .MoveFirst
                                        .Find "trs_obj = '" + nameCONTROLs + "'"
                                        If Not .EOF Then VB.Forms(i).Controls(j).Caption = !msg_text
                                    Else
                                        .MoveFirst
                                        .Find "trs_obj = '" + nameCONTROLs + "'"
                                        If Not .EOF Then VB.Forms(i).Controls(j).Caption = !msg_text
                                    End If
                                End If
                            End If
                        Next
                    End If
                    Exit For
                End If
            Next
        End With
    End If
End Sub




Public Function InitializeXevents(ObjXevents As ImsXevents, cn As ADODB.Connection)
 Set ObjXevents = New ImsXevents
 ObjXevents.ConnectionObject = cn
End Function

Public Sub SendEmails(Message As String, subject As String, DistributionCode As String, NameSpace As String, cn As ADODB.Connection, Attachments() As String)
Dim rsEmail As ADODB.Recordset
Dim rsPesys As ADODB.Recordset
Dim Address() As String
Dim X As Integer
Dim Utils As imsutilsx.imsmisc
'dim Message as String
On Error GoTo Handler

Set rsEmail = New ADODB.Recordset
Set rsPesys = New ADODB.Recordset

rsEmail.Source = "select dis_mail from distribution where dis_id='" & DistributionCode & "' and  dis_npecode='" & NameSpace & "'"
rsEmail.ActiveConnection = cn
rsEmail.Open

  
  X = 0
  Do While Not rsEmail.EOF
     ReDim Preserve Address(X)
     Address(X) = rsEmail!dis_mail
     rsEmail.MoveNext
     X = X + 1
  Loop
  

rsPesys.Source = "select psys_site from pesys where psys_npecode = '" & NameSpace & "'"
rsPesys.ActiveConnection = cn
rsPesys.Open
                    
                    
Set Utils = New imsutilsx.imsmisc
Message = Message & " On Station " & Trim$(rsPesys!psys_site) & "."

Call Utils.SendAttMail(Message, subject, Address, Attachments)
                    
 Exit Sub
Handler:
 MsgBox "Errors Occurred While sending an email. Error Description is " & Err.Description
 Err.Clear
End Sub



Public Function GetNameSpacesForUser(UserId As String, cn As ADODB.Connection, cboNameSpace As ComboBox, NameSpaces As Collection) As Boolean

Dim rs As ADODB.Recordset

On Error GoTo ErrHandler
GetNameSpacesForUser = False
    
    
    Set rs = New ADODB.Recordset
    'Rs.Source = "select npce_code ,npce_name from namespace inner join usr_npecode from xuserprofile where usr_userid ='" & USERID & "'"
    rs.Source = "select npce_code, npce_name from xuserprofile inner join namespace on npce_code = usr_npecode where usr_userid='" & UserId & "'"
    rs.Open , cn
    
    If rs.RecordCount > 0 Then
    
        cboNameSpace.Visible = True
        
                
        cboNameSpace.Clear
        
        Set NameSpaces = Nothing
        Set NameSpaces = New Collection
        
        Do While Not rs.EOF

            cboNameSpace.AddItem rs("npce_name")
            Call NameSpaces.Add(CStr(rs!npce_code & ""), CStr(rs!npce_name & ""))
            rs.MoveNext

        Loop
        
            cboNameSpace.ListIndex = 0
        
     ElseIf rs.RecordCount = 0 Then
            
            cboNameSpace.Visible = False
            cboNameSpace.Clear
            
            MsgBox "Please make sure that the UserId is valid.", vbInformation, "Ims"
     
     End If
    
GetNameSpacesForUser = True
Exit Function
ErrHandler:


MsgBox "Errors Occured while trying to get the namespaces registered by the user. Error Description :" & Err.Description, vbInformation, "Ims"
Err.Clear

End Function

