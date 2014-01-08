Attribute VB_Name = "Utils"
Option Explicit
Public Const WM_CLOSE = &H10

Public Const LB_ERR = (-1)
Public Const CB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)
Public Const LB_FINDSTRING = &H18F
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158


'// For startupinfo structure
Public Enum StartupInfoFlags
    STARTF_FORCEOFFFEEDBACK = &H80
    STARTF_FORCEONFEEDBACK = &H40
    STARTF_RUNFULLSCREEN = &H20
    STARTF_USECOUNTCHARS = &H8
    STARTF_USEFILLATTRIBUTE = &H10
    STARTF_USEPOSITION = &H4
    STARTF_USESHOWWINDOW = &H1
    STARTF_USESIZE = &H2
    STARTF_USESTDHANDLES = &H100
End Enum

Public Enum UserRights
    mdReadonly = 0
    mdReadWriteOnly
End Enum


Public Enum ShowWindowEnum
    SW_ERASE = &H4
    SW_HIDE = 0
    SW_INVALIDATE = &H2
    SW_MAX = 10
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
    SW_OTHERUNZOOM = 4
    SW_OTHERZOOM = 2
    SW_PARENTCLOSING = 1
    SW_PARENTOPENING = 3
    SW_RESTORE = 9
    SW_SCROLLCHILDREN = &H1
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWNORMAL = 1

End Enum


Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Enum FormMode
    mdNa = 0
    mdCreation
    mdModified
    mdModification
    mdvisualization
End Enum

'Added by muzammil



Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal HWND As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Long, lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'Added by Juan Gonzalez 8/29/2000 for Translation fix
Global msg1, msg2 As String
Global translator As imsTranslator
'----------------------------------------------------
Global fromAlphaSearch As Boolean
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii, r As Integer
Dim word As String
Dim found As Boolean
    With gridACTIVE
        If cellACTIVE = "" Then
            .row = 1
            .RowSel = 1
            .Col = 0
            .ColSel = 5
            .topROW = 1
        Else
            If Not .Visible Then .Visible = True
            If .Rows < val(.Tag) Then .Tag = 1
            If IsNumeric(.Tag) Then
                .Col = column
            End If
            If .Cols <= column Then Exit Sub
            .Col = column
            .Tag = ""
            found = False
            For r = 1 To .Rows - 1
                word = .TextMatrix(r, 0)
                If UCase(cellACTIVE) = UCase(Left(word, Len(cellACTIVE))) Then
                    found = True
                    Exit For
                End If
            Next
            If found Then
                .row = r
                .RowSel = r
                .Col = 0
                .ColSel = 5
                .topROW = r
            End If
            fromAlphaSearch = True
            cellACTIVE.SetFocus
            cellACTIVE.SelStart = Len(cellACTIVE)
        End If
    End With
End Sub

Public Function ToArrayFromRecordset(rs As ADODB.Recordset) As String()
Dim str() As String
Dim UpperBound As Integer

On Error GoTo ErrHandler
    ReDim str(0)
    UpperBound = -1
    If rs Is Nothing Then Exit Function
        
    rs.MoveFirst
    Do While Not rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve str(UpperBound)
        If InStr(UCase(rs(0)), "INTERNET!") > 0 Then
            str(UpperBound) = Mid(rs(0), 10)
        Else
            str(UpperBound) = rs(0)
        End If
        rs.MoveNext
    Loop
    ToArrayFromRecordset = str
    Exit Function
    
ErrHandler:
    Err.Raise Err.number, Err.Description
    Err.Clear
End Function

Public Function ToArrayFromRecO(rs As PoReceipients) As String()
Dim BK As Variant
Dim str() As String
Dim OldFilter As Variant
Dim UpperBound As Integer
On Error GoTo ErrHandler
    ReDim str(0)
    UpperBound = -1
    If rs Is Nothing Then Exit Function
    
    rs.MoveFirst
    Do While Not rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve str(UpperBound)
        str(UpperBound) = rs.Receipient
        rs.MoveNext
    Loop
    
    ToArrayFromRecO = str
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.number, Err.Description
    Err.Clear
End Function

'check load form status
Public Function IsLoaded(Formname As String) As Boolean
Dim frm As Form
    
    IsLoaded = False
    For Each frm In Forms
        DoEvents
        If frm.Name = Formname Then IsLoaded = True: Exit For
    Next frm
End Function

'************************************************************
' Procedure:    IndexOf
' Arguments:
' Created By:   Rohan Williams
' Date Created: 07/14/1999
' Purpose:
'************************************************************
'**********  Modification Log   *****************************
'Modified By:                                 Date Modified:
'
'************************************************************
Public Function IndexOf(cbo As ComboBox, Item As String) As Integer
    IndexOf = SendMessageStr(cbo.HWND, CB_FINDSTRING, 0, Item)
End Function

Public Function ListIndexOf(lstBox As ListBox, Item As String) As Long
    ListIndexOf = SendMessageStr(lstBox.HWND, LB_FINDSTRING, 0&, Item)
End Function

'Public Function IndexOfDataCombo(cbo As DataCombo, Item As String) As Integer
'   IndexOfDataCombo = -1
'End Function

'************************************************************
' Procedure:    LaunchApp
' Arguments:
' Created By:   Rohan Williams
' Date Created: 07/14/1999
' Purpose:
'************************************************************
'**********  Modification Log   *****************************
'Modified By:                                 Date Modified:
'
'************************************************************
Public Function LaunchApp(Filename As String, CmdLine As String, CmdShow As ShowWindowEnum, Optional CloseHandles As Boolean = False) As Boolean
On Error Resume Next
    Dim SI As STARTUPINFO, PI As PROCESS_INFORMATION
    Dim retval As Long

    'Make memory nullable
    
    With SI
        .cb = Len(SI)
        .wShowWindow = CmdShow
        .dwFlags = STARTF_USESHOWWINDOW
    End With
    
    LaunchApp = CreateProcess(Filename, CmdLine, ByVal 0, ByVal 0, 0, 0, ByVal 0, vbNullString, SI, PI)
    
    retval = CloseHandle(PI.hThread)
    retval = CloseHandle(PI.hProcess)
End Function

'function for loading application

Public Function LaunchAppAndWait(Filename As String, CmdLine As String, CmdShow As ShowWindowEnum, Optional CloseHandles As Boolean = False) As Boolean
On Error Resume Next
    Dim SI As STARTUPINFO, PI As PROCESS_INFORMATION
    Dim retval As Long

    'Make memory nullable
    
    With SI
        .cb = Len(SI)
        .wShowWindow = CmdShow
        .dwFlags = STARTF_USESHOWWINDOW
    End With
    
    
    LaunchAppAndWait = CreateProcess(Filename, CmdLine, ByVal 0, ByVal 0, 0, 0, ByVal 0, vbNullString, SI, PI)
    retval = WaitForSingleObjectEx(PI.hProcess, -1, -1)
    
    
    retval = CloseHandle(PI.hThread)
    retval = CloseHandle(PI.hProcess)
End Function

'
Function GetNearestComboItem(cbo As ComboBox, Optional KeyAscii As Integer, Optional sItem As String) As Boolean
On Error Resume Next
Dim y As Integer, i As Integer

    #If DBUG = 0 Then
        On Error Resume Next
    #End If
    

    If sItem = "" Then
            
            cbo.SelText = ""
            sItem = cbo.Text
            y = Len(cbo)
            i = cbo.SelLength
            
            If KeyAscii = 0 Then
            
            ElseIf KeyAscii > 31 Then
                cbo.SelText = Chr$(KeyAscii)
                
           
            Else
                y = y - 1
                i = i + 1
                
            
                cbo.SelStart = y
                cbo.SelLength = i
                cbo.SelText = ""
                cbo.SelStart = y
            End If
        
        sItem = cbo.Text: KeyAscii = 0:
        y = cbo.SelStart: i = cbo.SelLength
        cbo.SetFocus: cbo.SelStart = y: cbo.SelLength = i
    End If
    
    i = SendMessageStr(cbo.HWND, CB_FINDSTRING, CLng(-1), sItem)
                
    
        If i = CB_ERR Then i = cbo.ListIndex
        
        GetNearestComboItem = i <> CB_ERR
        Call SendMessage(cbo.HWND, &H14E, i, 0)
        
        
        If TypeName(cbo) = "ComboBox" And i <> CB_ERR Then
            cbo.SelStart = Len(sItem)
            cbo.SelLength = Len(cbo.Text) - cbo.SelStart
        End If
End Function

'populate recordset to data grid

Public Sub PopuLateFromRecordSet(cbo As ComboBox, ByRef rs As ADODB.Recordset, FieldName As String, bClear As Boolean)
On Error Resume Next
Dim BK As Variant, str As String, SField As String

    SField = cbo.DataField
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub
    
    cbo.DataField = ""
    If bClear Then cbo.Clear
    
    BK = rs.Bookmark
    
    'rs.MoveLast
    
    rs.MoveFirst
    cbo.Clear
    While Not rs.EOF
        
        str = Trim$(rs(FieldName).value)
        
        If IndexOf(cbo, str) = CB_ERR Then _
            If Len(str) <> 0 Then cbo.AddItem (str)
                                   
        rs.MoveNext
        If Err Then Err.Clear
        
        'DoEvents
    Wend
    
    cbo.DataField = SField
    rs.Bookmark = BK
End Sub


'sort data grid

Public Sub SortGrid(rs As ADODB.Recordset, Grid As DataGrid, Col As Integer)
On Error Resume Next
    Dim SortOrder As String
    Dim BK As Variant

    BK = rs.Bookmark
    SortOrder = Grid.Tag
    SortOrder = IIf(UCase(SortOrder) = "ASC", "ASC", "DESC")
    Grid.Tag = IIf(UCase(SortOrder) = "ASC", "DESC", "ASC")

    rs.Sort = ""
    rs.Sort = ((Grid.Columns(Col).DataField) + " " + SortOrder)
    rs.Bookmark = BK
    If Err Then Err.Clear
End Sub

'load imaging file to form

Public Function FileToField(ByVal fldFile As String, ByRef lSize As Long) As Variant

'On Error GoTo ErrHandler

Screen.MousePointer = vbHourglass

Dim bytData() As Byte
Dim strData As String
Dim intBlocks As Integer
Dim intBlocksLo As Integer
Dim lngSourceFileLen As Long
Dim intSourceFile As Integer
Dim intCnt As Integer
Dim lngBlockSize As Long
Dim rs As New ADODB.Recordset

    'Open BLOB file.
    intSourceFile = FreeFile
    Open fldFile For Binary Access Read As intSourceFile
    lngSourceFileLen = LOF(intSourceFile)

    'Create an empty recordset and append fields
    rs.Fields.Append "Picture", adLongVarBinary, lngSourceFileLen, adFldLong
    rs.Open
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
    End If

    lngBlockSize = 15000
    intBlocks = lngSourceFileLen \ lngBlockSize
    intBlocksLo = lngSourceFileLen Mod lngBlockSize
    
    'AppendChunk IMAGE Column
    ReDim bytData(intBlocksLo)
    Get intSourceFile, , bytData()
    rs.Fields(0).AppendChunk bytData()
    
    ReDim bytData(lngBlockSize)
    For intCnt = 1 To intBlocks
        Get intSourceFile, , bytData()
        rs.Fields(0).AppendChunk bytData()
    Next intCnt
    rs.Update
    
    GoTo Shutdown
    
ErrHandler:
    MsgBox "Error : " & Err.number & vbCrLf & Err.Description

Shutdown:
    Close intSourceFile
    lSize = rs.Fields(0).ActualSize
    FileToField = rs.Fields(0).value
    Screen.MousePointer = vbNormal
End Function

'check recordset

Public Function FilterRecords(rs As ADODB.Recordset, FilterStr As String) As ADODB.Recordset
    rs.Filter = adFilterNone
    rs.Filter = FilterStr
    Set FilterRecords = rs
End Function

'function make command object

Public Function MakeCommand(cn As ADODB.Connection, CommandType As ADODB.CommandTypeEnum) As ADODB.Command
    Set MakeCommand = Nothing
    
    Set MakeCommand = New ADODB.Command
    Set MakeCommand.ActiveConnection = cn
    MakeCommand.CommandType = CommandType
End Function

'function user MTS to set transaction start command

Public Function BeginTransaction(cn As ADODB.Connection)

    With MakeCommand(cn, adCmdText)
        .CommandText = "BEGIN TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
        
End Function

'functionuser MTS to set transaction end command

Public Function CommitTransaction(cn As ADODB.Connection)
On Error Resume Next

    With MakeCommand(cn, adCmdText)
        .CommandText = "COMMIT TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
        
    End With
        
        If Err Then Err.Clear
End Function

'function user MTS to set transaction rollback command

Public Function RollbackTransaction(cn As ADODB.Connection)
On Error Resume Next

    With MakeCommand(cn, adCmdText)
        .CommandText = "ROLLBACK TRANSACTION"
        Call .Execute(Options:=adExecuteNoRecords)
    End With
        
    If Err Then Err.Clear
End Function

'function to find item in data grid and assign values to it

Public Function FindInGrid(FindIn As Object, strFind As String, Optional Exact As Boolean = True, Optional ColumnIndex = 0) As Boolean
Dim i As Long, x As Long

    FindInGrid = -1
    If TypeOf FindIn Is SSOleDBCombo Or TypeOf FindIn Is SSOleDBGrid Then

        x = FindIn.Rows
        FindIn.MoveFirst

        For i = 0 To x
        
            If Trim$(FindIn.Columns(ColumnIndex).Text) = Trim$(strFind) Then Exit For
            FindIn.MoveNext
        Next i
                
    End If
End Function

'get form name then get data form controls

Public Sub BindAll(frm As Form, oDataSource As Object)
Dim str As String
On Error Resume Next

Dim ctl As Control

    For Each ctl In frm.Controls
        str = ctl.DataField
        
        str = str & ctl.DataMember
        
        If Len(str) Then
            Set ctl.DataSource = Nothing
            Set ctl.DataSource = oDataSource
        End If
        
        If Err Then Err.Clear
    Next ctl
            
End Sub
   
'set navbar buttom

Public Sub DisableButtons(frmname As Form, oNavbar As Object)
Dim bl As Boolean
Dim ctl As Control

On Error Resume Next

    bl = Getmenuuser(deIms.NameSpace, CurrentUser, frmname.Tag, deIms.cnIms)
    If UCase(deIms.NameSpace) = "TRNNG" Then 'M
          oNavbar.EMailVisible = False  'M
          If Err.number <> 0 Then Err.Clear
    End If
    
    If bl = True Then Exit Sub
    If Not (TypeOf oNavbar Is lrnavigators.NavBar _
    Or TypeOf oNavbar Is lrnavigators.LROleDBNavBar) Then Exit Sub

    oNavbar.NewEnabled = False
    oNavbar.SaveEnabled = False
    oNavbar.CancelEnabled = False
    oNavbar.EditEnabled = False
    
    'EmailandFax Should not be Visible if the Namespace is 'TRNNG'
      'M
  
    If Err Then Err.Clear
    For Each ctl In frmname.Controls

    'Modofied by Muzammil 03/26/01
    'Reason - It Disables the Gird and the user can not access or scroll throught it
    
    
      '  If TypeOf ctl Is SSOleDBGrid Or
      
       If TypeOf ctl Is CheckBox Or _
         TypeOf ctl Is DataGrid Or _
           TypeOf ctl Is textBOX Then

                ctl.Enabled = False

        ElseIf TypeOf ctl Is ComboBox Or _
                   TypeOf ctl Is DataCombo Then ctl.locked = True

        ElseIf TypeOf ctl Is SSOleDBCombo Then ctl.AllowInput = False
        
        ElseIf TypeOf ctl Is SSOleDBGrid Then ctl.AllowUpdate = False
        
        End If

    Next ctl

End Sub

'set seperator character

Public Sub get_seperator(ctl As Control)
    ctl.FieldSeparator = Chr$(1)
End Sub

'function to seacrh recordset

Public Function RecordsetFind(rs As ADODB.Recordset, Criteria As String) As Boolean
Dim i As Integer
Dim BK As Variant
    
    If rs.EOF And rs.BOF Then Exit Function
    If rs.RecordCount = 0 Then Exit Function
    
    BK = rs.Bookmark
    
    'rs.MoveFirst
    Call rs.Find(Criteria, 0, adSearchForward, adBookmarkFirst)
    Call rs.Find(Criteria, 0, adSearchForward, adBookmarkFirst)
    
    
    If rs.EOF Then
        rs.Bookmark = BK
    Else
        Call rs.Move(0)
        RecordsetFind = True
    End If
End Function

'bind text boxse to data source

Public Sub BindControlsToDataMenber(sDataMember As String, frm As Form)
On Error Resume Next
Dim ctl As Control
Dim str As String

    sDataMember = UCase(sDataMember)
    
    For Each ctl In frm.Controls
        str = UCase(ctl.DataMember)
        
        If Len(str) Then
            If str = sDataMember Then
                Set ctl.DataSource = Nothing
                Set ctl.DataSource = deIms
                ctl.DataMember = sDataMember
            End If
        End If
        
        If Err Then Err.Clear
    Next ctl

    If Err Then Err.Clear
End Sub

'call function to send message

Public Sub DateTime_SetFormat(HWND As Long, Format As String)
    Call SendMessageStr(HWND, &H1005, 0, Format)
End Sub

'Added by Muzammil
'Reason - To Store the UserId
Public Function StripConnectionString(connectionstring As String) As ConnInfo
'DBPassword=   (connectionstring,"password="
Dim variables() As String
Dim MainString() As String
Dim i As Integer
Dim DelimitedValues() As String
Dim ConnInfo As ConnInfo
MainString = Split(connectionstring, ";")

For i = 0 To UBound(MainString)
      
      If InStr(MainString(i), "Password=") > 0 Then
          DelimitedValues = Split(MainString(i), "=")
          ConnInfo.Pwd = Trim$(DelimitedValues(1))
          
         'DBPassword = Trim$(DelimitedValues(1)) 'M
         
      ElseIf InStr(MainString(i), "User ID=") > 0 Then
      
          DelimitedValues = Split(MainString(i), "=")
          ConnInfo.UId = Trim$(DelimitedValues(1))
          
         'UserId = Trim$(DelimitedValues(1)) M
      ElseIf InStr(MainString(i), "Initial Catalog=") > 0 Then
          DelimitedValues = Split(MainString(i), "=")
          ConnInfo.InitCatalog = Trim$(DelimitedValues(1))
          
         'InitialCatalog = Trim$(DelimitedValues(1)) M
      ElseIf InStr(MainString(i), "Data Source=") > 0 Then
          DelimitedValues = Split(MainString(i), "=")
          ConnInfo.DSource = Trim$(DelimitedValues(1))
          
         'Datasource = Trim$(DelimitedValues(1)) M
      End If
       
Next i
    
     ConnInfo.Dsnname = GetDSNNameFromFile
     
    StripConnectionString = ConnInfo
'DBPassword = Split(connectionstring, ";")
 'UserID
End Function

Public Function GetDSNNameFromFile() As String
Dim i As Integer
Dim CompletePath As String
Dim strrow As String
Dim Dsnname() As String
i = FreeFile
CompletePath = App.Path & "\ImsParam.Ims"

Open CompletePath For Input As #i
 Line Input #i, strrow
 Dsnname = Split(strrow, "=")
 Close #i
 
GetDSNNameFromFile = Dsnname(1)

End Function
'Added by Muzammil
Public Function GetNamespaceConfigurationValues(EmailClient As EmailClients, EmailOutFolder As String, EmailParameterFolder As String)

Dim RsPesys As New ADODB.Recordset
 
EmailClient = Unknown
 
 RsPesys.Source = "select psys_gateway, psys_gatewayParambskt, Psys_extendedcurcode, isnull(psys_usexport,0) psys_usexport, psys_eccnactivate, isnull(psys_eccnlength,0) psys_eccnlength from pesys where psys_npecode='" & deIms.NameSpace & "'"

 RsPesys.ActiveConnection = deIms.cnIms
 
 RsPesys.Open

 If Trim(UCase(RsPesys("psys_gateway")) & "") = "ATT" Then
 
    EmailClient = ATT
    
 ElseIf Trim(UCase(RsPesys("psys_gateway")) & "") = "OUTLOOK" Then
  
    EmailClient = Outlook
    
    EmailParameterFolder = Trim(RsPesys("psys_gatewayParambskt"))
  
    If Mid(EmailParameterFolder, Len(EmailParameterFolder), 1) <> "\" Then EmailParameterFolder = EmailParameterFolder & "\"
    
    EmailOutFolder = EmailParameterFolder & "OUT" & "\"
  
 ElseIf Trim(RsPesys("psys_gateway") & "") = 0 Or Trim(RsPesys("psys_gatewayParambskt") & "") = 0 Then
  
    MsgBox "Database is not set up for Email gateway. Please configure the database before sending emails.", vbCritical, "Imswin"
  
 End If
 
 GExtendedCurrency = Trim(RsPesys("Psys_extendedcurcode") & "")
 
 If IsNull(RsPesys("psys_usexport")) Or IsNull(RsPesys("psys_eccnactivate")) Or _
 (LCase(Trim(RsPesys("psys_eccnactivate"))) <> Constyes And LCase(Trim(RsPesys("psys_eccnactivate"))) <> Constno And _
 LCase(Trim(RsPesys("psys_eccnactivate"))) <> ConstOptional) Then
 
    MsgBox "Database is not set up for Eccn. Please configure the database before working on Eccn.", vbCritical, "Imswin"
        
 Else
 
    ConnInfo.usexport = RsPesys("psys_usexport")
    ConnInfo.Eccnactivate = RsPesys("psys_eccnactivate")
    ConnInfo.EccnLength = IIf(IsNull(RsPesys("psys_eccnlength")), 0, RsPesys("psys_eccnlength"))
    
 End If
 
 
End Function

'Added by Muzammil - 12/18/00/
'Reason - All the Master files  Should have the Three Mode (Same way as in PO)

Public Function ChangeModeOfForm(lblStatus As Label, FMode As FormMode) As FormMode
On Error Resume Next
Dim bl As Boolean

    
    If FMode = mdCreation Then
        lblStatus.ForeColor = vbRed
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00125") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Creation", msg1) 'J modified
        '---------------------------------------------
        
    ElseIf FMode = mdModification Then
        lblStatus.ForeColor = vbBlue
                
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00126") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Modification", msg1) 'J modified
        '---------------------------------------------
  
     ElseIf FMode = mdvisualization Then
        lblStatus.ForeColor = vbGreen
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("L00092") 'J added
        lblStatus.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
        '---------------------------------------------
    
    End If
    
       
    ChangeModeOfForm = FMode
    
    

End Function
Function FirstOfMonth() As String
    FirstOfMonth = Format(Now, "yyyy") + "/" + Format(Now, "mm") + "/1"
End Function
Public Function GetValuesFromControls(Controlnames() As String, Optional currentformname As Form) As String()
Dim n(7) As String
Dim j As Integer
n(2) = deIms.NameSpace

    
  If InStr(Controlnames(1), "buffy") Then
    
        n(1) = FrmRequisition.SSGridHeaderDetails.Columns(0).Text

  ElseIf InStr(Controlnames(1), "summer") Then
        
        n(1) = frmInvoice.cell(0)

  ElseIf InStr(Controlnames(1), "polist") Then
             
        n(1) = frmPackingList.POlist.TextMatrix(frmPackingList.POlist.row, 1)
        
  Else
        n(1) = IIf(Controlnames(1) = "", "", currentformname.Controls(Controlnames(1))) 'VBPrimaryKey3
        
  End If
  
     
If n(3) = "" Then
   
   For j = 3 To 7
  
    If Controlnames(j) <> "" Then

    n(j) = VB.Forms(3).Controls(Controlnames(j))    'VBPrimaryKey3
    j = j + 1
    
    End If

   Next

End If

GetValuesFromControls = n

End Function


'commented out by Jawdat 2.27.02
'
'Public Function ROWGUIDVALUE(lock_date As Date)
'
'
'Dim SQLstring4 As String
'
'        SQLstring4 = "select rowguid from LOCK where DateOpened = '" & lock_date & "'"
'
'    Set Locked_data = New ADODB.Recordset
'    Locked_data.Open SQLstring4, Connection
'
'    rowguid = Locked_data("rowguid")
'
'locked = True
'End Function


'Public Function IsArrayLoaded(ArrayToTest() As String, Optional DIMENSION As Integer) As Boolean
Public Function IsArrayLoaded(ArrayToTest() As String) As Boolean

Dim x As Integer

On Error GoTo ErrHandler

IsArrayLoaded = False

'If DIMENSION = 0 Then

    x = UBound(ArrayToTest)
'Else
'x = UBound(ArrayToTest, DIMENSION)
'End If

IsArrayLoaded = True

Exit Function

ErrHandler:

Err.Clear

End Function

Public Function CreateFolder(FolderPath As String)


End Function

Public Function FixTheFirstCarriageReturn(Text As String) As String

Text = Trim(Text)
     
Do While InStr(1, Text, vbCrLf) = 1
     
If InStr(1, Text, vbCrLf) = 1 Then Text = Mid$(Text, 3, Len(Text))

Loop

FixTheFirstCarriageReturn = Text

 

End Function

Public Sub ExportToExcel(Optional RsRecord As ADODB.Recordset, Optional Arr As Variant, Optional ArrColumnNames As Variant, Optional Progressbar As Progressbar, Optional Filename As String)

Dim Report As Excel.Application
Dim i As Integer
Dim j As Integer
Dim Sa As Scripting.FileSystemObject
Dim Fld As ADODB.Field
Dim x As Integer
Dim y As Integer
Dim Incr As Integer

    Set Report = New Excel.Application
    Set Sa = New Scripting.FileSystemObject
    
    i = Rnd(20)
    
    If Sa.FileExists(App.Path & "\" & Filename & i & ".xls") = False Then Sa.CreateTextFile App.Path + "\" & Filename & i & ".xls"
    
    Report.Workbooks.Open App.Path + "\" + Filename & i & ".xls", , , , "asweetkiss"
    
    
    With Report
       
       .WindowState = xlMinimized

    If IsNothing(RsRecord) Then
    
    'This is executed when an array is passed in here.
    
        x = UBound(ArrColumnNames)
        

        
                  For i = 0 To x
               
                   .Cells(1, i + 1) = ArrColumnNames(i)
                     
                    
                  Next i
                  
        .activeCELL.EntireRow.Font.Bold = True

        'MsgBox .ActiveCell.Table(
    
    
        x = UBound(Arr, 1)
        
        y = UBound(Arr, 2)
        
       Incr = y / 10
        
            For j = 0 To y
                      
              For i = 0 To x
               
                  .Cells(j + 2, i + 1) = Arr(i, j)
                       
              Next i
              
                If j > 0 And j Mod Incr = 0 Then
                 Call IncrementProgreesBar(1, Progressbar)
                End If
                    
            Next j
             
             
     ElseIf IsNothing(RsRecord) = False Then
        
     RsRecord.MoveFirst
    'This is executed when a recordset is passed
        
        i = 1
        j = 1
        
        
        'Writing the names of the Fields
        
            For Each Fld In RsRecord.Fields
                      
                    .Cells(i, j) = Fld.Name & ""
                     
                     j = j + 1
                    
                    
             Next Fld
        
        i = i + 1
        
        Do While Not RsRecord.EOF
        
            j = 1
        
             For Each Fld In RsRecord.Fields
             
                    .Cells(i, j) = Fld.value & ""
                     
                     j = j + 1
             
             Next Fld
            
            i = i + 1
            RsRecord.MoveNext
        
        Loop
        
   End If
            
        Report.Visible = True
        
    End With
    
    SetProgressBarToMax Progressbar
    
'Screen.MousePointer = vbArrow

End Sub


Public Function IncrementProgreesBar(value As Integer, Progressbar As Progressbar)

Progressbar.Max = 10

If Progressbar.Visible = False Then Progressbar.Visible = True

If Progressbar.value < Progressbar.Max Then

        Progressbar.value = Progressbar.value + value
        
ElseIf Progressbar.value = Progressbar.Max Or Progressbar.value > Progressbar.Max Then

        Progressbar.value = 1
        
End If

End Function

Public Sub SetProgressBarToMax(Progressbar As Progressbar)
Dim i As Integer

For i = Progressbar.value To Progressbar.Max - 1

    Progressbar.value = Progressbar.value + 1
    
Next
    
Progressbar.value = 0



'Progressbar.Visible = False

End Sub

Public Function CheckifFqaExist(FQA As String, Level As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo ErrHand
Level = UCase(Level)

rs.Source = "select count(*) countit from fqa where fqa ='" & FQA & "' and namespace ='" & deIms.NameSpace & "' and level ='" & Level & "'"
rs.Open , deIms.cnIms

If rs("countit").value > 0 Then
    CheckifFqaExist = True
End If
Set rs = Nothing

Exit Function
ErrHand:

MsgBox "Errors occurred while trying to access the FQA table." & Err.desc, vbCritical, "Ims"
Err.Clear

End Function

''Function GetStockOnHandWithSublocation()
''
''Dim Rs As New ADODB.Recordset
''deIms.cnIms.Open
''
''Rs.source = " SELECT"
''Rs.source = Rs.source & "     qs5_compcode,loc_name ,sublocation ,qs5_stcknumb,"
''Rs.source = Rs.source & " qs5_cond, sum(qs5_primqty) QTY, ( select sap_value from sap where sap_compcode =qs5_compcode and sap_loca=qs5_ware"
''Rs.source = Rs.source & " and sap_stcknumb =qs5_stcknumb and"
''Rs.source = Rs.source & " sap_cond =qs5_cond ) 'Unit Price' ,"
''Rs.source = Rs.source & " qs1_desc , curr_desc"
''Rs.source = Rs.source & " From"
''Rs.source = Rs.source & " PECTEN.dbo.StockOnHandmodified StockOnHandmodified"
''Rs.source = Rs.source & " WHERE"
''Rs.source = Rs.source & " curr_code = 'USD' AND"
''Rs.source = Rs.source & " qs5_npecode = 'PECT'"
''Rs.source = Rs.source & " and qs5_ware in ('drl','chm','dch','m&t','prd','Sur')"
''Rs.source = Rs.source & " Group By"
''Rs.source = Rs.source & " qs5_compcode , qs5_ware, loc_name, sublocation, qs5_stcknumb, qs5_cond, qs1_desc, curr_desc"
''
''Rs.Open , deIms.cnIms, 3, 3
''Call ExportToExcel(Rs, , , ProgressBar1, "StockDetails")
''
''End Function
Public Function IsNamespaceEccnActivated() As Boolean
On Error GoTo ErrHand

If ConnInfo.Eccnactivate = Constno Then IsNamespaceEccnActivated = False

ErrHand:
Exit Function

MsgBox Err.Description

Err.Clear
End Function

