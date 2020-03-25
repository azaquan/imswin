Attribute VB_Name = "mIKE"

Option Explicit

'MM : March 01 2009
Public Const Report_EmailFax_PO_name = "poforEmailFaxCR11.rpt"
'Public Const Report_EmailFax_PO_name = "po.rpt"
Public Const Report_EmailFax_FreightReceipt_name = "freception.rpt"
Public Const Report_EmailFax_TrackingPo = "obs.rpt"
Public Const Report_EmailFax_Supplier_name = "supplier.rpt"
Public Const Report_EmailFax_Stockmaster_name = "Stckmaster1ForEmailFaxCR11.rpt"
'Public Const Report_EmailFax_Stockmaster_name = "Stckmaster1.rpt"
Public Const Report_EmailFax_PackingManifest_name = "packinglistforEmailFaxCR11.rpt"

Public Enum EReportTypesForCR11
    PO = 0
    FreightRececipt
End Enum

Public Type RPTIFileInfo
    Login As String
    Password As String
    parameters() As String
    ReportFileName As String
End Type

Public Enum EmailClients
   Unknown = 0
   ATT
   Outlook
End Enum

Public Type ConnInfo
    
    UId As String
    Pwd As String
    InitCatalog As String
    DSource As String
    Dsnname As String
    EmailClient As EmailClients   'M 02/15/02
    EmailOutFolder As String 'M 02/26/02
    EmailParameterFolder As String 'M 02/26/02
    usexport As Boolean
    Eccnactivate As String * 1
    EccnLength As Integer
End Type

Public Const GW_OWNER = 4
Public Const GWL_STYLE = (-16)
Public Const GWL_HWNDPARENT = (-8)
Public Const WS_CHILD = &H40000000
'Public Const DBPassword = "saa"  'M
'Public DBPassword As String  'M
'Public Datasource As String  'M
'Public InitialCatalog As String  'M
'Added by Muzammil
'Reason - To Store the UserId
'Public UserId As String  'M
Public ConnInfo As ConnInfo

Public Language As String 'M
'Public MutexHandle As Long
Public reportPath As String
Public LogPath As String
Public CurrentUser As String
Public Const WAIT_TIMEOUT = &H102&
Public Const ERROR_ALIAS_EXISTS = 1379&
Public Const ERROR_ALREADY_EXISTS = 183&
Public Declare Sub DebugBreak Lib "kernel32" ()
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Global m_OutlookLocation As String
Global GExtendedCurrency As String

Public Const Constyes = "y"
Public Const Constno = "n"
Public Const ConstOptional = "o"

Dim myASPx As myASPx.Exec 'JCG 2008/7/6

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private PDFCreator1 As PDFCreator.clsPDFCreator 'JCG 2008/7/10
Private pErr As clsPDFCreatorError, opt As clsPDFCreatorOptions 'JCG 2008/7/10
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'JCG 2008/712
''Function and procedures created by Mike Lavery


Public Function Print_Crystal_Rpt(rpt_filename As String) As Integer
'Created By Mike Lavery
 On Error Resume Next
 
  Dim old_ptr As Integer
   
   
  old_ptr = Screen.ActiveForm.MousePointer
  Screen.ActiveForm.MousePointer = vbHourglass
   
  Call LaunchApp(rpt_filename, "", 1, True)
  Screen.ActiveForm.MousePointer = old_ptr
  
  If Err Then MsgBox Err.Description: Err.Clear
End Function

'set form text boxse back ground color

Public Function HighlightBackground(ThisObject As Control) As Integer
On Error Resume Next
    Dim ErrorCode As Integer

    ErrorCode = 0

    On Error GoTo errorHandler
    ThisObject.BackColor = &HC0FFFF

    HighlightBackground = ErrorCode
    Exit Function

errorHandler:
    ErrorCode = Err
    Resume Next
End Function

'set combobox, checkbox, back ground color

Public Function NormalBackground(ThisObject As Control) As Integer
On Error Resume Next
    Dim ErrorCode As Integer

    ErrorCode = 0

    On Error GoTo errorHandler

    If TypeOf ThisObject Is ComboBox Or TypeOf ThisObject Is textBOX Then
        ThisObject.BackColor = frm_Color.txt_textbox.BackColor


    ElseIf TypeOf ThisObject Is checkBox Or TypeOf ThisObject Is OptionButton Then
        ThisObject.BackColor = frm_Color.txt_WBackground.BackColor

    Else
        ThisObject.BackColor = frm_Color.txt_textbox.BackColor


    End If
    
    NormalBackground = ErrorCode
    Exit Function
errorHandler:
    ErrorCode = 0
    Resume Next
End Function

'set controls back ground colors

Public Sub gsb_fade_to_black(ctl_curr As Control)
On Error Resume Next
    'code is originally from the book "The Art of Programming with Visual Basic"
    'by Mark Warhol
    'Published by John Wiley & Sons Inc.
    'Copyright 1995

    'Heavily modified and changed by Michael S. Lavery
    'IEEE Affiliate member #40356121
    'for IMS Inc.
    'Danbury,CT 06811
    'May 20, 1999

    Dim TB As Control
    Dim BG As Control

    Set BG = frm_Color.txt_WBackground
    Set TB = frm_Color.txt_textbox

    If (TypeOf ctl_curr Is textBOX) Then
        ctl_curr.BackColor = TB.BackColor
    End If

    If TypeOf ctl_curr Is Label Then

        Select Case ctl_curr.Name

            Case "lbl_Namespace", "lbl_Username", "lbl_Password", "lbl_IMS"

                ctl_curr.BackColor = frm_Color.txt_Background.BackColor

            Case Else
                ctl_curr.BackColor = BG.BackColor

        End Select
    End If

    If TypeOf ctl_curr Is SSOleDBGrid Then
        ctl_curr.BackColor = BG.BackColor
        ctl_curr.BevelColorFace = BG.BackColor
    End If

    If TypeOf ctl_curr Is Frame Then
        ctl_curr.BackColor = BG.BackColor
    End If

    If TypeOf ctl_curr Is SSTab Then
        ctl_curr.BackColor = BG.BackColor
    End If

    If TypeOf ctl_curr Is ComboBox Then
        ctl_curr.BackColor = TB.BackColor
    End If

    If TypeOf ctl_curr Is ListBox Then
        ctl_curr.BackColor = TB.BackColor
    End If

    If TypeOf ctl_curr Is CommandButton Then
        ctl_curr.BackColor = BG.BackColor
    End If

    If TypeOf ctl_curr Is OptionButton Then
        ctl_curr.BackColor = BG.BackColor
    End If

     If TypeOf ctl_curr Is checkBox Then
        ctl_curr.BackColor = BG.BackColor
    End If
   Set TB = Nothing
    Set BG = Nothing
End Sub
'end functions created by Mike Lavery

'load form set crystal report path and check window execute status

Public Sub Main()
On Error Resume Next

    If App.PrevInstance Then _
        MsgBox "Ims for window is already running": Exit Sub ' _
            Call ReleaseMutex(MutexHandle): Call CloseHandle(MutexHandle): Exit Sub

    
    
    frm_Load.Show
    Set frm_Load = Nothing
    reportPath = FixDir(App.Path) & "CRreports\"
    
    If Err Then
        Call LogErr("Main", Err.Description, Err, True)
    End If
End Sub

'call store procedure to get po revision number
'Modified By Muzammil 01/08/00
'Reason -Added a function by the same name when Created a new PO
'''''Public Sub InsertPoRevision(Ponumb As String)
'''''Dim cmd As ADODB.Command
'''''
'''''    Set cmd = MakeCommand(deIms.cnIms, adCmdStoredProc)
'''''
'''''    cmd.Prepared = True
'''''    cmd.CommandText = "InsertPoRevision"
'''''    cmd.Execute , Array(deIms.NameSpace, Ponumb)
'''''End Sub

'set fax number format

Function adjust_address(ByVal address As String, ByVal for_fax As Boolean)
    address = Trim(address)
    If Len(address) <= 0 Then Exit Function
    If for_fax Then
        If LCase(Left(address, 5)) = "fax!+" Then
            address = Mid(address, 6)
        ElseIf LCase(Left(address, 4)) = "fax!" Then
            address = Mid(address, 5)
        ElseIf LCase(Left(address, 4)) = "fax:" Then
            address = Mid(address, 5)
        End If
        
        If IsNumeric(Left(address, 1)) Then
            address = "fax!" & address
            address = Replace(address, "/DELIVERY", "/RECEIPT", , , vbTextCompare)
            address = Replace(address, "(", "(/", , , vbTextCompare)
            address = Replace(address, " (/", "(/", , , vbTextCompare)
            address = Replace(address, "//", "/", , , vbTextCompare)
            address = Replace(address, " /", "/", , , vbTextCompare)
        Else
            address = ""
        End If
        
        If (Len(address) > 0) And Not (InStr(address, "(/") > 0) Then
            address = address & "(/no name)"
        End If
    
        If (Len(address) > 0) And Not (InStr(address, "/REPORT") > 0) Then
            address = Replace(address, "(/", "/REPORT(/", , , vbTextCompare)
        End If
    Else
       If LCase(Left(address, 9)) = "internet!" Then
            address = Mid(address, 10)
       ElseIf LCase(Left(address, 4)) = "mhs!" Or InStr(address, "@") Then
            address = address
       Else
            Dim lcl_copy As String: lcl_copy = adjust_address(address, True)
            If Len(lcl_copy) > 0 Then address = "" ' fax is not an e-mail
       End If
    End If
    
    adjust_address = address
End Function

'count input fax addresses

Function filterAddresses(ByRef inp() As String, ByVal for_fax As Boolean) As String()
    
    Dim addr_count As Integer
    
    addr_count = 0
    
    Dim outp() As String
    
    'ReDim outp(UBound(inp) - LBound(inp) + 1) 'M 02/26/02
    
    Dim address
    
    For Each address In inp
        
        Dim adjusted_address As String:  adjusted_address = adjust_address(address, for_fax)
        
        If Len(Trim(adjusted_address)) > 0 Then
              
            ReDim Preserve outp(addr_count) 'M 02/26/02
              
            outp(addr_count) = adjusted_address
            
            addr_count = addr_count + 1
            
        End If
        
    Next address
    
    'ReDim Preserve outp(addr_count) 'M 02/26/02
    
    filterAddresses = outp
    
End Function

Private Sub PDFCreator1_eError()
 Set pErr = PDFCreator1.cError
 Screen.MousePointer = vbNormal
End Sub
'function for send fax, get application path
Sub sendFaxOnly(subject As String, faxAddresses() As String, Attachment As String)
    Dim IFile As IMSFile
    Dim fo As att_FaxOptions
    Dim Attachments(0) As String, str As String
    
    On Error Resume Next

    Set IFile = New IMSFile
    
    If Attachment = "" Then
        Attachments(0) = FixDir(App.Path) & "Report.rpti"
    Else
        Attachments(0) = Attachment
    End If
    
    If Not IFile.FileExists(Attachments(0)) Then
        MsgBox "Error Preparing Fax Message"
        Exit Sub
    End If
    
    fo = att_FAX_FINE Or att_FAX_XETIFFEMAIL Or att_PROPORTIONAL_WIDTH
    Call SendAttFax("", "", subject, faxAddresses, fo, Attachments, Portrait)
    
    Set IFile = Nothing
    
    If Err Then Err.Clear
End Sub

'function for send email and get application path

Sub sendEmailOnly(subject As String, emailAddresses() As String, Attachment As String)
    Dim IFile As IMSFile
    Dim Attachments(0) As String, str As String

    On Error Resume Next

    Set IFile = New IMSFile
    
    Attachments(0) = FixDir(App.Path) & "Report.rtf"
    
    If Not FileExists(Attachments(0)) Then MDI_IMS.SaveReport Attachments(0), crptRTF
    
    If Not IFile.FileExists(Attachments(0)) Then
        MsgBox "Error Preparing Electronic Message"
        Exit Sub
    End If
    
    Dim converter As Object: Set converter = CreateObject("Imstools.rtftotxt")
    If converter Is Nothing Then
        MsgBox "The component IMSTOOLS was not properly registered. Contact IMS support."
        Exit Sub
    End If
    
    converter.sourceFile = Attachments(0)
        
    Dim nullAttachments() As String
    Call SendAttMail(converter.asText, subject, emailAddresses, nullAttachments)
    
    If IFile.FileExists(converter.sourceFile) Then IFile.DeleteFile (converter.sourceFile)
    
    Set IFile = Nothing
    
    If Err Then Err.Clear
End Sub

'set parameters for store procedure and send faxes, and emails

Public Sub SendEmailAndFax(Recipients As ADODB.Recordset, FieldName As String, _
                           subject As String, Message As String, Attachment As String, _
                           Optional Orientation As OrientationConstants)
    Dim address() As String
    Dim str As String
    Dim i As Integer

    On Error Resume Next

    address = ToArray(Recipients, FieldName, i, str)
    
    Dim faxAddresses() As String: faxAddresses = filterAddresses(address, True)
    
    If IsArrayLoaded(faxAddresses) Then
    
        'If UBound(faxAddresses) > 0 Then
            Call sendFaxOnly(subject, faxAddresses, Attachment)
        'End If
    
    End If
    
    Dim emailAddresses() As String: emailAddresses = filterAddresses(address, False)
    
    If IsArrayLoaded(emailAddresses) Then
    
        'If UBound(emailAddresses) > 0 Then
            Call sendEmailOnly(subject, emailAddresses, Attachment)
        'End If
        
    End If
    
    Kill Attachment

    If Not IsLoaded("MDI_IMS") Then End
    MDI_IMS.CrystalReport1.Reset
    
    If Err Then Err.Clear
End Sub

'call function to check file exists or not

Public Function GetFileContents(Filename As String) As String
Dim i As Integer
    
    If FileExists(Filename) Then
    
        i = FreeFile
        Open Filename For Input As #i
        GetFileContents = Input(LOF(i), i)
        
        Close #i
    End If
    
End Function

'function to check recepient number exist or not

Public Function IsInList(RecepientName As String, FieldName As String, rsReceptList As ADODB.Recordset) As Boolean
On Error Resume Next
Dim BK As Variant
    
    IsInList = False
    If rsReceptList.RecordCount = 0 Then Exit Function
    If Not (rsReceptList.EOF Or rsReceptList.BOF) Then BK = rsReceptList.Bookmark
    
    rsReceptList.MoveFirst
    Call rsReceptList.Find(FieldName & " = '" & RecepientName & "'", 0, adSearchForward, adBookmarkFirst)
    
    If Not (rsReceptList.EOF) Then
        
        IsInList = True
        MsgBox "Address Already exist in list"
    End If
    
    rsReceptList.Bookmark = BK
    If Err Then Err.Clear
End Function

'set store procedure parameters and call it to get OBS reciptients

Public Function GetObsRecipients(NameSpace As String, poNumber As String, MsgNumber As String) As ADODB.Recordset
Dim cmd As ADODB.Command
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdStoredProc)
    
    With cmd
        .CommandText = "GETOBSRECIPIENTS"
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 5, NameSpace)
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 15, poNumber)
        .parameters.Append .CreateParameter("", adVarChar, adParamInput, 15, MsgNumber)
        
        Set GetObsRecipients = .Execute
        Call GetObsRecipients.Close
        Call GetObsRecipients.Open(, , adOpenStatic, adLockReadOnly)
    End With

End Function

'function for set up attachment

Public Sub SendEmailAndFaxWithAttachments(Recipients As ADODB.Recordset, FieldName As String, subject As String, Message As String, Attachments() As String, Optional UseReport As Boolean = True)

Dim str As String
Dim IFile As IMSFile
Dim address() As String
Dim fo As att_FaxOptions
Dim i As Integer, l As Integer
On Error Resume Next

    Set IFile = New IMSFile
    
    i = UBound(Attachments) + 1
    If Err And i = 0 Then If Not UseReport Then Exit Sub
    
    If UseReport Then
        ReDim Preserve Attachments(i)
        
        Attachments(i) = FixDir(App.Path) & "Report.doc"
        
        Call MDI_IMS.SaveReport(Attachments(i), crptWinWord)
        
        If Not IFile.FileExists(Attachments(i)) Then
            MsgBox "Error Sending Electronic Message": Exit Sub
        End If
    End If
    
    DoEvents
    address = ToArray(Recipients, FieldName, l, str)
    
    If i >= 0 Then
        fo = att_FAX_FINE Or att_FAX_XETIFFEMAIL Or att_PROPORTIONAL_WIDTH
        Call SendAttFax("", "", subject, address, fo, Attachments, Portrait)
    End If
    
    On Error Resume Next
    
    For l = 0 To i
        DoEvents: DoEvents
        If IFile.FileExists(Attachments(l)) Then Call IFile.DeleteFile(Attachments(l))
    Next
    
    Set IFile = Nothing
    If Not IsLoaded("MDI_IMS") Then End
    
End Sub

'call function to get fax numbers

Private Function GetFaxNumbers(Recipients As ADODB.Recordset, FieldName As String, Upper As Integer) As String()
Dim im As imsmisc
Dim str As String, i As Long
    
    Set im = New imsmisc
    str = FieldName & " like 'FAX!%'"
    GetFaxNumbers = im.ToArray(Recipients, FieldName, Upper, str)
    
    Set im = Nothing
End Function

'function check error message and record error message to log file

Public Sub LogErr(RoutineName As String, ErrorDescription As String, ErrorNumber As Long, Optional Clear As Boolean = False)

'Dim str As String
Dim i As IMSFile
Dim ms As imsmisc

Dim Filename As String
Dim FileNumb As Integer
On Error Resume Next

        
    If Len(Trim$(ErrorDescription)) = 0 Then Exit Sub
    
    
    Set i = New IMSFile
    Set ms = New imsmisc
    
    If Not i.DirectoryExists(LogPath) Then Call MkDir(LogPath)
    
    
    FileNumb = FreeFile
    Filename = LogPath + i.ChangeFileExt(App.EXEName + Format$(Date, "ddmmyy"), "imserrlog")
    
    Open Filename For Append As 1
    
        Print #FileNumb, "Module:             " & App.EXEName
        Print #FileNumb, "Routine:            " & RoutineName
        Print #FileNumb, "Error Number:       " & ErrorNumber
        Print #FileNumb, "Error Source:       " & Err.Source
        Print #FileNumb, "Error Description:  " & ErrorDescription
        Print #FileNumb, "Error Date:         " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        
        Print #FileNumb, "": Print #FileNumb, ""
    Close #FileNumb
    
    Set i = Nothing
    Set ms = Nothing
    If Err Then Err.Clear
End Sub

Function sendFlatEmail()

    
End Function

Public Sub sendProcess(RecipientList As String, Attachments As String, subject As String, messageText As String) 'MM 022409
'Save the Email/ request to the Database

On Error GoTo errorHandler
    Dim strOut As String
    Dim programName As String
    Dim parameters As String
    Dim cmd As ADODB.Command
    
     Set cmd = MakeCommand(deIms.cnIms, ADODB.CommandTypeEnum.adCmdStoredProc)
            
   
    With cmd
        .CommandText = "InsertEmailFax"
        .parameters.Append .CreateParameter("@Subject", adVarChar, adParamInput, 4000, subject)
        .parameters.Append .CreateParameter("@Body", adVarChar, adParamInput, 8000, messageText)
        .parameters.Append .CreateParameter("@AttachmentFile", adVarChar, adParamInput, 2000, Attachments)
        
        
        .parameters.Append .CreateParameter("@recepientStr", adVarChar, adParamInput, 8000, RecipientList)
        .parameters.Append .CreateParameter("@creauser", adVarChar, adParamInput, 100, CurrentUser)
        
        Call .Execute(Options:=adExecuteNoRecords)
    
               
    End With
    
    
    Set cmd = Nothing
    
    LogExec ("Successfully saved email\ Fax request with Subject " & subject & " to the Database.")
    
Exit Sub

errorHandler:
    Call LogErr("sendProcess", "Érror Occured while trying to save Email\ Fax request to the DB for Subject " + subject + " Body " + messageText + " Attachment " + Attachments + " Recepient List " + RecipientList + ". " + Err.Description, Err.number, False)
    MsgBox "Errors Occured while trying to generate email\ Fax request. Please dont send any more emails and faxes and call the Administrator. " + Err.Description
    Err.Clear
    
End Sub


'çommented by MM
'Public Sub sendProcess(RecipientList As String, Attachments As String, subject As String, messageText As String) 'JCG 2008/7/6
'On Error GoTo errorHandler
'    Dim strOut As String
'    Dim programName As String
'    Dim parameters As String
'    Set myASPx = CreateObject("myASPx.Exec")
'
'    Sleep 4000
'
'    programName = App.Path + "\sendEmail\sendEmail.exe"
'    If Trim(Attachments) = "" Then
'        parameters = "-o timeout=1 -f pecten@groupgls.com -t " + RecipientList + " -xu pecten@groupgls.com -xp pectendla4312 -s smtpout.secureserver.net:80 -m " + messageText + " -u " + subject
'    Else
''MsgBox "@subject->" + subject
''MsgBox "@messageText->" + messageText
''MsgBox "@RecipientList->" + RecipientList
'        parameters = "-o timeout=1 -f pecten@groupgls.com -t " + RecipientList + " -xu pecten@groupgls.com -xp pectendla4312 -s smtpout.secureserver.net:80 -o -a " + Attachments + " -m " + messageText + " -u " + subject
'    End If
'
'  ' set the parameters
'    myASPx.AppName = programName
'    myASPx.Params = parameters
'
'  ' executing and waiting for the value
'    strOut = myASPx.DosExec
'
'    If Len(strOut) > 0 Then
'        If InStr(1, strOut, "Email was sent successfully!") > 0 Then
'              'MsgBox "Messages sent succesfully"
'        Else
'            MsgBox "There is an issue when sending the messages. " + strOut
'        End If
'    End If
'
'Exit Sub
'
'errorHandler:
'    MsgBox "Process sendProcess " + Err.Description
'    Err.Clear
'End Sub


'get warehouse recipients recordset for email function

Public Sub SendWareHouseMessage(NameSpace As String, _
                                Comment As String, _
                                subject As String, _
                                cn As ADODB.Connection, _
                                ReportInfo As RPTIFileInfo)
On Error Resume Next
Dim fo As att_FaxOptions
Dim rs As ADODB.Recordset
Dim FaxNumbers() As String
Dim EmailAddress() As String
Dim i As Integer, l As Integer
    
    'Added by Muzammil to Stop Trnng from sending Email or Fax.
  If UCase(deIms.NameSpace) <> "TRNNG" Then   'M
    
    
            ReDim FaxNumbers(0)
            ReDim EmailAddress(0)
            Set rs = GetWareHouseRecipients(NameSpace, cn)
            
            DoEvents
            Do Until rs.EOF
            
                DoEvents
                Do While ((Len(Trim$(rs("Address") & "")) = 0))
                    rs.MoveNext
                    DoEvents: DoEvents
                    If rs.EOF Then Exit Do
                Loop
                
                If rs.EOF Then Exit Do
                
                If LCase(rs("Type")) = "fax" Then
                
                    ReDim Preserve FaxNumbers(i)
                    
                    FaxNumbers(i) = rs("Address") & ""
                    
                    If (LCase(Left$(FaxNumbers(i), 4)) <> "fax!") Then
                        FaxNumbers(i) = "FAX!" & FaxNumbers(i)
                    End If
                    
                    i = i + 1
                Else
                    ReDim Preserve EmailAddress(l)
                    EmailAddress(i) = rs("Address") & ""
                    l = l + 1
                    
                End If
                
                rs.MoveNext
                DoEvents: DoEvents
            Loop
            
            If ((i = 0) And (l = 0)) Then _
                If ((Len(Trim$(EmailAddress(0))) = 0) And _
                    (Len(Trim$(FaxNumbers(0))) = 0)) Then Exit Sub
        
           
              
        
            Dim IFile As IMSFile, FileNames(0) As String
            
            Set IFile = New IMSFile
            
            Call WriteRPTIFile(ReportInfo, FileNames(0))
            If Not IFile.FileExists(FileNames(0)) Then Exit Sub
            
            fo = att_FAX_FINE Or att_FAX_XETIFFEMAIL Or att_PROPORTIONAL_WIDTH
            Call SendAttFax(Comment, "", subject, FaxNumbers, fo, FileNames, Portrait)
            Call SendAttFax(Comment, "", subject, EmailAddress, fo, FileNames, Portrait)
            
            
            Call IFile.DeleteFile(FileNames(0))
            
            Set IFile = Nothing
            
            If Not IsLoaded("MDI_IMS") Then End
            If Err Then MsgBox Err.Description: Err.Clear
            
  End If
            
End Sub

'call store procedure to get warehouse distribution recordset

Public Function GetWareHouseRecipients(NameSpace As String, _
                                      cn As ADODB.Connection) As ADODB.Recordset
Dim cmd As ADODB.Command


    Set cmd = New ADODB.Command
    
    With cmd
        Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "GetWareHouseDistribution"
        
        .parameters.Append .CreateParameter("NP", adVarChar, adParamInput, 5, NameSpace)
        Set GetWareHouseRecipients = .Execute
    End With
    
End Function

'set form
Public Function open_forms() As Integer
  Dim count_forms As Integer
  Dim frm As Form
  Dim str As String
  

    open_forms = 1 'Forms.Count
    For Each frm In Forms
    
        If frm.Visible = False Then
        
            If InStr(1, frm.Name, "Navigator", vbTextCompare) = 0 Then
            
                If frm.Name <> "frm_Color" Then
                
                    Unload frm
                    Set frm = Nothing
                    
                End If
                
            End If
            
        End If
        
    Next
 
End Function

'show navigator form

Public Sub ShowNavigator()
On Error Resume Next

    If frmNavigator.Visible = False Then
        frmNavigator.Show
        frmNavigator.SetFocus
    End If
    
    If Err Then Err.Clear
End Sub

'function to strip characters

Public Function StripNullChar(Source As String) As String
    StripNullChar = Left(Source, (InStr(1, Source, vbNullChar, vbTextCompare) - 1))
End Function

'function record execution error to log file

Public Sub LogExec(Message As String)

Dim str As String
Dim i As IMSFile
Dim ms As imsmisc
Dim FileNumb As Integer
Dim Filename As String
On Error Resume Next

        
    If Len(Trim$(Message)) = 0 Then Exit Sub
    
    
    

    Set i = New IMSFile
    Set ms = New imsmisc
    
    
    'If Not i.DirectoryExists(LogPath) Then Call MkDir(LogPath)
    
    'str = fixdir(str) & "ims\"
    If Not i.DirectoryExists(LogPath) Then Call MkDir(LogPath)
    
    
    FileNumb = FreeFile
    Filename = LogPath + i.ChangeFileExt(App.EXEName + Format$(Date, "ddmmyy"), "imsexeclog")
    
    Call SetAttr(Filename, vbNormal)
    Open Filename For Append As FileNumb
    
        Print #FileNumb, "Module:            " & App.EXEName
        Print #FileNumb, "Description:       " & Message
        Print #FileNumb, "Error Date:        " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        
        Print #FileNumb, "": Print #FileNumb, ""
    Close #FileNumb
    
    Set i = Nothing
    Set ms = Nothing
    Call SetAttr(Filename, vbReadOnly)
    
    If Err Then Err.Clear
End Sub

'check file exists or not, if file exist delete it
'then create new file

Public Function CreateFile(Filename As String) As Integer

Dim i As IMSFile

    Set i = New IMSFile
    
    If i.FileExists(Filename) Then Call i.DeleteFile(Filename)
    
    CreateFile = FreeFile
    Open Filename For Append As CreateFile
End Function

'print file

Public Sub WriteToFile(str As String, FileNumber As Integer)
    Print #FileNumber, str
End Sub

'close file

Public Sub CloseFile(FileNumber As Integer)
    Close #FileNumber
End Sub

'get file information

Public Sub WriteRPTIFile(FileInfo As RPTIFileInfo, Optional Filename As String)
Dim FileNumb As Integer
Dim i As Integer, l As Integer

    FileNumb = FreeFile
    Filename = Trim$(Filename)
    l = UBound(FileInfo.parameters)
    FileInfo.Login = Trim$(FileInfo.Login)
    FileInfo.Password = Trim$(FileInfo.Password)
    If Len(Filename) = 0 Then Filename = FixDir(App.Path) & "Report.rpti"
    
    
'    If Len(FileInfo.Login) = 0 Then FileInfo.Login = "sa"  'M
    If Len(FileInfo.Login) = 0 Then FileInfo.Login = ConnInfo.UId 'userid 'M
    If Len(FileInfo.Password) = 0 Then FileInfo.Password = ConnInfo.Pwd 'DBPassword 'M
    
    Open Filename For Output As FileNumb
    
        Print #FileNumb, "[report]"
        Print #FileNumb, "file=" & Trim$(FileInfo.ReportFileName)
        
        Print #FileNumb, ""
        Print #FileNumb, "[data]"
        Print #FileNumb, "login=" & Trim$(FileInfo.Login)
        Print #FileNumb, "password=" & Trim$(FileInfo.Password)
        
        Print #FileNumb, ""
        Print #FileNumb, "[params]"
        
        For i = 0 To l
            Print #FileNumb, Trim$(FileInfo.parameters(i))
        Next
        
    Close #FileNumb
        
End Sub
'DidFieldChange(Trim(oldValue), Trim(ssdbgKeys.Columns(ColIndex).text))

'Added by muzammil to run the App as directed by Francois.
Public Function NotValidLen(Code As String) As Boolean

On Error Resume Next
If Len(Trim(Code)) > 0 Then
    NotValidLen = False
Else
    NotValidLen = True
End If
End Function
''''''
'Added by muzammil to run the App as directed by Francois.
Public Function DidFieldChange(strOldValue As String, strNewValue As String)

Dim ret
    ret = StrComp(Trim(strOldValue), Trim(strNewValue), vbTextCompare)
            If ret <> 0 Then
                DidFieldChange = True
            Else
                DidFieldChange = False
            End If

End Function

Public Function WriteParameterFiles(Recepients() As String, sender As String, Attachments() As String, subject As String, attention As String)
 
 Dim l
 Dim x
 Dim y
 Dim i
 Dim Email() As String
 Dim fax() As String
 Dim rs As New ADODB.Recordset
 
 If Len(Trim(sender)) = 0 Then
 
    rs.Source = "select com_name from company where com_compcode = ( select psys_compcode from pesys where psys_npecode ='" & deIms.NameSpace & "')"
    rs.ActiveConnection = deIms.cnIms
    rs.Open
    
    If rs.RecordCount > 0 Then
        If Len(rs("com_name") & "") > 0 Then sender = rs("com_name")
    End If
    rs.Close
    
    
End If
 
On Error GoTo errMESSAGE
  
'Splitting the address into Emails and Faxes.
 l = UBound(Recepients)
 
 
 
    x = 0
    y = 0
 
 
 For i = 0 To l
    Recepients(i) = Replace(Recepients(i), " ", "") 'JCG 2008/10/12
    
     If InStr(Recepients(i), "@") > 0 Then
       
       ReDim Preserve Email(x)
       Email(x) = Recepients(i)
       x = x + 1
       
    Else
      
       ReDim Preserve fax(y)
       fax(y) = Recepients(i)
       y = y + 1
       
    End If
       
       
       
 Next i

    If IsArrayLoaded(Email) Then 'M 02/23/02
    
        If Not (UBound(Email) = 0 And Email(0) = "") Then
            If UBound(Email) >= 0 Then Call WriteParameterFileEmail(Attachments, Email, subject, sender, attention)
            'If UBound(Email) >= 0 Then Call WriteParameterFileEmailUsingPDFCreator(Attachments, Email, subject, sender, attention)
        End If
        
    End If                      'M 02/23/02
    
    If IsArrayLoaded(fax) Then 'M 02/23/02
    
        If Not (UBound(fax) = 0 And fax(0) = "") Then
        
' JCG 06/14/2008, change made for eFax
'            If UBound(fax) >= 0 Then Call WriteParameterFileFax(Attachments, fax, subject, sender, attention)  'Dont use this, this is old
            If UBound(fax) >= 0 Then Call WriteParameterEfax(Attachments, fax, subject, sender, attention)
            'If UBound(fax) >= 0 Then Call WriteParameterEfaxUsingPDFCreator(Attachments, fax, subject, sender, attention)
'---------------------------------------------
            
        End If
        
    End If 'M 02/23/02

errMESSAGE:
    
    If Err.number <> 0 And Err.number <> 9 Then
        
        MsgBox "Process WriteParameterFiles " + Err.Description
    
    Else
        
        Err.Clear
    
    End If

End Function

Public Sub SendAttFaxAndEmail(reportNAME As String, ParamsForRPTI() As String, CrystalControl As Crystal.CrystalReport, ParamsForCrystalReport() As String, rsReceptList As ADODB.Recordset, subject As String, Message As String, FieldName As String)

Dim rptinf As RPTIFileInfo

Dim i As Integer

On Error Resume Next
       
 'Sets the Crystal component with all the required information
       
    With CrystalControl
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\" & reportNAME

        For i = 0 To UBound(ParamsForCrystalReport)

        .ParameterFields(i) = ParamsForCrystalReport(i)

        Next i

    End With
       
  'Generates and RPTINF file with all the parameters specified.
       
    With rptinf
        
        .ReportFileName = reportPath & reportNAME
        
        Call translator.Translate_Reports(reportNAME)
        
        Call translator.Translate_SubReports
        
        .parameters = ParamsForRPTI
        
    End With
    
    ParamsForRPTI(0) = ""
    
    'Writes to the RPTI file all the parameter information,NOT the data generated from the report
    
    Call WriteRPTIFile(rptinf, ParamsForRPTI(0))
    
    'Send out Emails and Faxes , this is where the real report is generated.
    
    Call SendEmailAndFax(rsReceptList, FieldName, subject, Message, ParamsForRPTI(0))
 
 If Err.number > 0 Then Err.Clear

End Sub



Public Function generateattachments(reportNAME As String, ReportCaption As String, ParamsForCrystalReport() As String, CrystalControl As Crystal.CrystalReport) As String()
  Dim Attachments(0) As String
  
  Dim IFile As IMSFile
  
  Dim Filename As String
  
  Dim i As Integer
  
  Set IFile = New IMSFile

On Error GoTo errMESSAGE
  
    With CrystalControl
        
        .Reset
        
        .ReportFileName = reportPath & reportNAME
        
        Call translator.Translate_Reports(reportNAME)
        
        Call translator.Translate_SubReports
       
       For i = 0 To UBound(ParamsForCrystalReport)

        .ParameterFields(i) = ParamsForCrystalReport(i)

        Next i
        
    End With
    
     Attachments(0) = "Report-" & ReportCaption & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"  'JCG 2008/7/6
     'Attachments(0) = "PO-" + deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf" 'JCG 2008/7/6
     
     'FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
     
     Filename = ConnInfo.EmailOutFolder & Attachments(0)
     'FileName = App.Path + "\messages\" + Attachments(0) 'JCG 2008/7/6
     If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
     Attachments(0) = Filename
     ' If Not FileExists(Filename) Then MDI_IMS.SaveReport Filename, crptRTF 'JCG 2008/7/6
     MDI_IMS.SaveReport Filename, crptRTF 'JCG 2008/7/6
     
     
     generateattachments = Attachments
errMESSAGE:

    If Err.number <> 0 Then
    
        MsgBox "Process generateattachments " + Err.Description
        
    End If

End Function



Public Function generateattachmentswithCR11(reportNAME As String, ReportCaption As String, ParamsForCrystalReport() As String, CrystalControl As Crystal.CrystalReport) As String()
  
  Dim Attachments(0) As String
  
  Dim IFile As IMSFile
  
  Dim Filename As String
  
  Dim i As Integer
  
  Set IFile = New IMSFile

On Error GoTo errMESSAGE

'
'    With CrystalControl
'
'        .Reset
'
'        .ReportFileName = ReportPath & reportNAME
'
'        Call translator.Translate_Reports(reportNAME)
'
'        Call translator.Translate_SubReports
'
'       For I = 0 To UBound(ParamsForCrystalReport)
'
'        .ParameterFields(I) = ParamsForCrystalReport(I)
'
'        Next I
'
'    End With
'
     

     Attachments(0) = "Report-" & ReportCaption & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".Pdf"  'JCG 2008/7/6
     'Attachments(0) = "PO-" + deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf" 'JCG 2008/7/6
     
     'FileName = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
     
     Filename = ConnInfo.EmailOutFolder & Attachments(0)

    Dim x As New clsexport
    
    'x.ParamsForCrystalReport = ParamsForCrystalReport()
    'x.Namespace = "Pect" ' ParamsForCrystalReport(0)
    x.ExportFilePath = Filename
    x.reportNAME = reportNAME

   ' x.ReporttypesCr11 = EReportTypesForCR11.PO


     If IFile.FileExists(Filename) Then IFile.DeleteFile (Filename)
     Attachments(0) = Filename
    
     Call x.GeneratePdf(ParamsForCrystalReport)
       
     generateattachmentswithCR11 = Attachments

Exit Function

errMESSAGE:

    If Err.number <> 0 Then
    
        MsgBox "Process generateattachments " + Err.Description
        
    End If

End Function

Public Function ParseMiddleValue(str As String) As String

Dim loc1 As String
Dim loc2 As String

Dim Arr() As String

Arr = Split(str, ";")

'loc1 = InStr(0, str, ";")
'loc2 = InStr(loc1 + 1, str, ";")

ParseMiddleValue = Arr(1) 'Mid(ParamsForCrystalReport(1), loc1, loc2 - loc1)

End Function


Public Function generateattachmentsPDF(reportNAME As String, ReportCaption As String, ParamsForCrystalReport() As String, CrystalControl As Crystal.CrystalReport, poNum As String, Optional docKind As String) As String()
  Dim Attachments(0) As String
  Dim IFile As IMSFile
  Dim fileNameString As String
  Dim i As Integer
  Set IFile = New IMSFile
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sql, DocType, confirm, msg As String
  Dim Flag As Integer
On Error GoTo errMESSAGE
    sql = "select * from po where po_npecode='" + deIms.NameSpace + "' and po_ponumb = '" + poNum + "' "
    rs.Source = sql
    rs.ActiveConnection = deIms.cnIms
    rs.Open
    If rs.RecordCount > 0 Then
        DocType = rs!po_docutype
        confirm = rs!po_confordr
    Else
        DocType = ""
        confirm = ""
    End If
    ' JCG 2008/7/10
    Attachments(0) = ""
    If docKind = "receipt" Then
        Attachments(0) = "Receipt-" + poNum + "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".pdf"
    ElseIf docKind = "document" Then
        Attachments(0) = "Document-" + poNum + "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".pdf"
    Else
        Attachments(0) = "PO-" + poNum + "-" + Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") + ".pdf"
    End If
     'Filename = App.Path + "\messages\" + Attachments(0)
     fileNameString = Attachments(0)
     'If IFile.FileExists(ConnInfo.EmailOutFolder + fileNameString) Then IFile.DeleteFile (ConnInfo.EmailOutFolder + fileNameString)
    
    Call pdfStuff(fileNameString)
    Attachments(0) = ConnInfo.EmailOutFolder + Attachments(0)

    Dim oldPrinter As String
    oldPrinter = Printer.DeviceName
    Dim w As New WshNetwork
    w.SetDefaultPrinter ("PDFCreator")
    Set w = Nothing

    With CrystalControl
        .Reset
        .Destination = crptToPrinter
        .ReportFileName = reportPath & reportNAME
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + poNum + ";true"

       'Modified by Juan 2016-02-10
        msg1 = translator.Trans("M00392")
        .WindowTitle = IIf(msg1 = "", "Transaction", msg1)
        Call translator.Translate_Reports("po.rpt")

        msg1 = translator.Trans("M00091")
        If msg1 = "" Then msg1 = "Total Price of"
        msg2 = translator.Trans("M00093")
        If msg2 = "" Then msg2 = "in"
        Dim curr
        curr = " : "
        .Formulas(99) = "gttext = ' " + msg1 + " ' + {DOCTYPE.doc_desc} + ' " + msg2 + " ' + {CURRENCY.curr_desc} + ' " + curr + "' + totext(Sum ({@total}, {PO.po_ponumb}))"

        Dim lbl_doc_desc As String
        lbl_doc_desc = translator.TranslateObject(deIms.cnIms, "doctype", DocType)
        If lbl_doc_desc <> "" Then
            .Formulas(101) = "lbl_doc_desc = '" + lbl_doc_desc + "'"
        End If

        If confirm And translator.TR_LANGUAGE <> "US" Then
            msg = translator.Trans("M00881")
            If msg <> "" Then
                .Formulas(100) = "confirmingorder = '" + msg + "'"
            End If
        End If

        Call translator.Translate_SubReports

        '---------------------------------------------
    End With
    MDI_IMS.PrintDirectReport ConnInfo.EmailOutFolder + fileNameString

'    Do Until PDFCreator1.cCountOfPrintjobs = 0
'    DoEvents
'    Loop
    
    
    
    Dim ww As New WshNetwork
    ww.SetDefaultPrinter (oldPrinter)
    Set ww = Nothing
    rs.Close
    '---------------
    Sleep 3000
     generateattachmentsPDF = Attachments
'     Do While PDFCreator1.cCountOfPrintjobs > 0
'        DoEvents
'        If PDFCreator1.cCountOfPrintjobs = 0 Then
'           If PDFCreator1.cVisible Then PDFCreator1.cVisible = False
'        End If
'     Loop
errMESSAGE:

    If Err.number <> 0 Then
    
        MsgBox "Process generateattachmentsPDF -flag:" + Err.Description
        
    End If

End Function

Public Sub pdfStuff(Filename As String)
On Error GoTo ErrHandler
 Set PDFCreator1 = New clsPDFCreator
 Set pErr = New clsPDFCreatorError
 
    With PDFCreator1
        .cVisible = True
        If .cStart("/NoProcessingAtStartup") = False Then
            If .cStart("/NoProcessingAtStartup", True) = False Then
                Exit Sub
            End If
            .cVisible = True
        End If
        .cErrorClear
        If .cPrinterStop = True Then .cPrinterStop = False
        Set opt = .cOptions
        .cClearCache
    End With

    opt.AutosaveFilename = Filename
    opt.AutosaveDirectory = ConnInfo.EmailOutFolder
    opt.UseAutosave = 1

    Set PDFCreator1.cOptions = opt
Exit Sub

ErrHandler:
 If Err.number <> 0 Then
    MsgBox "Error on pdfStuff -" + Err.Description
    Err.Clear
 End If
End Sub
Private Sub PDFCreator1_eReady()
    PDFCreator1.cPrinterStop = True
End Sub


Private Function PrinterIndex(Printername As String) As Long ' JCG 2008/7/10
 Dim i As Long
 For i = 0 To Printers.Count - 1
  If UCase(Printers(i).DeviceName) = UCase$(Printername) Then
   PrinterIndex = i
   Exit For
  End If
 Next i
End Function

'Muzammil;s code
Public Function sendOutlookEmailandFax(reportNAME As String, ReportCaption As String, CrystalControl As Crystal.CrystalReport, ParamsForCrystalReports() As String, rsReceptList As ADODB.Recordset, subject As String, attention As String, Optional sender As String, Optional FieldName As String, Optional PO As String)
Dim Params(1) As String
Dim i As Integer
Dim Attachments() As String
Dim Recepients() As String
'Recepients = Null 'JCG 2008/8/29

Dim str As String

On Error GoTo errMESSAGE

     If rsReceptList.RecordCount > 0 Then


        ' attention = "Attention Please " ' JCG 2008/7/12
        'JCG 2008/7/14
        Dim poNum As String
        Dim size As Integer

        'poNum = subject 'JCG 2008/9/1
        'size = InStr(poNum, "PO Number") + 10 'JCG 2008/9/1
        'poNum = LTrim(Mid(poNum, size)) 'JCG 2008/9/1
        '---------
        poNum = PO
        If poNum = "" Then poNum = subject
        
        Dim Text As String
        Text = "Please find here attached PO #"
        If translator.TR_LANGUAGE <> "US" Then
            Text = translator.Trans("M00928")
            Text = IIf(Text = "", "Please find here attached PO #", Text)
        End If
        
        'added by Juan 2020/02/20
        If deIms.NameSpace = "JA414" Then
            attention = "Buenos días estimados," + Chr(13) + Chr(10) + Chr(13) + Chr(10) _
                + "Adjunta encontrarán la Orden de Compra número: " + poNum + Chr(13) + Chr(10) + Chr(13) + Chr(10) _
                + "Por Favor proceder a iniciar la gestión de entrega de inmediato." + Chr(13) + Chr(10) + Chr(13) + Chr(10) _
                + "Un cordial saludo."
        Else
            attention = "Please find here attached PO #" + poNum  ' JCG 2008/7/12
        End If
        
         'Attachments = generateattachments(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl)  ' JCG 2008/7/10
         
         ' JCG 2016-02-10
         'Attachments = generateattachmentswithCR11(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl)  ' JCG 2008/7/10
        
'
        If Left(subject, 2) = "PO" Then
            Attachments = generateattachmentsPDF(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl, poNum)
        Else
            Attachments = generateattachmentsPDF(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl, poNum, "document")
        End If


     If Len(Trim(FieldName)) = 0 Then

        Recepients = ToArrayFromRecordset(rsReceptList) ' This is just to keep with the old compatiblity. It thinks
                                                        ' the First field in the recordset are the email Addresses.

     Else
        Recepients = ToArray(rsReceptList, FieldName, i, str)
     End If

     'MsgBox "@recs->" + Format(rsReceptList.RecordCount)

        Call WriteParameterFiles(Recepients, sender, Attachments, subject, attention)
    Else

         MsgBox "No Recipients to Send", , "Imswin"

    End If

errMESSAGE:

    If Err.number <> 0 Then

        MsgBox "Process sendOutlookEmailandFax " + Err.Description

    End If

End Function

'Juan's code
'Public Function sendOutlookEmailandFax(reportNAME As String, ReportCaption As String, CrystalControl As Crystal.CrystalReport, ParamsForCrystalReports() As String, rsReceptList As adodb.Recordset, subject As String, attention As String, Optional sender As String, Optional FieldName As String, Optional PO As String)
'Dim Params(1) As String
'subject = Replace(subject, " ", "") 'JCG 2008/10/08
'Dim i As Integer
'
'Dim Attachments() As String
'
'Dim Recepients() As String
''Recepients = Null 'JCG 2008/8/29
'
'Dim str As String
'
'On Error GoTo errMESSAGE
'
'     If rsReceptList.RecordCount > 0 Then
'
'
'        ' attention = "Attention Please " ' JCG 2008/7/12
'        'JCG 2008/7/14
'        Dim poNum As String
'        Dim size As Integer
'
'        'poNum = subject 'JCG 2008/9/1
'        'size = InStr(poNum, "PO Number") + 10 'JCG 2008/9/1
'        'poNum = LTrim(Mid(poNum, size)) 'JCG 2008/9/1
'        '---------
'        poNum = PO
'        'reportNAME = reportNAME + "-" + Trim(PO)
'
'        If poNum = "" Then poNum = subject
'        attention = "Please find here attached PO #" + poNum + " From Pecten Cameroon company" ' JCG 2008/7/12
'
'        ' Attachments = generateattachments(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl) ' JCG 2008/7/10
'        ' JCG 2008/7/13
'
'        If Left(subject, 2) = "PO" Then
'            Attachments = generateattachmentsPDF(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl, poNum)
'        Else
'            Attachments = generateattachmentsPDF(reportNAME, ReportCaption, ParamsForCrystalReports, CrystalControl, poNum, "document")
'        End If
'        '----------------
'
'     If Len(Trim(FieldName)) = 0 Then
'
'        Recepients = ToArrayFromRecordset(rsReceptList) ' This is just to keep with the old compatiblity. It thinks
'                                                        ' the First field in the recordset are the email Addresses.
'
'     Else
'        Recepients = ToArray(rsReceptList, FieldName, i, str)
'     End If
'
'     'MsgBox "@recs->" + Format(rsReceptList.RecordCount)
'
'        Call WriteParameterFiles(Recepients, sender, Attachments, subject, attention)
'    Else
'
'         MsgBox "No Recipients to Send", , "Imswin"
'
'    End If
'
'errMESSAGE:
'
'    If Err.number <> 0 Then
'
'        MsgBox "Process sendOutlookEmailandFax " + Err.Description
'
'    End If
'
'End Function

Private Function FixFaxNumber(Faxnumber As String) As String
On Error Resume Next

    If Len(Faxnumber) < 7 Then Exit Function

    If Left$(Faxnumber, 1) = "+" Then
        Faxnumber = Right$(Faxnumber, Len(Faxnumber) - 1)
    End If
    
    If Mid$(Faxnumber, 1, 4) <> "" Then
        FixFaxNumber = "" & Faxnumber
    End If
End Function

'Juan's code
Public Function WriteParameterFileEmailUsingPDFCreator(Attachments() As String, Recipients() As String, subject As String, sender As String, attention As String) As Integer
On Error GoTo errMESSAGE
     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String

     Filename = "Email" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     'FileName = "c:\IMSRequests\IMSRequests\" & FileName

     Filename = ConnInfo.EmailParameterFolder & Filename

    For i = 0 To UBound(Recipients)
            'recepientSTR = recepientSTR & Trim$(Recipients(i) & ";") 'JCG 2008/7/6
            recepientSTR = recepientSTR + Trim$(Recipients(i)) + "  " ' JCG 2008/7/6
    Next

      i = 0

    For i = 0 To UBound(Attachments)
            ' reports = reports & Trim$(Attachments(i) & ";") ' JCG 2008/7/6
            reports = reports + App.Path + "\messages\" + Trim$(Attachments(i)) + "  " ' JCG 2008/7/6
    Next

'JCG 2008/7/5
    'Open Filename For Output As FileNumb

    '    Print #FileNumb, "[Email]"
    '    Print #FileNumb, "Recipients=" & recepientSTR
    '    Print #FileNumb, "Reports=" & reports
    '    Print #FileNumb, "Subject=" & subject
    '    Print #FileNumb, "Sender=" & sender
    '    Print #FileNumb, "Attention=" & Trim$(attention)

    'Close #FileNumb
'-----------------

If Len(recepientSTR) > 0 Then Call sendProcess(recepientSTR, reports, subject, attention) 'JCG 2008/7/5
'JCG 20088/8/29
recepientSTR = ""
reports = ""
'subject = ""
'attention = ""
'--------------------
WriteParameterFileEmailUsingPDFCreator = 1

Exit Function
errMESSAGE:
    If Err.number <> 0 Then
        MsgBox "Error in WriteParameterFileEmailUsingPDFCreator : " + Err.Description
    End If
End Function

'muzammil;s code
Public Function WriteParameterFileEmail(Attachments() As String, Recipients() As String, subject As String, sender As String, attention As String) As Integer

On Error GoTo errMESSAGE
     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String

     'FileName = "Email" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     'FileNumb = FreeFile

     'FileName = ConnInfo.EmailParameterFolder & FileName

    For i = 0 To UBound(Recipients)
            recepientSTR = recepientSTR & Trim$(Recipients(i) & ",")

    Next

      i = 0

If UBound(Attachments) > 0 Then
    For i = 0 To UBound(Attachments)
             reports = reports & Trim$(Attachments(i) & ";") ' JCG 2008/7/6
            'reports = reports + App.Path + "\messages\" + Trim$(Attachments(i)) + "  " ' JCG 2008/7/6
    Next

    ElseIf UBound(Attachments) = 0 Then
       reports = reports & Trim$(Attachments(i))
End If
'JCG 2008/7/5
    'Open Filename For Output As FileNumb

    '    Print #FileNumb, "[Email]"
    '    Print #FileNumb, "Recipients=" & recepientSTR
    '    Print #FileNumb, "Reports=" & reports
    '    Print #FileNumb, "Subject=" & subject
    '    Print #FileNumb, "Sender=" & sender
    '    Print #FileNumb, "Attention=" & Trim$(attention)

    'Close #FileNumb
'-----------------

If Len(recepientSTR) > 0 Then Call sendProcess(recepientSTR, reports, subject, attention) 'JCG 2008/7/5
'JCG 20088/8/29
recepientSTR = ""
reports = ""
'subject = ""
'attention = ""
'--------------------
WriteParameterFileEmail = 1

Exit Function
errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If
End Function

'Juan's code
Public Function WriteParameterEfaxUsingPDFCreator(Attachments, Recipients, subject, sender, attention) 'JCG 6/14/2008 added for eFax
    On Error GoTo errMESSAGE

     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String
     Dim sql, companyName
     Dim datax As New ADODB.Recordset

     Filename = "Email" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     Filename = ConnInfo.EmailParameterFolder & Filename


    For i = 0 To UBound(Recipients)
            'recepientSTR = recepientSTR & Trim$(Recipients(i) & ";") 'JCG 2008/7/6
            'MsgBox "@-->" + Recipients(i)
            If Trim(Recipients(i)) <> "" Then
                recepientSTR = recepientSTR + FixFaxNumber(Trim$(Recipients(i))) + "@efaxsend.com" + "  " ' JCG 2008/7/6
            End If
    Next

      i = 0

    For i = 0 To UBound(Attachments)
            'reports = reports & Trim$(attachments(i) & ";") ' JCG 2008/7/6
            reports = reports + App.Path + "\messages\" + Trim$(Attachments(i)) + "  " ' JCG 2008/7/6
    Next

    'JCG 2008/7/6
    'Open Filename For Output As FileNumb

    '    Print #FileNumb, "[EFAX]"
    '    Print #FileNumb, "Recipients=" & recepientSTR + "@efaxsend.com"
    '    Print #FileNumb, "Reports=" & reports
    '    Print #FileNumb, "Subject=" & subject
    '    Print #FileNumb, "Sender=" & "acourtaud@groupgls.com"
    '    Print #FileNumb, "Attention=" & Trim$(attention)

    'Close #FileNumb
    '---------------------
    Dim subjectText As String
    Dim bodyText As String
    subjectText = subject
    bodyText = attention

    Call sendProcess(recepientSTR, reports, subjectText, bodyText) 'JCG 2008/7/5

'JCG 20088/8/29
recepientSTR = ""
reports = ""
subjectText = ""
bodyText = ""
'--------------------


errMESSAGE:
    If Err.number <> 0 Then
        MsgBox "Error occured in WriteParameterEfaxUsingPDFCreator : " + Err.Description
    End If
End Function

'muzammil's code
Public Function WriteParameterEfax(Attachments, Recipients, subject, sender, attention) 'JCG 6/14/2008 added for eFax
    On Error GoTo errMESSAGE

     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String
     Dim sql, companyName
     Dim datax As New ADODB.Recordset

     'FileName = "Email" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     'FileNumb = FreeFile

     'FileName = ConnInfo.EmailParameterFolder & FileName


    For i = 0 To UBound(Recipients)
            'recepientSTR = recepientSTR & Trim$(Recipients(i) & ";") 'JCG 2008/7/6
            'MsgBox "@-->" + Recipients(i)
            If Trim(Recipients(i)) <> "" Then
                recepientSTR = recepientSTR + FixFaxNumber(Trim$(Recipients(i))) + "@efaxsend.com" + "," ' JCG 2008/7/6
            End If
    Next

      i = 0

If UBound(Attachments) > 0 Then

    For i = 0 To UBound(Attachments)
            reports = reports & Trim$(Attachments(i) & ";") ' JCG 2008/7/6
            'reports = reports + App.Path + "\messages\" + Trim$(Attachments(i)) + "  " ' JCG 2008/7/6
    Next


    ElseIf UBound(Attachments) = 0 Then
        reports = reports & Trim$(Attachments(i))
End If


    'JCG 2008/7/6
    'Open Filename For Output As FileNumb

    '    Print #FileNumb, "[EFAX]"
    '    Print #FileNumb, "Recipients=" & recepientSTR + "@efaxsend.com"
    '    Print #FileNumb, "Reports=" & reports
    '    Print #FileNumb, "Subject=" & subject
    '    Print #FileNumb, "Sender=" & "acourtaud@groupgls.com"
    '    Print #FileNumb, "Attention=" & Trim$(attention)

    'Close #FileNumb
    '---------------------
    Dim subjectText As String
    Dim bodyText As String
    subjectText = subject
    bodyText = attention

    Call sendProcess(recepientSTR, reports, subjectText, bodyText) 'JCG 2008/7/5

'JCG 20088/8/29
recepientSTR = ""
reports = ""
subjectText = ""
bodyText = ""
'--------------------


errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If
End Function


Public Function WriteParameterFileFax(Attachments, Recipients, subject, sender, attention)
    On Error GoTo errMESSAGE
    
     Dim Filename As String
     Dim FileNumb As Integer
     Dim i As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String
     Dim sql, companyName
     Dim datax As New ADODB.Recordset

     Filename = "Fax" & "-" & deIms.NameSpace & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     'FileName = "c:\IMSRequests\IMSRequests\" & FileName
     
     Filename = ConnInfo.EmailParameterFolder & Filename

    For i = 0 To UBound(Recipients)
            recepientSTR = recepientSTR & Trim$(Recipients(i) & ";")
    Next

      i = 0

    For i = 0 To UBound(Attachments)
            reports = reports & Trim$(Attachments(i) & ";")
    Next
    
    Open Filename For Output As FileNumb

        Print #FileNumb, "[WINFAX]"
        Print #FileNumb, "Recipients=" & recepientSTR
        Print #FileNumb, "Reports=" & reports
        Print #FileNumb, "Subject=" & subject
        Print #FileNumb, "Sender=" & sender
        Print #FileNumb, "Attention=" & Trim$(attention)

    Close #FileNumb

errMESSAGE:
    If Err.number <> 0 Then
        MsgBox Err.Description
    End If
End Function




