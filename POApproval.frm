VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmPOApproval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Approval"
   ClientHeight    =   5835
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   Tag             =   "02020500"
   Begin VB.TextBox txtsearch 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Text            =   "Hit enter to see results"
      ToolTipText     =   "Hit enter to see results"
      Top             =   240
      Width           =   1770
   End
   Begin VB.CommandButton CmdPrint 
      Cancel          =   -1  'True
      Caption         =   "&Print"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton CmdApprove 
      Caption         =   "Approval"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancal 
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
      Width           =   1500
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGLine 
      Height          =   4635
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8505
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ";"
      Col.Count       =   6
      UseGroups       =   -1  'True
      HeadFont3D      =   4
      DefColWidth     =   5292
      CheckBox3D      =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   212
      Groups(0).Width =   14023
      Groups(0).Caption=   "Transaction Approval"
      Groups(0).Columns.Count=   6
      Groups(0).Columns(0).Width=   2646
      Groups(0).Columns(0).Caption=   "PO# "
      Groups(0).Columns(0).Name=   "ponumb"
      Groups(0).Columns(0).CaptionAlignment=   2
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   2
      Groups(0).Columns(0).FieldLen=   15
      Groups(0).Columns(0).Locked=   -1  'True
      Groups(0).Columns(1).Width=   2884
      Groups(0).Columns(1).Caption=   "Transaction Type"
      Groups(0).Columns(1).Name=   "transtype"
      Groups(0).Columns(1).CaptionAlignment=   2
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   25
      Groups(0).Columns(1).Locked=   -1  'True
      Groups(0).Columns(2).Width=   2831
      Groups(0).Columns(2).Caption=   "Total Amount"
      Groups(0).Columns(2).Name=   "total"
      Groups(0).Columns(2).CaptionAlignment=   2
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(2).Locked=   -1  'True
      Groups(0).Columns(3).Width=   1958
      Groups(0).Columns(3).Caption=   "Approve"
      Groups(0).Columns(3).Name=   "approve"
      Groups(0).Columns(3).CaptionAlignment=   2
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   11
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(3).Style=   2
      Groups(0).Columns(4).Width=   1508
      Groups(0).Columns(4).Caption=   "Sent"
      Groups(0).Columns(4).Name=   "Sent"
      Groups(0).Columns(4).CaptionAlignment=   2
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   11
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(4).Style=   2
      Groups(0).Columns(5).Width=   2196
      Groups(0).Columns(5).Caption=   "Us Export"
      Groups(0).Columns(5).Name=   "Us Export"
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   11
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(5).Style=   2
      TabNavigation   =   1
      _ExtentX        =   15002
      _ExtentY        =   8176
      _StockProps     =   79
      DataMember      =   "DOCTYPE"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "To view PO double click on the line"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmPOApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cmd As ADODB.Command
Dim cmdItem As ADODB.Command
Dim rs As ADODB.Recordset

Private Type DocumentType

    Ponumb As String
    Docutype As String

End Type


Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Dim TableLocked As Boolean, currentformname As String   'jawdat
Dim FDocumentTypes() As DocumentType
Dim GPOnumbs() As String

'set store procedure parameters and call it to update po and
'po line item  statuts

Private Sub CmdApprove_Click()
Dim PONumbers() As String
Dim porejected() As String
Dim l As Integer, y As Integer, x As Integer
Dim str As String
Dim i As Integer
Dim countarray As Integer

    Screen.MousePointer = 11
    SSDBGLine.Enabled = False
  Load FrmShowApproving
  Screen.MousePointer = 11
  FrmShowApproving.Top = 4620
  FrmShowApproving.Left = 3330
  FrmShowApproving.Width = 3000
  FrmShowApproving.Height = 1140
  
  
  FrmShowApproving.Show
  Screen.MousePointer = 11
    FrmShowApproving.Refresh
    Screen.MousePointer = 11
    Set cmd = New ADODB.Command
    
    With cmd
        .Prepared = True
        .CommandText = "ApprovePo"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = deIms.cnIms
        Screen.MousePointer = 11
        'DoEvents: DoEvents
        If .parameters.Count = 0 Then
            .parameters.Append .CreateParameter("RT", adInteger, adParamReturnValue)
            .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, deIms.NameSpace)
            
            .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, Null)
            .parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 20, CurrentUser)
            .parameters.Append .CreateParameter("@USEXPORT", adBoolean, adParamInput, , Null)
            .parameters.Append .CreateParameter("@returnresult", adBoolean, adParamOutput, , Null)
            
        End If
        
        Screen.MousePointer = 11
        .parameters("@PONUMB") = Null
        .parameters("@USER") = CurrentUser
        .parameters("@NAMESPACE") = deIms.NameSpace
        
    End With

Screen.MousePointer = 11
    With SSDBGLine
        .MoveFirst
        
        y = 0
        l = .Rows
        
        Do While y <= l
            y = y + 1
            
            'DoEvents: DoEvents
            If .Columns("approve").value Then
            
                Screen.MousePointer = 11
                Call ApprovePo(.Columns("ponumb").Text, .Columns(5).value, porejected, countarray)
                Screen.MousePointer = 11
                Call MDI_IMS.WriteStatus("Approving PO Number " & .Columns("ponumb").Text, 1)
                Screen.MousePointer = 11
                'MDI_IMS.WriteStatus ("Getting Po Numbers to be approved")
                
                'I = I + 1
                'Call .RemoveItem(y - I)
                
                '.MovePrevious
                If Err Then MsgBox Err.Description: Err.Clear
            End If
                
            If y = l Then Exit Do
            
            .MoveNext
            'DoEvents: DoEvents: DoEvents
            
        Loop
        countarray = countarray - 1
        
        If PoApprovalRejection(porejected, 1, countarray) = True Then
            
            str = "The following POs can not be approved because the Po line items did not have valid Eccn codes." & vbCrLf
                    
                For i = 0 To countarray  'UBound(porejected, 1)
        
                    If porejected(1, i) = 1 Then str = str & porejected(0, i) & ", "
        
                Next
                
                str = str & vbCrLf
                    
        End If
        
        If PoApprovalRejection(porejected, 2, countarray) = True Then
        
            str = str & "The following POs could not be approved because there were some unspecified errors." & vbCrLf
                    
                For i = 0 To countarray 'UBound(porejected, 1)
        
                    If porejected(1, i) <> 1 Then str = str & porejected(0, i) & ", "
        
                Next
                
        End If
        
        If IsArrayLoaded(porejected) Then MsgBox str, vbCritical
        
    End With
    
    Screen.MousePointer = 11
    Set cmd = Nothing
    Set cmdItem = Nothing
    Call MDI_IMS.WriteStatus("", 1)
    Screen.MousePointer = 11
    Unload FrmShowApproving
    
    SSDBGLine.Enabled = True
    Screen.MousePointer = 0
    
    Unload Me

    
End Sub

Public Function PoApprovalRejection(porejected() As String, ErrorType As Integer, ArrayMax As Integer) As Boolean

Dim i As Integer
On Error GoTo ErrHand

If IsArrayLoaded(porejected) = False Then
    
    PoApprovalRejection = False
    Exit Function
    
End If

If ErrorType = 1 Then

        For i = 0 To ArrayMax 'UBound(porejected, 1)
            
            If porejected(1, i) = ErrorType Then PoApprovalRejection = True
            Exit Function
            
        Next

ElseIf ErrorType = 2 Then

        For i = 0 To ArrayMax 'UBound(porejected, 1)
            
            If porejected(1, i) <> 1 Then PoApprovalRejection = True
            Exit Function
            
        Next


End If
Exit Function
ErrHand:
MsgBox Err.Description

End Function

'unload form

Private Sub CmdCancal_Click()
    Unload FrmShowApproving
    Unload Me
End Sub

'get crystal report parameters

Private Sub CmdPrint_Click()
    On Error GoTo ErrHandler
    
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Transtoap&send.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "doctype;ALL;TRUE"
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00458") 'J added
        .WindowTitle = "Transaction Approval" 'J modified
        Call translator.Translate_Reports("Transtoap&send.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
                   Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
End Sub

'load form call function to get po number and populate combo

Private Sub Form_Load()

Dim currentformname

    currentformname = Forms(3).Name

    Call translator.Translate_Forms("frmPOApproval")

    SSDBGLine.DataMode = ssDataModeAddItem
    
    Call AddPos(GetPOsForApproval(deIms.NameSpace, CurrentUser, deIms.cnIms))

    frmPOApproval.Caption = frmPOApproval.Caption + " - " + frmPOApproval.Tag
    SSDBGLine.HeadFont.size = 10
    SSDBGLine.HeadFont.Bold = True
    SSDBGLine.HeadFont.Weight = 1
    
Me.Left = Round((Screen.Width - Me.Width) / 2)
Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

Private Sub AddPos(rs As ADODB.Recordset)
Dim str As String
Dim i As Integer
    If rs Is Nothing Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    
    SSDBGLine.FieldSeparator = Chr(1)
    i = 0
    
    Do While Not rs.EOF
    
        If ConnInfo.Eccnactivate = "y" Or ConnInfo.Eccnactivate = "o" Then
        
            SSDBGLine.AddItem rs!PO_PONUMB & "" & str & rs!doc_desc & "" & str & rs!po_currcode & " " & Format(IIf(Len(Trim(rs!po_totacost & "")) = 0, 0, rs!po_totacost), "0.00") & str & 0 & str & 0 & str & IIf(rs!po_usexport = True, 1, 0)
            SSDBGLine.Columns(5).Visible = True
            
        Else
        
            SSDBGLine.AddItem rs!PO_PONUMB & "" & str & rs!doc_desc & "" & str & rs!po_currcode & " " & Format(IIf(Len(Trim(rs!po_totacost & "")) = 0, 0, rs!po_totacost), "0.00") & str & 0 & str & 0
            SSDBGLine.Columns(5).Visible = False
            SSDBGLine.Columns(5).value = False
            
        End If
    
        ReDim Preserve FDocumentTypes(i)
        
        FDocumentTypes(i).Ponumb = Trim(rs!PO_PONUMB)
        FDocumentTypes(i).Docutype = Trim(rs!po_docutype)
        
        rs.MoveNext
        
        i = i + 1
        
    Loop
    
End Sub

Public Function ApprovePo(PO As String, usexport As Boolean, ByRef PosRejected() As String, ByRef ArrayMax As Integer) As Boolean
    
    Dim rs As ADODB.Recordset
    Dim Max As Integer
    Dim returnresult As Integer
    On Error GoTo ErrHand
    
    
''    If IsArrayLoaded(PosRejected) = False Then
''
''        Max = 0
''
''    Else
''
''        Max = UBound(PosRejected, 1)
''
''    End If


    cmd.parameters("@PONUMB") = PO
    cmd.parameters("@usexport") = usexport
    
    'Set Rs = cmd.Execute(Options:=adExecuteNoRecords)
    Set rs = cmd.Execute
     returnresult = cmd.parameters("@returnresult").value
   ' If Rs.Fields(0) = 0 Then
    If returnresult = 0 Then
        
    
        'If Not UCase(deIms.NameSpace) = "TRNNG" And IsTransactionARequsition(PO) = False And SSDBGLine.Columns("sent").value = True Then 'M  'Commented by JCG 2008/1/20
        If Not UCase(deIms.NameSpace) = "TRNNG" And SSDBGLine.Columns("sent").value = True Then 'JCG 2008/1/20
        
           Call SendPO(PO, deIms.NameSpace)
        
        End If
        
    ElseIf returnresult = 1 Then
    
        ReDim Preserve PosRejected(1, ArrayMax)
        PosRejected(0, ArrayMax) = PO
        PosRejected(1, ArrayMax) = 1  '"Po\ Poline items does not have Eccn values even though this po is a US Export"
        ArrayMax = ArrayMax + 1
        
    End If
      'M
    
    ApprovePo = True
Exit Function
ErrHand:

        ReDim Preserve PosRejected(1, ArrayMax)
        PosRejected(0, ArrayMax) = PO
        PosRejected(1, ArrayMax) = 2 '"Po\ Poline items does not have Eccn values even though this po is a US Export"
        ArrayMax = ArrayMax + 1
        
MsgBox Err.Description
Err.Clear
End Function

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    Set cmd = Nothing
    If open_forms <= 5 Then ShowNavigator

'If TableLocked = True Then    'jawdat
'Dim imsLock As imsLock.Lock
'Set imsLock = New imsLock.Lock
'currentformname = Forms(3).Name
'Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
'End If



End Sub

'delete a recordset

Private Sub SSDBGLine_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    DispPromptMsg = False
End Sub

'get store procedure parameters
'and call function send email and fax

Public Sub SendPO(PO As String, NameSpace As String)
On Error GoTo Handled

'Dim Rs As ADODB.Recordset 'JCG 2008/8/28

Dim Filename As String
Dim cmd As ADODB.Command
Dim sql As String

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = deIms.cnIms
    
    sql = "SELECT porc_rec FROM POREC WHERE porc_ponumb = '" + Trim(PO) + "' AND porc_npecode = '" + NameSpace + "' " _
        + "union select dis_mail as porc_rec from distribution where dis_npecode='" + NameSpace + "'  "
    cmd.CommandText = sql
    If cmd.parameters.Count = 0 Then
        cmd.parameters.Append cmd.CreateParameter("PONUMB", adVarChar, adParamInput, 15)
        cmd.parameters.Append cmd.CreateParameter("NameSpace", adVarChar, adParamInput, 5)
    End If
    
    cmd.parameters(0) = PO
    cmd.parameters(1) = NameSpace
    
    Set rs = cmd.Execute
    
    rs.Close
    rs.CursorLocation = adUseClient
    
    rs.Open
    
    If rs.RecordCount > 0 Then
    
    
            Dim ParamsForRPTI(1) As String
            
            Dim rptinf As RPTIFileInfo
            
            Dim ParamsForCrystalReports(1) As String
            
            Dim subject As String
            
            Dim FieldName As String
            
            Dim Message As String
            
            Dim attention As String
            
            On Error Resume Next

''If rsReceptList Is Nothing Then Exit Sub
                

   With MDI_IMS.CrystalReport1
        .ReportFileName = reportPath & "po.rpt"
        .ParameterFields(0) = "namespace;" + NameSpace + ";true"
         .ParameterFields(1) = "ponumb;" + PO + ";true"


        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00458") 'J added
        .WindowTitle = "Transaction Approval" 'J modified
        Call translator.Translate_Reports("Transtoap&send.rpt") 'J added
        '---------------------------------------------

    End With


'    With BeforePrint
'        ReDim .parameters(1)
'        .parameters(1) = "ponumb=" & Ponumb
'        .ReportFileName = reportPath & "po.rpt"
'        .parameters(0) = "namespace=" & NameSpace
'    End With

            


            ParamsForCrystalReports(0) = "namespace;" + deIms.NameSpace + ";true"

            ParamsForCrystalReports(1) = "ponumb;" + PO + ";true"

            ParamsForRPTI(0) = "ponumb=" & PO

            ParamsForRPTI(1) = "namespace=" & deIms.NameSpace

            FieldName = "porc_rec"

            'added by Juan 2020/02/20
            If deIms.NameSpace = "JA414" Then
                sql = "select * from PO where po_ponumb = '" + PO + "' and " _
                    + "po_npecode='" + deIms.NameSpace + "'"
                Dim datax As New ADODB.Recordset
                Set datax = New ADODB.Recordset
                datax.Open sql, deIms.cnIms, adOpenForwardOnly
                Dim SupplierCode As String
                SupplierCode = ""
                Dim companyCode As String
                SupplierCode = ""
                If datax.RecordCount > 0 Then
                    SupplierCode = datax!po_suppcode
                    SupplierCode = Trim(SupplierCode)
                    companyCode = datax!po_compcode
                    companyCode = Trim(companyCode)
                End If
                sql = ""

                sql = "select * from company " _
                    + "where com_npecode='" + deIms.NameSpace + "' " _
                    + "and com_compcode = '" + companyCode + "'"
                Set datax = New ADODB.Recordset
                datax.Open sql, deIms.cnIms, adOpenForwardOnly
                Dim companyName As String
                companyName = ""
                If datax.RecordCount > 0 Then
                    companyName = datax!com_name
                    companyName = Trim(companyName)
                End If
                
                sql = "select * from supplier " _
                    + "where sup_npecode='" + deIms.NameSpace + "'" _
                    + "and sup_code = '" + SupplierCode + "'"
                Set datax = New ADODB.Recordset
                datax.Open sql, deIms.cnIms, adOpenForwardOnly
                Dim supplierName As String
                supplierName = ""
                If datax.RecordCount > 0 Then
                    supplierName = datax!sup_name
                    supplierName = Trim(supplierName)
                End If
                
                
                datax.Close
                subject = companyName + " - " + PO + " - " + supplierName
            Else
                subject = "Po Approval for PO Number " & PO
            End If

            attention = "Attention Please "

            Message = "PO Approval"
            
            Dim Text As String
            Text = "Transaction Approval"
            If translator.TR_LANGUAGE <> "US" Then
                Text = translator.Trans("L00458")
                Text = IIf(Text = "", "Transaction Approval", Text)
            End If
            If ConnInfo.EmailClient = Outlook Then
                'Call sendOutlookEmailandFax("po.rpt", "Transaction Approval", MDI_IMS.CrystalReport1, ParamsForCrystalReports, Rs, subject, attention) 'JCG 2008/9/1
                Call sendOutlookEmailandFax("po.rpt", Text + "-" & PO & "-", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rs, subject, attention, , , PO) 'JCG 2008/9/1
                'Call sendOutlookEmailandFax(Report_EmailFax_PO_name, "Transaction Approval-" & PO & "-", MDI_IMS.CrystalReport1, ParamsForCrystalReports, rs, subject, attention, , , PO)  'JCG 2008/9/1

            ElseIf ConnInfo.EmailClient = ATT Then

                Call SendAttFaxAndEmail("po.rpt", ParamsForRPTI, MDI_IMS.CrystalReport1, ParamsForCrystalReports, rs, subject, Message, FieldName)

            ElseIf ConnInfo.EmailClient = Unknown Then

                MsgBox "Email is not set up Properly. Please configure the database for emails.", vbCritical, "Imswin"

            End If


    
'' This is the previous piece of code
''''''''''
''''''''       Call WriteRPTIFile(BeforePrint(po, NameSpace), FileName)
''''''''
''''''''       Call SendEmailAndFax(rs, "porc_rec", "Po Approval for PO Number " & po, "PO Approval", FileName)
''''''''''
''''''''''
''''''''''
''''''''''
    End If
    
    Set rs = Nothing
    Set cmd = Nothing
    Exit Sub
    
Handled:
    Call LogErr(Name & "::BeforePrint", Err.Description, Err.number, True)
End Sub

'get parameters for crystal report

Private Function BeforePrint(Ponumb As String, NameSpace As String) As RPTIFileInfo
On Error GoTo Handled
Dim LOGINKEY
 'added the next 5 line 08/08/00  / Muzammil
   With MDI_IMS.CrystalReport1
        .ReportFileName = reportPath & "po.rpt"
        .ParameterFields(0) = "namespace;" + NameSpace + ";true"
         .ParameterFields(1) = "ponumb;" + Ponumb + ";true"
         'LOGINKEY = .LogOnServer("pdsodbc.dll", "imsO", "SAKHALIN", "sa", "2r2m9k3")
        
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("L00458") 'J added
        .WindowTitle = "Transaction Approval" 'J modified
       ' Call translator.Translate_Reports("Transtoap&send.rpt") 'J added
        Call translator.Translate_Reports("po.rpt") 'J added
        '---------------------------------------------
         
    End With
         

    With BeforePrint
        ReDim .parameters(1)
        .parameters(1) = "ponumb=" & Ponumb
        .ReportFileName = reportPath & "po.rpt"
        .parameters(0) = "namespace=" & NameSpace
    End With
    
    Exit Function
    
Handled:
    Call LogErr(Name & "::BeforePrint", Err.Description, Err.number, True)
End Function


Public Function IsTransactionARequsition(PO As String) As Boolean

Dim i As Integer

IsTransactionARequsition = False

For i = 0 To UBound(FDocumentTypes)

    If UCase(Trim(PO)) = UCase(FDocumentTypes(i).Ponumb) Then
    
     If FDocumentTypes(i).Docutype = "R" Then IsTransactionARequsition = True
        
        Exit Function
        
    End If
        

Next



End Function

Private Sub SSDBGLine_Change()

Select Case SSDBGLine.Col

Case 3
    
     SSDBGLine.Columns(4).value = SSDBGLine.Columns(3).value
Case 5
    'case of Eccn
    If ConnInfo.Eccnactivate = Constyes Then
        
        If SSDBGLine.Columns(5).value = False Then
            
            MsgBox "Us Export Option is set to True for the System and can not be false.", vbInformation
            SSDBGLine.Columns(5).value = True
            Exit Sub
        End If
        
    ElseIf ConnInfo.Eccnactivate = Constno Then
        
        If SSDBGLine.Columns(5).value = True Then
            
            MsgBox "Us Export Option is set to False for the System and can not be True.", vbInformation
            SSDBGLine.Columns(5).value = False
            Exit Sub
        End If
    
    
    
    ElseIf ConnInfo.Eccnactivate = ConstOptional Then
        
    End If
    
End Select

End Sub

Public Function MoveToPoinPoDetails(Ponumb As String)

Dim i As Integer
Dim x As Integer

x = Len(Trim(Ponumb))

If Len(Trim(Ponumb)) = 0 Then Exit Function
For i = 0 To UBound(FDocumentTypes)

    If UCase(Trim(Ponumb)) = UCase(Left(FDocumentTypes(i).Ponumb, x)) Then
    
        SSDBGLine.MoveFirst
        
        SSDBGLine.MoveRecords i
        
        Exit Function
        
     End If
    

Next i
SSDBGLine.SetFocus
MsgBox "Record does not exit.", vbInformation, "Ims"

End Function

Private Sub SSDBGLine_DblClick()
    On Error GoTo ErrHandler
    Dim PO As String
    PO = SSDBGLine.Columns("ponumb").Text
    
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath + "po.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "ponumb;" + PO + ";true"
        
        'Modified by Juan (8/28/2000) for Multilingual
        msg1 = translator.Trans("M00392") 'J added
        .WindowTitle = IIf(msg1 = "", "Transaction", msg1) 'J modified
        Call translator.Translate_Reports("po.rpt") 'J added
        msg1 = translator.Trans("M00091") 'J added
        If msg1 = "" Then msg1 = "Total Price of"
        msg2 = translator.Trans("M00093") 'J added
        If msg2 = "" Then msg2 = "in"
        Dim curr
        curr = " : "
        .Formulas(99) = "gttext = ' " + msg1 + " ' + {DOCTYPE.doc_desc} + ' " + msg2 + " ' + {CURRENCY.curr_desc} + ' " + curr + "' + totext(Sum ({@total}, {PO.po_ponumb}))" 'J modified
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::NavBar1_OnPrintClick", Err.Description, Err.number, True)
    End If

End Sub

Private Sub txtsearch_GotFocus()
If Trim(txtsearch.Text) = "Hit enter to see results" Then txtsearch = ""
End Sub

Private Sub txtsearch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call MoveToPoinPoDetails(txtsearch)
End Sub
