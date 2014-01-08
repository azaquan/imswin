VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_EditReception 
   Caption         =   "Edit Freight Forwarder Receipt"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   10860
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cboPoNumb 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ComboBox cboRecepTion 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   600
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   10455
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   10
      Columns(0).Width=   714
      Columns(0).Caption=   "Item"
      Columns(0).Name =   "Item"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HasForeColor=   -1  'True
      Columns(0).ForeColor=   8421504
      Columns(1).Width=   1958
      Columns(1).Caption=   "Stock #"
      Columns(1).Name =   "Stock #"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(1).HasForeColor=   -1  'True
      Columns(1).ForeColor=   8421504
      Columns(2).Width=   3069
      Columns(2).Caption=   "Description"
      Columns(2).Name =   "Description"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(2).HasForeColor=   -1  'True
      Columns(2).ForeColor=   8421504
      Columns(3).Width=   1402
      Columns(3).Caption=   "Unit"
      Columns(3).Name =   "Unit"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(3).HasForeColor=   -1  'True
      Columns(3).ForeColor=   8421504
      Columns(4).Width=   1826
      Columns(4).Caption=   "Qty. PO"
      Columns(4).Name =   "Qty. PO"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(4).HasForeColor=   -1  'True
      Columns(4).ForeColor=   8421504
      Columns(5).Width=   3440
      Columns(5).Caption=   "Qty. Recieved to Date"
      Columns(5).Name =   "Qty. Recieved to Date"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      Columns(5).HasForeColor=   -1  'True
      Columns(6).Width=   3678
      Columns(6).Caption=   "Qty. Recieved Reception"
      Columns(6).Name =   "Qty. Recieved Reception"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   5
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(6).HasForeColor=   -1  'True
      Columns(6).ForeColor=   8421504
      Columns(7).Width=   1482
      Columns(7).Caption=   "UnitPrice"
      Columns(7).Name =   "UnitPrice"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(7).HasForeColor=   -1  'True
      Columns(7).ForeColor=   8421504
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "Price"
      Columns(8).Name =   "Price"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   3
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(8).HasForeColor=   -1  'True
      Columns(8).ForeColor=   8421504
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "test"
      Columns(9).Name =   "test"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   18441
      _ExtentY        =   3413
      _StockProps     =   79
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "&frmReception.cboRecepTion"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order #"
      Height          =   315
      Index           =   0
      Left            =   4500
      TabIndex        =   3
      Top             =   540
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reception #"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1800
   End
End
Attribute VB_Name = "frm_EditReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'''''''''''''        Option Explicit
Dim cboRecepTion1
Dim sRecNum As String
Dim vPKValues As Variant
Dim FNamespace As String
Dim rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim Recpdelt As imsReceptionDetail
Dim Reception As imsReception
Dim RsPOITEMS As ADODB.Recordset
Dim notREADY As Boolean
Dim SaveEnabled As Boolean
Dim beginning As Boolean
Dim currentPO, currentRECEPTION, currentRECEPTION_1, currentPO1
Dim rowguid, locked As Boolean, idleStateEngagedFlag As Boolean

Sub fillPO()
    On Error Resume Next

    Dim str As String
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    str = "po_ponumb = '" & cboPoNumb & "'"

    ReceptionDate = ""
    cboRecepTion = ""


    Set rst = deIms.rsGETPOITEMFORRECEPTION_SP

    If (rst.State And adStateOpen) = adStateOpen Then rst.Close
    Set cmd = deIms.Commands("GETPOITEMFORRECEPTION_SP")

    cmd.parameters("@PONUMB").value = cboPoNumb
     cmd.parameters("@NAMESPACE").value = FNamespace

    Set rst = cmd.Execute
    

    ReceptionDate = Format(Date, "mm/dd/yyyy")
    GetReceptions
    If Err Then Err.Clear

End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Public Sub cboPoNumb_Validate(Cancel As Boolean)
  'Added by Juan 11/18/200
    On Error Resume Next
    Dim text, i, sql
    Dim searcher As ADODB.Recordset

    'If deIms.rsGETPOITEMFORRECEPTION_SP.State = 1 Then Exit Sub
    With cboPoNumb
        text = .text
        If text <> "" Then
            For i = 0 To .ListCount - 1
                If text Like .list(i) Then
                    Call fillPO
                    Exit Sub
                End If
            Next
            sql = "SELECT po_ponumb, po_stas FROM PO WHERE po_npecode = '" + FNamespace + "' " _
                & "AND LTRIM(po_ponumb) = '" + Trim(.text) + "'"
            Set searcher = New ADODB.Recordset
            searcher.Open sql, deIms.cnIms, adOpenForwardOnly
            If searcher.RecordCount > 0 Then
                If Err.number > 0 Then
                    .text = "Error"
                    .SetFocus
                    Exit Sub
                End If
                If searcher!po_stas = "OP" Then
                    Call fillPO
                    Exit Sub
                Else
                    msg1 = translator.Trans("M00697")
                    MsgBox IIf(msg1 = "", "This transaction is not open", msg1)
                End If
            Else
                msg1 = translator.Trans("M00698")
                MsgBox IIf(msg1 = "", "This transaction doesn't exist", msg1)
            End If
        End If
        .text = ""
        .SetFocus
    End With
    '------------------------
End Sub
'call function get recordset for reception
'and populate data grid and format date data type

Private Sub cboRecepTion_Click()

On Error Resume Next

Dim rst As ADODB.Recordset
    If cboRecepTion = "" Then Exit Sub
    Screen.MousePointer = 11
    Set rst = deIms.rsGet_Reception_Info_From_PONumb
    Screen.MousePointer = 11
    If ((rst.State And adStateOpen) = adStateClosed) Then _
        Call deIms.Get_Reception_Info_From_PONumb(cboPoNumb, FNamespace)
    Screen.MousePointer = 11
    rst.Filter = 0
    rst.Filter = "recd_recpnumb = '" & cboRecepTion & "'"
    Screen.MousePointer = 11
    dgDetl.DataMember = ""
    Set dgDetl.DataSource = Nothing
    dgDetl.DataMember = "Get_Reception_Info_From_PONumb"

    If rs.RecordCount = 0 Then
     Screen.MousePointer = 0
    Exit Sub
    End If
    Screen.MousePointer = 11
    If ((Not IsNull(rst!rec_date)) Or IsEmpty(rst!rec_date)) Then
        'Label1(3) = Format(rst!rec_date, "mm/dd/yyyy")
        ReceptionDate = Format(rst!rec_date, "mm/dd/yyyy")
    End If
    Screen.MousePointer = 11
   Set dgDetl.DataSource = deIms

'''''''''''''       Call MakeGridReadonly(Not SaveEnabled)
'   dgDetl.Columns(6).Visible = False
   dgDetl.Columns(6).Visible = True
   dgDetl.Refresh


    Screen.MousePointer = 11
    If Len(Trim$(cboRecepTion)) > 0 Then
        NavBar1.PrintEnabled = True
        NavBar1.EMailEnabled = True
        NavBar1.SaveEnabled = False
    End If
    currentRECEPTION = cboRecepTion
    If Err Then Err.Clear
    Screen.MousePointer = 0
End Sub

Private Sub cboPoNumb_Click()


Screen.MousePointer = 11


Call fillPO

Screen.MousePointer = 0

Set SSOleDBGrid1.DataSource = Nothing

Dim i
For i = 0 To SSOleDBGrid1.Rows - 1
        SSOleDBGrid1.RemoveItem SSOleDBGrid1.Rows - 1
Next i

Call SSOleDBGrid1_populate
    
Screen.MousePointer = 0
End Sub

'unlock reception data combo

Private Sub cboRecepTion_DropDown()

End Sub

'do not allow enter data to reception data combo

Private Sub cboRecepTion_KeyPress(KeyAscii As Integer)
'''''''If NavBar1.SaveEnabled = False Then KeyAscii = 0

    'Added by Juan 11/17/2000
    Dim i, text
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        Exit Sub
    End If
    With cboRecepTion
        text = .text
        For i = 0 To .ListCount - 1
            If text Like .list(i) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    '-------------------------
End Sub
'load form and set navbar buttom




Private Sub dgDetl_ColEdit(ByVal ColIndex As Integer)
''''''jawdat, start copy
End Sub



Private Sub cboPoNumb000000_Change()

End Sub

Private Sub Form_Load()
   
 cboPoNumb = frmReception.cboPoNumb

cboRecepTion1 = frmReception.cboRecepTion

Label2.Caption = frmReception.cboRecepTion
Label3.Caption = frmReception.cboPoNumb

'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Me.Name
currentformname1 = Me.Name
Dim imsLock As imsLock.Lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock
ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)
Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid)   'lock should be here, added by jawdat, 2.1.02
If locked = True Then                                        'sets locked = true because another user has this record open in edit mode
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
locked = True
End If                                                       'without this End if the form will get compilation errors

'jawdat, end copy

'
Me.Height = "4260"
Me.Width = "10980"







Dim datax As ADODB.Recordset
Dim sql
On Error Resume Next
    notREADY = False
    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmReception")
    '------------------------------------------

    FNamespace = deIms.NameSpace: GetPoNumb


    cboPoNumb_Click


    SaveEnabled = Getmenuuser(FNamespace, CurrentUser, Me.Tag, deIms.cnIms)

    frmReception.Caption = frmReception.Caption + " - " + frmReception.Tag
     With RecipientList
        .ColWidth(0) = 300
        .ColWidth(1) = 9095
        .Rows = 2
        .Clear
        .TextMatrix(0, 1) = "Recipient List"
    End With

    With frm_EditReception
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub
'validate data format and show messege

Private Function ValidateData() As Boolean
Dim str As Double
Dim msg, Style, Title


End Function
'get po number recordset and populate data combo

Private Sub GetPoNumb()
On Error Resume Next

    deIms.rsGETPONUMBERSFORRECEPTION_SP.Close

    If Err Then Err.Clear: deIms.cnIms.Errors.Clear
    Call deIms.GETPONUMBERSFORRECEPTION_SP(FNamespace)

    Set rs = deIms.rsGETPONUMBERSFORRECEPTION_SP

    rs.Filter = 0
    Call PopuLateFromRecordSet(frmReception.cboPoNumb, rs, "po_ponumb", False)

    Set rs = deIms.rsGETPOITEMFORRECEPTION_SP


    If Err Then Err.Clear
End Sub

'call function get reception number recordset

Private Sub GetReceptions()
'On Error Resume Next
Dim l As Long
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim i  As Integer
Dim goSAMEreception As Boolean
    If currentRECEPTION = cboRecepTion1 Then goSAMEreception = True

    Set rst = deIms.rsGet_Reception_Info_From_PONumb
    Set cmd = deIms.Commands("Get_Reception_Info_From_PONumb")
    If rst.State And adStateOpen = adStateOpen Then rst.Close

    With cmd
        rst.Filter = 0
        Set rst = Nothing
        cboRecepTion.Clear
        .parameters("PONUMB").value = frmReception.cboPoNumb '''cboPoNumb
        .parameters("NAMESPACE").value = FNamespace

    End With
    If currentPO = frmReception.cboPoNumb Then ''cboPoNumb Then
        If goSAMEreception Then
            For i = 0 To cboRecepTion.ListCount - 1
                If cboRecepTion.list(i) = currentRECEPTION Then
                    cboRecepTion.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub



Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'call function get recordset and set store procedure parameters
'validata data format, check reception number exist or not

Private Sub AddItems()

Screen.MousePointer = 11
Dim lng As Long
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim Result As Boolean
Dim Check As Boolean

    Result = False



    Screen.MousePointer = 11

        Screen.MousePointer = 11
       Call UpdateReceptiontable
 
        
Dim i


        Screen.MousePointer = 0
        Screen.MousePointer = 11
        Me.Refresh

        cboRecepTion.Tag = cboRecepTion

    Screen.MousePointer = 0
    Exit Sub

    deIms.cnIms.RollbackTrans
    Screen.MousePointer = 0

End Sub

'set store procedure parameters and call it to update po status

Private Function UpdateReceptiontable() As String

Dim j, i
Dim SQL_QTY As New ADODB.Recordset

     poi_primreqdqty_string = "SELECT poi_qtydlvd, poi_primreqdqty FROM POITEM "
     poi_primreqdqty_string = poi_primreqdqty_string & "WHERE poi_npecode = '" & FNamespace & "' AND poi_ponumb = '" & cboPoNumb & "' "  ''cboPoNumb & "' "frmReception.cboPoNumb
     poi_primreqdqty_string = poi_primreqdqty_string & "AND poi_liitnumb = '" & SSOleDBGrid1.Columns(0).text & "'"

    Set SQL_QTY = New ADODB.Recordset
    SQL_QTY.Open poi_primreqdqty_string, deIms.cnIms
 
    poi_primreqdqty_value = SQL_QTY("poi_primreqdqty")
    poi_qtydlvd_value = SQL_QTY("poi_qtydlvd")


     SQL_QTY_RESULT = (poi_primreqdqty_value - poi_qtydlvd_value)

If (SQL_QTY_RESULT <= 0) Then

    SQL_Update = "UPDATE POITEM SET poi_qtytobedlvd = '0' "
    SQL_Update = SQL_Update & "WHERE poi_npecode = '" & FNamespace & "' AND poi_ponumb = '" & frmReception.cboPoNumb & "' " ''cboPoNumb & "' "
    SQL_Update = SQL_Update & "AND poi_liitnumb = '" & SSOleDBGrid1.Columns(0).text & "' AND poi_qtydlvd  > '0' AND poi_qtytobedlvd <= '0'"

Else

    SQL_Update = "UPDATE POITEM SET poi_qtytobedlvd = '" & SQL_QTY_RESULT & "' "
    SQL_Update = SQL_Update & "WHERE poi_npecode = '" & FNamespace & "' AND poi_ponumb = '" & frmReception.cboPoNumb & "' " ''cboPoNumb & "' "
    SQL_Update = SQL_Update & "AND poi_liitnumb = '" & SSOleDBGrid1.Columns(0).text & "' AND poi_qtydlvd  > '0'"

End If


    Set SQL_QTY = New ADODB.Recordset
    SQL_QTY.Open SQL_Update, deIms.cnIms
    



Dim SQL_ReceptionDetl As New ADODB.Recordset

''''put loop in here
SSOleDBGrid1.MoveFirst
For j = 0 To SSOleDBGrid1.Rows - 1


recd_unitpric = SSOleDBGrid1.Columns(9).text

recd_primqtydlvd = SSOleDBGrid1.Columns(6).text
QTY_REC_ToDate = SSOleDBGrid1.Columns(5).text
QTY_PO = SSOleDBGrid1.Columns(4).text



qty_tobe = Int(Int(QTY_PO) - Int(QTY_REC_ToDate))

If qty_tobe = "" Then qty_tobe = "0"
If qty_tobe < 0 Then
    
    msg = "Stock# " & Trim$(SSOleDBGrid1.Columns(1).text) & " is being over received, Do you want to continue ?"
  
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Imswin"

    If MsgBox(msg, Style, Title) = vbNo Then
    Exit Function: SSOleDBGrid1.SetFocus
    End If
    
End If

recd_totapric = (Int(recd_unitpric) * Int(recd_primqtydlvd))


If QTY_REC_ToDate = QTY_PO Then
poi_stasdlvy = "RC"
End If

If QTY_REC_ToDate = "0" Then
poi_stasdlvy = "NR"
End If

If ((QTY_REC_ToDate > 0) And (QTY_REC_ToDate < QTY_PO)) Then
poi_stasdlvy = "RP"
End If



'''''  Update ReceptionDetl
     
     UPDATE_ReceptionDetl_string = "Update RECEPTIONDETL SET recd_primqtydlvd = '" & recd_primqtydlvd & "' , recd_totapric = '" & recd_totapric & "' "
     UPDATE_ReceptionDetl_string = UPDATE_ReceptionDetl_string & " "
     UPDATE_ReceptionDetl_string = UPDATE_ReceptionDetl_string & "WHERE recd_npecode = '" & FNamespace & "' AND recd_recpnumb = '" & cboRecepTion1 & "' "
     UPDATE_ReceptionDetl_string = UPDATE_ReceptionDetl_string & "AND recd_liitnumb = '" & SSOleDBGrid1.Columns(0).text & "'"



    Set SQL_ReceptionDetl = New ADODB.Recordset
    SQL_ReceptionDetl.Open UPDATE_ReceptionDetl_string, deIms.cnIms
    
'''''  Update POITEM

     UPDATE_POItem_string = "Update POITEM SET poi_qtydlvd = '" & QTY_REC_ToDate & "' , poi_totaprice = '" & recd_totapric & "' , poi_secototaprice = '" & recd_totapric & "',  poi_qtytobedlvd = '" & qty_tobe & "', poi_stasdlvy = '" & poi_stasdlvy & "'"
     UPDATE_POItem_string = UPDATE_POItem_string & "WHERE poi_npecode = '" & FNamespace & "' AND poi_liitnumb = '" & SSOleDBGrid1.Columns(0).text & "' and poi_ponumb = '" & cboPoNumb & "'"

    Set SQL_ReceptionDetl = New ADODB.Recordset
    SQL_ReceptionDetl.Open UPDATE_POItem_string, deIms.cnIms
    
  SSOleDBGrid1.MoveNext
Next j

 Set SSOleDBGrid1.DataSource = Nothing



''''  Update Updaterepstatus

sql1 = "SELECT poi_stasdlvy1 = COUNT(*) FROM POITEM WHERE poi_ponumb =  '" & cboPoNumb & "' AND "
sql1 = sql1 & "poi_npecode = '" & FNamespace & "' AND poi_stasdlvy = 'RC'"

    Set SQL_ReceptionDetl = New ADODB.Recordset
    SQL_ReceptionDetl.Open sql1, deIms.cnIms

COUNT1 = SQL_ReceptionDetl("poi_stasdlvy1")

poi_stasdlvy = ""


If COUNT1 = SSOleDBGrid1.Rows Then
poi_stasdlvy = "RC"
End If

If COUNT1 = "0" Then
poi_stasdlvy = "NR"
End If

If ((COUNT1 > "0") And (COUNT1 < SSOleDBGrid1.Rows)) Then
poi_stasdlvy = "RP"
End If






   SQL_PO_Update = "UPDATE PO SET po_tbs='1', po_stasdelv = '" & poi_stasdlvy & "'"
   SQL_PO_Update = SQL_PO_Update & " WHERE po_ponumb =  '" & cboPoNumb & "' AND po_npecode = '" & FNamespace & "'"

    Set SQL_ReceptionDetl = New ADODB.Recordset
    SQL_ReceptionDetl.Open SQL_PO_Update, deIms.cnIms
    
    
For i = 0 To SSOleDBGrid1.Rows - 1
        SSOleDBGrid1.RemoveItem SSOleDBGrid1.Rows - 1
Next i



Call SSOleDBGrid1_populate
    
MsgBox "Reception # " & cboRecepTion1 & " on Purchase Order # " & cboPoNumb & " has been updated."

CancelButton.Caption = "Close"
frmReception.Refresh

End Function
Private Sub dgDetl_Error(ByVal DataError As Integer, response As Integer)
    If DataError <> 0 Then
        notREADY = True
    End If
End Sub

'set store procedurer paratmeters and call it to update po line item
'status

Private Function UPDATEPOITENTOBE() As Boolean
UPDATEPOITENTOBE = False
On Error GoTo ErrHandler
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "UPDATEPOITENTOBE"
        Set .ActiveConnection = deIms.cnIms

        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, frmReception.cboPoNumb) '' cboPoNumb)
        .parameters.Append .CreateParameter("@RECPNUMB", adVarChar, adParamInput, 15, sRecNum)
        .Execute

    End With

   Set cmd = Nothing
   UPDATEPOITENTOBE = True
   Exit Function
ErrHandler:
  MsgBox Err.Description
  Err.Clear
End Function


Private Sub cboPoNumb_GotFocus()
    cboPoNumb.BackColor = &HC0FFFF
End Sub


'set store procedure parameters and call it to update
'reception statues


Private Sub SSOleDBGrid1_populate()
    Dim FillGrid As New ADODB.Recordset
 ' cboPoNumb = frmReception.cboPoNumb
If cboPoNumb <> "" Then

SQL_FillGrid = "SELECT poi_liitnumb, poi_comm, poi_desc, poi_unitprice,"
SQL_FillGrid = SQL_FillGrid & "poi_primreqdqty, poi_qtydlvd, poi_qtytobedlvd, poi_totaprice,"
SQL_FillGrid = SQL_FillGrid & "poi_primuom, uni_desc FROM POITEM LEFT OUTER JOIN UNIT ON "
SQL_FillGrid = SQL_FillGrid & "poi_primuom = uni_code AND poi_npecode = uni_npecode "
SQL_FillGrid = SQL_FillGrid & "WHERE (poi_ponumb = '" & cboPoNumb & "' AND poi_npecode = '" & FNamespace & "')"
SQL_FillGrid = SQL_FillGrid & "ORDER BY CONVERT(integer, poi_liitnumb)"

    Set FillGrid = New ADODB.Recordset
    FillGrid.Open SQL_FillGrid, deIms.cnIms


Do While Not FillGrid.EOF

SSOleDBGrid1.AddItem FillGrid("poi_liitnumb") & vbTab & FillGrid("poi_comm") & vbTab & FillGrid("poi_desc") & vbTab & FillGrid("uni_desc") & vbTab & FillGrid("poi_primreqdqty") & vbTab & FillGrid("poi_qtydlvd") & vbTab & FillGrid("poi_qtytobedlvd") & vbTab & FillGrid("poi_totaprice") & vbTab & FillGrid("poi_primuom") & vbTab & FillGrid("poi_unitprice")


FillGrid.MoveNext

Loop
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid)  'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
   
End Sub

Private Sub SaveButton_Click()
   
Call UpdateReceptiontable
    
    Screen.MousePointer = 0

End Sub
