VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frmElecDistribution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electronic Distribution"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   6060
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSOleDBDDDisCode 
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5318
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2170
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3413
      _ExtentY        =   1720
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGridList 
      Height          =   3015
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   4
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   2
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1535
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   5
      Columns(1).Width=   2249
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   30
      Columns(2).Width=   3200
      Columns(2).Caption=   "Mail"
      Columns(2).Name =   "Mail"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   59
      Columns(3).Width=   2646
      Columns(3).Caption=   "Fax"
      Columns(3).Name =   "Fax"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   50
      _ExtentX        =   9763
      _ExtentY        =   5318
      _StockProps     =   79
      Caption         =   "List of System Distribution"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frmElectronicDistribution.frx":0000
      DeleteVisible   =   -1  'True
      CloseToolTipText=   ""
      PrintToolTipText=   ""
      EmailToolTipText=   ""
      NewToolTipText  =   ""
      SaveToolTipText =   ""
      CancelToolTipText=   ""
      NextToolTipText =   ""
      LastToolTipText =   ""
      FirstToolTipText=   ""
      PreviousToolTipText=   ""
      DeleteToolTipText=   ""
      EditToolTipText =   ""
      EmailEnabled    =   -1  'True
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "System Distribution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmElecDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rstlist As ADODB.Recordset


Public Sub GetDocumentCode()
On Error Resume Next

Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
       
        .CommandText = " SELECT doc_code, doc_desc "
        .CommandText = .CommandText & " From DOCTYPE "
        .CommandText = .CommandText & " WHERE doc_npecode = '" & deIms.Namespace & "'"
        
        Set rst = .Execute
    End With

    
    
    str = Chr$(1)
    SSOleDBDDDisCode.FieldSeparator = str

    SSOleDBDDDisCode.RemoveAll
    If rst.BOF And rst.EOF Then Exit Sub
    If rst Is Nothing Then Exit Sub
    If rst.RecordCount = 0 Then GoTo CleanUp

    rst.MoveFirst

    Do While ((Not rst.EOF))
        SSOleDBDDDisCode.AddItem rst!doc_desc & str & (rst!doc_code & "")
         rst.MoveNext
    Loop
    
  
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
    
    
End Sub

Public Sub GetTranstypeCode()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        
        .CommandText = " SELECT tty_code, tty_desc "
        .CommandText = .CommandText & " From TRANSACTYPE "
        .CommandText = .CommandText & " WHERE tty_npecode = '" & deIms.Namespace & "'"

        Set rst = .Execute
    End With

    str = Chr$(1)
   SSOleDBDDDisCode.FieldSeparator = str
    
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    Do While ((Not rst.EOF))
        SSOleDBDDDisCode.AddItem rst!tty_desc & str & (rst!tty_code & "")
         rst.MoveNext
    Loop
    
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
        

End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset

    'deIms.cnIms.Open
    'deIms.Namespace = "SAKHA"
    
   
    Call GetDocumentCode
    Call GetTranstypeCode
    Call GetDistributionCode

    SSDBGridList.DataMode = ssDataModeAddItem
    Call Addtogrid(Getlistforgrid(deIms.Namespace, deIms.cnIms))
    Call Addtogridtran(Getlistforgridtran(deIms.Namespace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.Namespace, deIms.cnIms))


    'Call DisableButtons(Me, NavBar1)
End Sub

Public Sub GetDistributionCode()
Dim str As String

    str = Chr(1)
    SSOleDBDDDisCode.FieldSeparator = str

    SSOleDBDDDisCode.AddItem "Update Database" & str & "UD"
    SSOleDBDDDisCode.AddItem "Delivery" & str & "DL"
    SSOleDBDDDisCode.AddItem "Login Security" & str & "LO"
    SSOleDBDDDisCode.AddItem "Shipping" & str & "SH"

End Sub

Public Sub refreshGrid()
    SSDBGridList.RemoveAll
    Call Addtogrid(Getlistforgrid(deIms.Namespace, deIms.cnIms))
    Call Addtogridtran(Getlistforgridtran(deIms.Namespace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.Namespace, deIms.cnIms))
    SSDBGridList.MoveLast
    
End Sub

Private Sub NavBar1_OnCancelClick()
     Call refreshGrid
     SSDBGridList.MoveLast

End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub


Private Sub NavBar1_OnDeleteClick()
    If Len(Trim$(SSDBGridList.Columns("mail").Text)) Then
        Call DeleteUserMail(SSDBGridList.Columns("Mail").Text)
        Call Clearform
        
    ElseIf Len(Trim$(SSDBGridList.Columns("fax").Text)) Then
        Call DeleteUserFax(SSDBGridList.Columns("fax").Text)
        Call Clearform

    End If
        Call refreshGrid

End Sub


Private Sub NavBar1_OnFirstClick()
On Error Resume Next
    SSDBGridList.MoveFirst
End Sub

Private Sub NavBar1_OnLastClick()
On Error Resume Next
    SSDBGridList.MoveLast
End Sub

Private Sub NavBar1_OnNextClick()
On Error Resume Next

  '  With SSDBGridList
        If Not SSDBGridList.EOF Then
        SSDBGridList.MoveNext

        Else
            With SSDBGridList
                If .EOF Then Exit Sub
            End With
        End If
  ' End With

   
End Sub
Private Sub NavBar1_OnPreviousClick()
On Error Resume Next

   
        If Not SSDBGridList.BOF Then
            SSDBGridList.MovePrevious
        Else
            With SSDBGridList
                If .BOF Then Exit Sub
            End With
        End If
   
    
End Sub
Private Sub NavBar1_OnNewClick()
   
       
    SSDBGridList.RemoveAll
    Call Addtogrid(Getlistforgrid(deIms.Namespace, deIms.cnIms))
    Call Addtogridtran(Getlistforgridtran(deIms.Namespace, deIms.cnIms))
    Call AddtogridCode(GetlistforgridCode(deIms.Namespace, deIms.cnIms))

     SSDBGridList.AddNew
'     Call NavBar1_OnPreviousClick
'     SSDBGridList.MoveNext
     'Call Clearform

End Sub


Public Sub Clearform()
    SSDBGridList.Columns(0).Text = ""
    SSDBGridList.Columns(1).Text = ""
    SSDBGridList.Columns("mail").Text = ""
    SSDBGridList.Columns("fax").Text = ""
'    txt"mail" = ""
'    txtfaxNumb = ""
    
End Sub

Public Function DataValidate()
    If Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
          'SSDBGridList.Columns("mail").Text.SetFocus:
          Exit Function
       ElseIf Len(Trim$(SSDBGridList.Columns("fax").Text)) > 0 Then
          MsgBox "You only can select one Email or Fax"
       End If
    Else
        'txtfaxNumb.SetFocus:
        Exit Function
    End If
    
End Function


Private Sub InsertElecDistribution()
On Error GoTo Noinsert
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "INSERT_DISTRIBUTION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.Namespace)
        .Parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "SYS")
        .Parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 5, SSOleDBDDDisCode.Columns("Code").Text)
        .Parameters.Append .CreateParameter("@MAIL", adVarChar, adParamInput, 59, SSDBGridList.Columns("Mail").Text)
        .Parameters.Append .CreateParameter("@FAXNUMB", adVarChar, adParamInput, 50, SSDBGridList.Columns("fax").Text)
        
        .Execute , , adExecuteNoRecords

    End With

    Set cmd = Nothing
    
    SSDBGridList.MovePrevious
    Exit Sub

Noinsert:
        If Err Then Err.Clear
        MsgBox "Insert into Distribution is failure "

End Sub

Private Sub NavBar1_OnSaveClick()

    If Not Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
       If Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
            Call txtmailValidate(True)

       ElseIf Not Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
          MsgBox "You only can select one Email or Fax"
          'txtMail.SetFocus:
          Exit Sub
       End If
    Else
        If Not Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
            Call txtfaxnumber_validate(True)

         ElseIf Not Len(Trim$(SSDBGridList.Columns("mail").Text)) = 0 Then
             MsgBox "You only can select one Email or Fax"
             'txtfaxNumb.SetFocus:
             Exit Sub
        End If
    End If

 If Len(Trim$(SSDBGridList.Columns("MAIL").Text)) = 0 And Len(Trim$(SSDBGridList.Columns("fax").Text)) = 0 Then
        MsgBox "You cannot leave Email and Fax empty"
 End If
    


End Sub

Public Function GetListofdistribution(Namespace As String, Gender As String, cn As ADODB.Connection) As Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " WHERE dis_npecode = '" & Namespace & "'"
        '.CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("MAIL").Text & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " ORDER BY dis_id "
        Set Rstlist = .Execute

    End With

    If Rstlist Is Nothing Then Exit Function
        

        
        Set cmd = Nothing
        'Set Rstlist = Nothing
        
End Function

Public Function GetDistributionMail(EmailNunber As String) As Boolean
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").Text & "'"
        .CommandText = .CommandText & " AND dis_gender =  'SYS' "
        .CommandText = .CommandText & " AND dis_mail = '" & SSDBGridList.Columns("MAIL").Text & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionMail = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
End Function

Public Function GetDistributionFaxnumb(FaxNumber As String) As Boolean

Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns("code").Text & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & SSDBGridList.Columns("FAX").Text & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Set rst = .Execute
        GetDistributionFaxnumb = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing

End Function

Public Function txtmailValidate(Cancel As Boolean) As Boolean
        
    Cancel = False
    
    If Len(SSDBGridList.Columns("mail").Text) Then
        
        If GetDistributionMail(SSDBGridList.Columns("Mail").Text) Then
'            SSDBGridList.Columns("Mail").Text = ""
            MsgBox "This Mail number is exist"
            'txtMail.SetFocus:
            Exit Function
        Else
            Cancel = False
            Call txtfaxnumber_validate(True)
            Call InsertElecDistribution
            'txtfaxNumb.SetFocus:
            Exit Function
        End If
    
    End If
End Function

Public Function txtfaxnumber_validate(Cancel As Boolean) As Boolean
    Cancel = False
    
       If Len(SSDBGridList.Columns("fax").Text) Then
        
        If GetDistributionFaxnumb(SSDBGridList.Columns("fax").Text) Then
            'txtfaxNumb.Text = ""
            MsgBox "This Fax number is exist"
            'txtfaxNumb.SetFocus: Exit Function
            'SSDBGridList.Columns(fax).Text Exit Function
        Else
            Cancel = False
            Call InsertElecDistribution
            'txtMail.SetFocus:
            Exit Function
        End If
    
    End If
    

End Function


Public Function DeleteUserMail(Mail As String) As Boolean
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS'"
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_mail ='" & Mail & "'"

        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        
        Call .Execute(0, 0, adExecuteNoRecords)
        
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            MsgBox "Delete from Distribution is failure "
        Else
            SSDBGridList.Columns("Mail").Text = ""
            'txtMail.SetFocus
        End If
End Function

Public Function DeleteUserFax(fax As String) As Boolean
On Error GoTo NoDelete
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "DELETE FROM DISTRIBUTION"
        .CommandText = .CommandText & " Where dis_npecode = '" & deIms.Namespace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " AND dis_id = '" & SSDBGridList.Columns(0).Text & "'"
        .CommandText = .CommandText & " AND dis_faxnumb ='" & fax & "'"
        .Parameters.Append .CreateParameter("RT", adInteger, adParamOutput, 4)
        

        
        Call .Execute(0, 0, adExecuteNoRecords)
       
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
NoDelete:
        If Err Then
            Err.Clear
            MsgBox "Delete from Distribution is failure "
        Else
            SSDBGridList.Columns("FAX").Text = ""
            
        End If
End Function


Private Sub LoadValues()
On Error Resume Next

    SSDBGridList.Columns("Mail").Text = Rstlist!dis_mail & ""
    SSDBGridList.Columns("FAx").Text = Rstlist!dis_mail & ""

    If Err Then Err.Clear
End Sub


Public Sub Addtogrid(Rstlist As ADODB.Recordset)
Dim str As String
    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
    
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!doc_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop
End Sub

Public Sub Addtogridtran(Rstlist As ADODB.Recordset)
Dim str As String
    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
    
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & Rstlist!tty_desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop
End Sub
Public Sub AddtogridCode(Rstlist As ADODB.Recordset)
Dim str As String
Dim desc As String

    If Rstlist Is Nothing Then Exit Sub
    If Rstlist.EOF And Rstlist.BOF Then Exit Sub
    If Rstlist.RecordCount = 0 Then Exit Sub
    
    
    str = Chr(1)
    SSDBGridList.FieldSeparator = Chr(1)
    
    Do While Not Rstlist.EOF
        If Trim$(Rstlist!dis_id) = "UD" Then desc = "Update Database"
        If Trim$(Rstlist!dis_id) = "DL" Then desc = "Delivery"
        If Trim$(Rstlist!dis_id) = "LO" Then desc = "Login Security"
        If Trim$(Rstlist!dis_id) = "SH" Then desc = "Shipping"
        SSDBGridList.AddItem Rstlist!dis_id & "" & str & desc & "" & str & Rstlist!dis_mail & "" & str & Rstlist!dis_faxnumb & ""
        
        Rstlist.MoveNext
    Loop

End Sub


Public Function Getlistforgrid(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, doc_desc, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION, doctype "
        .CommandText = .CommandText & " WHERE dis_npecode = doc_npecode "
        .CommandText = .CommandText & " and dis_id = doc_code and "
        .CommandText = .CommandText & " dis_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgrid = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

Public Function Getlistforgridtran(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, tty_desc, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION,  TRANSACTYPE "
        .CommandText = .CommandText & " WHERE dis_npecode =  tty_npecode "
        .CommandText = .CommandText & " and dis_id = tty_code and "
        .CommandText = .CommandText & " dis_npecode = '" & Namespace & "'"
        .CommandText = .CommandText & " AND dis_gender = 'SYS' "
        .CommandText = .CommandText & " order by dis_id "
        Set Getlistforgridtran = .Execute

    End With

       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

Public Function GetlistforgridCode(Namespace As String, cn As ADODB.Connection) As ADODB.Recordset
Dim str As String
Dim cmd As ADODB.Command

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT  dis_id, dis_mail, dis_faxnumb "
        .CommandText = .CommandText & " From DISTRIBUTION "
        .CommandText = .CommandText & " where ((dis_id IN ( 'ud', 'sh', 'lo', 'dl')) "
        .CommandText = .CommandText & " AND (dis_gender = 'SYS') "
        .CommandText = .CommandText & " AND (dis_npecode = '" & Namespace & "'))"
        .CommandText = .CommandText & " order by dis_id "
        

        Set GetlistforgridCode = .Execute

    End With


       
    If Rstlist Is Nothing Then Exit Function
    
    Set cmd = Nothing
   
        
End Function

Private Sub SSDBGridList_InitColumnProps()

    SSDBGridList.Columns(0).DropDownHwnd = SSOleDBDDDisCode.hwnd

End Sub

Private Sub SSOleDBDDDisCode_Click()
    
            SSDBGridList.MoveLast
            SSDBGridList.MoveNext
            SSDBGridList.Columns("code").Text = SSOleDBDDDisCode.Columns("code").Text
            SSDBGridList.Columns("description").Text = SSOleDBDDDisCode.Columns("description").Text
   
End Sub



