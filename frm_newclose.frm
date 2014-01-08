VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close/cancel Transaction"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   8235
   Tag             =   "02020300"
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close/Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtReason 
      Height          =   1215
      Left            =   2400
      MaxLength       =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   5415
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboPoNumb 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   2160
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frm_newclose.frx":0000
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frm_newclose.frx":001C
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      RowSelectionStyle=   1
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2302
      Columns(0).Caption=   "PO-Number"
      Columns(0).Name =   "PO-Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   3836
      Columns(1).Caption=   "Buyer"
      Columns(1).Name =   "Buyer"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1826
      Columns(2).Caption=   "Status"
      Columns(2).Name =   "Status"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).HeadStyleSet=   "ColHeader"
      Columns(2).StyleSet=   "RowFont"
      Columns(3).Width=   2037
      Columns(3).Caption=   "Date"
      Columns(3).Name =   "Date"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   7
      Columns(3).FieldLen=   256
      Columns(3).HeadStyleSet=   "ColHeader"
      Columns(3).StyleSet=   "RowFont"
      _ExtentX        =   3810
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label5 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Close"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Reason"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Transaction Number"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Close/Cancel Transaction"
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rowguid, locked As Boolean, dbtablename As String, j1 As Integer, rowguid1 As String

'SQL statement get po information and populate data grid

Public Sub GetTransactionNumber()
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT po_ponumb, po_buyr, po_stas, po_date "
        .CommandText = .CommandText & " From PO "
        .CommandText = .CommandText & " WHERE po_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND (po_stas = 'OP' OR po_stas = 'OH') "
        .CommandText = .CommandText & " order by po_ponumb "
         Set rst = .Execute
    End With
    
    str = Chr$(1)
   ssdcboPoNumb.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
    
    rst.MoveFirst
    
    
    Do While ((Not rst.EOF))
        ssdcboPoNumb.AddItem rst!PO_PONUMB & "" & str & rst!po_buyr & "" & str & rst!po_stas & "" & str & rst!PO_Date & ""
        
        rst.MoveNext
    Loop
  
    
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

Private Sub Form_Activate()
Dim bl As Boolean
    Screen.MousePointer = 11
    Me.Refresh
    DoEvents
    GetTransactionNumber
    bl = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    
    'If BL = False Then
          ssdcboPoNumb.AllowInput = bl
            txtReason.Enabled = bl
            Option1.Enabled = bl
            Option2.Enabled = bl
            cmdClose.Enabled = bl
    Screen.MousePointer = 0
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'set values to option

Private Sub Option1_Click()
    If Option1.value = 0 Then
        Option2.value = 1
    End If
End Sub

Private Sub Option1_GotFocus()
Call HighlightBackground(Option1)
End Sub

Private Sub Option1_LostFocus()
Call NormalBackground(Option1)
End Sub

'Private Sub Option1_Validate(Cancel As Boolean)
'    If Option1.Value = 0 Then
'        Option2.Value = 1
'    End If
'End Sub

'set values to option

Private Sub Option2_Click()
    If Option2.value = 0 Then
        Option1.value = 1
    End If
End Sub
'
'Private Sub Option2_Validate(Cancel As Boolean)
'    If Option2.Value = 0 Then
'        Option1.Value = 1
'    End If
'End Sub

'check option value and validate data, save record

Private Sub cmdClose_Click()
Dim str As String
Dim Result As Boolean
    If Option2.value = Option1.value Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00349") ' J added
        MsgBox IIf(msg1 = "", "Check Close or Cancel.", msg1) 'J modified
        '---------------------------------------------
        
        Option1.SetFocus: Exit Sub
    End If
    
    Call FieldValidate
    
    If Len(Trim$(txtReason)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00350") ' J added
        MsgBox IIf(msg1 = "", "The Reason field cannot be left empty.", msg1) 'J modified
        '---------------------------------------------
        
        txtReason.SetFocus: Exit Sub
    End If
    
    
     If UpdatePotable Then
     
        
        Result = InsertintoPorem
        
        If Result = True Then
           str = IIf(Option1 = True, "closed", "Cancelled")
           MsgBox "The transaction number has been " & str & " sucessfully.", vbInformation, "Imwin"
        End If
     
    End If
      
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode


    
End Sub

'close form

Private Sub cmdExit_Click()
    Unload Me
End Sub

'call function get transaction number

Private Sub Form_Load()
Dim bl As Boolean
    'Added by Juan (9/15/2000) for Multilingual
    Call translator.Translate_Forms("frmClose")
    '------------------------------------------

    
    frmClose.Caption = frmClose.Caption + " - " + frmClose.Tag
    
    
    
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    

Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

    
    
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub Option2_GotFocus()
Call HighlightBackground(Option2)
End Sub

Private Sub Option2_LostFocus()
Call NormalBackground(Option2)
End Sub

Private Sub ssdcboPoNumb_Change()
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

End Sub

'call function

Private Sub ssdcboPoNumb_Click()
    
    If ssdcboPoNumb <> "" Then
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode
 
End If

    
    
    
    
'jawdat, start copy
Dim currentformname, currentformname1
currentformname = Forms(3).Name
currentformname1 = Forms(3).Name
'Dim imsLock As imsLock.lock
Dim ListOfPrimaryControls() As String
Set imsLock = New imsLock.Lock

ListOfPrimaryControls = imsLock.GetPrimaryControls(currentformname, deIms.cnIms)

Call imsLock.Check_Lock(locked, deIms.cnIms, CurrentUser, GetValuesFromControls(ListOfPrimaryControls, Me), currentformname1, rowguid, dbtablename)  'lock should be here, added by jawdat, 2.1.02
If locked = True Then 'sets locked = true because another user has this record open in edit mode



   ssdcboPoNumb = ""

 
  Option1.Enabled = False
  cmdClose.Enabled = False
  Option2.Enabled = False
  txtReason.Enabled = False
Exit Sub                                                     'Exit Edit sub because theres nothing the user can do


  


' 'Dim imsLock As imsLock.lock
'Set imsLock = New imsLock.lock
'Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid1, , dbtablename) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

                                             'Exit Edit sub because theres nothing the user can do
Else

    
  Option1.Enabled = True
  cmdClose.Enabled = True
  Option2.Enabled = True
  txtReason.Enabled = True
    
    
    Call GetTransactionNumber
    
    locked = True
End If
    
End Sub

'depend on option value select, call function to insert data

Public Function UpdatePotable() As Boolean
On Error GoTo Noinsert

UpdatePotable = False

Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    With cmd
        If Option1.value = True Then
            .CommandText = "UPDATEPOCL"
        Else
            .CommandText = "UPDATEPO"
        End If
        
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, ssdcboPoNumb)
        .Execute , , adExecuteNoRecords
    
    End With
    
    Set cmd = Nothing
    
    'Modified by Juan (9/15/2000) for Multilingual
'    Dim msg3 As String 'J added
'    msg1 = translator.Trans("M00343") 'J added
'    msg2 = translator.Trans("L00523") 'J added
'    msg3 = translator.Trans("L00524") 'J added
'    MsgBox IIf(msg1 = "", "This transaction is", msg1 + " ") + IIf(Option1 = True, IIf(msg2 = "", " closed.", msg2), IIf(msg3 = "", " Cancelled", msg3)) 'J modified
    '---------------------------------------------
    UpdatePotable = True
    Exit Function
    
Noinsert:

    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00344") 'J added
    MsgBox IIf(msg1 = "", "Error during the closing of the transction.", msg1) 'J modiefied
    '---------------------------------------------
     Err.Clear
End Function

'call store procedure insert data

Public Function InsertintoPorem() As Boolean
On Error GoTo Noinsert
Dim cmd As ADODB.Command

InsertintoPorem = False

    Set cmd = New ADODB.Command
    
    With cmd
        .CommandText = "INSERTPOREM"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms
        
        
        .parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@ponumb", adVarChar, adParamInput, 15, ssdcboPoNumb)
        .parameters.Append .CreateParameter("@por_remk", adVarChar, adParamInput, 3000, txtReason)
        .parameters.Append .CreateParameter("@USER", adVarChar, adParamInput, 15, CurrentUser)
        
         If Option1.value = True Then
             .parameters.Append .CreateParameter("@CLOSE", adVarChar, adParamInput, 1, 1)
         Else
            .parameters.Append .CreateParameter("@CLOSE", adVarChar, adParamInput, 1, 0)
         End If
                
        Call .Execute(Options:=adExecuteNoRecords)
    
    End With
    
    Set cmd = Nothing
    
    'Modified by Juan (9/15/2000) for Multilingual
    'msg1 = translator.Trans("M00345") 'J aadded
    'MsgBox IIf(msg1 = "", "Insert into PO Remark is completed successfully", msg1) 'J modified
    '---------------------------------------------
    InsertintoPorem = True
    Exit Function
    
Noinsert:

    'Modified by Juan (9/15/2000) for Multilingual
    msg1 = translator.Trans("M00346") 'J added
    MsgBox IIf(msg1 = "", "Insert into PO Remark is failed ", msg1) 'J modified
    '---------------------------------------------
    Err.Clear
End Function

'validate data

Private Function FieldValidate() As Boolean
On Error Resume Next

    FieldValidate = False
    
    If Option1.value = 0 And Option2.value = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00347") 'J added
        MsgBox IIf(msg1 = "", "Select Close or Cancel.", msg1) 'J modified
        '---------------------------------------------
        
             Exit Function
    End If
    
    If Len(Trim(ssdcboPoNumb)) = 0 Then
    
        'Modified by Juan (9/15/2000) for Multilingual
        msg1 = translator.Trans("M00348") 'J added
        MsgBox IIf(msg1 = "", "Transaction Number cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        ssdcboPoNumb.SetFocus: Exit Function
    End If
    
End Function

Private Sub ssdcboPoNumb_GotFocus()
Call HighlightBackground(ssdcboPoNumb)
End Sub

Private Sub ssdcboPoNumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboPoNumb.DroppedDown Then ssdcboPoNumb.DroppedDown = True
End Sub

Private Sub ssdcboPoNumb_LostFocus()
Call NormalBackground(ssdcboPoNumb)
End Sub

Private Sub ssdcboPoNumb_Scroll(Cancel As Integer)
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
Call imsLock.Unlock_Row(locked, deIms.cnIms, CurrentUser, rowguid, , dbtablename, , True) 'jawdat, if user hits neither Cancel nor Save, but just closes the form while in edit mode

End Sub

Private Sub ssdcboPoNumb_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboPoNumb)) > 0 Then
    If Not ssdcboPoNumb.IsItemInList Then
       MsgBox "Please select a valid Transaction number.", vbInformation, "Imswin"
       Cancel = True
       ssdcboPoNumb.SetFocus
    End If
End If
End Sub

Private Sub txtReason_GotFocus()
Call HighlightBackground(txtReason)
End Sub

Private Sub txtReason_LostFocus()
Call NormalBackground(txtReason)
End Sub
