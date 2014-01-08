VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_StockSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Master Search"
   ClientHeight    =   7665
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   9270
   Tag             =   "02010200"
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   5160
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   396
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtStockNumb 
      Height          =   315
      Left            =   2220
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin VB.TextBox txt_ShortDescript 
      Height          =   288
      Left            =   2220
      TabIndex        =   3
      Top             =   1740
      Width           =   6915
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   600
      Width           =   2325
   End
   Begin VB.ComboBox cboStockType 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Top             =   930
      Width           =   2325
   End
   Begin MSDataGridLib.DataGrid dgSearchList 
      Height          =   4545
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8017
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   50
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "stk_stcknumb"
         Caption         =   "Stock Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "stk_desc"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column01 
            WrapText        =   -1  'True
            ColumnWidth     =   7515.213
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_Close 
      Caption         =   "Cancel"
      Height          =   396
      Left            =   7800
      TabIndex        =   7
      Top             =   2505
      Width           =   1335
   End
   Begin VB.CommandButton cmd_Search 
      Caption         =   "Search"
      Height          =   396
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblDNUmb 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2220
      TabIndex        =   14
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label LabelNumb 
      Caption         =   "Number of records found :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1365
      Width           =   2055
   End
   Begin VB.Label lbl_Waiting 
      Caption         =   "Messages"
      Height          =   390
      Left            =   3120
      TabIndex        =   11
      Top             =   2475
      Visible         =   0   'False
      Width           =   4470
   End
   Begin VB.Label lbl_ShortDescript 
      BackStyle       =   0  'Transparent
      Caption         =   "Short Description"
      Height          =   225
      Left            =   105
      TabIndex        =   10
      Top             =   1740
      Width           =   1995
   End
   Begin VB.Label lbl_StockNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Number"
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   285
      Width           =   2000
   End
   Begin VB.Label lbl_StockType 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Type"
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   915
      Width           =   2000
   End
   Begin VB.Label lbl_Category 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   600
      Width           =   2000
   End
End
Attribute VB_Name = "frm_StockSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim ShowOnlyActive As Boolean
Dim retval As Boolean, fshoweditor As Boolean

Event Unloading(Cancel As Integer)
Event Completed(Cancelled As Boolean, sStockNumber As String)

Private Sub cboCategory_GotFocus()
Call HighlightBackground(cboCategory)
End Sub

'if data entry not character and number get out function
'if return key enter start search

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 8) Or (KeyAscii > 31)) Then _
        Call GetNearestComboItem(cboCategory, KeyAscii)
        
    If KeyAscii = vbKeyReturn Then
        Call cmd_Search_Click
    End If
End Sub

Private Sub cboCategory_LostFocus()
Call NormalBackground(cboCategory)
End Sub

Private Sub Command1_Click()
On Error Resume Next
Screen.MousePointer = 11
Dim list As String
Dim i As Integer
    list = ""
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                
                list = IIf(list = "", Chr(34), list + ", " + Chr(34)) + RTrim(.Fields(0)) + Chr(34)
                .MoveNext
            Loop
            
            If Len(list) > 255 Then
                MsgBox "Your selection is too long, please reduce the number of records to print"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        .MoveFirst
    End With
    With dgSearchList
    End With
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = reportPath & "Stckmaster1.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stcknumb;a;TRUE"
        .ReplaceSelectionFormula ("{STOCKMASTER.stk_npecode} = {?namespace} and  " _
            & "{STOCKMASTER.stk_stcknumb} in [" + list + "]")
        msg1 = translator.Trans("L00119")
        .WindowTitle = IIf(msg1 = "", "Stock Master", msg1)
        Call translator.Translate_Reports("Stckmaster1.rpt")
        .Action = 1
        .Reset
    End With
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If List1.list(i) = txtStockNumb Then Exit Sub
    Next
    List1.AddItem txtStockNumb
End Sub


Private Sub Command3_Click()
    If List1.ListIndex >= 0 Then List1.RemoveItem List1.ListIndex
End Sub


Private Sub txt_ShortDescript_GotFocus()
Call HighlightBackground(txt_ShortDescript)
End Sub

Private Sub txt_ShortDescript_LostFocus()
Call NormalBackground(txt_ShortDescript)
End Sub

'validate stock number

Private Sub txtStockNumb_Change()
On Error Resume Next

Dim i As Integer

    i = Len(txtStockNumb)
    If i = 0 Then Exit Sub
    rst.Filter = adFilterNone
    rst.Filter = "stk_stcknumb like '" & Trim(txtStockNumb) & "%" & "'"
       
    If Not rst.EOF Then
        txtStockNumb = Trim$(rst!stk_stcknumb)
        
        txtStockNumb.SelStart = i
        txtStockNumb.SelLength = Len(txtStockNumb) - i
    End If
    If Err Then Err.Clear
End Sub

'on stock type combo, if enter not character and number  get out function
'if enter equal to return key start search

Private Sub cboStockType_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 8) Or (KeyAscii > 31)) Then _
        Call GetNearestComboItem(cboStockType, KeyAscii)
        
    If KeyAscii = vbKeyReturn Then
        Call cmd_Search_Click
    End If
End Sub

'close form, free memory

Private Sub cmd_Close_Click()
    Hide
    Set rs = Nothing
    Set rst = Nothing
    
    retval = False
    If open_forms <= 5 Then ShowNavigator
End Sub

'depend on condition search stock numbers
'uppcase equal to lowcase

Private Sub cmd_Search_Click()
On Error Resume Next

    Dim str() As String
    Dim i As Integer, x As Integer

    
    lbl_Waiting.Visible = True
    lbl_Waiting.Caption = LoadResString(6)

    Dim sSqlWhere As String
        
    rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    
    rs.LockType = adLockReadOnly
    rs.CursorType = adOpenForwardOnly
    rs.ActiveConnection = deIms.cnIms
     
    sSqlWhere = " where stk_npecode = '" & deIms.NameSpace & "'"

    If Len(cboCategory) > 0 Then
        rs.Open ("select cate_catecode from CATEGORY where cate_catename = '" & cboCategory & "' and cate_npecode = '" & deIms.NameSpace & "'")
        
        If Not IsNull(rs!cate_catecode) Then _
             sSqlWhere = " where stk_catecode = '" & Trim$(rs!cate_catecode) & "' and stk_npecode = '" & deIms.NameSpace & "' "

          '  sSqlWhere = " and stk_catecode = '" & Trim$(rs!cate_catecode) & "' and stk_npecode = '" & deIms.NameSpace & "' "
'            sSqlWhere = " where stk_catecode = '" & Trim$(rs!cate_catecode) & "'"

        rs.Close

        End If
    
    If Len(cboStockType) > 0 Then
        rs.Open ("select sty_stcktype from STOCKTYPE where sty_desc = '" & cboStockType & "'")
        
        If Not IsNull(rs!sty_stcktype) Then
        
            If Len(sSqlWhere) Then
                sSqlWhere = sSqlWhere & " and"
            Else: sSqlWhere = " where "
            End If
            
            sSqlWhere = sSqlWhere & " stk_stcktype = '" & Trim$(rs!sty_stcktype) & "'"
        
        End If

        
        rs.Close
    End If
    
    If Len(txtStockNumb) > 0 Then
    
        If Len(sSqlWhere) Then
            sSqlWhere = sSqlWhere & " and"
        Else: sSqlWhere = " where "
        End If
            
        sSqlWhere = sSqlWhere & " stk_stcknumb like '" & Trim$(txtStockNumb) & "' and stk_npecode = '" & deIms.NameSpace & "' "

    End If
    
    If Len(txt_ShortDescript) Then
    
        If Len(sSqlWhere) Then
            sSqlWhere = sSqlWhere & " and"
           Else: sSqlWhere = " where "
        End If
        
        'sSqlWhere = sSqlWhere & " stk_desc like '%" & Replace(txt_ShortDescript, ",", "%") & "%'"
    
        str = Split(txt_ShortDescript, ",", -1, vbTextCompare)
        
        On Error Resume Next
        
        Dim P As Integer, Y As Integer, s As String
        
        Err.Clear
        
        x = UBound(str)
        i = LBound(str)
        
        P = IIf(x >= i, 1, -1)
        
        'x = x + 1
        For Y = i To x Step P
            s = s & " stk_desc like " & "'%" & str(Y) & "%'" & IIf(Y = x, "", " AND ")
        Next
        Debug.Print s
         sSqlWhere = sSqlWhere & s

        sSqlWhere = sSqlWhere & " and stk_npecode = '" & deIms.NameSpace & "'"
        
            
    End If
    
    
        
    
    If ShowOnlyActive Then sSqlWhere = sSqlWhere & " and stk_flag <> 0 "
    sSqlWhere = sSqlWhere & " ORDER BY stk_stcknumb"

    If Len(cboCategory) = 0 And Len(cboStockType) = 0 And Len(txtStockNumb) = 0 And Len(txt_ShortDescript) = 0 Then
        rs.Open ("select stk_stcknumb, stk_desc, stk_npecode, stk_flag from STOCKMASTER " & sSqlWhere)
       
    Else
       rs.Open ("select stk_stcknumb, stk_desc, stk_npecode, stk_flag from STOCKMASTER " & sSqlWhere)
    End If
    
    lbl_Waiting.Visible = False
    Set dgSearchList.DataSource = rs.DataSource
    
        LblDNUmb = rs.RecordCount
    
    If Err Then Call LogErr(Name & "::cmd_Search_Click", Err.Description, Err.number, True)
    
End Sub

'on search if error cause show message

Private Sub dgSearchList_DblClick()
On Error Resume Next

    If rs!stk_flag Then
        'Hide
        retval = True
    
        RaiseEvent Completed(False, rs!stk_stcknumb)
    Else
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00542") 'J added
        MsgBox IIf(msg1 = "", "Selected Stock record is inactive", msg1) 'J modified
        '---------------------------------------------
        
    End If
    
    If Err Then Call LogErr(Name & "::dgSearchList_DblClick", Err.Description, Err.number, True)
End Sub

Private Sub dgSearchList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
'Dim sFilter As String, BK As Variant
'
'    If Button = vbRightButton Then
'        frm_StockRecord.Show
'
'        BK = rs.Bookmark
'        deIms.rsSTOCKMASTER.Filter = adFilterNone
'
'        rs.MoveLast
'        rs.Bookmark = BK
'        If rs.RecordCount < 1 Then Exit Sub
'        sFilter = "stk_stcknumb = '" & rs!stk_stcknumb & "'"
'        sFilter = sFilter & " and stk_npecode ='" & rs!stk_npecode & "'"
'
'        Debug.Print "StockSearch.dgSearchList_MouseUp"
'        Debug.Print sFilter
'        deIms.rsSTOCKMASTER.Filter = sFilter
'    End If
'
End Sub

'SQL statement get stocktype and populate combo
'and stock master number

Private Sub Form_Load()

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_StockSearch")
    '------------------------------------------

    'deIms.cnIms.Open
    Set rs = New ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    Set rst.ActiveConnection = deIms.cnIms
    rst.Open ("Select sty_desc From STOCKTYPE order by sty_desc")
    
    Call PopuLateFromRecordSet(cboStockType, rst, "sty_desc", False)
    
    rst.Close
        
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    Set rst.ActiveConnection = deIms.cnIms
    Call rst.Open("Select cate_catename from CATEGORY where cate_npecode = '" & deIms.NameSpace & "' ")
    Call PopuLateFromRecordSet(cboCategory, rst, "cate_catename", False)
    
    rst.Close
    
    rst.LockType = adLockReadOnly
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    Set rst.ActiveConnection = deIms.cnIms
    'NO DISABLE BUTTONS BECAUSE THIS FORM IS FOR READ ONLY
    'Call DisableButtons(Me, Nothing)
    Call rst.Open("Select stk_stcknumb from STOCKMASTER where stk_npecode = '" & deIms.NameSpace & "'")
    
    Caption = Caption + " - " + Tag

End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Unloading(Cancel)
    
    Hide
    Set rs = Nothing
    Set rst = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'enter return and comon start search function

Private Sub txt_ShortDescript_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call cmd_Search_Click
    ElseIf KeyAscii = Asc(",") Then
        Call cmd_Search_Click
    End If
    
End Sub

'call active function

Public Function Execute() As Boolean
    ShowOnlyActive = True
    Call Show: WindowState = 0: Call Move(0, 0)
End Function

'assign value

Public Property Get StockNumber() As String
    StockNumber = rs!stk_stcknumb
End Property

'assign value

Public Property Get ShowEditor() As Boolean
    ShowEditor = fshoweditor
End Property

'assign value

Public Property Let ShowEditor(ByVal NewVal As Boolean)
    fshoweditor = NewVal
End Property

'assign value

Public Function Description() As String
    If Not (rs Is Nothing) Then _
        If Not IsNull(rs!stk_desc) Then _
            Description = IIf(IsNull(rs!stk_desc), "", rs!stk_desc)
End Function

Private Sub txtStockNumb_GotFocus()
Call HighlightBackground(txtStockNumb)
End Sub

'depend keybroad  entry start search or clear text

Private Sub txtStockNumb_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 8 Then
        txtStockNumb.SelStart = txtStockNumb.SelStart - 1
        txtStockNumb.SelLength = Len(txtStockNumb) - txtStockNumb.SelStart
        
        KeyAscii = 0
        txtStockNumb.SelText = ""
    
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call cmd_Search_Click
        
    Else
        KeyAscii = Asc(Chr(UCase(KeyAscii)))
        
    End If

    If Err Then Err.Clear
End Sub

Private Sub txtStockNumb_LostFocus()
Call NormalBackground(txtStockNumb)
End Sub
