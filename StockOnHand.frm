VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form StockOnHand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StockOnHand "
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   4725
   Tag             =   "03030400"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "StockOnHand.frx":0000
      CancelVisible   =   0   'False
      PreviousVisible =   0   'False
      NewVisible      =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      SaveVisible     =   0   'False
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
      NewEnabled      =   0   'False
      SaveEnabled     =   0   'False
      CancelEnabled   =   0   'False
      NextEnabled     =   0   'False
      LastEnabled     =   0   'False
      FirstEnabled    =   0   'False
      PreviousEnabled =   0   'False
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcbolocation 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   780
      Width           =   2760
      DataFieldList   =   "Column 0"
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
      stylesets(0).Picture=   "StockOnHand.frx":001C
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
      stylesets(1).Picture=   "StockOnHand.frx":0038
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2540
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5292
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4868
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCompa 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2760
      DataFieldList   =   "Column 0"
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
      stylesets(0).Picture=   "StockOnHand.frx":0054
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
      stylesets(1).Picture=   "StockOnHand.frx":0070
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2858
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5212
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4868
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   2760
      DataFieldList   =   "Column 1"
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
      stylesets(0).Picture=   "StockOnHand.frx":008C
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
      stylesets(1).Picture=   "StockOnHand.frx":00A8
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2540
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5054
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4868
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Output Currency"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   1700
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   1700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1700
   End
End
Attribute VB_Name = "StockOnHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SQL statement get all location list for location combo

Private Sub GetalllocationName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT loc_locacode,loc_name "
        .CommandText = .CommandText & " From location "
        .CommandText = .CommandText & " WHERE loc_npecode = '" & deIms.NameSpace & "'"
         .CommandText = .CommandText & " and (UPPER(loc_gender) <> 'OTHER') "
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by loc_locacode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    ssdcbolocation.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    ssdcbolocation.RemoveAll
    
    rst.MoveFirst
      
    'ssdcbolocation.AddItem (("ALL" & STR) & "ALL" & "")
    Do While ((Not rst.EOF))
        ssdcbolocation.AddItem rst!loc_locacode & str & (rst!loc_name & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetalllocationName", Err.Description, Err.number, True)
End Sub




'SQL statement get all currency list for currency combo

Private Sub GetCurrencylist()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset
Dim flagCURR
        
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT curr_code, curr_desc "
        .CommandText = .CommandText & " FROM CURRENCY "
        .CommandText = .CommandText & " WHERE curr_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " and loc_compcode = '" & SSOleDBCompany.Columns(0).Text & "'"
        .CommandText = .CommandText & " order by curr_code"
         Set rst = .Execute
    End With


    str = Chr$(1)
    SSOleDBCurrency.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
       
    SSOleDBCurrency.RemoveAll
    
    'rst.MoveFirst
      
    'FG 9/2 stupid to add ALL !
    'SSOleDBCurrency.AddItem (("ALL" & STR) & "ALL" & "")
    'Do While ((Not rst.EOF))
    '    SSOleDBCurrency.AddItem rst!curr_code & STR & (rst!curr_desc & "")
    '    rst.MoveNext
    'Loop
      
    '9/2 Juan USD by default
    Do While ((Not rst.EOF))
        SSOleDBCurrency.AddItem (rst!curr_code & "") & str & rst!curr_desc
        If rst!curr_code = "USD" Then flagCURR = SSOleDBCurrency.Rows - 1
        rst.MoveNext
    Loop
    SSOleDBCurrency.Bookmark = flagCURR
    SSOleDBCurrency.Text = SSOleDBCurrency.Columns("Description").Text
    
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::Getcurrencylist", Err.Description, Err.number, True)
End Sub

'SQL statement get company list for company combo

Private Sub GetCampanyName()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = " SELECT com_compcode, com_name "
        .CommandText = .CommandText & " From Company "
        .CommandText = .CommandText & " WHERE com_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " order by com_compcode"
         Set rst = .Execute
    End With



    str = Chr$(1)
    SSOleDBCompa.FieldSeparator = str
    If rst.RecordCount = 0 Then GoTo CleanUp
       
    rst.MoveFirst
      
    'SSOleDBCompa.AddItem (("ALL" & STR) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCompa.AddItem rst!com_compcode & str & (rst!com_name & "")
        
        rst.MoveNext
    Loop
    
    SSOleDBCompa.Bookmark = 0
    'SSOleDBCompa.text = "ALL"
    
    ssdcbolocation.Bookmark = 0
    'ssdcbolocation.text = "ALL"
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::GetCampanyName", Err.Description, Err.number, True)
End Sub

'SQL statement get all company list for company combo
'
'Private Sub GetALLCampanyName()
'On Error Resume Next
'Dim STR As String
'Dim cmd As ADODB.Command
'Dim rst As ADODB.Recordset
'
'
'    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
'
'    With cmd
'        .CommandText = " SELECT com_compcode, com_name "
'        .CommandText = .CommandText & " From Company "
''        .CommandText = .CommandText & " WHERE com_npecode = '" & deIms.NameSpace & "'"
'        .CommandText = .CommandText & " order by com_compcode"
'         Set rst = .Execute
'    End With
'
'
'
'    STR = Chr$(1)
'    SSOleDBCompa.FieldSeparator = STR
'    If rst.RecordCount = 0 Then GoTo CleanUp
'
'    rst.MoveFirst
'
'    SSOleDBCompa.AddItem (("ALL" & STR) & "ALL" & "")
'    Do While ((Not rst.EOF))
'        SSOleDBCompa.AddItem rst!com_compcode & STR & (rst!com_name & "")
'
'        rst.MoveNext
'    Loop
'
'
'
'CleanUp:
'    rst.Close
'    Set cmd = Nothing
'    Set rst = Nothing
'If Err Then Call LogErr(Name & "::GetALLCampanyName", Err.Description, Err.Number, True)
'End Sub

'Load from and form caption text

Private Sub Form_Load()
Dim rs As ADODB.Recordset

    'Added by Juan (9/27/2000) for Multilingual
    Call translator.Translate_Forms("StockOnHand")
    '------------------------------------------

    'Me.Height = 2650
    'Me.Width = 4425
    
    Call GetCampanyName
    Call GetCurrencylist
    Call GetalllocationName
    'SSOleDBCurrency = "USD"

    StockOnHand.Caption = StockOnHand.Caption + " - " + StockOnHand.Tag
    
    With StockOnHand
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'get crystal report parameters and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    
    With MDI_IMS.CrystalReport1
        .Reset
        Printer.Orientation = 2  'Juan 2012/6/5
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "locacode;" + Trim$(ssdcbolocation.Columns("Code").Text) + ";TRUE"
        .ParameterFields(2) = "compcode;" + IIf(Trim$(SSOleDBCompa.Columns("code").Text) = "ALL", "ALL", SSOleDBCompa.Columns("code").Text) + ";TRUE"
        .ParameterFields(3) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(3) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(2) = "compcode;" + Trim$(SSOleDBCompa.Columns("code").Text) + ";TRUE"
        .ReportFileName = reportPath & "onhand.rpt"
        
        'Modified by Juan (9/27/2000) for Multilingual
        msg1 = translator.Trans("M00184") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock On hand", msg1) 'J modified
        Call translator.Translate_Reports("onhand.rpt") 'J added
        Call translator.Translate_SubReports 'J added
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

'set location combo format

Private Sub ssdcbolocation_DropDown()

    'Modified by Juan (9/27/2000) for Multilingual
    msg1 = translator.Trans("L00050") 'J added
    msg2 = translator.Trans("L00028") 'J modified
    ssdcbolocation.Columns(1).Caption = IIf(msg1 = "", "Name", msg1) 'J modified
    ssdcbolocation.Columns(0).Caption = IIf(msg2 = "", "Code", msg2) 'J modified
    '---------------------------------------------
    
    ssdcbolocation.Columns(0).Width = 900
    ssdcbolocation.Columns(1).Width = 2000
End Sub

Private Sub ssdcboLocation_GotFocus()
Call HighlightBackground(ssdcbolocation)
End Sub

Private Sub ssdcboLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcbolocation.DroppedDown Then ssdcbolocation.DroppedDown = True
End Sub

Private Sub ssdcboLocation_KeyPress(KeyAscii As Integer)
'ssdcbolocation.MoveNext
End Sub

Private Sub ssdcboLocation_LostFocus()
Call NormalBackground(ssdcbolocation)
End Sub

Private Sub ssdcbolocation_Validate(Cancel As Boolean)
If Len(Trim$(ssdcbolocation)) > 0 Then
         If Not ssdcbolocation.IsItemInList Then
                ssdcbolocation.Text = ""
            End If
            If Len(Trim$(ssdcbolocation)) = 0 Then
            ssdcbolocation.SetFocus
            Cancel = True
            End If
            End If
End Sub



'SQL statement get location information

Private Sub SSOleDBCompa_Click()
Dim rs As ADODB.Recordset
Dim str As String


 Set rs = New ADODB.Recordset
 
    ssdcbolocation = ""
    
    If Trim$(SSOleDBCompa) = "ALL" Then
    
        With rs
            .LockType = adLockReadOnly
            .CursorLocation = adUseServer
            .CursorType = adOpenForwardOnly
            Set .ActiveConnection = deIms.cnIms
            
            .Source = "Select loc_locacode, loc_name from LOCATION"
            .Source = .Source & " where loc_npecode = '" & deIms.NameSpace & "'"
            .Source = .Source & " and (UPPER(loc_gender) <> 'OTHER') "
            .Source = .Source & " order by loc_locacode "
            .Open
            ssdcbolocation.DataMode = ssDataModeAddItem
            
            ssdcbolocation.RemoveAll
            
            
            str = Chr$(1)
            ssdcbolocation.FieldSeparator = str
            rs.MoveFirst
              
            'ssdcbolocation.AddItem (("ALL" & STR) & "ALL" & "")
            Do While Not .EOF
                ssdcbolocation.AddItem rs!loc_locacode & str & (rs!loc_name & "")
                .MoveNext
            Loop
            

        
        End With

    Else
    
        With rs
            .LockType = adLockReadOnly
            .CursorLocation = adUseServer
            .CursorType = adOpenForwardOnly
            Set .ActiveConnection = deIms.cnIms
            
            .Source = "Select loc_locacode, loc_name from LOCATION"
            .Source = .Source & " where loc_npecode = '" & deIms.NameSpace & "'"
            .Source = .Source & " and (UPPER(loc_gender) <> 'OTHER') "
            .Source = .Source & " and loc_compcode = '" & SSOleDBCompa & "'"
            .Source = .Source & " order by loc_locacode "
            .Open
            
            ssdcbolocation.DataMode = ssDataModeAddItem
            
            ssdcbolocation.RemoveAll
            
            str = Chr$(1)
            ssdcbolocation.FieldSeparator = str
            
            rs.MoveFirst
            
            'ssdcbolocation.AddItem (("ALL" & STR) & "ALL" & "")
            Do While Not .EOF
                ssdcbolocation.AddItem rs!loc_locacode & str & (rs!loc_name & "")
                .MoveNext
            Loop
            
        End With
    End If
    ssdcbolocation.Bookmark = 0
    'ssdcbolocation.text = "ALL"
    
End Sub


Private Sub SSOleDBCompa_GotFocus()
Call HighlightBackground(SSOleDBCompa)
End Sub

Private Sub SSOleDBCompa_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCompa.DroppedDown Then SSOleDBCompa.DroppedDown = True
End Sub

Private Sub SSOleDBCompa_KeyPress(KeyAscii As Integer)
'SSOleDBCompa.MoveNext
End Sub

Private Sub SSOleDBCompa_LostFocus()
Call NormalBackground(SSOleDBCompa)
End Sub

Private Sub SSOleDBCompa_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCompa)) > 0 Then
         If Not SSOleDBCompa.IsItemInList Then
                SSOleDBCompa.Text = ""
            End If
            If Len(Trim$(SSOleDBCompa)) = 0 Then
            SSOleDBCompa.SetFocus
            Cancel = True
            End If
            End If
End Sub

Private Sub SSOleDBCurrency_GotFocus()
Call HighlightBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
If Not SSOleDBCurrency.DroppedDown Then SSOleDBCurrency.DroppedDown = True
End Sub

Private Sub SSOleDBCurrency_KeyPress(KeyAscii As Integer)
'SSOleDBCurrency.MoveNext
End Sub

Private Sub SSOleDBCurrency_LostFocus()
Call NormalBackground(SSOleDBCurrency)
End Sub

Private Sub SSOleDBCurrency_Validate(Cancel As Boolean)
If Len(Trim$(SSOleDBCurrency)) > 0 Then
         If Not SSOleDBCurrency.IsItemInList Then
                SSOleDBCurrency.Text = ""
            End If
            If Len(Trim$(SSOleDBCurrency)) = 0 Then
            SSOleDBCurrency.SetFocus
            Cancel = True
            End If
            End If
End Sub

