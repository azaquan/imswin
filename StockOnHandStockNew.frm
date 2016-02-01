VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmStockOnHandStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock On Hand per Stock Number"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   5085
   Tag             =   "03030900"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   1860
      TabIndex        =   2
      Top             =   1080
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "StockOnHandStock.frx":0000
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboStockNumb 
      Bindings        =   "StockOnHandStock.frx":001C
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   180
      Width           =   2760
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   "(Space)"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "StockOnHandStock.frx":0027
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "StockOnHandStock.frx":0043
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   5292
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   4868
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCurrency 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2760
      DataFieldList   =   "Column 1"
      AllowInput      =   0   'False
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
      stylesets(0).Picture=   "StockOnHandStock.frx":005F
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
      stylesets(1).Picture=   "StockOnHandStock.frx":007B
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2434
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4974
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Output Currency"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Number"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmStockOnHandStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim rs As ADODB.Recordset
    Screen.MousePointer = 11
    Me.Refresh
    Set rs = New ADODB.Recordset
    With rs
        .LockType = adLockReadOnly
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        Set .ActiveConnection = deIms.cnIms
        
        '.Source = "Select distinct(qs1_stcknumb), qs1_desc from QTYST1"
        '.Source = .Source & " where qs1_npecode = '" & deIms.NameSpace & "'"
        '.Source = .Source & "order by qs1_stcknumb"
        'DoEvents
        '.Open
        
        With ssdcboStockNumb
            Set .DataSourceList = deIms.Commands("getStockOnHandQTYST1").Execute(100, Array(0, deIms.NameSpace))
            .DataFieldToDisplay = "qs1_stcknumb"
            .DataFieldList = "qs1_stcknumb"
            .Refresh
        End With
        
        
        
        'ssdcboStockNumb.DataMode = ssDataModeAddItem
        
        
   '     Do While Not .EOF
'            ssdcboStockNumb.AddItem !qs1_stcknumb & "" & ";" & !qs1_desc & "'"
 '           .MoveNext
  '      Loop
        
      Call GetCurrencylist
       SSOleDBCurrency = "USD"
      
    End With
    Screen.MousePointer = 0
End Sub

'SQL statement get stock numbers list and populate data grids

Private Sub Form_Load()
Dim rs As ADODB.Recordset

    Screen.MousePointer = 0
    
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmStockOnHandStock")
    '------------------------------------------

    frmStockOnHandStock.Caption = frmStockOnHandStock.Caption + " - " + frmStockOnHandStock.Tag
    
    With frmStockOnHandStock
        .Left = Round((Screen.Width - Me.Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub


'SQL statement get all currency list for currency combo

Private Sub GetCurrencylist()
On Error Resume Next
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    
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
    
    rst.MoveFirst
      
'    SSOleDBCurrency.AddItem (("ALL" & str) & "ALL" & "")
    Do While ((Not rst.EOF))
        SSOleDBCurrency.AddItem rst!curr_code & str & (rst!curr_desc & "")
        
        rst.MoveNext
    Loop
      
CleanUp:
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
If Err Then Call LogErr(Name & "::Getcurrencylist", Err.Description, Err.number, True)
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

'get crystal report parmeters
'and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler
    
    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\stckohstck.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .ParameterFields(1) = "stcknumb;" + ssdcboStockNumb.Columns(0).Text + ";TRUE"
        '.ParameterFields(2) = "curr;" + Trim$(SSOleDBCurrency.Columns("code").Text) + ";TRUE"
'        .ParameterFields(2) = "curr;" + IIf(Trim$(SSOleDBCurrency.Columns("code").Text) = "ALL", "ALL", SSOleDBCurrency.Columns("code").Text) + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("M00188") 'J added
        .WindowTitle = IIf(msg1 = "", "Stock On Hand per Stock Number", msg1) 'J modified
        Call translator.Translate_Reports("stckohstck.rpt") 'J added
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

'set stock number combo size

Private Sub ssdcboStockNumb_DropDown()

    'Modified by Juan (9/25/2000) for Multilingual
    msg1 = translator.Trans("L00538") 'J added
    msg2 = translator.Trans("L00029") 'J added
    ssdcboStockNumb.Columns(0).Caption = IIf(msg1 = "", "Number", msg1) 'J modified
    ssdcboStockNumb.Columns(1).Caption = IIf(msg2 = "", "Description", msg2) 'J modified
    '---------------------------------------------

    ssdcboStockNumb.Columns(0).Width = 1000
    ssdcboStockNumb.Columns(1).Width = 4000

    With ssdcboStockNumb
        .MoveNext
    End With

End Sub

Private Sub ssdcboStockNumb_GotFocus()
Call HighlightBackground(ssdcboStockNumb)
End Sub

Private Sub ssdcboStockNumb_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ssdcboStockNumb.DroppedDown Then ssdcboStockNumb.DroppedDown = True
End Sub

Private Sub ssdcboStockNumb_KeyPress(KeyAscii As Integer)
'ssdcboStockNumb.MoveNext
End Sub

Private Sub ssdcboStockNumb_LostFocus()
Call NormalBackground(ssdcboStockNumb)
End Sub

Private Sub ssdcboStockNumb_Validate(Cancel As Boolean)
If Len(Trim$(ssdcboStockNumb)) > 0 Then
         If Not ssdcboStockNumb.IsItemInList Then
                'ssdcboStockNumb.Text = ""
            End If
            If Len(Trim$(ssdcboStockNumb)) = 0 Then
            ssdcboStockNumb.SetFocus
            Cancel = True
            End If
            End If
End Sub


