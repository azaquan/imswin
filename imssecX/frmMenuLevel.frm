VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmMenuLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Level"
   ClientHeight    =   5475
   ClientLeft      =   5715
   ClientTop       =   2925
   ClientWidth     =   6570
   Icon            =   "frmMenuLevel.frx":0000
   LinkTopic       =   "frmMenuLevel"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Tag             =   "04010700"
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   840
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4920
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgMenuLevel 
      Height          =   3795
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   6075
      _Version        =   196617
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
      stylesets(0).Picture=   "frmMenuLevel.frx":000C
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmMenuLevel.frx":0028
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
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
      ExtraHeight     =   159
      Columns.Count   =   3
      Columns(0).Width=   2196
      Columns(0).Caption=   "Menu Level"
      Columns(0).Name =   "menulevel"
      Columns(0).DataField=   "ml_melvid"
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   8017
      Columns(1).Caption=   "Level Name"
      Columns(1).Name =   "name"
      Columns(1).DataField=   "ml_melvname"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   50
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NameSpace"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "ml_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   5
      _ExtentX        =   10716
      _ExtentY        =   6694
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Menu Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6105
   End
End
Attribute VB_Name = "frmMenuLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim cn As ADODB.Connection

Private Sub Form_Load()
On Error Resume Next



    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmMenuLevel")
    '--------------------------------------------------

    Set NavBar.Recordset = New ADODB.Recordset
    'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
    
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err, True)
End Sub

'Set conncetion to database, SQL statement to get menu level

Public Sub SetConnection(con As ADODB.Connection)
On Error Resume Next

    Set cn = con
    With NavBar.Recordset
        Set .ActiveConnection = con
        
        .CursorType = adOpenStatic
        .CursorLocation = adUseServer
        .LockType = adLockBatchOptimistic
        
        .Source = "SELECT ml_melvid, ml_melvname,ml_npecode FROM MENULEVEL"
        .Source = .Source & " WHERE ml_npecode = '" & np & "'"
        .Source = .Source & " ORDER BY ml_melvid"
        
        .Open
        
        Set ssdbgMenuLevel.DataSource = NavBar
        Call DisableButtons(Me, NavBar, np, CurrentUser, con)
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End With
    
    Set NavBar.Recordset = NavBar.Recordset
    
    Caption = Caption + " - " + Tag
    If Err Then Call LogErr(Name & "::SetConnection", Err.Description, Err, True)
End Sub

'unload form cancel recordset update and set memory free

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    'NavBar.update
    ssdbgMenuLevel.CancelUpdate
    NavBar.Recordset.CancelUpdate
    NavBar.Recordset.CancelBatch
    'COMMENTED OUT BY MUZAMMIL
    'rEASON - WOULD GENERATE ALL THE TEPLATES AT FORM_UNLOAD WHICH TOOK SOME TIME
    'dATE - 26/01/01
   ' UpdateTemplates
    Set NavBar.Recordset = Nothing
    If Err Then Call LogErr(Name & "::Form_Unload", Err.Description, Err, True)
End Sub

'set recordset cancel update

Private Sub NavBar_BeforeCancelClick()
On Error Resume Next

    ssdbgMenuLevel.CancelUpdate
    If Err Then Call LogErr(Name & "::NavBar_BeforeCancelClick", Err.Description, Err, True)
End Sub

'check menu level exist or not, if it existed show the message
'else make new menu level

Private Sub NavBar_BeforeNewClick()
Dim str As String
Dim BK As Variant
Dim rs As ADODB.Recordset
On Error Resume Next

    Set rs = NavBar.Recordset
    NavBar.AllowAddNew = False
    str = Trim$(InputBox(LoadResString(10000), "Menu Level"))
    
    If rs.EOF Then rs.MoveLast
    BK = rs.Bookmark
    
    If Len(str) = 1 Then
    
        Call rs.Find("ml_melvid = '" & str & "'", 0, adSearchForward, adBookmarkFirst)
        
        If rs.EOF Then
            rs.Bookmark = BK
            ssdbgMenuLevel.AddNew
            
            ssdbgMenuLevel.Columns(2).Text = np
            ssdbgMenuLevel.Columns(0).Text = str
            
        Else
            rs.Bookmark = BK
            Call MsgBox(LoadResString(1004))
        End If
    
    ElseIf Len(str) > 1 Then
        Call MsgBox(LoadResString(1005))
    
    End If
        
    If Err Then Call LogErr(Name & "::NavBar_BeforeNewClick", Err.Description, Err, True)
End Sub

'close form and free memory

Private Sub NavBar_OnCloseClick()
On Error Resume Next

    Unload Me
    Err.Clear
End Sub

'get crystal report parameters and application path

Private Sub NavBar_OnPrintClick()
On Error Resume Next

    With CrystalReport1
        .ReportFileName = ReportPath + "menulevel.rpt"
        .ParameterFields(0) = "namespace;" + np + ";TRUE"
        
        'Modified by Juan (10/23/00) for Multilingual 'J added
        msg1 = Trans("M00208") 'J added
        .WindowTitle = IIf(msg1 = "", "Menu Level", msg1) 'J modified
        Call translate_reports(Me.Name, "menulevel.rpt", True, cn, CrystalReport1) 'J added
        '--------------------------------------------------
        
        .Action = 1
    End With
   If Err Then Call LogErr(Name & "::NavBar_OnPrintClick", Err.Description, Err, True)
End Sub

'cancel recordset update

Private Sub NavBar_OnSaveClick()

Dim cmd As ADODB.Command
Dim id_numb As Integer
On Error Resume Next

If ssdbgMenuLevel.IsAddRow Then
   
   id_numb = ssdbgMenuLevel.Columns(0).Text
    
    ssdbgMenuLevel.update
    
    If Err.Number <> 0 Then
    
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = cn
            
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POPULATEMENUTEMPLATE"
            
            cmd.Prepared = True
            cmd.Parameters.Refresh
            cmd.Parameters(1).Value = np
            cmd.Parameters("@LEVELID").Value = id_numb
                
            
            Call cmd.Execute(0, 0, adExecuteNoRecords)
        
   End If
        
    If Err Then Call LogErr(Name & "::NavBar_OnSaveClick", Err.Description, Err, True)
   
End If

End Sub

'set name space equal to cuurrent name space

Public Sub SetNameSpace(NameSpace As String)
    np = NameSpace
End Sub

'call store procedure to get menutemplate and populate data grid

Private Sub UpdateTemplates()
On Error Resume Next

Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POPULATEMENUTEMPLATE"
    
    cmd.Prepared = True
    cmd.Parameters.Refresh
    cmd.Parameters(1).Value = np
    
    NavBar.Recordset.MoveFirst
    
    Do While Not NavBar.Recordset.EOF
        cmd.Parameters("@LEVELID").Value = NavBar.Recordset!ml_melvid
        
        NavBar.Recordset.MoveNext
        Call cmd.Execute(0, 0, adExecuteNoRecords)
        
    Loop
    
    Set cmd = Nothing
    
   If Err Then Call LogErr(Name & "::UpdateTemplates", Err.Description, Err, True)
End Sub

