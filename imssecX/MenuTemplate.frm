VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmMenuTemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Template"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "MenuTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5745
   Tag             =   "04010800"
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      NewVisible      =   0   'False
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgMenuTemplate 
      Height          =   3135
      Left            =   180
      TabIndex        =   3
      Top             =   1440
      Width           =   5475
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
      stylesets(0).Picture=   "MenuTemplate.frx":000C
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
      stylesets(1).Picture=   "MenuTemplate.frx":0028
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
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
      Columns(0).Width=   5292
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   9657
      _ExtentY        =   5530
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboUserName 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   960
      Width           =   3795
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnHeaders   =   0   'False
      ForeColorEven   =   8388608
      BackColorOdd    =   16771818
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4842
      Columns(0).Caption=   "ml_melvname"
      Columns(0).Name =   "Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3281
      Columns(1).Caption=   "ID"
      Columns(1).Name =   "ID"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6694
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Menu Template"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Level  ID"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmMenuTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim cn As ADODB.Connection

'load form get caption and size

Private Sub Form_Load()

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmMenuTemp")
    '--------------------------------------------------

    Set NavBar.Recordset = New ADODB.Recordset
    Caption = Caption + " - " + Tag
    
    'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub

'set database conncetion and call function to get menulevels
'recordset and set navbar buttom

Public Sub SetConnection(con As ADODB.Connection)
    Set cn = con
    GetMenuLevels
    NavBar.Recordset.ActiveConnection = cn
    Call DisableButtons(Me, NavBar, np, CurrentUser, con)
    ssdcboUserName.AllowInput = True
End Sub

'unload form and close recordset, free memory

Private Sub Form_Unload(Cancel As Integer)
    Set NavBar.Recordset = Nothing
End Sub

'cancel recordset update

Private Sub NavBar_BeforeCancelClick()
    ssdbgMenuTemplate.CancelUpdate
End Sub

'cancel recordset update

Private Sub NavBar_BeforeSaveClick()
    ssdbgMenuTemplate.update
End Sub

'close form

Private Sub NavBar_OnCloseClick()
    Unload Me
End Sub

'get ccrystal report parameters and application path

Private Sub NavBar_OnPrintClick()
    With CrystalReport1
        .ReportFileName = ReportPath + "menutemplate.rpt"
        .ParameterFields(0) = "namespace;" + np + ";TRUE"
        .ParameterFields(1) = "levelid;" + ssdcboUserName.Columns(1).Text + ";TRUE"
        
        'Modified by Juan (10/23/00) for Multilingual 'J added
        Call translate_reports(Me.Name, "menutemplate.rpt", True, cn, CrystalReport1) 'J added
        msg1 = Trans("M00209") 'J added
        .WindowTitle = IIf(msg1 = "", "Menu Template", msg1) 'J modified
        '--------------------------------------------------
        
        .Action = 1
    End With
End Sub

'set recordset to update position

Private Sub NavBar_OnSaveClick()
    NavBar.Recordset.UpdateBatch
End Sub

'set name space to current name space

Public Sub SetNameSpace(NameSpace As String)
    np = NameSpace
End Sub

'SQL statement to get menu level recordset
'and populate recordset

Private Sub GetMenuLevels()
Dim str As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    str = Chr(1)
    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = cn
    ssdcboUserName.FieldSeparator = str
    
    With cmd
        .CommandText = "SELECT LTRIM(RTRIM(ml_melvname)), ml_melvid"
        .CommandText = .CommandText & " FROM MENULEVEL"
        .CommandText = .CommandText & " WHERE ml_npecode = '" & np & "'"
        
        If UserLevel(CurrentUser, np, cn) > 0 Then _
            .CommandText = .CommandText & " AND ml_melvid >= 3"
            
        .CommandText = .CommandText & " Order by ml_melvid, ml_melvname"
        
        Set rs = .Execute
    End With
    
    Do While Not rs.EOF
        ssdcboUserName.AddItem rs(0) & str & rs(1)
        rs.MoveNext
    Loop
End Sub

'SQL statement to get recordset and populate to data grid

Private Sub GetMenuAccess(MenuLevel As String)
Dim rs As ADODB.Recordset


    Set rs = NavBar.Recordset
    If rs.State And adStateOpen Then rs.Close
    
    With rs
    
        .CursorType = adOpenStatic
        Set .ActiveConnection = cn
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        
        
        .Source = "SELECT ma_npecode, ma_melvid, ma_meopid, ma_accsflag, "
        .Source = .Source & " ma_accsflagwrit, mo_meopname"
        
        .Source = .Source & " From MENUACCESS, MENUOPTION"
        .Source = .Source & " WHERE ((ma_melvid ='" & MenuLevel & "')"
        
        .Source = .Source & " AND(mo_meopid =  ma_meopid)"
        .Source = .Source & " AND(mo_npecode = ma_npecode)"
        .Source = .Source & " AND(mo_npecode = '" & np & "'))"
        .Source = .Source & " ORDER BY mo_meopid"
        .Open

        Set ssdbgMenuTemplate.DataSource = Nothing

        Set NavBar.Recordset = rs
        Set ssdbgMenuTemplate.DataSource = rs
    End With

End Sub

'check menu template column values

Private Sub ssdbgMenuTemplate_Change()
Dim str As String

    With ssdbgMenuTemplate.Columns(ssdbgMenuTemplate.Col)
    
        str = LCase$(.Name)
        If str = "id" Or str = "menu" Then Exit Sub
        
        If str = "show" Then
        
            If .Value = True Then
                'ssdbgMenuTemplate.Columns("Read").Value = True
            Else
                'ssdbgMenuTemplate.Columns("Read").Value = False
                ssdbgMenuTemplate.Columns("Write").Value = False
            End If
            
                
        ElseIf str = "write" Then
        
            If .Value = True Then
            
                ssdbgMenuTemplate.Columns("Show").Value = True
                'ssdbgMenuTemplate.Columns("Read").Value = True
                
            Else
                
                If ssdbgMenuTemplate.Columns("Show").Value = True Then
                    'ssdbgMenuTemplate.Columns("Read").Value = True
                Else
                    'ssdbgMenuTemplate.Columns("Read").Value = False
                End If
           
            End If
            
                ElseIf str = "read" Then
       
            If .Value = True Then
                ssdbgMenuTemplate.Columns("Show").Value = True
            Else
                ssdbgMenuTemplate.Columns("Show").Value = False
            End If
            
        End If
    End With
End Sub

'set data grid column size and caption texts

Private Sub ssdbgMenuTemplate_InitColumnProps()
    With ssdbgMenuTemplate
    
        .Columns.RemoveAll
        .AllowUpdate = True
        
        .Columns.Add 0
        .Columns.Add 1
        .Columns.Add 2
        .Columns.Add 3
        .Columns.Add 4
       ' .Columns.Add 5
        
        'Modified by Juan (10/23/2000) for Multilingual
        .Columns(0).Width = 3200
        msg1 = Trans("M00694") 'J added
        .Columns(0).Caption = IIf(msg1 = "", "Option Name", msg1) 'J modified
        .Columns(0).Name = "id"
        .Columns(0).DataField = "mo_meopname"
        .Columns(0).Locked = -1   'True
        .Columns(0).HeadStyleSet = "ColHeader"
        .Columns(0).StyleSet = "RowFont"
        
        .Columns(1).Width = 675
        .Columns(1).Visible = False
        msg1 = Trans("M00692") 'J added
        .Columns(1).Caption = IIf(msg1 = "", "Menu", msg1) 'J modified
        .Columns(1).Name = "menu"
        .Columns(1).DataField = "Column 1"
        .Columns(1).DataType = 8
        .Columns(1).Locked = -1   'True
        .Columns(1).Style = 2
        .Columns(1).HeadStyleSet = "ColHeader"
        .Columns(1).StyleSet = "RowFont"
        
        .Columns(2).Width = 675
        msg1 = Trans("M00693") 'J added
        .Columns(2).Caption = IIf(msg1 = "", "Show", msg1) 'J modified2
        .Columns(2).Name = "Show"
        .Columns(2).DataField = "ma_accsflag"
        .Columns(2).Style = 2
       .Columns(2).DataType = 11
        .Columns(2).HeadStyleSet = "ColHeader"
        .Columns(2).StyleSet = "RowFont"
                
      '  .Columns(3).Width = 675
      '  msg1 = Trans("M00649") 'J added
       ' .Columns(1).Caption = IIf(msg1 = "", "Read", msg1) 'J modified
       ' .Columns(3).Name = "Read"
       ' .Columns(3).DataField = "ma_accsflagread"
       ' .Columns(3).DataType = 11
       ' .Columns(3).Style = 2
       ' .Columns(3).HeadStyleSet = "ColHeader"
       ' .Columns(3).StyleSet = "RowFont"
        
        .Columns(3).Width = 675
        msg1 = Trans("M00652") 'J added
        .Columns(3).Caption = IIf(msg1 = "", "Write", msg1) 'J modified
        .Columns(3).Name = "Write"
        .Columns(3).DataField = "ma_accsflagwrit"
        .Columns(3).Style = 2
        .Columns(3).HeadStyleSet = "ColHeader"
        .Columns(3).StyleSet = "RowFont"
        
        .Columns(4).Width = 675
        .Columns(4).Visible = 0    'False
        msg1 = Trans("L00309") 'J added
        .Columns(4).Caption = IIf(msg1 = "", "np", msg1) 'J modified
        .Columns(4).Name = "np"
        .Columns(4).DataField = "ma_npecode"
        .Columns(4).DataType = 8
        '------------------------------------------------------------
        
    End With
End Sub

'call function to menu access recordset

Private Sub ssdcboUserName_Click()
    Call LockWindowUpdate(Hwnd)
    Call GetMenuAccess(ssdcboUserName.Columns(1).Text)
    
    Call LockWindowUpdate(0)
End Sub



