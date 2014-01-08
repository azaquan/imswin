VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmMenuOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Options"
   ClientHeight    =   5010
   ClientLeft      =   4545
   ClientTop       =   3120
   ClientWidth     =   7260
   HasDC           =   0   'False
   Icon            =   "frmMenuOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Tag             =   "04010600"
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Tag             =   "04010600"
      Top             =   4380
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgMenuOption 
      Height          =   3255
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   6885
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
      stylesets(0).Picture=   "frmMenuOption.frx":000C
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
      stylesets(1).Picture=   "frmMenuOption.frx":0028
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
      ExtraHeight     =   106
      Columns.Count   =   3
      Columns(0).Width=   2805
      Columns(0).Caption=   "Option ID"
      Columns(0).Name =   "id"
      Columns(0).DataField=   "mo_meopid"
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   5292
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "Menu"
      Columns(1).Name =   "menu"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(1).Style=   2
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   8837
      Columns(2).Caption=   "Option Title"
      Columns(2).Name =   "title"
      Columns(2).DataField=   "mo_meopname"
      Columns(2).FieldLen=   256
      Columns(2).HeadStyleSet=   "ColHeader"
      Columns(2).StyleSet=   "RowFont"
      _ExtentX        =   12144
      _ExtentY        =   5741
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
      Caption         =   "Menu Options"
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
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   6915
   End
End
Attribute VB_Name = "frmMenuOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim cn As ADODB.Connection

'set new recordset

Private Sub Form_Load()

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmMenuOption")
    '--------------------------------------------------

    Set NavBar.Recordset = New ADODB.Recordset
    'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
End Sub

'SQL statement to get menu option recordset and set buttom

Public Sub SetConnection(con As ADODB.Connection)

    Set cn = con
    With NavBar.Recordset
        Set .ActiveConnection = con
        
        .CursorType = adOpenStatic
        .CursorLocation = adUseServer
        .LockType = adLockBatchOptimistic
        
        .Source = "SELECT mo_meopid, mo_meopname FROM MENUOPTION"
        .Source = .Source & " WHERE mo_npecode = '" & np & "'"
        
        .Open
        
        Set ssdbgMenuOption.DataSource = NavBar
    End With
    
    Set NavBar.Recordset = NavBar.Recordset
    Call DisableButtons(Me, NavBar, np, CurrentUser, cn)
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Caption = Caption + " - " + Tag
    
End Sub

'unload form and free memory

Private Sub Form_Unload(Cancel As Integer)
    Set NavBar.Recordset = Nothing
End Sub

'cancel recordset update

Private Sub NavBar_BeforeCancelClick()
    ssdbgMenuOption.CancelUpdate
End Sub

'cancel recordset update

Private Sub NavBar_OnCancelClick()
    NavBar.Recordset.CancelUpdate
    Call NavBar.Recordset.CancelBatch(adAffectCurrent)
End Sub

'close form

Private Sub NavBar_OnCloseClick()
    Unload Me
End Sub

'get crystal report parameters and application path

Private Sub NavBar_OnPrintClick()
    With CrystalReport1
        .ReportFileName = ReportPath + "menuoption.rpt"
        .ParameterFields(0) = "namespace;" + np + ";TRUE"
        
        'Modified by Juan (10/23/00) for Multilingual 'J added
        Call translate_reports(Me.Name, "menuoption.rpt", True, cn, CrystalReport1) 'J added
        msg1 = Trans("M00207") 'J added
        .WindowTitle = IIf(msg1 = "", "Menu Option", msg1) 'J modified
        '--------------------------------------------------
        
        .Action = 1
    End With

End Sub

'save recordset to database

Private Sub NavBar_OnSaveClick()
    ssdbgMenuOption.update
    NavBar.Recordset.UpdateBatch
End Sub

'set name space equal to current name space

Public Sub SetNameSpace(NameSpace As String)
    np = NameSpace
End Sub

