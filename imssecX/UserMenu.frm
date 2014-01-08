VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmUserMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Menu"
   ClientHeight    =   5790
   ClientLeft      =   5640
   ClientTop       =   3060
   ClientWidth     =   7080
   Icon            =   "UserMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Tag             =   "04010900"
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdcboUserName 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   1140
      Width           =   3735
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
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin LRNavigators.LROleDBNavBar NavBar 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5220
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      NewVisible      =   0   'False
      AllowAddNew     =   0   'False
      AllowCancel     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgUserMenu 
      Height          =   2970
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   6705
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
      stylesets(0).Picture=   "UserMenu.frx":000C
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
      stylesets(1).Picture=   "UserMenu.frx":0028
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
      _ExtentX        =   11827
      _ExtentY        =   5239
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
   Begin VB.Label lblMenuLevel 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ml_melvname"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   1500
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "User ID"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   1140
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "User Menu"
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
      Left            =   300
      TabIndex        =   6
      Top             =   300
      Width           =   6555
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Level  ID"
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   1500
      Width           =   1800
   End
End
Attribute VB_Name = "frmUserMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim np As String
Dim cn As ADODB.Connection
Dim ObjXevents As ImsXevents
Dim TRANSACTIONNUBMER As Integer



'set form size and caption

Private Sub Form_Load()

 Set ObjXevents = New ImsXevents 'Shakir
    ObjXevents.ConnectionObject = cn

    'Added by Juan (10/23/00) for Multilingual 'J added
    Translate_Forms ("frmUserMenu")
    '--------------------------------------------------

    Set NavBar.Recordset = New ADODB.Recordset
    NavBar.Recordset.LockType = adLockBatchOptimistic
    cn.BeginTrans
    
    'Call CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Caption = Caption + " - " + Tag
    
End Sub

'set database conncetion and call function to get recordset

Public Sub SetConnection(con As ADODB.Connection)
    Set cn = con
    GetUserNames
    Set NavBar.Recordset.ActiveConnection = cn
    Call DisableButtons(Me, NavBar, np, CurrentUser, con)
    ssdcboUserName.AllowInput = True
End Sub

Private Sub Form_Paint()
        If Not NavBar.SaveEnabled Then ssdbgUserMenu.AllowUpdate = False
End Sub

'unload form free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set NavBar.Recordset = Nothing
    cn.RollbackTrans
End Sub

'cancel user menu update

Private Sub NavBar_BeforeCancelClick()
If Not ObjXevents Is Nothing Then Set ObjXevents = Nothing
    ssdbgUserMenu.CancelUpdate
    cn.RollbackTrans
    cn.BeginTrans
End Sub

'before save record set recordset to update position

Private Sub NavBar_BeforeSaveClick()
    ssdbgUserMenu.update
End Sub

'get crystal report parameters and application path

Private Sub NavBar_OnPrintClick()
    With CrystalReport1
        .ReportFileName = ReportPath + "menuuser.rpt"
        .ParameterFields(0) = "namespace;" + np + ";TRUE"
        .ParameterFields(1) = "userid;" + ssdcboUserName.Columns(1).Text + ";TRUE"
        
        'Modified by Juan (10/23/00) for Multilingual 'J added
        Call translate_reports(Me.Name, "menuuser.rpt", True, cn, CrystalReport1) 'J added
        msg1 = Trans("M00690") 'J added
        .WindowTitle = IIf(msg1 = "", "User Menu", msg1) 'J modified
        '--------------------------------------------------
        
        .Action = 1
    End With
End Sub

'save reecordset to database

Private Sub NavBar_OnSaveClick()
On Error Resume Next
    NavBar.Recordset.UpdateBatch
    
    If Err Then MsgBox Err.Description: Err.Clear
    
  Dim OldVAlueShow As Boolean
  Dim NewVAlueShow As Boolean
Dim OldVAlueWrite As Boolean
  Dim NewVAlueWrite As Boolean
  Dim Text As String
  Dim oldOPtions As String
  Dim newoptions As String
  
  
  
   
If ssdbgUserMenu.IsAddRow = True Then
    OldVAlueShow = 0
Else
    Set ObjXevents = New ImsXevents
    ObjXevents.ConnectionObject = cn
   
    If ObjXevents Is Nothing Then Set ObjXevents = New ImsXevents
 
    ObjXevents.AddNew
    ObjXevents.NameSpace = np
    ObjXevents.MyLoginId = CurrentUser
    ObjXevents.HisLoginId = Trim$(ssdcboUserName.Columns(1).Text)
   
 
      
    If ssdbgUserMenu.Columns(2).ColChanged = True Then
        OldVAlueShow = ssdbgUserMenu.Columns(2).CellText(ssdbgUserMenu.Bookmark)
        NewVAlueShow = IIf(ssdbgUserMenu.Columns(2).Text = -1, True, False)
    End If
    If ssdbgUserMenu.Columns(3).ColChanged = True Then
        OldVAlueWrite = ssdbgUserMenu.Columns(3).CellText(ssdbgUserMenu.Bookmark)
        NewVAlueWrite = IIf(ssdbgUserMenu.Columns(3).Text = -1, True, False)
    End If
       
    If OldVAlueShow = True And OldVAlueWrite = True Then
        oldOPtions = "Read/Write"
    ElseIf OldVAlueShow = False And OldVAlueWrite = True Then
        oldOPtions = "Write only"
        ElseIf OldVAlueShow = True And OldVAlueWrite = False Then
            oldOPtions = "Read Only"
        ElseIf OldVAlueShow = False And OldVAlueWrite = False Then
            oldOPtions = "No rights"
        End If
       
       
        If NewVAlueShow = True And NewVAlueWrite = True Then
            newoptions = "Read/Write"
        ElseIf NewVAlueShow = False And NewVAlueWrite = True Then
            newoptions = "Write only"
        ElseIf NewVAlueShow = True And NewVAlueWrite = False Then
            newoptions = "Read Only"
        ElseIf NewVAlueShow = False And NewVAlueWrite = False Then
            newoptions = "Read Only"
        End If
              
        Text = "The Record with option Name " & ssdbgUserMenu.Columns(0).Text & " have been changed from " & oldOPtions & " to " & newoptions & " for the USER " & ssdcboUserName.Value
        ObjXevents.STAs = "A"
        ObjXevents.EventDetail = Text
        ObjXevents.NameSpace = np
End If
    cn.COMMITTRANS
  
    
End Sub

'unload form

Private Sub NavBar_OnCloseClick()
    Unload Me
End Sub

'set name space value aquel to current name space

Public Sub SetNameSpace(NameSpace As String)
    np = NameSpace
End Sub

'SQL statement to get user recordset  and populate data grid

Private Sub GetUserNames()
Dim str As String
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    str = Chr(1)
    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = cn
    ssdcboUserName.FieldSeparator = str
    
    With cmd
        
        .CommandText = "SELECT usr_username, usr_userid"
        .CommandText = .CommandText & " FROM XUSERPROFILE"
        .CommandText = .CommandText & " WHERE usr_npecode = '" & np & "'"
        If ImsSecX.UserLevel(CurrentUser, np, cn) = 0 Then
            .CommandText = .CommandText & " AND usr_leve in (1, 2)"
        End If
        .CommandText = .CommandText & " Order by usr_username, usr_userid"
        
        
        'Set rs = .Execute
        Set rs = New ADODB.Recordset
        rs.Open .CommandText, cn, adOpenDynamic, adLockBatchOptimistic
        
    End With
    
    
    
    Do While Not rs.EOF
        ssdcboUserName.AddItem rs(0) & str & rs(1)
        rs.MoveNext
    Loop
End Sub

'SQL statement toget menu access recordset
'and populate recordset

Private Sub GetMenuAccess(UserName As String)
Dim rs As ADODB.Recordset


    Set rs = NavBar.Recordset
    If rs.State And adStateOpen Then rs.Close
    
    With rs
    
        .CursorType = adOpenStatic
        Set .ActiveConnection = cn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        
        .Source = "SELECT mu_accsflag, "
        .Source = .Source & " mu_accsflagwrit, mu_npecode,"
        .Source = .Source & " mu_userid, ml_melvname mu_melvid, mu_meopid,"
        
        .Source = .Source & " ml_melvname, mo_meopname"
        .Source = .Source & " FROM MENUUSER, MENULEVEL, MENUOPTION"
        
        .Source = .Source & " WHERE( (mu_npecode = '" & np & "')"
        .Source = .Source & " AND(mu_userid = '" & UserName & "')"
        
        
        .Source = .Source & " AND(ml_melvid =  (SELECT usr_menuleve FROM XUSERPROFILE"
        .Source = .Source & " WHERE usr_userid = mu_userid AND usr_npecode = mu_npecode))"
        
        .Source = .Source & " AND(ml_npecode = '" & np & "') AND(mo_meopid =  mu_meopid)"
        .Source = .Source & " AND(mo_npecode = '" & np & "'))"
        .Source = .Source & " ORDER BY mu_meopid"
        
        
        .Open
        
        Set lblMenuLevel.DataSource = Nothing
        Set ssdbgUserMenu.DataSource = Nothing

        Set NavBar.Recordset = rs
        Set lblMenuLevel.DataSource = NavBar
        Set ssdbgUserMenu.DataSource = NavBar
    End With

End Sub

Private Sub ssdbgUserMenu_AfterUpdate(RtnDispErrMsg As Integer)
 If Not ObjXevents Is Nothing Then
      ObjXevents.update
    Set ObjXevents = Nothing
   End If  'RtnDispErrMsg = 0
End Sub

Private Sub ssdbgUserMenu_BeforeUpdate(Cancel As Integer)
  Dim OldVAlueShow As Boolean
  Dim NewVAlueShow As Boolean
Dim OldVAlueWrite As Boolean
  Dim NewVAlueWrite As Boolean
  Dim Text As String
  Dim oldOPtions As String
  Dim newoptions As String
  
  
  
   
If ssdbgUserMenu.IsAddRow = True Then
   OldVAlueShow = 0
ElseIf ssdbgUserMenu.RowChanged = True Then
      
      Set ObjXevents = New ImsXevents
   ObjXevents.ConnectionObject = cn
      
   
If ObjXevents Is Nothing Then Set ObjXevents = New ImsXevents
 
 ObjXevents.AddNew
  ObjXevents.NameSpace = np
  ObjXevents.MyLoginId = CurrentUser
  ObjXevents.HisLoginId = Trim$(ssdcboUserName.Columns(1).Text)
   
 
      
      If ssdbgUserMenu.Columns(2).ColChanged = True Then
           OldVAlueShow = ssdbgUserMenu.Columns(2).CellText(ssdbgUserMenu.Bookmark)
           NewVAlueShow = IIf(ssdbgUserMenu.Columns(2).Text = -1, True, False)
       End If
      If ssdbgUserMenu.Columns(3).ColChanged = True Then
           OldVAlueWrite = ssdbgUserMenu.Columns(3).CellText(ssdbgUserMenu.Bookmark)
           NewVAlueWrite = IIf(ssdbgUserMenu.Columns(3).Text = -1, True, False)
      End If
       
        If OldVAlueShow = True And OldVAlueWrite = True Then
           oldOPtions = "Read/Write"
        ElseIf OldVAlueShow = False And OldVAlueWrite = True Then
           oldOPtions = "Write only"
        ElseIf OldVAlueShow = True And OldVAlueWrite = False Then
          oldOPtions = "Read Only"
       ElseIf OldVAlueShow = False And OldVAlueWrite = False Then
         oldOPtions = "No rights"
       End If
       
       
       If NewVAlueShow = True And NewVAlueWrite = True Then
           newoptions = "Read/Write"
        ElseIf NewVAlueShow = False And NewVAlueWrite = True Then
           newoptions = "Write only"
        ElseIf NewVAlueShow = True And NewVAlueWrite = False Then
          newoptions = "Read Only"
       ElseIf NewVAlueShow = False And NewVAlueWrite = False Then
         newoptions = "NOTHING"
      End If
       
       Text = "The Record with option Name " & ssdbgUserMenu.Columns(0).Text & " have been changed from " & oldOPtions & " to " & newoptions & " for the USER " & ssdcboUserName.Value
       ObjXevents.STAs = "A"
       ObjXevents.EventDetail = Text
       ObjXevents.NameSpace = np
    
       
End If

 
 
 'Shakir 12-17-00
  'The Object is Being Fed over here with the Data to be fed in the Xevents Table.
  'Update method on this Object is Exceuted when the User Exist the Form.
  


'set user menu column values

 
 
 'Shakir 12-17-00
  'The Object is Being Fed over here with the Data to be fed in the Xevents Table.
  'Update method on this Object is Exceuted when the User Exist the Form.
  


'set user menu column values
End Sub

Private Sub ssdbgUserMenu_Change()
Dim str As String

    With ssdbgUserMenu.Columns(ssdbgUserMenu.Col)
    
        str = LCase$(.Name)
        If str = "id" Or str = "menu" Then Exit Sub
        
        If str = "show" Then
        
            If .Value = True Then
                'ssdbgUserMenu.Columns("Read").Value = True
            Else
                'ssdbgUserMenu.Columns("Read").Value = False
                ssdbgUserMenu.Columns("Write").Value = False
            End If
            
                
        ElseIf str = "write" Then
        
            If .Value = True Then
            
                ssdbgUserMenu.Columns("Read").Value = True
                'ssdbgUserMenu.Columns("Read").Value = True
                
            Else
                
                If ssdbgUserMenu.Columns("Read").Value = True Then
                    'ssdbgUserMenu.Columns("Read").Value = True
                Else
                    'ssdbgUserMenu.Columns("Read").Value = False
                End If
           
            End If
            
        ElseIf str = "Read" Then
       
            If .Value = True Then
                ssdbgUserMenu.Columns("Read").Value = True
            Else
                ssdbgUserMenu.Columns("Read").Value = False
            End If
            
        End If
    End With

End Sub


 
  


Private Sub ssdbgUserMenu_GotFocus()
    If Not NavBar.SaveEnabled Then ssdbgUserMenu.AllowUpdate = False
End Sub

'set column size and caption
Private Sub ssdbgUserMenu_InitColumnProps()
    With ssdbgUserMenu
        .Columns.RemoveAll
        
        .Columns.Add 0
        .Columns.Add 1
        .Columns.Add 2
        .Columns.Add 3
        .Columns.Add 4
        '.Columns.Add 5
        
        'Modified by Juan (10/23/2000) for Multilingual
        .Columns(0).Width = 3200
        msg1 = Trans("M00691") 'J added
        .Columns(0).Caption = IIf(msg1 = "", "Option Name", msg1) 'J modified
        .Columns(0).Name = "id"
        .Columns(0).DataField = "mo_meopname"
        .Columns(0).Locked = -1   'True
        .Columns(0).HeadStyleSet = "ColHeader"
        .Columns(0).StyleSet = "RowFont"
        
        .Columns(1).Visible = False
        .Columns(1).Width = 950
        msg1 = Trans("M00692") 'J added
        .Columns(1).Caption = IIf(msg1 = "", "Menu", msg1) 'J modified
        .Columns(1).Name = "menu"
        .Columns(1).DataField = "Column 1"
        .Columns(1).DataType = 8
        .Columns(1).Locked = -1   'True
        .Columns(1).Style = 2
        .Columns(1).HeadStyleSet = "ColHeader"
        .Columns(1).StyleSet = "RowFont"
        
        .Columns(2).Width = 950
        msg1 = Trans("M00693") 'J added
        .Columns(2).Caption = IIf(msg1 = "", "Read", msg1) 'J modified
        .Columns(2).Name = "Read"
        .Columns(2).DataField = "mu_accsflag"
        .Columns(2).Style = 2
        .Columns(2).HeadStyleSet = "ColHeader"
        .Columns(2).StyleSet = "RowFont"
        
        '.Columns(3).Width = 950
        'msg1 = Trans("M00649") 'J added
        '.Columns(3).Caption = IIf(msg1 = "", "Read", msg1) 'J modified
        '.Columns(3).Name = "Read"
        '.Columns(3).DataField = "mu_accsflagread"
        '.Columns(3).DataType = 11
        '.Columns(3).Style = 2
        '.Columns(3).HeadStyleSet = "ColHeader"
        '.Columns(3).StyleSet = "RowFont"
        
        .Columns(3).Width = 950
        msg1 = Trans("M00652") 'J added
        .Columns(3).Caption = IIf(msg1 = "", "Write", msg1) 'J modified
        .Columns(3).Name = "Write"
        .Columns(3).DataField = "mu_accsflagwrit"
        .Columns(3).Style = 2
        .Columns(3).HeadStyleSet = "ColHeader"
        .Columns(3).StyleSet = "RowFont"
        
        .Columns(4).Width = 950
        .Columns(4).Visible = 0    'False
        msg1 = Trans("L00309") 'J added
        .Columns(4).Caption = IIf(msg1 = "", "np", msg1) 'J modified
        .Columns(4).Name = "np"
        .Columns(4).DataField = "ma_npecode"
        .Columns(4).DataType = 8
        '-------------------------------------------------------------
        
    End With
    
    
End Sub

'call function get  menuaccess recordset and set data grid update

Private Sub ssdcboUserName_Click()
    Call LockWindowUpdate(Hwnd)
    Call GetMenuAccess(ssdcboUserName.Columns(1).Text)
    
    Call LockWindowUpdate(0)
End Sub

