VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmServiceCate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Code Category"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6465
   Tag             =   "01011900"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frmServiceCata.frx":0000
      EmailEnabled    =   -1  'True
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGServiceCate 
      Height          =   2955
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   5775
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "Colls"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmServiceCata.frx":001C
      stylesets(1).Name=   "Rows"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmServiceCata.frx":0038
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorFrame =   -2147483630
      BevelColorHighlight=   14737632
      BevelColorShadow=   -2147483633
      AllowDelete     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3413
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "scs_code"
      Columns(0).FieldLen=   4
      Columns(1).Width=   5927
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "scs_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   40
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "scs_npecode"
      Columns(2).FieldLen=   5
      TabNavigation   =   1
      _ExtentX        =   10186
      _ExtentY        =   5212
      _StockProps     =   79
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
   Begin VB.Label lbl_ServiceCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Service Category"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   5910
   End
End
Attribute VB_Name = "frmServiceCate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat

'call function to get recordset and populate data grid

Private Sub Form_Load()

'
''copy begin here
'
'If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)


   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode


   SSDBGServiceCate.Columns("code").locked = True
   SSDBGServiceCate.Columns("description").locked = True
 '  SSDBGServiceCate.Columns("category").locked = True

NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

'Exit Sub                                                     'Exit Edit sub because theres nothing the user can do
Else
TableLocked = True
End If
'NavBar1.SaveEnabled = False
'NavBar1.NewEnabled = False
'NavBar1.CancelEnabled = False
'
'    Dim textboxes As Control
'
'    For Each textboxes In Controls
'        If (TypeOf textboxes Is textBOX) Then
'            textboxes.Enabled = False
'        End If
'
'    Next textboxes
'    Else
'    TableLocked = True
'    End If
'End If

'end copy


Dim ctl As Control
     Screen.MousePointer = vbHourglass
     
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmServiceCate")
    '------------------------------------------
    
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.SERVCODECAT(deIms.NameSpace)
    SSDBGServiceCate.DataMember = "SERVCODECAT"
    
    Screen.MousePointer = vbDefault
    Set SSDBGServiceCate.DataSource = deIms
    Call DisableButtons(Me, NavBar1)
    
    frmServiceCate.Caption = frmServiceCate.Caption + " - " + frmServiceCate.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)
End Sub

'unload form and close recordset

Private Sub Form_Unload(Cancel As Integer)



If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If


On Error Resume Next
    
    Hide
    deIms.rsSERVCODECAT.Update
    deIms.rsSERVCODECAT.CancelUpdate
    
    deIms.rsSERVCODECAT.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator

End Sub

'delete a record form recordset

Private Sub NavBar1_OnDeleteClick()
    SSDBGServiceCate.DeleteSelected
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
'
    SSDBGServiceCate.MoveFirst
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
'
    SSDBGServiceCate.MoveLast
End Sub

'before add new move recordset to add position
'and set name space to current name space

Private Sub NavBar1_OnNewClick()
    
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
'
    
' NavBar1.SaveEnabled = True
' NavBar1.CancelEnabled = True
    
    SSDBGServiceCate.AddNew
    SSDBGServiceCate.Columns("np").value = deIms.NameSpace
End Sub

'cancel recordset update

Private Sub NavBar1_OnCancelClick()
    
'If TableLocked = True Then    'jawdat
'Dim imsLock As imsLock.lock
'Set imsLock = New imsLock.lock
'currentformname = Forms(3).Name
'Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
'End If
    
    
    SSDBGServiceCate.CancelUpdate
End Sub

'close from

Private Sub NavBar1_OnCloseClick()
    
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
    Unload Me
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
    
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
'
    SSDBGServiceCate.MoveNext
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
    
'If TableLocked = True Then    'Added for locking rows, user was allowed to view edit more records while having the current record locked, Jawdat 2.5.02
'MsgBox "You must save the information, or cancel modification before moving to any other record."
'Exit Sub                'cancel movement if they still have it locked, until they save or cancel
'End If
    
    SSDBGServiceCate.MovePrevious
End Sub

'set crystal report parameters and get application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo Handler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\servcodecat.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("M00103") 'J added
        .WindowTitle = IIf(msg1 = "", "Service Code Category", msg1) 'J modified
        Call translator.Translate_Reports("servcodecat.rpt") 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
    Exit Sub

Handler:
    If Err Then MsgBox Err.Description: Err.Clear
    
End Sub

'save recordset to database

Private Sub NavBar1_OnSaveClick()

       
'If TableLocked = True Then    'jawdat
'Dim imsLock As imsLock.lock
'Set imsLock = New imsLock.lock
'currentformname = Forms(3).Name
'Call imsLock.UNLOCK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)
'End If


On Error Resume Next
    SSDBGServiceCate.Update
    Call SSDBGServiceCate.Update
    If Err Then Call LogErr(Name & "::NavBar1_OnSaveClick", Err.Description, Err.number, True)
End Sub





