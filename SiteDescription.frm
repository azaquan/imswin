VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_SiteDescript 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Description"
   ClientHeight    =   4830
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6225
   Tag             =   "01040200"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "SiteDescription.frx":0000
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGSite 
      Height          =   3195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5745
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
      stylesets(0).Picture=   "SiteDescription.frx":001C
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
      stylesets(1).Picture=   "SiteDescription.frx":0038
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
      Columns(0).Width=   3281
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "sit_code"
      Columns(0).FieldLen=   10
      Columns(1).Width=   6033
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "sit_name"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   30
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "sit_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   5
      TabNavigation   =   1
      _ExtentX        =   10134
      _ExtentY        =   5636
      _StockProps     =   79
      DataMember      =   "SITE"
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
   Begin VB.Label lbl_SiteDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Site Description"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frm_SiteDescript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat

'cancel recordset update

Private Sub NavBar1_BeforeCancelClick()
    SSDBGSite.CancelUpdate
End Sub

'set data grid name space equal to current name space

Private Sub NavBar1_BeforeNewClick()
    SSDBGSite.AddNew
    SSDBGSite.Columns("np").value = deIms.NameSpace
End Sub

'unload form close recordsset

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Hide
    deIms.rsSITE.Update
    deIms.rsSITE.CancelUpdate
    
    deIms.rsSITE.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
    
    
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
      
    
End Sub

'cancel update

Private Sub NavBar1_OnCancelClick()
    SSDBGSite.CancelUpdate
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    
     
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
        
        
    
    Unload Me
End Sub

'delete select a record

Private Sub NavBar1_OnDeleteClick()
    SSDBGSite.DeleteSelected
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    SSDBGSite.MoveFirst
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    SSDBGSite.MoveLast
End Sub

'
Private Sub NavBar1_OnNewClick()
    SSDBGSite.AddNew
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
    SSDBGSite.MoveNext
End Sub

'move recordset to provious position

Private Sub NavBar1_OnPreviousClick()
    SSDBGSite.MovePrevious
End Sub

'get crystal report parameter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Site.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00270") 'J added
        .WindowTitle = IIf(msg1 = "", "Site Description", msg1) 'J modified
        Call translator.Translate_Reports("Site.rpt") 'J added
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

'save recordset

Private Sub NavBar1_BeforeSaveClick()
    SSDBGSite.Update
End Sub

'get data for data grid and set buttom

Private Sub Form_Load()

'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
   
      
   SSDBGSite.Columns("code").locked = True
   SSDBGSite.Columns("description").locked = True

   
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
    Else
    TableLocked = True
    End If
End If

'end copy




Dim ctl As Control

    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_SiteDescript")
    '------------------------------------------

    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    

    Call deIms.Site(deIms.NameSpace)
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    Set SSDBGSite.DataSource = deIms
    
    frm_SiteDescript.Caption = frm_SiteDescript.Caption + " - " + frm_SiteDescript.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'set data grid name space equal to current name space

Private Sub NavBar1_OnSaveClick()
    SSDBGSite.Columns("np").value = deIms.NameSpace

    SSDBGSite.Update
    Call SSDBGSite.MoveRecords(0)
End Sub

'set data grid name space equal to current name space

Private Sub SSDBGSite_BeforeUpdate(Cancel As Integer)
    SSDBGSite.Columns("np").value = deIms.NameSpace
End Sub

