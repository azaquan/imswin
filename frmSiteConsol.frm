VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frmSiteConsolidation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Consolidation"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   6225
   Tag             =   "01040300"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   1035
      TabIndex        =   2
      Top             =   3720
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frmSiteConsol.frx":0000
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
      EmailEnabled    =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGSiteConsol 
      Height          =   2955
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5760
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
      stylesets(0).Picture=   "frmSiteConsol.frx":001C
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
      stylesets(1).Picture=   "frmSiteConsol.frx":0038
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
      Columns(0).Width=   3969
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "ste_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   10
      Columns(1).Width=   5371
      Columns(1).Caption=   "Consolidate"
      Columns(1).Name =   "Consolidate"
      Columns(1).DataField=   "ste_codecons"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   10
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "ste_npecode"
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10160
      _ExtentY        =   5212
      _StockProps     =   79
      DataMember      =   "SITECONSOL"
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
      Caption         =   "Site Consolidation"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5805
   End
End
Attribute VB_Name = "frmSiteConsolidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Sub Form_Load()

'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
   SSDBGSiteConsol.Columns("code").locked = True
   SSDBGSiteConsol.Columns("consolidate").locked = True
   
   
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

    Screen.MousePointer = vbHourglass
    
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmSiteConsolidation")
    '------------------------------------------

    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.SITECONSOL(deIms.NameSpace)
    Set SSDBGSiteConsol.DataSource = deIms
    
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    frmSiteConsolidation.Caption = frmSiteConsolidation.Caption + " - " + frmSiteConsolidation.Tag
    
    With frmSiteConsolidation
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsSITECONSOL.Update
    deIms.rsSITECONSOL.UpdateBatch
    
    deIms.rsSITECONSOL.Close
    If Err Then Err.Clear
    Set frmSiteConsolidation = Nothing
    If open_forms <= 5 Then ShowNavigator
    
    
    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
    
End Sub

Private Sub NavBar1_OnDeleteClick()
    SSDBGSiteConsol.DeleteSelected
End Sub

Private Sub NavBar1_OnFirstClick()
    SSDBGSiteConsol.MoveFirst
End Sub

Private Sub NavBar1_OnLastClick()
    SSDBGSiteConsol.MoveLast
End Sub

Private Sub NavBar1_OnNewClick()
    SSDBGSiteConsol.AddNew
    SSDBGSiteConsol.Columns("np").value = deIms.NameSpace
End Sub

Private Sub NavBar1_OnCancelClick()
    SSDBGSiteConsol.CancelUpdate
End Sub

Private Sub NavBar1_OnCloseClick()
    
     
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
        
    
    Unload Me
End Sub

Private Sub NavBar1_OnNextClick()
    SSDBGSiteConsol.MoveNext
End Sub

Private Sub NavBar1_OnPreviousClick()
    SSDBGSiteConsol.MovePrevious
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\conso.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("frmSiteConsolidation") 'J added
        .WindowTitle = "Site Consolidation" 'J modified
        Call translator.Translate_Reports("conso.rpt") 'J added
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

Private Sub NavBar1_OnSaveClick()
    SSDBGSiteConsol.Update
    Call SSDBGSiteConsol.MoveRecords(0)
End Sub



Private Sub SSDBGSiteConsol_BeforeUpdate(Cancel As Integer)
   SSDBGSiteConsol.Columns("np").value = deIms.NameSpace
End Sub

