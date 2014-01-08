VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_ServiceCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Code"
   ClientHeight    =   4140
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   6855
   Tag             =   "01010800"
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown ssdbddCategory 
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
      DataFieldList   =   "scs_code"
      _Version        =   196617
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
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
      stylesets(0).Picture=   "frm_ServiceCode.frx":0000
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
      stylesets(1).Picture=   "frm_ServiceCode.frx":001C
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5292
      Columns(0).Caption=   "Description"
      Columns(0).Name =   "Description"
      Columns(0).DataField=   "scs_desc"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   40
      Columns(1).Width=   2117
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).DataField=   "scs_code"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   4
      _ExtentX        =   6165
      _ExtentY        =   1296
      _StockProps     =   77
      DataFieldToDisplay=   "scs_desc"
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   1590
      TabIndex        =   2
      Top             =   3600
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBServiceCode 
      Height          =   2925
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6620
      _Version        =   196617
      BorderStyle     =   0
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
      stylesets(0).Picture=   "frm_ServiceCode.frx":0038
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
      stylesets(1).Picture=   "frm_ServiceCode.frx":0054
      stylesets(1).AlignmentText=   1
      DefColWidth     =   5292
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   1508
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "srvc_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   2
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   4895
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "srvc_desc"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   40
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "srvc_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2461
      Columns(3).Caption=   "Category"
      Columns(3).Name =   "Category"
      Columns(3).DataField=   "srvc_cate"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   4
      Columns(3).HeadStyleSet=   "ColHeader"
      Columns(3).StyleSet=   "RowFont"
      Columns(4).Width=   1455
      Columns(4).Caption=   "Active"
      Columns(4).Name =   "Active"
      Columns(4).DataField=   "srvc_actvflag"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(4).HeadStyleSet=   "ColHeader"
      _ExtentX        =   11677
      _ExtentY        =   5159
      _StockProps     =   79
      DataMember      =   "SERVCODE"
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Codes"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   6555
   End
End
Attribute VB_Name = "frm_ServiceCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat

'load form to get data for combo and set back ground color

Private Sub Form_Load()
Dim ctl As Control


If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
 '  ssdbddCategory.Columns("category").locked = True
   ssdbddCategory.Columns("code").locked = True
   ssdbddCategory.Columns("description").locked = True
      
   SSDBServiceCode.Columns("code").locked = True
   SSDBServiceCode.Columns("description").locked = True
   SSDBServiceCode.Columns("category").locked = True
   
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


'end copy

   TableLocked = True
    End If
End If


    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_ServiceCode")
    '------------------------------------------

    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call GetCategory
    deIms.SERVCODE (deIms.NameSpace)
    Set SSDBServiceCode.DataSource = deIms
    Set NavBar1.Recordset = deIms.rsSERVCODE
    Call DisableButtons(Me, NavBar1)
    
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
    NavBar1.EditVisible = True 'Juan Gonzalez 12/29/2006
    
    frm_ServiceCode.Caption = frm_ServiceCode.Caption + " - " + frm_ServiceCode.Tag


 
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    deIms.rsSERVCODE.Close
    Set frm_ServiceCode = Nothing
   If open_forms <= 5 Then ShowNavigator
   
   If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
   
   
   
End Sub

'cancel uppdate

Private Sub NavBar1_BeforeCancelClick()
    SSDBServiceCode.CancelUpdate
End Sub

'set combe name space equal to current name space

Private Sub NavBar1_BeforeNewClick()
    NavBar1.EditEnabled = False 'Juan Gonzalez 12/29/2006
    SSDBServiceCode.AddNew
    SSDBServiceCode.Columns("np").value = deIms.NameSpace
End Sub

'set recordset to update

Private Sub NavBar1_BeforeSaveClick()
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
    SSDBServiceCode.Update
End Sub

Private Sub NavBar1_OnCancelClick()
    NavBar1.EditEnabled = True 'Juan Gonzalez 12/29/2006
End Sub


'close form

Private Sub NavBar1_OnCloseClick()
     
If TableLocked = True Then    'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If
    
        
        



On Error Resume Next
    Unload Me
End Sub

Private Sub NavBar1_OnEditClick()
    NavBar1.CancelEnabled = True 'Juan Gonzalez 12/29/2006
    NavBar1.EditEnabled = False 'Juan Gonzalez 12/29/2006
    NavBar1.SaveEnabled = True 'Juan Gonzalez 12/29/2006
End Sub

'get crystal report parameters and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Servcode.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("L00199") 'J added
        .WindowTitle = IIf(msg1 = "", "Service Code", msg1) 'J modified
        Call translator.Translate_Reports("Servcode.rpt") 'J added
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

'load data to combo

Private Sub GetCategory()
    If deIms.rsSERVCODECAT.State And adStateOpen Then
        Set ssdbddCategory.DataSource = deIms.rsSERVCODECAT.Clone(adLockReadOnly)
    Else
        Call deIms.SERVCODECAT(deIms.NameSpace)
        Set ssdbddCategory.DataSource = deIms.rsSERVCODECAT.Clone(adLockReadOnly)
        deIms.rsSERVCODECAT.Close
    End If
End Sub

Private Sub NavBar1_OnSaveClick()
On Error Resume Next 'Juan Gonzalez 12/29/2006
    Call deIms.rsSERVCODE.Move(0) 'Juan Gonzalez 12/29/2006
    
End Sub

'set recordset to link data grid

Private Sub SSDBServiceCode_InitColumnProps()
    SSDBServiceCode.Columns("Category").DropDownHwnd = ssdbddCategory.HWND
End Sub
