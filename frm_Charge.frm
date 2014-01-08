VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_Charge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charge To"
   ClientHeight    =   3975
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   5535
   Tag             =   "01012000"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   870
      TabIndex        =   2
      Top             =   3480
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowUpdate     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGCharge 
      Height          =   2895
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   5175
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
      stylesets(0).Picture=   "frm_Charge.frx":0000
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
      stylesets(1).Picture=   "frm_Charge.frx":001C
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1905
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "cha_acctcode"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   3
      Columns(1).Width=   6403
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "cha_acctname"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   15
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "cha_npecode"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   9128
      _ExtentY        =   5106
      _StockProps     =   79
      DataMember      =   "CHARGE"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Account"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   1
      Top             =   60
      Width           =   5055
   End
End
Attribute VB_Name = "frm_Charge"
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
   
   
   
  
   SSDBGCharge.Columns("code").locked = True
   SSDBGCharge.Columns("description").locked = True
   SSDBGCharge.Columns("np").locked = True

   
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

    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Charge")
    '------------------------------------------

    Screen.MousePointer = vbHourglass

    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.CHARGE(deIms.NameSpace)
    Set SSDBGCharge.DataSource = deIms
    Set NavBar1.Recordset = deIms.rsCHARGE
    
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    
    Caption = Caption + " - " + Tag
    
    With frm_Charge
        .Left = Round((Screen.Width - .Width) / 2)
        .Top = Round((Screen.Height - .Height) / 2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If

On Error Resume Next

    Hide
    deIms.rsCHARGE.Update
    deIms.rsCHARGE.UpdateBatch
    
    deIms.rsCHARGE.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub NavBar1_BeforeCancelClick()
    SSDBGCharge.CancelUpdate
End Sub

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGCharge.Update
End Sub

Private Sub NavBar1_BeforeNewClick()
    SSDBGCharge.Update
    SSDBGCharge.AddNew
End Sub

Private Sub NavBar1_BeforeSaveClick()
    SSDBGCharge.Update
    deIms.rsCHARGE!cha_modiuser = CurrentUser
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

Private Sub NavBar1_OnNewClick()
    deIms.rsCHARGE!cha_creauser = CurrentUser
End Sub

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Chargeacct.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00519") 'J added
        .WindowTitle = IIf(msg1 = "", "Charge", msg1) 'J modified
        Call translator.Translate_Reports("Chargeacct.rpt") 'J added
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

Private Sub SSDBGCharge_BeforeUpdate(Cancel As Integer)
    SSDBGCharge.Columns("np").value = deIms.NameSpace
End Sub

