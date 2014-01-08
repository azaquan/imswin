VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#8.0#0"; "LRNavigators.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5280
   Tag             =   "01020400"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Top             =   3540
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      NewEnabled      =   -1  'True
      AllowAddNew     =   0   'False
      AllowUpdate     =   0   'False
      AllowCancel     =   0   'False
      AllowDelete     =   0   'False
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGDestination 
      Bindings        =   "DESTINATIONNEW.frx":0000
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5190
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
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
      stylesets(0).Picture=   "DESTINATIONNEW.frx":0014
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
      stylesets(1).Picture=   "DESTINATIONNEW.frx":0030
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorHighlight=   16777215
      AllowAddNew     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
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
      Columns(0).Width=   2699
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "des_destcode"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   3
      Columns(0).HeadStyleSet=   "Colls"
      Columns(0).StyleSet=   "Rows"
      Columns(1).Width=   5689
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "des_destname"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   20
      Columns(1).HeadStyleSet=   "Colls"
      Columns(1).StyleSet=   "Rows"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "des_npecode"
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      BatchUpdate     =   -1  'True
      _ExtentX        =   9155
      _ExtentY        =   5212
      _StockProps     =   79
      BackColor       =   -2147483638
      DataMember      =   "Destination"
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
      Caption         =   "Destination"
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
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsDestination As ADODB.Recordset
'load form,populate combo data,set navbar button

Private Sub Form_Load()
Dim ctl As Control
    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Destination")
    '------------------------------------------
    
    Screen.MousePointer = vbHourglass
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.Destination(deIms.NameSpace)
    ' Set rsDestination = deIms.rsDestination 'M
    Set NavBar1.Recordset = deIms.rsDestination
    
    
    Visible = True
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    Set SSDBGDestination.DataSource = deIms
    Caption = Caption + " - " + Tag
    
    
End Sub

' unload form,free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Hide
    deIms.rsDestination.CancelBatch
    
    deIms.rsDestination.Close
    If open_forms <= 5 Then ShowNavigator
    
    If Err Then Err.Clear
End Sub

'cancel recordset update

Private Sub NavBar1_BeforeCancelClick()
    SSDBGDestination.CancelUpdate
End Sub

'set record sset update

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGDestination.Update
End Sub

'set recordset add new

Private Sub NavBar1_BeforeNewClick()
    SSDBGDestination.AddNew
End Sub

'before save records set record update

Private Sub NavBar1_BeforeSaveClick()
    SSDBGDestination.Update
     'SSDBGDestination.Refresh
    Call SSDBGDestination.MoveRecords(0)
    
End Sub

'cancel recordset update

Private Sub NavBar1_OnCancelClick()
    SSDBGDestination.CancelUpdate
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'set name space equal to current name space

Private Sub NavBar1_OnNewClick()
    deIms.rsDestination!des_npecode = deIms.NameSpace
End Sub

'get crystal report paramenter and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Destination.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        msg1 = translator.Trans("L00053") 'J added
        .WindowTitle = IIf(msg1 = "", "Destination", msg1) 'J modified
        Call translator.Translate_Reports("Destination.rpt") 'J added
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

'save record set

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
    Call deIms.rsDestination.Move(0)
    If Err Then Err.Clear
    
End Sub

Private Sub SSDBGDestination_AfterUpdate(RtnDispErrMsg As Integer)
'SSDBGDestination.ReBind
MsgBox "Changes Saved"
  'SSDBGDestination.Move (0)
End Sub

Private Sub SSDBGDestination_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldValue As Variant, Cancel As Integer)
Dim Recchanged As Boolean
Dim ret As Integer
      
'   If TransCancelled = False Then
   
     If SSDBGDestination.IsAddRow And ColIndex = 0 Then 'And TMPCTL.RecordToProcess.editmode = adEditAdd Then
             If NotValidLen(SSDBGDestination.Columns(ColIndex).Value) Then
                MsgBox ("Required field, please enter value.")
                Cancel = 1
                SSDBGDestination.SetFocus
                SSDBGDestination.Columns(ColIndex).Value = oldValue
                SSDBGDestination.col = 0
                'GoodColMove = False
              ElseIf CheckDesCode(SSDBGDestination.Columns(ColIndex).Value) Then
                MsgBox ("This code already exists. Please choose a unique value.")
                Cancel = 1
                SSDBGDestination.SetFocus
                SSDBGDestination.Columns(ColIndex).Value = oldValue
                SSDBGDestination.col = 0
                'GoodColMove = False
             End If
        Else
          '  RecChanged = NavBar1.Recordset.Fields.s
            Recchanged = DidFieldChange(Trim(oldValue), Trim(SSDBGDestination.Columns(ColIndex).Value))
        End If
    
        
End Sub


'set data grip value to current name space
Private Sub SSDBGDestination_BeforeUpdate(Cancel As Integer)
 Dim Response As String
  
    
  '  Cancel = 0
  'Else
     Response = MsgBox("Record have been Modified,Would you Like to Save the Changes", vbOKCancel, "Imswin")
     If Response = vbOK Then
       If SSDBGDestination.IsAddRow Then SSDBGDestination.Columns("np").text = deIms.NameSpace
     Cancel = 0
     Else
     Cancel = -1
   End If
  
End Sub

Private Function NotValidLen(Code As String) As Boolean

On Error Resume Next
If Len(Trim(Code)) > 0 Then
    NotValidLen = False
Else
    NotValidLen = True
End If
End Function



'Added 11/20/00 by S. McMorrow to check for duplicate key valuePrivate Function CheckDesCode(Code As String) As Boolean
Private Function CheckDesCode(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
         .CommandText = .CommandText & " From destination "
        .CommandText = .CommandText & " Where des_npecode = '" & deIms.NameSpace & "'"
       .CommandText = .CommandText & " AND des_destcode = '" & Code & "'"
       
        Set rst = .Execute
        CheckDesCode = rst!rt
    End With
       
     Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckDesCode", Err.Description, Err.number, True)
End Function

Private Function DidFieldChange(strOldValue As String, strNewValue As String)
Dim ret
    ret = StrComp(Trim(strOldValue), Trim(strNewValue), vbTextCompare)
            If ret <> 0 Then
                DidFieldChange = True
            Else
                DidFieldChange = False
            End If

End Function

