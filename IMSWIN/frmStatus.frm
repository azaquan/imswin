VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   6030
   Tag             =   "01040100"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frmStatus.frx":0000
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGStatua 
      Height          =   2955
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   5295
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
      stylesets(0).Picture=   "frmStatus.frx":001C
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
      stylesets(1).Picture=   "frmStatus.frx":0038
      HeadFont3D      =   4
      DefColWidth     =   5292
      BevelColorFrame =   -2147483630
      BevelColorHighlight=   14737632
      BevelColorShadow=   -2147483633
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
      Columns(0).Width=   1746
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "sts_code"
      Columns(0).FieldLen=   2
      Columns(0).HeadStyleSet=   "Colls"
      Columns(0).StyleSet=   "Rows"
      Columns(1).Width=   6641
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "sts_name"
      Columns(1).FieldLen=   30
      Columns(1).HeadStyleSet=   "Colls"
      Columns(1).StyleSet=   "Rows"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "np"
      Columns(2).Name =   "np"
      Columns(2).DataField=   "sts_npecode"
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   9340
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
      Caption         =   "Status"
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
      Left            =   420
      TabIndex        =   1
      Top             =   240
      Width           =   5190
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'load form and set back ground color

Private Sub Form_Load()
Dim ctl As Control
Dim rst As ADODB.Recordset

    
    Screen.MousePointer = vbHourglass
    
    'Added by Juan (9/25/2000) for Multilingual
    Call translator.Translate_Forms("frmStatus")
    '------------------------------------------
    
    'color the controls and form backcolor
    Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        Call gsb_fade_to_black(ctl)
    Next ctl
        
    If deIms.rssStatus.State And adStateOpen Then
        Set rst = deIms.rssStatus.Clone
    Else
        
        deIms.sStatus (deIms.NameSpace)
        Set rst = deIms.rssStatus.Clone
    End If
    
    
    Set SSDBGStatua.DataSource = rst.Clone
    
    rst.Close
    Set rst = Nothing
    
    Screen.MousePointer = vbDefault
    Call DisableButtons(Me, NavBar1)
    
    Caption = Caption + " - " + Tag
End Sub

'cancel update

Private Sub NavBar1_BeforeCancelClick()
    SSDBGStatua.CancelUpdate
End Sub

'before save set add new position and name space to current name space

Private Sub NavBar1_BeforeNewClick()
    SSDBGStatua.AddNew
    SSDBGStatua.Columns("np").Value = deIms.NameSpace
End Sub

'unload form set recordset close and free memory

Private Sub Form_Unload(Cancel As Integer)
    Hide
    SSDBGStatua.Update
    Set frmStatus = Nothing
    Set SSDBGStatua.DataSource = Nothing
    If open_forms <= 5 Then ShowNavigator
End Sub

'cancel update

Private Sub NavBar1_OnCancelClick()
    SSDBGStatua.CancelUpdate
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'delete a record from recordset

Private Sub NavBar1_OnDeleteClick()
    SSDBGStatua.DeleteSelected
End Sub

'move recordset to first position

Private Sub NavBar1_OnFirstClick()
    SSDBGStatua.MoveFirst
End Sub

'move recordset to last position

Private Sub NavBar1_OnLastClick()
    SSDBGStatua.MoveLast
End Sub

'before add new set recordset to add new positon and
'set user to current user and name space to current name space

Private Sub NavBar1_OnNewClick()
    SSDBGStatua.AddNew
    deIms.rssStatus!sts_creauser = CurrentUser
    SSDBGStatua.Columns("np").Value = deIms.NameSpace
End Sub

'move recordset to next position

Private Sub NavBar1_OnNextClick()
    SSDBGStatua.MoveNext
End Sub

'move recordset to previous position

Private Sub NavBar1_OnPreviousClick()
    SSDBGStatua.MovePrevious
End Sub

'set crystal report parameters

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

   With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Status.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/25/2000) for Multilingual
        msg1 = translator.Trans("L00110") 'J added
        .WindowTitle = "Status" 'J modified
        Call translator.Translate_Reports("Status.rpt") 'J added
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

'before save recordset check code exist or not
'if code is existed, show message

Private Sub NavBar1_OnSaveClick()
On Error Resume Next
Dim Numb As Integer
Dim number As Integer
Dim numbe As Integer

  
    
    Numb = SSDBGStatua.Rows
    number = SSDBGStatua.GetBookmark(-1)
    numbe = SSDBGStatua.Bookmark
    
    If (Numb - number) = 1 And (Numb > numbe) Then
        If Len(Trim$(SSDBGStatua.Columns(0).text)) <> 0 Then
            If CheckStatus(SSDBGStatua.Columns(0).text) Then
            
                'Modified by Juan (9/25/2000) for Multilingual
                msg1 = translator.Trans("M00013") 'J added
                MsgBox IIf(msg1 = "", "Code exist, Please make new one", msg1) 'J modified
                '---------------------------------------------
                
                SSDBGStatua.CancelUpdate: Exit Sub
            End If
        End If
            
    End If
           
    deIms.rssStatus!sts_modiuser = CurrentUser
    
    SSDBGStatua.Update
    Call SSDBGStatua.MoveRecords(0)
    
    
End Sub

'check status code before save recordset

Private Sub SSDBGStatua_BeforeColUpdate(ByVal ColIndex As Integer, ByVal oldVALUE As Variant, Cancel As Integer)
Dim oldstr As String
Dim newstr As String
Dim Numb As Integer
Dim number As Integer
Dim numbe As Integer

    oldstr = SSDBGStatua.Columns(0).CellText(SSDBGStatua.Bookmark)
    newstr = SSDBGStatua.Columns(0).text
    
    Numb = SSDBGStatua.Rows
    number = SSDBGStatua.GetBookmark(-1)
    numbe = SSDBGStatua.Bookmark
    
    If (Numb - number) = 1 And (Numb > numbe) Then
            Exit Sub
        Else
        If ColIndex = 0 Then
            If oldstr <> newstr Then
                Cancel = True
                
                'Modified by Juan (9/25/2000) for Multilingual
                msg1 = translator.Trans("M00015") 'J added
                MsgBox IIf(msg1 = "", "Code can not changed once it is saved, Please make new one.", msg1) 'J modified
                '---------------------------------------------
                
            End If
        End If
    End If
    
End Sub

'SQL statement check status code exist or not

Private Function CheckStatus(Code As String) As Boolean
On Error Resume Next
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        .CommandText = "SELECT count(*) RT"
        .CommandText = .CommandText & " From Status"
        .CommandText = .CommandText & " Where sts_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND sts_code = '" & Code & "'"
               
        Set rst = .Execute
        CheckStatus = rst!rt
    End With
        
    Set cmd = Nothing
    Set rst = Nothing
    If Err Then Call LogErr(Name & "::CheckStatus", Err.Description, Err.number, True)

End Function

