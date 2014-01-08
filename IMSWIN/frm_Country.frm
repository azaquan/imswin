VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Begin VB.Form frm_Country 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Country"
   ClientHeight    =   4095
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   5445
   Tag             =   "01030300"
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   690
      TabIndex        =   2
      Top             =   3540
      Width           =   3525
      _ExtentX        =   6218
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
      DisableSaveOnSave=   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
      DeleteToolTipText=   ""
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGCountry 
      Height          =   3015
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   5115
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      stylesets(0).Picture=   "frm_Country.frx":0000
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
      stylesets(1).Picture=   "frm_Country.frx":001C
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
      ExtraHeight     =   185
      Columns.Count   =   3
      Columns(0).Width=   1905
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "ctry_code"
      Columns(0).DataType=   8
      Columns(0).Case =   2
      Columns(0).FieldLen=   3
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   6324
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "ctry_name"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   40
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NP"
      Columns(2).Name =   "NP"
      Columns(2).DataField=   "CTRY_NPECODE"
      Columns(2).FieldLen=   256
      _ExtentX        =   9022
      _ExtentY        =   5318
      _StockProps     =   79
      DataMember      =   "COUNTRY"
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
   Begin VB.Label lbl_Title 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Left            =   165
      TabIndex        =   1
      Top             =   60
      Width           =   5130
   End
End
Attribute VB_Name = "frm_Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'load form get data for combo, set each control, and button

Private Sub Form_Load()
Dim ctl As Control

    'Added by Juan (9/11/2000) for Multilingual
    Call translator.Translate_Forms("frm_Country")
    '------------------------------------------

    Screen.MousePointer = vbHourglass
    'color the controls and form backcolor
    'Me.BackColor = frm_Color.txt_WBackground.BackColor
    
    For Each ctl In Controls
        'If Not (TypeOf ctl Is Toc) Then Call gsb_fade_to_black(ctl)
    Next ctl
    
    Call deIms.Country(deIms.NameSpace)
    Set SSDBGCountry.DataSource = deIms
    Set NavBar1.Recordset = deIms.rsCOUNTRY
    Screen.MousePointer = vbDefault
     Call DisableButtons(Me, NavBar1)
     Caption = Caption + " - " + Tag
     

End Sub

'unload form free memory

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Hide
    deIms.rsCOUNTRY.Update
    deIms.rsCOUNTRY.CancelUpdate
    
    deIms.rsCOUNTRY.Close
    If Err Then Err.Clear
    If open_forms <= 5 Then ShowNavigator
End Sub

'cancel update

Private Sub NavBar1_BeforeCancelClick()
    SSDBGCountry.CancelUpdate
End Sub

'update recordset

Private Sub NavBar1_BeforeMove(bCancel As Boolean)
    SSDBGCountry.Update
End Sub

'before update new record set record update

Private Sub NavBar1_BeforeNewClick()
    SSDBGCountry.Update
    SSDBGCountry.AddNew
End Sub

'before save update record set

Private Sub NavBar1_BeforeSaveClick()
    SSDBGCountry.Update
End Sub

'close form

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

'get parameter for crystal report and application path

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\Country.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/11/2000) for Multilingual
        Call translator.Translate_Reports("Country.rpt") 'J added
        msg1 = translator.Trans("L00006") 'J added
        .WindowTitle = IIf(msg1 = "", "Country", msg1) 'J modified
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

Private Sub SSDBGCountry_AfterUpdate(RtnDispErrMsg As Integer)

MsgBox "Record saved successfully.", vbInformation, "Imswin"

End Sub

Private Sub SSDBGCountry_BeforeRowColChange(Cancel As Integer)
Dim good_field As Boolean
    good_field = validate_fields(SSDBGCountry.Col)
    If Not good_field Then
       Cancel = True
    End If

End Sub

Private Sub SSDBGCountry_BeforeUpdate(Cancel As Integer)
'deIms.rsCOUNTRY!CTRY_NPECODE = deIms.NameSpace
SSDBGCountry.Columns(2).text = deIms.NameSpace
End Sub

Private Function validate_fields(colnum As Integer) As Boolean
Dim x As Boolean

validate_fields = True
If SSDBGCountry.IsAddRow Then
   
   
   If colnum = 0 Or colnum = 1 Or colnum = 2 Then
    x = NotValidLen(SSDBGCountry.Columns(colnum).text)
      
      
          If x = True Then
    
             msg1 = translator.Trans("M00702")
             MsgBox IIf(msg1 = "", "Required field, please enter value", msg1)
             SSDBGCountry.SetFocus
             SSDBGCountry.Col = colnum
             validate_fields = False
             Exit Function
          End If
      
    End If
      
      If colnum = 0 Then
            x = CheckDesCode(SSDBGCountry.Columns(0).text)
            If x <> False Then
    '             RecSaved = False
                 msg1 = translator.Trans("M00703")
                 MsgBox IIf(msg1 = "", "This code already exists. Please choose a unique value.", msg1)
                 SSDBGCountry.SetFocus
                 SSDBGCountry.Col = 0
                 SSDBGCountry.Columns(0).text = ""
                validate_fields = False
             End If
      End If
         
End If

End Function

Public Function CheckDesCode(Code As String) As Boolean
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Code = Trim$(Code)

rs.ActiveConnection = deIms.cnIms
rs.Source = "select count(*) NUMBER from country where ctry_CODE='" & Code & "' and ctry_npecode='" & deIms.NameSpace & "'"
rs.Open

 CheckDesCode = IIf(rs!number = 0, False, True)
 

End Function
