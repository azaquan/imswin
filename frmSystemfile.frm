VERSION 5.00
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form frm_systemfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System File"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   7230
   Tag             =   "01040700"
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSystemfile.frx":0000
      Left            =   5640
      List            =   "frmSystemfile.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   1920
      TabIndex        =   36
      Top             =   5160
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "frmSystemfile.frx":001A
      EditEnabled     =   -1  'True
   End
   Begin VB.CheckBox Chkresend 
      Caption         =   "Check2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox Chkmodifi 
      BackColor       =   &H8000000A&
      Caption         =   "Check1"
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   5640
      TabIndex        =   32
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtfreout 
      Height          =   315
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtfrein 
      Height          =   315
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3600
      Width           =   495
   End
   Begin VB.CheckBox Chksend 
      Height          =   315
      Left            =   6360
      TabIndex        =   24
      Top             =   3960
      Width           =   375
   End
   Begin VB.CheckBox chkback 
      Height          =   315
      Left            =   6360
      TabIndex        =   22
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox txtCompany 
      Height          =   315
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtSite 
      Height          =   315
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtshipcode 
      Height          =   315
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Mns"
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   3990
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Mns"
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   3630
      Width           =   855
   End
   Begin VB.Label Lblgateway 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Gateway"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   3150
      Width           =   1800
   End
   Begin VB.Label Lblbackout 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label Label22 
      Caption         =   "Out Basket"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4710
      Width           =   1800
   End
   Begin VB.Label Lblbackin 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label Label20 
      Caption         =   "In Basket"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   4350
      Width           =   1800
   End
   Begin VB.Label Label19 
      Caption         =   "Allow Resending"
      Height          =   255
      Left            =   3720
      TabIndex        =   28
      Top             =   1470
      Width           =   2000
   End
   Begin VB.Label Label18 
      Caption         =   "Allow Modification"
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   1110
      Width           =   2000
   End
   Begin VB.Label Label17 
      Caption         =   "Frequency Out"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3990
      Width           =   1800
   End
   Begin VB.Label Label16 
      Caption         =   "Frequency In"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3630
      Width           =   1800
   End
   Begin VB.Label Label15 
      Caption         =   "Send Update Database"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   4000
      Width           =   2565
   End
   Begin VB.Label Label14 
      Caption         =   "Receive Update Database"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   3630
      Width           =   2565
   End
   Begin VB.Label Lbllocation 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Unit of Measure"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   1830
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label LblConsolidation 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Consolidation"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2550
      Width           =   1800
   End
   Begin VB.Label Lblstation 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label LblshipName 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Company Code"
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   2550
      Width           =   2000
   End
   Begin VB.Label Label6 
      Caption         =   "Location"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   2190
      Width           =   2000
   End
   Begin VB.Label Label5 
      Caption         =   "Site"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2190
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Station"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "Shipper Code"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1470
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "Shipper Name"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1110
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "System File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frm_systemfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableLocked As Boolean, currentformname As String   'jawdat

Private Sub chkback_GotFocus()
Call HighlightBackground(chkback)
End Sub

Private Sub chkback_LostFocus()
Call NormalBackground(chkback)
End Sub

Private Sub Chkmodifi_GotFocus()
Call HighlightBackground(Chkmodifi)
End Sub

Private Sub Chkmodifi_LostFocus()
Call NormalBackground(Chkmodifi)
End Sub

Private Sub Chkresend_GotFocus()
Call HighlightBackground(Chkresend)
End Sub

Private Sub Chkresend_LostFocus()
Call NormalBackground(Chkresend)
End Sub

Private Sub Chksend_GotFocus()
Call HighlightBackground(Chksend)
End Sub

Private Sub Chksend_LostFocus()
Call NormalBackground(Chksend)
End Sub

Private Sub Combo1_GotFocus()
Call HighlightBackground(Combo1)
End Sub

Private Sub Combo1_LostFocus()
Call NormalBackground(Combo1)
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

'SQL statement get information for form
'and load data to form

Public Sub Getrecordfortable()
On Error GoTo CleanUp
Dim str As String
Dim cmd As ADODB.Command
Dim rst As ADODB.Recordset

    Set cmd = MakeCommand(deIms.cnIms, adCmdText)
    
    With cmd
        
        .CommandText = " SELECT psys_shipname,psys_allwmodi, psys_allwresd, psys_shipcode, "
        .CommandText = .CommandText & " psys_sttn, psys_uom, psys_cons, psys_site, "
        .CommandText = .CommandText & " psys_sendupdt, psys_recvupdt, psys_inbskt, "
        .CommandText = .CommandText & " psys_outbskt, psys_gateway, psys_udinfreq, "
        .CommandText = .CommandText & " psys_udoutfreq , psys_compcode, psys_ware "
        .CommandText = .CommandText & " From PESYS "
        .CommandText = .CommandText & " WHERE psys_npecode = '" & deIms.NameSpace & "'"
        .CommandText = .CommandText & " AND psys_usercode =  'PE' "
    

        Set rst = .Execute
    End With
    
    If rst.RecordCount = 0 Then GoTo CleanUp
        
          
        LblshipName.Caption = rst!psys_shipname & ""
        txtshipcode.Text = rst!psys_shipcode & ""
        Lblstation.Caption = rst!psys_sttn & ""
        txtSite.Text = rst!psys_site & ""
        LblConsolidation = rst!psys_cons & ""
        Lbllocation = rst!psys_ware & ""
        txtCompany = rst!psys_compcode & ""
        Lblgateway = rst!psys_gateway & ""
        Lblbackin = rst!psys_inbskt & ""
        Lblbackout = rst!psys_outbskt & ""
        txtfrein = rst!psys_udinfreq & ""
        txtfreout = rst!psys_udoutfreq & ""
        chkback.value = IIf((rst!psys_sendupdt), 1, 0)
        Chksend.value = IIf((rst!psys_recvupdt), 1, 0)
        Chkmodifi.value = IIf((rst!psys_allwmodi), 1, 0)
        Chkresend.value = IIf((rst!psys_allwresd), 1, 0)
        Combo1.ListIndex = IndexOf(Combo1, rst!psys_uom & "")

    
CleanUp:
    'Rst.Close
    Set cmd = Nothing
    Set rst = Nothing

End Sub

' close form

Private Sub cmdClose_Click()
    Unload Me
End Sub

'get crystal report parameter

Private Sub NavBar1_OnPrintClick()
On Error GoTo ErrHandler

    With MDI_IMS.CrystalReport1
        .Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\system.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00127") 'J added
        .WindowTitle = IIf(msg1 = "", "System File", msg1) 'J modified
        Call translator.Translate_Reports("system.rpt") 'J added
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

'before save check data format

Private Sub NavBar1_OnSaveClick()
    
    If Len(Trim$(txtshipcode)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00324") 'J added
        MsgBox IIf(msg1 = "", "The Shipping Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtshipcode.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txtSite)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00325") 'J added
        MsgBox IIf(msg1 = "", "The Site cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtSite.SetFocus: Exit Sub
    End If
    
    If Len(Trim$(txtCompany)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00326") 'J added
        MsgBox IIf(msg1 = "", "The Company Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtCompany.SetFocus: Exit Sub
    End If
        
    If Not Len(Trim$(txtfrein)) Then
        If Not IsNumeric(txtfrein) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00321") 'J added
            MsgBox IIf(msg1 = "", "Frequency In must be numeric", msg1) 'J modified
            '---------------------------------------------
           
           txtfrein.SetFocus: Exit Sub
        ElseIf txtfrein.Text = 0 Then
            
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00322") 'J added
            MsgBox IIf(msg1 = "", "The value must be bigger than 0", msg1) 'J modified
            '---------------------------------------------
            
            txtfrein.SetFocus: Exit Sub
        End If
    End If
    
    If Not Len(Trim$(txtfreout)) Then
        If Not IsNumeric(txtfreout) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00323") 'J added
            MsgBox IIf(msg1 = "", "Frequency out must be numeric", msg1) 'J modified
            '---------------------------------------------
           
           txtfreout.SetFocus: Exit Sub
        ElseIf txtfreout.Text = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00322") 'J added
            MsgBox IIf(msg1 = "", "The value must be bigger than 0", msg1) 'J modified
            '---------------------------------------------
           
           txtfreout.SetFocus: Exit Sub
        End If
    End If
    
     Call UpdateTable

End Sub

'call function get data and set buttons

Private Sub Form_Load()
   
'copy begin here

If NavBar1.SaveEnabled = True Then          ''jawdat, to be put into every form with similar navbar

Dim currentformname
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.CHECK_TABLE(TableLocked, currentformname, deIms.cnIms, CurrentUser)

   If TableLocked = True Then    'sets locked = true because another user has this record open in edit mode
   
NavBar1.SaveEnabled = False
NavBar1.NewEnabled = False
NavBar1.CancelEnabled = False

    Dim textboxes As Control

    For Each textboxes In Controls
        If (TypeOf textboxes Is textBOX) Then
            textboxes.Enabled = False
        End If

    Next textboxes
    
      Dim checkboxes As Control

    For Each checkboxes In Controls
        If (TypeOf checkboxes Is CheckBox) Then
            checkboxes.Enabled = False
        End If

    Next checkboxes
    
    Else
    TableLocked = True
    End If
End If

'end copy
   
   
   
   
   
    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_systemfile")
    '------------------------------------------
   
    Call Getrecordfortable
    Call DisableButtons(Me, NavBar1)
    frm_systemfile.Caption = frm_systemfile.Caption + " - " + frm_systemfile.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub

'call stock procedure to save a record to database

Public Sub UpdateTable()
On Error GoTo Noupdate
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .CommandText = "UPDATEE_PESYS"
        .CommandType = adCmdStoredProc
        .ActiveConnection = deIms.cnIms


        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        .parameters.Append .CreateParameter("@gender", adVarChar, adParamInput, 5, "PE")
        .parameters.Append .CreateParameter("@shipcode", adVarChar, adParamInput, 10, txtshipcode)
        .parameters.Append .CreateParameter("@site", adVarChar, adParamInput, 10, txtSite)
        .parameters.Append .CreateParameter("@udinfreq", adInteger, adParamInput, 2, txtfrein)
        .parameters.Append .CreateParameter("@udoutfreq", adInteger, adParamInput, 2, txtfreout)
        .parameters.Append .CreateParameter("@compcode", adVarChar, adParamInput, 10, txtCompany)
        .parameters.Append .CreateParameter("@sendupdt", adInteger, adParamInput, 1, chkback)
        .parameters.Append .CreateParameter("@recvupdt", adInteger, adParamInput, 1, Chksend)
        .parameters.Append .CreateParameter("@UOM", adVarChar, adParamInput, 4, Combo1.Text)
        .Execute , , adExecuteNoRecords

    End With
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00319") 'J added
        MsgBox IIf(msg1 = "", "Your modifications have been saved", msg1) 'J modified
        '---------------------------------------------
        
    Set cmd = Nothing
    
    
    Exit Sub

Noupdate:
        If Err Then Err.Clear
        
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00320") 'J added
        MsgBox IIf(msg1 = "", "Update System File failed.", msg1) 'J modified
        '---------------------------------------------

End Sub

'before save data check data format

Public Function ValidateData() As Boolean

     ValidateData = False
     
    If Not Len(Trim$(txtfrein)) Then
        If Not IsNumeric(txtfrein) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00321") 'J added
            MsgBox IIf(msg1 = "", "Frequency In must be numeric", msg1) 'J modified
            '---------------------------------------------
           
           txtfrein.SetFocus: Exit Function
        ElseIf txtfrein.Text = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00322") 'J added
            MsgBox IIf(msg1 = "", "The value must be bigger than 0", msg1) 'J modified
            '---------------------------------------------
            
            txtfrein.SetFocus: Exit Function
        End If
    End If
    
    If Not Len(Trim$(txtfreout)) Then
        If Not IsNumeric(txtfreout) Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00323") 'J added
            MsgBox IIf(msg1 = "", "Frequency out must be numeric", msg1) 'J modified
            '---------------------------------------------
            
           txtfreout.SetFocus: Exit Function
        ElseIf txtfreout.Text = 0 Then
        
            'Modified by Juan (9/14/2000) for Multilingual
            msg1 = translator.Trans("M00322") 'J added
            MsgBox IIf(msg1 = "", "The value must be bigger than 0", msg1) 'J modified
            '---------------------------------------------

           txtfreout.SetFocus: Exit Function
        End If
    End If
    
    If Len(Trim$(txtshipcode)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00324") 'J added
        MsgBox IIf(msg1 = "", "The Shipping Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtshipcode.SetFocus:
    End If
    
    If Len(Trim$(txtSite)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00325") 'J added
        MsgBox IIf(msg1 = "", "The Site cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtSite.SetFocus: Exit Function
    End If
    
    If Len(Trim$(txtCompany)) = 0 Then
    
        'Modified by Juan (9/14/2000) for Multilingual
        msg1 = translator.Trans("M00326") 'J added
        MsgBox IIf(msg1 = "", "The Company Code cannot be left empty", msg1) 'J modified
        '---------------------------------------------
        
        txtCompany.SetFocus: Exit Function
    End If

End Function

'unload form

Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator

    
If TableLocked = True Then   'jawdat
Dim imsLock As imsLock.Lock
Set imsLock = New imsLock.Lock
currentformname = Forms(3).Name
Call imsLock.UnLock_table(TableLocked, currentformname, deIms.cnIms, CurrentUser)
End If


End Sub

Private Sub txtCompany_GotFocus()
Call HighlightBackground(txtCompany)
End Sub

Private Sub txtCompany_LostFocus()
Call NormalBackground(txtCompany)
End Sub

Private Sub txtfrein_GotFocus()
Call HighlightBackground(txtfrein)
End Sub

Private Sub txtfrein_LostFocus()
Call NormalBackground(txtfrein)
End Sub

Private Sub txtfreout_GotFocus()
Call HighlightBackground(txtfreout)
End Sub

Private Sub txtfreout_LostFocus()
Call NormalBackground(txtfreout)
End Sub

Private Sub txtshipcode_GotFocus()
Call HighlightBackground(txtshipcode)
End Sub

Private Sub txtshipcode_LostFocus()
Call NormalBackground(txtshipcode)
End Sub
