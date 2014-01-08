VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNavigatorsX.ocx"
Begin VB.Form FrmManifestPOD 
   Caption         =   "Manifest POD"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   5625
   Tag             =   "02030400"
   Begin LRNavigators.NavBar NavBar1 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      Style           =   1
      MouseIcon       =   "frmmanifestpod.frx":0000
      CancelVisible   =   0   'False
      PreviousVisible =   0   'False
      NewVisible      =   0   'False
      PrintVisible    =   0   'False
      LastVisible     =   0   'False
      NextVisible     =   0   'False
      FirstVisible    =   0   'False
      EditVisible     =   -1  'True
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
      NewEnabled      =   0   'False
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
      DisableSaveOnSave=   0   'False
   End
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "12/31/9999 12:00:00 AM"
      Format          =   60489729
      CurrentDate     =   40112
   End
   Begin VB.TextBox lblTxt 
      Height          =   375
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBManifest 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      DataFieldList   =   "Column 0"
      AllowNull       =   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   6165
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin MSComCtl2.DTPicker dtptime 
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60489730
      CurrentDate     =   36494
   End
   Begin VB.Label lbldatetime 
      Caption         =   "Date/ Time"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblmanifest 
      Caption         =   "Packing/Manifest"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmmanifestpod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub fillupCombo()
Dim rsManifestList As ADODB.Recordset
Set rsManifestList = GetManifestList

    With rsManifestList
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                 Call SSOleDBManifest.AddItem(rsManifestList!pl_manfnumb, 0)
                 
                .MoveNext
            Loop
        End If

    End With
    
    SSOleDBManifest.Columns(0).Width = SSOleDBManifest.Width
End Sub

Private Sub Form_Load()
Me.Height = 2805
Me.Width = 5745
NavBar1.LastPrintSepVisible = False
NavBar1.PrintSaveSepVisible = False
NavBar1.CancelLastSepVisible = False
NavBar1.EditVisible = False

Call fillupCombo
    

Me.dtpdate.value = FormatDateTime(Now, vbShortDate)
Me.dtptime.value = FormatDateTime(Now, vbShortTime)

End Sub

Private Function GetManifestList() As Recordset

Dim Sql As String
Dim cmd As ADODB.Command
Dim rsManifestList As ADODB.Recordset

On Error GoTo ErrHandler

    Set rsManifestList = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    SSOleDBManifest.RemoveAll
   
    With cmd
        .CommandText = "PODPackingListNumbers"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = deIms.cnIms
        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)
        Set rsManifestList = .Execute
        
    End With
    
    Set cmd = Nothing
    
Set GetManifestList = rsManifestList
'SSOleDBManifest.Columns.Add (0)

    Exit Function
ErrHandler:

    MsgBox "Errors occured while trying to retrieve packing list."
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetPODDetailsForManifest() As Recordset

End Function



Private Sub NavBar1_OnCloseClick()
Unload Me
End Sub

Private Sub NavBar1_OnSaveClick()
On Error GoTo Error

If Len(Trim(lblTxt.Text)) = 0 Then
Call MsgBox("You cannot save without entering a valid name.", vbOKOnly, "IMS")
Exit Sub

End If

Dim cmd As ADODB.Command

Dim Name As String
Dim packinglist As String
Dim datetime As Date
Dim dt As String
Name = lblTxt.Text
packinglist = Trim(SSOleDBManifest.Text)
dt = dtpdate.value & " " & dtptime.value

    Set cmd = MakeCommand(deIms.cnIms, adCmdStoredProc)

    With cmd
        .Prepared = True
        .CommandText = "pod_save"
        
        .parameters.Append .CreateParameter("@packinglist", adVarChar, adParamInput, 50, Trim(packinglist))
        .parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 100, Name)
        .parameters.Append .CreateParameter("@datetime", adVarChar, adParamInput, 30, dt)
        .parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, deIms.NameSpace)


        Call .Execute

    End With

    Set cmd = Nothing

Call MsgBox("Saved successfully", vbOKOnly, "IMS")

Exit Sub
Error:

MsgBox ("Errors occurred while trying to save, Error Descrioption : " & Err.Description)

End Sub


Private Sub SSOleDBManifest_DropDown()
    fillupCombo
End Sub

Private Sub SSOleDBManifest_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SSOleDBManifest.DroppedDown Then SSOleDBManifest.DroppedDown = True
End Sub

Private Sub SSOleDBManifest_GotFocus()
 Call HighlightBackground(SSOleDBManifest)
End Sub

Private Sub SSOleDBManifest_LostFocus()
Call NormalBackground(SSOleDBManifest)
SSOleDBManifest_Validate (False)
End Sub


Private Sub SSOleDBManifest_Validate(Cancel As Boolean)

Dim rsManifestList As Recordset

If Len(Trim(SSOleDBManifest.Text)) = 0 Then Exit Sub

Set rsManifestList = GetManifestList()

        rsManifestList.MoveFirst
        rsManifestList.Find "pl_manfnumb='" & Trim$(SSOleDBManifest.Text) & "'"
        If rsManifestList.EOF Then
        
          MsgBox "Please enter a valid manifest, the one that you entered does not exists."
          SSOleDBManifest.Text = ""
          Cancel = True
            SSOleDBManifest.SetFocus
        End If
 

End Sub
