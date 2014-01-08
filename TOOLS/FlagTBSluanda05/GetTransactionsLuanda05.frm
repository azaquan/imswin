VERSION 5.00
Begin VB.Form FrmTrasmission 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flag Transactions"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "GetTransactionsLuanda05.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Flag Transactions"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Running ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   465
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Flaging the transmissions for update database."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmTrasmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gcn As ADODB.Connection
Dim NAMESPACE As String
''Type Data
''    Tablename As String
''    Datastring As String
''End Type

Private Sub Command1_Click()

Dim Gcn As ADODB.Connection
Dim str As String
Dim Cmd As New ADODB.Command
On Error GoTo handler

Label1.Visible = True
DoEvents
Screen.MousePointer = vbHourglass


Set Gcn = New ADODB.Connection
Gcn.CursorLocation = adUseClient
Gcn.CommandTimeout = 1000
Gcn.Open "driver={SQL Server};server=IMSSQL001;uid=sa;pwd=0eGxPx4;database=sakhalin"

Gcn.Errors.Clear

Gcn.BeginTrans

str = " update  invtreceipt set ir_tbs=1 where ir_trannumb in ()"

str = " update  invtreceipt set ir_tbs=1 where ir_trannumb in ('I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"
str = str & " update  invtreceiptdetl set ird_tbs=1 where ird_trannumb in ('I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"
str = str & " update  invtreceiptrem set irr_tbs=1 where irr_trannumb in ('152','I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"

str = str & " update  invtissue set ii_tbs=1 where ii_trannumb in ('I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"
str = str & " update  invtissuedetl set iid_tbs=1 where iid_trannumb in ('I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"
str = str & " update  invtissuerem set iir_tbs=1 where iir_trannumb in ('I-65667','RT-65668','RT-65669','IT-65670','IT-65671','RT-65672','I-65673','R-65674','IT-65675','AE-65676','R-65677','R-65678','R-65679','I-65680','RT-65681','I-65682','I-65683','RT-65684','RT-65685','IT-65686','R-65687','IT-65688','IT-65689','IT-65690','IT-65691','IT-65692','IT-65693','I-65694','R-65695','I-65696','RT-65697','I-65698','I-65699','IT-65700','IT-65701','I-65702','R-65703','I-65704','RT-65705')"

Cmd.CommandText = str
Cmd.CommandType = adCmdText
Cmd.ActiveConnection = Gcn
Cmd.Execute

If Gcn.Errors.Count = 0 Then
    Gcn.CommitTrans
Else
    Gcn.RollbackTrans
End If

MsgBox "The transactions have been Flaged successfully.", vbInformation, "Ims"

Gcn.Close

Label1.Visible = False

Unload Me

Exit Sub

handler:

    Gcn.RollbackTrans
    MsgBox "Errors occurred while trying to flag the transactions. Error Description :" & Err.Description, vbCritical, "Ims"
    Err.Clear
    
Label1.Visible = False
Unload Me
End Sub

Public Function LogMessage(MessageToLog As String)
  
Dim sa As New Scripting.FileSystemObject

Dim t As Scripting.TextStream

logfile = App.Path & "\TransactionDetails.txt"

If sa.FileExists(logfile) = False Then

    sa.CreateTextFile logfile, True
    
End If

Set t = sa.OpenTextFile(logfile, ForAppending)

t.WriteLine MessageToLog

t.Close

Set t = Nothing

End Function

''''Public Function ReadFromTbs() As Boolean
''''Dim Data() As Data
''''Dim sa As New Scripting.FileSystemObject
''''Dim t As Scripting.TextStream
''''Dim line As String
''''On Error GoTo ErrHand
''''
''''If sa.FileExists(App.Path & "\flagtbs.txt") = False Then GoTo ErrHand
''''
'''' Set t = sa.OpenTextFile(App.Path & "\flagtbs.txt")
''''
'''' Do While Not t.AtEndOfStream
''''
''''    line = Trim(t.ReadLine)
''''
''''    If InStr(line, "*") > 0 Then
''''
''''        ReDim Preserve Data(i)
''''        Data(i).Tablename = Mid(line, 2, Len(line))
''''
''''    End If
''''
''''
''''
''''
'''' Loop
''''
''''
''''Exit Function
''''ErrHand:
''''
''''Err.Clear
''''End Function
