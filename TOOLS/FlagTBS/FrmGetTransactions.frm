VERSION 5.00
Begin VB.Form FrmTrasmission 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Transactions"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "FrmGetTransactions.frx":0000
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
      Caption         =   "Flagging the transmissions for update database."
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
Gcn.Open "driver={SQL Server};server=ims;uid=sa;pwd=0eGxPx4;database=sakhalin"

Gcn.Errors.Clear

Gcn.BeginTrans


str = " update invtreceipt set ir_tbs=1 where ir_trannumb like 'AE-20048' and ir_npecode ='libya' "
str = str & " update invtreceiptdetl set ird_tbs=1 where ird_trannumb like 'AE-20048' and ird_npecode ='libya' "
 
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
    MsgBox "Errors occurred while trying to flag the Transactions. Error Description :" & Err.Description, vbCritical, "Ims"
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
