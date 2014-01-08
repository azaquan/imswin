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
      Caption         =   "Extract"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
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
      Caption         =   "Extracts a list of all the transactions and there line items."
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

Private Sub Command1_Click()

Dim Gcn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim str As String
Dim sa As New Scripting.FileSystemObject
Dim Logfile As String
On Error GoTo handler

Label1.Visible = True
DoEvents
Screen.MousePointer = vbHourglass

Logfile = App.Path & "\TransactionDetails.txt"

If sa.FileExists(Logfile) Then sa.DeleteFile Logfile, True

Call LogMessage(Now)

Set Gcn = New ADODB.Connection
Gcn.CursorLocation = adUseClient
Gcn.CommandTimeout = 1000
Gcn.Open "driver={SQL Server};server=.;uid=sa;pwd=0eGxPx4;database=sakhalin"

Set Rs = New ADODB.Recordset

str = " select distinct psys_sttn 'Machine name' from pesys "

Rs.Source = str
Rs.ActiveConnection = Gcn
Rs.Open
Call ThrowDatainFlatFiles(Rs)

str = " select chr_code , chr_fld3 from chrono where chr_npecode = 'Angol'"

Set Rs = Nothing
Set Rs = New ADODB.Recordset
Rs.Source = str
Rs.ActiveConnection = Gcn
Rs.Open
Call ThrowDatainFlatFiles(Rs)


str = " SELECT IR_TRANNUMB 'Transaction No', SUBSTRING(IR_TRANNUMB, CHARINDEX('-',IR_TRANNUMB) +1, LEN(IR_TRANNUMB)) 'TRANNUMB' , IR_TRANDATE ,"
str = str & " (select count(*) from invtreceiptdetl where IRD_NPECODE =IR_NPECODE AND IR_TRANNUMB =IRD_TRANNUMB) LineItems"
str = str & " FROM INVTRECEIPT WHERE CHARINDEX('-',IR_TRANNUMB) > 0 AND IR_TRANDATE > '12/31/2001' AND IR_NPECODE='ANGOL'"
str = str & " Union"

str = str & " SELECT II_TRANNUMB 'Transaction No', SUBSTRING(II_TRANNUMB, CHARINDEX('-',II_TRANNUMB) +1, LEN(II_TRANNUMB)) 'TRANNUMB', II_TRANDATE   ,"
str = str & " (select count(*) from invtissuedetl where IiD_NPECODE =Ii_NPECODE AND Ii_TRANNUMB =IiD_TRANNUMB) LineItems"
str = str & " FROM INVTISSUE WHERE CHARINDEX('-',II_TRANNUMB) > 0 AND II_TRANDATE  > '12/31/2001' AND II_NPECODE='ANGOL'"
str = str & " ORDER BY   TRANNUMB ASC"


Set Rs = Nothing
Set Rs = New ADODB.Recordset
Rs.Source = str
Rs.ActiveConnection = Gcn
Rs.Open
Call ThrowDatainFlatFiles(Rs)


'Rs.Source = str
'Rs.ActiveConnection = Gcn
'Rs.Open
'Rs.GetString

'Call UnbindRule(Gcn)

MsgBox "The transactions have been extracted successfully.", vbInformation, "Ims"

Gcn.Close

Label1.Visible = False

Unload Me

Exit Sub

handler:

        
    MsgBox "Errors occurred while trying to extract the Transactions. Error Description :" & Err.Description, vbCritical, "Ims"
    Err.Clear
    
Label1.Visible = False
Unload Me
End Sub

Public Function ThrowDatainFlatFiles(Rs As ADODB.Recordset) As String
Dim x As String
On Error GoTo ErrHand


    
      x = Rs.GetString(, , vbTab, vbCrLf)
    
      Call LogMessage(x)

Exit Function
ErrHand:

End Function

Public Function LogMessage(MessageToLog As String)
  
Dim sa As New Scripting.FileSystemObject

Dim t As Scripting.TextStream

Logfile = App.Path & "\TransactionDetails.txt"

If sa.FileExists(Logfile) = False Then

    sa.CreateTextFile Logfile, True
    
End If

Set t = sa.OpenTextFile(Logfile, ForAppending)

t.WriteLine MessageToLog

t.Close

Set t = Nothing

End Function

Public Function UnbindRule(Cn As ADODB.Connection)
On Error GoTo Errhandler

Dim Rs As New ADODB.Recordset
Rs.Source = "exec sp_unbindrule N'[CREATEUSER]'"
Rs.Open , Cn
Call LogMessage("UnB D S")
Exit Function

Errhandler:
 Call LogMessage("UnB D Us. " & Err.Description)
 Err.Clear

End Function
