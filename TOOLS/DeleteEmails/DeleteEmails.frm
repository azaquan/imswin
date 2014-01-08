VERSION 5.00
Begin VB.Form FrmTrasmission 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete cancelled Emails"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
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
      Caption         =   "Deleting !houemec, !luanda03 Email Accounts"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   3615
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
Dim str As String
Dim Cmd As New ADODB.Command
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Message As String
Dim subject As String
Dim emailAddresses() As String
Dim Attachment As String
Dim Machinename As String
Dim InTransaction As Boolean

On Error GoTo handler

Label1.Visible = True
DoEvents
Screen.MousePointer = vbHourglass

Set Gcn = New ADODB.Connection
Gcn.CursorLocation = adUseClient
Gcn.CommandTimeout = 1000
Gcn.Open "driver={SQL Server};server=.;uid=sa;pwd=0eGxPx4;database=sakhalin"

'Gcn.Open "driver={SQL Server};server=IMSSQL001;uid=sa;pwd=0eGxPx4;database=TEST"

Gcn.Errors.Clear

Gcn.BeginTrans
InTransaction = True

str = "select psys_sttn from pesys"
Rs.Source = str
Rs.Open , Gcn

Machinename = UCase(Trim(Rs!psys_sttn))

str = "delete from distribution where dis_mail in ('!HOUEMEC','!LUANDA03')"
Rs1.Source = str
Rs1.Open , Gcn, 3, 3

str = "delete from ud_common where com_mail in ('!HOUEMEC','!LUANDA03')"
Rs2.Source = str
Rs2.Open , Gcn, 3, 3

Gcn.CommitTrans

Gcn.Close
subject = "!HOUEMEC, !LUANDA03 email accounts has been successfully deleted from " & Machinename & "."
Message = "!HOUEMEC, !LUANDA03 email accounts has been successfully deleted from " & Machinename & "."

ReDim emailAddresses(4)
emailAddresses(0) = "MUZAMMIL@IMS-SYS.COM"
emailAddresses(1) = "ddegrazia@IMS-SYS.COM"
emailAddresses(2) = "FARMSTRONG@IMS-SYS.COM"
emailAddresses(3) = "jeb.burch@exxonmobil.com"

Call sendEmailOnly(Message, subject, emailAddresses(), Attachment)

MsgBox "!HOUEMEC, !LUANDA03 email accounts has been deleted successfully.", vbInformation, "Ims"

Label1.Visible = False

Unload Me

Exit Sub

handler:

    If InTransaction = True Then Gcn.RollbackTrans
    
    Message = "Errors occurred while trying to delete the list of cancelled email accounts (!HOUEMEC,!LUANDA03) from mailing list. Error desc :" & Err.Description
    
    subject = "Deleting Cancelled Email accounts."
    
    Call sendEmailOnly(Message, subject, emailAddresses(), Attachment)
    
    MsgBox "Errors occurred while trying to delete the list of cancelled email accounts  (!HOUEMEC,!LUANDA03) from mailing list. Error desc :" & Err.Description, vbCritical, "Ims"
    
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
Private Sub Form_Load()

End Sub
Sub sendEmailOnly(Message As String, subject As String, emailAddresses() As String, Attachment As String)

    On Error GoTo ErrHand
    
    Dim nullAttachments() As String
    Call SendAttMail(Message, subject, emailAddresses, nullAttachments)
    
    Exit Sub
ErrHand:
    
    MsgBox "Could not sent out Success/Failure Email. " & Err.Description
    Err.Clear
End Sub
