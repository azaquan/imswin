VERSION 5.00
Begin VB.Form FrmTrasmission 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert SAkha01 Email    V1.013"
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
      Caption         =   "Add SAKHA01 to the distribution mailing list in SAKHALIN namespace."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   5295
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

Gcn.Errors.Clear

Gcn.BeginTrans
InTransaction = True

    str = "select psys_sttn from pesys"
    Rs2.Source = str
    Rs2.Open , Gcn

    Machinename = UCase(Trim(Rs2!psys_sttn))

'DISTRIBUTION TABLE
str = "select * from distribution where dis_mail='!Sakha01'"
Rs.Source = str
Rs.Open , Gcn, 3, 3

If Trim(UCase(Rs.RecordCount)) = 0 Then

    str = "insert into distribution select top 1 dis_npecode, dis_gender, dis_id, '!SAKHA01', dis_faxnumb from distribution where dis_npecode ='sakha' and dis_mail in ('!yuzhn02','!yuzhn01')"

    Cmd.CommandText = str
    Cmd.CommandType = adCmdText
    Cmd.ActiveConnection = Gcn
    Cmd.Execute
    Set Cmd = Nothing
    Set Cmd = New ADODB.Command
    
End If

Set Rs = Nothing
Set Rs = New ADODB.Recordset

str = "select *  from ud_common where com_tonpecode = 'sAkha' and com_mail='!SAKHA01' order by com_npecode"
Rs.Source = str
Rs.Open , Gcn, 3, 3

If Trim(UCase(Rs.RecordCount)) = 0 Then

        str = "select psys_sttn from pesys"
        Rs1.Source = str
        Rs1.Open , Gcn
        
        If Trim(UCase(Rs1!psys_sttn)) = "HOUGSC" Or Trim(UCase(Rs1!psys_sttn)) = "HOUEMEC" Or Trim(UCase(Rs1!psys_sttn)) = "IMSCT1" Then
        
            str = "INSERT INTO UD_COMMON select com_npecode, '!sakha01', com_tablname, com_tonpecode  from ud_common where com_tonpecode = 'sAkha' and com_mail='!yuzhn02' order by com_npecode"
            Cmd.CommandText = str
            Cmd.CommandType = adCmdText
            Cmd.ActiveConnection = Gcn
            Cmd.Execute
        
        End If

End If

If Gcn.Errors.Count = 0 Then
    Gcn.CommitTrans
    InTransaction = False
Else
    Gcn.RollbackTrans
    InTransaction = False
    
End If

Set Rs = Nothing
Set Rs = New ADODB.Recordset

str = "select * from distribution where dis_npecode='sakha' and dis_id='ud'"
Rs.Source = str
Rs.Open , Gcn, 3, 3

Message = "The email id !SAKHA01 has been added successfully to the mailing list on the machine " & Machinename & ". V 1.013"
Message = Message & vbCrLf & "-------------------------------------------------" & vbCrLf
Message = Message & "Distribution table records for namespace SAKHALIN" & vbCrLf
Message = Message & vbCrLf & Rs.GetString(, , vbTab, vbCrLf)
Message = Message & vbCrLf & "-------------------------------------------------" & vbCrLf
Message = Message & "UD_COMMON : Records to be sent to SAKHALIN" & vbCrLf
Message = Message & "Namespace      Emails                                   CountofRecords  " & vbCrLf
Message = Message & "--------------------------------------------------------------------------------" & vbCrLf

Set Rs = Nothing
Set Rs = New ADODB.Recordset

str = "select com_npecode, com_mail, count(*) 'Count of Records' From ud_common where com_tonpecode = 'sakha' group by com_npecode, com_mail order by com_npecode"
Rs.Source = str
Rs.Open , Gcn, 3, 3

If Rs.RecordCount > 0 Then

    Message = Message & Rs.GetString(, , vbTab, vbCrLf)

Else

    Message = Message & " No Records found. "

End If

subject = "email id !SAKHA01 on " & Machinename & "."
ReDim emailAddresses(3)
emailAddresses(0) = "MUZAMMIL@IMS-SYS.COM"
emailAddresses(1) = "JCGONZALEZ@IMS-SYS.COM"
emailAddresses(2) = "ddegrazia@IMS-SYS.COM"
emailAddresses(3) = "FARMSTRONG@IMS-SYS.COM"

Call sendEmailOnly(Message, subject, emailAddresses(), Attachment)

MsgBox "The email id !SAKHA01 has been added successfully to the mailing list.", vbInformation, "Ims"

Gcn.Close

Label1.Visible = False

Unload Me

Exit Sub

handler:

    If InTransaction = True Then Gcn.RollbackTrans
    Message = "Errors occurred while adding email id !SAKHA01 to the mailing list. Error desc :" & Err.Description
    subject = "Email id !SAKHA01 "
    Call sendEmailOnly(Message, subject, emailAddresses(), Attachment)
    MsgBox "Errors occurred while trying to add the email id !SAKHA01 to the distribution list. Error Description :" & Err.Description, vbCritical, "Ims"
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
