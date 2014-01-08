VERSION 5.00
Begin VB.Form FrmTrasmission 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flag Transactions"
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
      Caption         =   "Flag the transmissions for update database."
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
Dim str1 As String
Dim Cmd As New ADODB.Command
On Error GoTo handler

Label1.Visible = True
DoEvents
Screen.MousePointer = vbHourglass


Set Gcn = New ADODB.Connection
Gcn.CursorLocation = adUseClient
Gcn.CommandTimeout = 0
Gcn.Open "driver={SQL Server};server=.;uid=sa;pwd=0eGxPx4;database=sakhalin"

Gcn.Errors.Clear

Gcn.BeginTrans


str1 = "'I-20332','I-20333','I-20334','RT-20335','AI-20336','AE-20337','I-20338','AE-20339','AI-20340','AE-20341','I-20342','RT-20343','I-20344','AI-20345','R-20346','RT-20347','AI-20348',"
str1 = str1 & "'SI-20349','SE-20350','AI-20351','AE-20352','AI-20353','AE-20354','AE-20355','RT-20356','R-20357','R-20358','AI-20359','AE-20360','RT-20361','AI-20362','RT-20363','RT-20364','RT-20365','I-20366','AI-20367','AE-20368','AI-20369','AE-20370','R-20371','AI-20372','AE-20373','RT-20374','RT-20375','RT-20376','I-20377','RT-20378','AE-20379','AE-20380','AE-20381','I-20382','AI-20383','RT-20384','RT-20385','AI-20386','AE-20387','IT-20388','IT-20389','IT-20390','RT-20391','R-20392','I-20393','I-20394','R-50000','AI-50001','AE-50002','AI-50003','AE-50004','AI-50005','AE-50006','AI-50007','SI-50008','SE-50009','R-50010','R-50011','SI-50012','SE-50013','SI-50014','SE-50015','SI-50016','SE-50017','I-50018','R-50019','I-50020','SI-50021','SE-50022','SI-50023','SE-50024','R-50025','SI-50026','SE-50027','RT-50028','TI-50029','SI-50030','SE-50031','SI-50032','SE-50033','SI-50034','SE-50035','SI-50036','SE-50037','AE-50038','SI-50039','SE-50040','SI-50041','SE-50042','SI-50043','SE-50044','SI-50045','SE-50046',"
str1 = str1 & "'SI-50047','SE-50048','SI-50049','SE-50050','SI-50051','SE-50052','AI-50053','RT-50054','I-50055','RT-50056','RT-50057','I-50058','I-50059','I-50060','I-50061','I-50062','AE-50063','R-50064','AI-50065','AE-50066','RT-50067','I-50068','I-50069','I-50070','I-50071','AI-50072','I-50073','I-50074','AE-50075','I-50076','AI-50077','AE-50078','I-50079','I-50080','RT-50081','RT-50082','I-50083','I-50084','I-50085','RT-50086','RT-50087','AI-50088','AE-50089','AI-50090','AE-50091','R-50092','AI-50093','AE-50094','AI-50095','AI-50096','AE-50097','R-50098','R-50099','AE-50100','AE-50101','AI-50102','AI-50103','AE-50104','R-50105','AI-50106','AI-50107','AE-50108','R-50109','AI-50110','AI-50111','AE-50112','AI-50113','AI-50114','AI-50115','AE-50116','AI-50117','AE-50118','RT-50119','AI-50120','AI-50121','AI-50122','AE-50123','SI-50124','SE-50125','SI-50126','SE-50127','SI-50128','SE-50129','AI-50130','AE-50131','AE-50132','RT-50133',"
str1 = str1 & "'I-50134','AI-50135','AE-50136','AE-50137','AI-50138','AE-50139','I-50140','I-50141','I-50142','I-50143','I-50144','I-50145','I-50146','I-50147','RT-50148','I-50149','RT-50150','I-50151','I-50152','I-50153','I-50154','R-50155','I-50156','I-50157','AI-50158','AE-50159','AI-50160','AE-50161','AI-50162','AI-50163','AE-50164','AI-50165','AE-50166','AI-50167','AE-50168','IT-50169','AI-50170','AE-50171','IT-50172','IT-50173','I-50174','I-50175','I-50176','AE-50177','IT-50178','AI-50179','AE-50180','AI-50181','AE-50182','R-50183','R-50184','AI-50185','AI-50186','R-50187','AI-50188','R-50189','AI-50190','AE-50191','AI-50192','R-65001','R-65002','AI-65003','AI-65004','AE-65005','R-65006','R-65007','AI-65008','AE-65009','R-65010','AI-65011','AE-65012','AE-65013','R-65014','R-65015','AI-65016','AE-65017','AE-65018','AE-65019','AE-65020','R-65021','AI-65022','R-65023','AI-65024','AE-65025','R-65026','AI-65027','AE-65028','R-65029','AI-65030','AI-65031','AE-65032','R-65033','AI-65034',"
str1 = str1 & "'AE-65035','AI-65036',"
str1 = str1 & "'AE-65037','AI-65038','AE-65039','I-65040','AI-65041','AI-65042','R-65043','R-65044','R-65045','IT-65046','AE-65047','IT-65048','R-65049','AI-65050','AE-65051','R-65052','AI-65053','AE-65054','I-65055','I-65056','AE-65057','I-65058','I-65059','I-65060','I-65061','I-65062','I-65063','I-65064','I-65065','R-65066','AI-65067','AE-65068','I-65069','AE-65070','AE-65071','IT-65072','RT-65073','AI-65074','AE-65075','I-65076','I-65077','I-65078','I-65079','R-65080','AE-65081','AE-65082','I-65083','I-65084','SI-65085','SE-65086','SI-65087','SE-65088','RT-65089','RT-65090','RR-65091','AI-65092','AE-65093','AI-65094','R-65095','AI-65096','AE-65097','I-65098','RT-65099','AI-65100','AE-65101','AI-65102','I-65103','AE-65104','AI-65105','AE-65106','AE-65107','I-65108',"
str1 = str1 & "'I-65109','I-65110','I-65111','SI-65112','SE-65113','SI-65114','SE-65115','R-65116','R-65117','R-65118','I-65729','RT-65730','RT-65731','IT-65732','RT-65733','IT-65734','I-65735','I-65736','I-65737'"


str = " update  invtreceipt set ir_tbs=1 where ir_trannumb in (" & str1 & ")"
str = str & " update  invtreceiptdetl set ird_tbs=1 where ird_trannumb in (" & str1 & ")"
str = str & " update  invtreceiptrem set irr_tbs=1 where irr_trannumb in (" & str1 & ")"
str = str & " update  invtissue set ii_tbs=1 where ii_trannumb in (" & str1 & ")"
str = str & " update  invtissuedetl set iid_tbs=1 where iid_trannumb in (" & str1 & ")"
str = str & " update  invtissuerem set iir_tbs=1 where iir_trannumb in (" & str1 & ")"

Cmd.CommandText = str
Cmd.CommandType = adCmdText
Cmd.ActiveConnection = Gcn
Cmd.CommandTimeout = 0
Cmd.Execute

If Gcn.Errors.Count = 0 Then
    Gcn.CommitTrans
Else
    Gcn.RollbackTrans
End If

MsgBox "The transactions have been flaged successfully.", vbInformation, "Ims"

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
