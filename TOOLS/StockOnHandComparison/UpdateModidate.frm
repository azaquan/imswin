VERSION 5.00
Begin VB.Form batch 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update StoredProcedures"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "UpdateModidate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Compare IMS with IDEAS"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compare IDEAS with IMS"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   2415
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
      Caption         =   "Compared The Stock OnHand report from IDEAS with IMS"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GCn As ADODB.Connection
Dim NAMESPACE As String
Private Type IMSDATA

    Location As String
    StockNo As String
    ImsCond_code As String
    QOH As Double
    unitPrice As Double

End Type
Private Sub Command1_Click()

Dim dOESTRAnsactionExist As Boolean
Dim Cmd As ADODB.Command
Dim RsIdeas As ADODB.Recordset
Dim RsIms As ADODB.Recordset
Dim ErrLocation As Integer
Dim Query As String
Dim Query1 As String
Dim Pcount As Integer
Dim Acount As Integer
Dim Difference As Integer
Dim count As Integer
Dim GCnIDEAS As ADODB.Connection
Dim GCnIMS As ADODB.Connection
Dim I As Integer
Dim Ximsdata() As IMSDATA
Dim RecordsNotInInvoentory As Integer
Dim TotalRecords As Integer
Dim Recordsuccess As Integer

On Error GoTo handler

DoEvents
Screen.MousePointer = vbHourglass

Set GCnIDEAS = New ADODB.Connection
Set GCnIMS = New ADODB.Connection

GCnIDEAS.CommandTimeout = 1000
GCnIMS.CommandTimeout = 1000

'GCnIMS.Open "driver={SQL Server};server=MUZAMMIL-TP\MUZAMMILTP2000;uid=sa;pwd=;database=pecten"
GCnIMS.Open "driver={SQL Server};server=pecten001;uid=sa;pwd=0egxpx4;database=pecten"
GCnIDEAS.Open "driver={SQL Server};server=(local);uid=sa;pwd=0eGxPx4;database=PECTENSTOCKONHAND"


Set RsIdeas = New ADODB.Recordset
RsIdeas.Source = "SELECT lOCATION, Stockno, ImsCond_code, QOH, UnitPrice FROM stockonhand WHERE unitprice is not null ORDER BY Stockno"
RsIdeas.Open , GCnIDEAS, adOpenKeyset

Set RsIms = New ADODB.Recordset
RsIms.Source = "SELECT qs5_ware 'Location', qs5_stcknumb 'StockNo', qs5_cond 'ImsCond_code', sum(qs5_primqty) 'QOH',sap_value 'UnitPrice' from stockonhand   group by qs5_stcknumb,qs5_ware, qs5_cond,sap_value order by qs5_stcknumb"
RsIms.Open , GCnIMS, adOpenKeyset

ReDim Ximsdata(RsIms.RecordCount)

Do While Not RsIms.EOF

    Ximsdata(I).Location = Trim(RsIms("Location"))
    Ximsdata(I).StockNo = Trim(RsIms("StockNo"))
    Ximsdata(I).unitPrice = Trim(RsIms("unitPrice"))
    Ximsdata(I).QOH = Trim(RsIms("QOH"))
    Ximsdata(I).ImsCond_code = Trim(RsIms("ImsCond_code"))

    I = I + 1

    RsIms.MoveNext

Loop

Do While Not RsIdeas.EOF
    
    DoEvents
        
    RecordCount = 0
    
    For t = 0 To I
            
            If Ximsdata(t).Location = Trim(RsIdeas("location")) And Ximsdata(t).StockNo = Trim(RsIdeas("StockNo")) And Ximsdata(t).ImsCond_code = Trim(RsIdeas("ImsCond_code")) Then 'And Ximsdata(t).QOH = Trim(RsIdeas("QOH")) And Ximsdata(t).unitPrice = Trim(RsIdeas("UnitPrice")) Then
    
                       RecordCount = RecordCount + 1
                       Exit For
    
            End If
    
    Next t
    
    
    If RecordCount = 1 Then
        
           If Ximsdata(t).QOH = CDbl(Trim(RsIdeas("QOH"))) And Ximsdata(t).unitPrice = CDbl(Trim(RsIdeas("UnitPrice"))) Then
           
                    Call LogSuccess(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")))
                    Recordsuccess = Recordsuccess + 1
                    
                    
           Else
                    Call LogfailureDueToQorSAP(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Ximsdata(t).QOH & vbTab & Trim(RsIdeas("UnitPrice")) & vbTab & Ximsdata(t).unitPrice)
                    RecordsFailedDueToQorSAP = RecordsFailedDueToQorSAP + 1
           End If
           
    ElseIf RecordCount > 1 Then
           
           Call Logfailure("----------------------------------------------------------")
           
           Call Logfailure("Mulitple records found")
           
           Call Logfailure(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")))
           
''           Do While Not RsIms.EOF
''
''                    Call Logfailure(RsIms("location") & vbTab & RsIms("StockNo") & vbTab & RsIms("ImsCond_code") & vbTab & RsIms("QOH") & vbTab & RsIms("UnitPrice") & " IMS ")
''
''                    RsIms.MoveNext
''
''           Loop
''
    ElseIf RecordCount = 0 Then
            
            Call Logfailure(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")))
            RecordsNotInInvoentory = RecordsNotInInvoentory + 1
            
    End If

    
    
    
    
''''    If RsIms.RecordCount = 1 Then
''''
''''           If RsIms("QOH") = CDbl(Trim(RsIdeas("QOH"))) And RsIms("UnitPrice") = CDbl(Trim(RsIdeas("UnitPrice"))) Then
''''
''''                    Call LogSuccess(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")))
''''
''''           Else
''''                    Call Logfailure("----------------------------------------------------------")
''''                    Call Logfailure("Either Quanity or the SAP is different")
''''                    Call Logfailure(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")) & " IDEAS ")
''''                    Call Logfailure(RsIms("location") & vbTab & RsIms("StockNo") & vbTab & RsIms("ImsCond_code") & vbTab & RsIms("QOH") & vbTab & RsIms("UnitPrice") & " IDEAS ")
''''           End If
''''
''''    ElseIf RsIms.RecordCount > 1 Then
''''
''''           Call Logfailure("----------------------------------------------------------")
''''
''''           Call Logfailure("Mulitple records found")
''''
''''           Call Logfailure(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")) & " IDEAS ")
''''
''''           Do While Not RsIms.EOF
''''
''''                    Call Logfailure(RsIms("location") & vbTab & RsIms("StockNo") & vbTab & RsIms("ImsCond_code") & vbTab & RsIms("QOH") & vbTab & RsIms("UnitPrice") & " IMS ")
''''
''''                    RsIms.MoveNext
''''
''''           Loop
''''
''''    ElseIf RsIms.RecordCount = 0 Then
''''
''''            Call Logfailure("----------------------------------------------------------")
''''
''''            Call Logfailure("No Associated Records in IMS")
''''
''''            Call Logfailure(Trim(RsIdeas("location")) & vbTab & Trim(RsIdeas("StockNo")) & vbTab & Trim(RsIdeas("ImsCond_code")) & vbTab & Trim(RsIdeas("QOH")) & vbTab & Trim(RsIdeas("UnitPrice")) & " IDEAS ")
''''
''''    End If

TotalRecords = TotalRecords + 1

RsIdeas.MoveNext

Loop



MsgBox "Ran Successfully.Out of " & TotalRecords & "  " & Recordsuccess & "  Exist in the invenotory, " & RecordsNotInInvoentory & " do not exist and " & RecordsFailedDueToQorSAP & " have different quantities or SAP."

Unload Me

Exit Sub

handler:

        
    MsgBox "Errors occurred. Error Description :" & Err.Description, vbCritical, "Ims"
    Err.Clear
    

Unload Me

End Sub

Public Function Logfailure(MessageToLog As String)
  
Dim sa As New Scripting.FileSystemObject

Dim t As Scripting.TextStream

FileName = App.Path & "\IDEAS-IMS-STOCKONHAND-FAILURE.txt"

If sa.FileExists(FileName) = False Then

    sa.CreateTextFile FileName, True
    
End If

Set t = sa.OpenTextFile(FileName, ForAppending)

t.WriteLine MessageToLog

t.Close

Set t = Nothing

End Function

Public Function LogSuccess(MessageToLog As String)
  
Dim sa As New Scripting.FileSystemObject

Dim t As Scripting.TextStream

FileName = App.Path & "\IDEAS-IMS-STOCKONHAND-SUCCESS.txt"

If sa.FileExists(FileName) = False Then

    sa.CreateTextFile FileName, True
    
End If

Set t = sa.OpenTextFile(FileName, ForAppending)

t.WriteLine MessageToLog

t.Close

Set t = Nothing

End Function

Public Function LogfailureDueToQorSAP(MessageToLog As String)
  
Dim sa As New Scripting.FileSystemObject

Dim t As Scripting.TextStream

FileName = App.Path & "\IDEAS-IMS-STOCKONHAND-FAILUREDUETO-QUANITY-OR-SAP.txt"

If sa.FileExists(FileName) = False Then

    sa.CreateTextFile FileName, True
    
End If

Set t = sa.OpenTextFile(FileName, ForAppending)

t.WriteLine MessageToLog

t.Close

Set t = Nothing

End Function

Private Sub Command2_Click()

Dim dOESTRAnsactionExist As Boolean
Dim Cmd As ADODB.Command
Dim RsIdeas As ADODB.Recordset
Dim RsIms As ADODB.Recordset
Dim ErrLocation As Integer
Dim Query As String
Dim Query1 As String
Dim Pcount As Integer
Dim Acount As Integer
Dim Difference As Integer
Dim count As Integer
Dim GCnIDEAS As ADODB.Connection
Dim GCnIMS As ADODB.Connection
Dim I As Integer
Dim Ximsdata() As IMSDATA
Dim RecordsNotInInvoentory As Integer
Dim TotalRecords As Integer
Dim Recordsuccess As Integer
Dim XiDEASdata() As IMSDATA
On Error GoTo handler

DoEvents
Screen.MousePointer = vbHourglass

Set GCnIDEAS = New ADODB.Connection
Set GCnIMS = New ADODB.Connection

GCnIDEAS.CommandTimeout = 1000
GCnIMS.CommandTimeout = 1000

'GCnIMS.Open "driver={SQL Server};server=MUZAMMIL-TP\MUZAMMILTP2000;uid=sa;pwd=;database=pecten"
GCnIMS.Open "driver={SQL Server};server=pecten001;uid=sa;pwd=0egxpx4;database=pecten"
GCnIDEAS.Open "driver={SQL Server};server=(local);uid=sa;pwd=0eGxPx4;database=PECTENSTOCKONHAND"


Set RsIdeas = New ADODB.Recordset
RsIdeas.Source = "SELECT lOCATION, Stockno, ImsCond_code, QOH, UnitPrice FROM stockonhand WHERE unitprice is not null ORDER BY Stockno"
RsIdeas.Open , GCnIDEAS, adOpenKeyset

Set RsIms = New ADODB.Recordset
RsIms.Source = "SELECT qs5_ware 'Location', qs5_stcknumb 'StockNo', qs5_cond 'ImsCond_code', sum(qs5_primqty) 'QOH',sap_value 'UnitPrice' from stockonhand  where qs5_ware in ('CHM' ,'D96','PRD','SUR','M&t','DRL') group by qs5_stcknumb,qs5_ware, qs5_cond,sap_value order by qs5_stcknumb"
RsIms.Open , GCnIMS, adOpenKeyset

ReDim Ximsdata(RsIms.RecordCount)

Do While Not RsIms.EOF

    Ximsdata(I).Location = Trim(RsIms("Location"))
    Ximsdata(I).StockNo = Trim(RsIms("StockNo"))
    Ximsdata(I).unitPrice = Trim(RsIms("unitPrice"))
    Ximsdata(I).QOH = Trim(RsIms("QOH"))
    Ximsdata(I).ImsCond_code = Trim(RsIms("ImsCond_code"))

    I = I + 1

    RsIms.MoveNext

Loop
''
''For t = 0 To I
''
'' If Ximsdata(t).Location = "SUR" And Ximsdata(t).StockNo = "4440468" And Ximsdata(t).ImsCond_code = "01" Then
''    Stop
'' ElseIf Ximsdata(t).Location = "PRD" And Ximsdata(t).StockNo = "3376229" And Ximsdata(t).ImsCond_code = "01" Then
''    Stop
'' ElseIf Ximsdata(t).Location = "SUR" And Ximsdata(t).StockNo = "4440448" And Ximsdata(t).ImsCond_code = "01" Then
''    Stop
''End If
''
''Next t



ReDim Ximsdata(RsIms.RecordCount)

RsIdeas.Filter = ""
RsIdeas.MoveFirst
ReDim XiDEASdata(RsIdeas.RecordCount)
Do While Not RsIdeas.EOF

    XiDEASdata(j).Location = Trim(RsIdeas("Location"))
    XiDEASdata(j).StockNo = Trim(RsIdeas("StockNo"))
    XiDEASdata(j).unitPrice = Trim(RsIdeas("unitPrice"))
    XiDEASdata(j).QOH = Trim(RsIdeas("QOH"))
    XiDEASdata(j).ImsCond_code = Trim(RsIdeas("ImsCond_code"))

    j = j + 1

    RsIdeas.MoveNext

Loop



For t = 0 To I
    
    DoEvents
        
    RecordCount = 0
    
    
    'RsIdeas.Filter = ""
    'RsIdeas.Filter = "lOCATION ='" & Ximsdata(t).Location & "' and  StockNo = '" & Ximsdata(t).StockNo & "' and ImsCond_code = '" & Ximsdata(t).ImsCond_code & "'"
    
   ' If RsIdeas.RecordCount > 0 Then RecordCount = 1
    
    For k = 0 To j
            
            If Ximsdata(t).Location = Trim(XiDEASdata(k).Location) And Ximsdata(t).StockNo = Trim(XiDEASdata(k).StockNo) And Ximsdata(t).ImsCond_code = Trim(XiDEASdata(k).ImsCond_code) Then
    
                       RecordCount = RecordCount + 1
                       Exit For
    
            End If
    
    Next k
    
    If RecordCount = 1 Then
        
           If Ximsdata(t).QOH = CDbl(Trim(XiDEASdata(k).QOH)) And Ximsdata(t).unitPrice = CDbl(Trim(XiDEASdata(k).unitPrice)) Then
           
                    Call LogSuccess(Trim(XiDEASdata(k).Location) & vbTab & Trim(XiDEASdata(k).StockNo) & vbTab & Trim(XiDEASdata(k).ImsCond_code) & vbTab & Trim(XiDEASdata(k).QOH) & vbTab & Trim(XiDEASdata(k).unitPrice))
                    Recordsuccess = Recordsuccess + 1
                    
                    
           Else
                    Call LogfailureDueToQorSAP(Trim(XiDEASdata(k).Location) & vbTab & Trim(XiDEASdata(k).StockNo) & vbTab & Trim(XiDEASdata(k).ImsCond_code) & vbTab & Trim(XiDEASdata(k).QOH) & vbTab & Ximsdata(t).QOH & vbTab & Trim(XiDEASdata(k).unitPrice) & vbTab & Ximsdata(t).unitPrice)
                    RecordsFailedDueToQorSAP = RecordsFailedDueToQorSAP + 1
           End If
           
    ElseIf RecordCount > 1 Then
           
           Call Logfailure("----------------------------------------------------------")
           
           Call Logfailure("Mulitple records found")
           
           Call Logfailure(Trim(XiDEASdata(k).Location) & vbTab & Trim(XiDEASdata(k).StockNo) & vbTab & Trim(XiDEASdata(k).ImsCond_code) & vbTab & Trim(XiDEASdata(k).QOH) & vbTab & Trim(XiDEASdata(k).unitPrice))
           

    ElseIf RecordCount = 0 Then
            
            Call Logfailure(Trim(Ximsdata(t).Location) & vbTab & Trim(Ximsdata(t).StockNo) & vbTab & Trim(Ximsdata(t).ImsCond_code) & vbTab & Trim(Ximsdata(t).QOH) & vbTab & Trim(Ximsdata(t).unitPrice))
            RecordsNotInInvoentory = RecordsNotInInvoentory + 1
            
    End If

TotalRecords = TotalRecords + 1


Next t


MsgBox "Ran Successfully.Out of " & TotalRecords & "  " & Recordsuccess & "  Exist in the invenotory, " & RecordsNotInInvoentory & " do not exist and " & RecordsFailedDueToQorSAP & " have different quantities or SAP."

Unload Me

Exit Sub

handler:

        
    MsgBox "Errors occurred. Error Description :" & Err.Description, vbCritical, "Ims"
    Err.Clear
    

Unload Me

End Sub
