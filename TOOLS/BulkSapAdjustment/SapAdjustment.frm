VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   370
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update SAP"
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Recorddata
    
    namespace As String
    CURRENTUSER As String
    COMPANYCODE As String
    REMARKS As String
    LOCATION As String
    NEWSAP As Double
    STOCKNO As String
    CONDITION As String
    
End Type


Private Sub Command1_Click()

Dim SAPRecord() As Recorddata
Dim Sa As Scripting.FileSystemObject
Dim T As Scripting.TextStream
Dim line As String
Dim namespace As String
Dim CURRENTUSER As String
Dim COMPANYCODE As String
Dim REMARKS As String
Dim LOCATION As String
Dim NEWSAP As Double
Dim STOCKNO As String
Dim CONDITION As String
Dim i As Integer

On Error GoTo ERRHNAND

de.cnims.Open
Set Sa = New Scripting.FileSystemObject
Set T = Sa.OpenTextFile(Text1.Text)


 Do While Not T.AtEndOfStream
 
    ReDim Preserve SAPRecord(i)
    
     line = T.ReadLine
     
    If Len(Trim(line)) > 0 Then
     
        SAPRecord(i).namespace = StripLine(line, 1, vbTab, "Before")
        SAPRecord(i).COMPANYCODE = StripLine(line, 2, vbTab, "Before")
        SAPRecord(i).LOCATION = StripLine(line, 3, vbTab, "Before")
        SAPRecord(i).CONDITION = StripLine(line, 4, vbTab, "BEFORE")
        SAPRecord(i).NEWSAP = CDbl(StripLine(line, 5, vbTab, "Before"))
        SAPRecord(i).STOCKNO = StripLine(line, 5, vbTab, "afteR")
        SAPRecord(i).REMARKS = " Adjustment made on request of PECTEN to synchronize IDEAS and IMS on DEC' 02"
    
    End If
    
    i = i + 1
    
Loop

T.Close

Set T = Nothing

Set Sa = Nothing

LogMessage ("Processed Started." & Now)
LogMessage ("-------------------------------------------------")
Call LogMessage(UBound(SAPRecord) & " STOCKNUMBERS TO BE PROCESS ")

For i = 0 To UBound(SAPRecord)

    Call SapAdjustment(SAPRecord(i).namespace, "IMSUSA", SAPRecord(i).COMPANYCODE, SAPRecord(i).REMARKS, SAPRecord(i).LOCATION, SAPRecord(i).NEWSAP, SAPRecord(i).STOCKNO, SAPRecord(i).CONDITION)

Next i

de.cnims.Close

LogMessage ("Processed completed Successfully.")
MsgBox ("Processed completed Successfully.")

Exit Sub

ERRHNAND:

LogMessage ("Unexpected Error Occurred. Process Stopped." & Err.Description)


Err.Clear

de.cnims.Close

MsgBox ("Unexpected Error Occurred. Process Stopped." & Err.Description)

End Sub
Private Function SapAdjustment(namespace As String, CURRENTUSER As String, COMPANYCODE As String, REMARKS As String, LOCATION As String, NEWSAP As Double, STOCKNO As String, CONDITION As String) As Boolean
Dim db As Double
Dim Cmd As New ADODB.Command
Dim Params As ADODB.Parameters
Dim CmdTrannumb As New ADODB.Command
Dim SItrannumb As String
Dim SEtrannumb As String
On Error GoTo ERRHAND
    
    SapAdjustment = False
    
    de.cnims.BeginTrans
    
'    Set CmdTrannumb = New ADODB.Command
'
'    With CmdTrannumb
'
'        .CommandType = adCmdStoredProc
'        .CommandText = "GET_INVTNUMB"
'        .Parameters.Append .CreateParameter("@namespace", adVarChar, adParamInput, 15, namespace)
'        .Parameters.Append .CreateParameter("@TRANSSERL", adVarChar, adParamOutput, 15, SItrannumb)
'        .Execute
'         SItrannumb = .Parameters("@TRANSSERL").Value
'
'    End With
'
'        SItrannumb = "SI-" & SItrannumb
'
'    Set CmdTrannumb = Nothing
'    Set CmdTrannumb = New ADODB.Command
'
'    With CmdTrannumb
'
'        .CommandType = adCmdStoredProc
'        .CommandText = "GET_INVTNUMB"
'        .Parameters.Append .CreateParameter("@namespace", adVarChar, adParamInput, 15, namespace)
'        .Parameters.Append .CreateParameter("@TRANSSERL", adVarChar, adParamOutput, 15, SEtrannumb)
'        .Execute
'         SEtrannumb = .Parameters("@TRANSSERL").Value
'
'    End With
'
'        SEtrannumb = "SE-" & SEtrannumb
'
'    Set CmdTrannumb = Nothing
    
    
    Set Params = Cmd.Parameters
    Cmd.CommandText = "SAPADJUSTMENT"
    Cmd.CommandType = adCmdStoredProc
    Set Cmd.ActiveConnection = de.cnims
    
        Params.Append Cmd.CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, namespace)
        Params.Append Cmd.CreateParameter("@USER", adVarChar, adParamInput, 20, CURRENTUSER)
        Params.Append Cmd.CreateParameter("@COMPCODE", adVarChar, adParamInput, 10, COMPANYCODE)
        Params.Append Cmd.CreateParameter("@REMARKS", adVarChar, adParamInput, 7000, REMARKS)
        Params.Append Cmd.CreateParameter("@LOCATION", adVarChar, adParamInput, 10, LOCATION)
        Params.Append Cmd.CreateParameter("@NEWSAP", adDouble, adParamInput, 20, NEWSAP)
        Params.Append Cmd.CreateParameter("@STOCKNUMBER", adVarChar, adParamInput, 20, STOCKNO)
        Params.Append Cmd.CreateParameter("@CONDITION", adVarChar, adParamInput, 2, CONDITION)
        Params.Append Cmd.CreateParameter("@SETRANSNUMB", adVarChar, adParamOutput, 15, SEtrannumb)
        Params.Append Cmd.CreateParameter("@SITRANSNUMB", adVarChar, adParamOutput, 15, SItrannumb)
    
    
    Call Cmd.Execute
    SItrannumb = Trim$(Params("@SITRANSNUMB") & "")
    SEtrannumb = Trim$(Params("@SETRANSNUMB") & "")
    
    Set Cmd = Nothing
    Set Params = Nothing
    
    LogMessage ("TRANSACTIONNO:" & SItrannumb & "," & SEtrannumb & " ---> namespace :" & namespace & ", COMPANYCODE :" & COMPANYCODE & ", LOCATION :" & LOCATION & ", NEWSAP :" & NEWSAP & ", STOCKNO :" & STOCKNO & ", CONDITION :" & CONDITION)
    
    de.cnims.CommitTrans
    
    SapAdjustment = True
    
Exit Function
    
ERRHAND:
    
    de.cnims.RollbackTrans
    
    Call LogMessage("ERROR" & Err.Description & " ---> Namespace :" & namespace & ",CURRENTUSER :" & CURRENTUSER & ",COMPANYCODE :" & COMPANYCODE & ", LOCATION :" & LOCATION & ", NEWSAP :" & NEWSAP & ", STOCKNO :" & STOCKNO & ", CONDITION :" & CONDITION)
    
    Err.Clear
    
End Function
Public Function LogMessage(MessageToLog As String)
  
Dim Sa As New Scripting.FileSystemObject

Dim T As Scripting.TextStream
Dim LOGFILEPATH As String

LOGFILEPATH = App.Path & " SAPAdjustmentLog.txt"

If Sa.FileExists(LOGFILEPATH) = False Then

    Sa.CreateTextFile LOGFILEPATH, True
    
End If

Set T = Sa.OpenTextFile(LOGFILEPATH, ForAppending)

T.WriteLine MessageToLog

T.Close

Set T = Nothing

End Function

Public Function StripLine(line As String, FieldDelimiterNumber As Integer, FileDelimiter As String, Before_After As String) As String

Dim LineCopy As String
    
Dim locationOfTAb As Integer

Dim LocationOfSecondTab As Integer

Dim count As Integer

Before_After = UCase(Before_After)

LineCopy = line

count = 1

Do While Not InStr(line, """" & vbTab & """") = 0

    locationOfTAb = InStr(LineCopy, """" & vbTab & """")
    
    If count < FieldDelimiterNumber Then
        
        LineCopy = Mid(LineCopy, locationOfTAb + 2, Len(LineCopy))
        
    ElseIf count = FieldDelimiterNumber Then
            
        If Before_After = "BEFORE" Then
            
            StripLine = Mid(LineCopy, 2, locationOfTAb - 2)
            
        ElseIf Before_After = "AFTER" Then
        
            LocationOfSecondTab = InStr(locationOfTAb + 1, LineCopy, """" & vbTab & """")
        
            If LocationOfSecondTab = 0 Then LocationOfSecondTab = InStrRev(LineCopy, """", -1)
        
            StripLine = Mid(LineCopy, locationOfTAb + 3, LocationOfSecondTab - locationOfTAb - 3)
        
        End If
          
     ElseIf count > FieldDelimiterNumber Then
     
     Exit Function
          
    End If
    
    count = count + 1
     
Loop
    
End Function

Private Sub Command2_Click()
cd1.Filter = "Text Files (*.txt)|*.txt"
cd1.ShowOpen

    Text1.Text = cd1.FileName

End Sub
