VERSION 5.00
Begin VB.Form fmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "fmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnMonitor 
      Caption         =   "Monitor"
      Visible         =   0   'False
      Begin VB.Menu opMonitor 
         Caption         =   "&About..."
         Index           =   0
      End
      Begin VB.Menu opMonitor 
         Caption         =   "E&xit monitor"
         Index           =   1
      End
      Begin VB.Menu opMonitor 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu opMonitor 
         Caption         =   "&Process to monitor"
         Index           =   3
      End
      Begin VB.Menu opMonitor 
         Caption         =   "&Stop Monitor"
         Index           =   4
      End
      Begin VB.Menu opMonitor 
         Caption         =   "&Configuration"
         Index           =   5
      End
   End
End
Attribute VB_Name = "fmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minutes As Integer
Dim minOc As Integer
Dim FIniInfo As fileinfo
Dim StatusFile As String


Private Sub Form_Load()

FIniInfo = GetIniInformation
StatusFile = App.Path & "\EmailSend"
Call ChecktheProcess
End Sub
Private Sub ChecktheProcess()
    
    Dim i As Long
    Dim j As Byte
    Dim aux As String
    Dim Servidor As String
    Dim direccion As String
    Dim chequear As Byte
    Dim message As String
    Dim sa As New Scripting.FileSystemObject
    Dim Attmail As New ImsUtils.imsmisc
    Dim Address() As String
    Dim attachments(0) As String
    Dim s As String
On Error GoTo Errors
    
    FIniInfo.Email = Trim(FIniInfo.Email)
    
    If Len(FIniInfo.Email) = 0 Then
        
        ReDim Address(0)
        
        Address(0) = "muzammil@ims-sys.com"
        
    Else
    
    Address = Split(FIniInfo.Email, ";")
    
    End If
    
    
    minutes = Val(FIniInfo.WaitTime)
    Email = FIniInfo.Email
    EXE = FIniInfo.ExeFullPath
    aux = EXE

    If aux = "" Then Call Attmail.SendAttMail("No process has been specified to run, the applciation will shut down.", "Start UpdateDB :Errors Occurred", Address(), attachments())

    If Not isActive(aux) Then
    
        Shell aux
    
        Call Attmail.SendAttMail("Update DB had crashed, had to start it up.", "Start UpdateDB :Restarted UpdateDB", Address(), attachments())
    
    End If
    
    minutes = 0
    
    If DatePart("h", FormatDateTime(Now, vbShortTime)) = 12 And sa.FileExists(StatusFile) = False Then
    
        Call Attmail.SendAttMail("Re Boot UpdateDb is up and Running.", "Start UpdateDB :Daily Email", Address(), attachments())
        
        sa.CreateTextFile StatusFile
    
    ElseIf DatePart("h", FormatDateTime(Now, vbShortTime)) <> 12 Then
    
        If sa.FileExists(StatusFile) Then sa.DeleteFile StatusFile
    
    End If
    
    Unload Me
    
    Exit Sub
    
Errors:
    
    Call Attmail.SendAttMail("Unhandled error occurred. Error Description: " & Err.Description, "Start UpdateDB :Errors Occurred", Address(), attachments())
    
    Err.Clear
    
    Unload Me
    
End Sub
