VERSION 5.00
Begin VB.Form frm_Load 
   BorderStyle     =   0  'None
   Caption         =   "Loading IMS for Windows®"
   ClientHeight    =   3390
   ClientLeft      =   1710
   ClientTop       =   1425
   ClientWidth     =   6705
   Icon            =   "Load.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Load.frx":000C
   ScaleHeight     =   3390
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Load 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4725
   End
End
Attribute VB_Name = "frm_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents sec As imsSecMod
Attribute sec.VB_VarHelpID = -1



''muzmamil June 24th 2008
Private Function Getconnectionstring() As String
On Error Resume Next

Dim Sa As New Scripting.FileSystemObject
Dim t  As TextStream
Dim strconn As String
Set t = Sa.OpenTextFile(App.Path & "\" & "imsudl.txt", ForReading, False)

strconn = t.ReadLine
'deIms.cnIms.connectionstring = Trim(strconn)
  'MsgBox (strconn)
Getconnectionstring = Trim(strconn)
    
End Function

'load MDI form if error cause show error message

Private Sub Form_Load()



On Error Resume Next

Dim str As String
    Show
    Call StayOnTop(HWND, True)
    
    If Err Then Call LogErr(Name & "::Form_Load", "error launching imsutils", Err)
    
    
    Err.Clear

    deIms.cnIms.connectionstring = Getconnectionstring()
    'deIms.cnIms.Open ("File name=" + App.Path + "\ims.udl") 'M
    'MsgBox (deIms.cnIms.connectionstring)
deIms.cnIms.Open
    
    If Err Then
        
        MsgBox Err.Description
        Call LogErr(Name & "::Form_Load", "error launching Data Env", Err) 'M
        MsgBox "A connection could not be established to the database. Please make sure the connection file exist and is valid." & Err.Description, vbCritical, "Ims"
        Unload Me
        Exit Sub
        
    End If
        
    If deIms.cnIms.State = 1 Then  'M
       ConnInfo = StripConnectionString(deIms.cnIms.connectionstring)
    End If
    
    
    
    Set sec = New imsSecMod
    If Err Then Call LogErr(Name & "::Form_Load", "error creating imssec", Err)
    'Modified By Muzammil
    'Passing a connection to the Security Dll
    Set sec.Connection = deIms.cnIms    'M
    
    'CurrentUser = GetUserName
    
    If sec.Login(CurrentUser, str) = lgsSuccess Then
        
        DoEvents
        Err.Clear
        
        deIms.NameSpace = sec.NameSpace
        
        Call GetNamespaceConfigurationValues(ConnInfo.EmailClient, ConnInfo.EmailOutFolder, ConnInfo.EmailParameterFolder)
        
        
        
        Set translator = New imsTranslator
        Set translator.thisREPO = MDI_IMS.CrystalReport1
        If sec.languageSELECTED = "US" Then
            translator.TR_LANGUAGE = "*"
        Else
            translator.TR_LANGUAGE = sec.languageSELECTED
        End If
        
        Language = sec.languageSELECTED 'M
        
        translator.changeLANGUAGE
        translator.mainREPORT = True
    
    
        
        Err.Clear
        Load MDI_IMS
        
        Hide
        MDI_IMS.Show
        MDI_IMS.StatusBar1.Panels(5).Text = str
    End If
    Set sec = Nothing
    Unload Me
    
End Sub

'call function stayontop

Private Sub sec_InfoMessage(Message As String)
On Error Resume Next

    If Message = "0" Then
        Call StayOnTop(HWND, False): Refresh
    Else
        lblInfo = Message
        lblInfo.Refresh
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub sec_OnError(Message As String)
    Call sec_InfoMessage(Message)
End Sub

Public Sub ShowMessage(Message As String)

End Sub


Public Function GetUserName() As String
Dim Sa As New Scripting.FileSystemObject
Dim t As Scripting.TextStream
Dim CompletePath As String
Dim CurrentUser As String
Dim I As Integer

On Error GoTo Errhandler

CompletePath = App.Path & "\ImsAutomaticLogin.Ims"

If Sa.FileExists(CompletePath) Then
          
   Set t = Sa.OpenTextFile(CompletePath, ForWriting, True)
   
   CurrentUser = Trim(t.ReadLine)
    
End If

GetUserName = CurrentUser

Exit Function
Errhandler:

MsgBox "Errors Occured while trying to retrive the username of the Last Logon. Error Description : " & Err.Description, vbCritical, "Ims"
Err.Clear

End Function
