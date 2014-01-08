VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsSecMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum LoginSuccess
    lgsSuccess = 0
    lgsCancelled = -1
End Enum
    
Public Enum MenuForms
    mfUser
    mfLevel
    mfOption
    mfTemplate
End Enum
    
Dim fUser As String
Dim fPassword As String
Dim FNameSpace As String
Dim Rptneeded As Boolean
Dim LoginKey(1) As Integer
Public WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

Public FConnection As ADODB.Connection
Public Event OnError(Message As String)
Public Event InfoMessage(Message As String)

Public Property Let languageSELECTED(ByRef newLANGUAGE As String)
    TR_LANGUAGE = newLANGUAGE
End Property

Friend Sub RaiseErrorEvent(msg As String, Optional Id As Long = 0)


    If Id = 0 Then
        RaiseEvent OnError(msg)
    Else
        RaiseEvent InfoMessage(msg)
    End If
        
End Sub

Public Function Login(User As String, Optional namespacedescription) As LoginSuccess
On Error Resume Next
    
    CurrentUser = User
    Rptneeded = False
    Set frm = New frmLogon
    
    Set frm.sec = Me
    Login = lgsCancelled
    
    frm.Show vbModal
    
    Login = frm.Result
    
    If Login = lgsSuccess Then
    
        fUser = frm.User
    
        User = fUser
        FNameSpace = frm.NameSpace
        namespacedescription = frm.namespacedescription
    End If
    
    Unload frm
    Set frm = Nothing
End Function

Public Property Get languageSELECTED() As String
    languageSELECTED = TR_LANGUAGE
End Property
Public Function CanChangePassword(ByVal NameSpace As String, ByVal User As String, cn As ADODB.Connection) As Boolean
On Error Resume Next

    If ((cn Is Nothing) And (Not (FConnection Is Nothing))) Then _
        Set cn = FConnection
        
    CanChangePassword = _
        PASSWORDAGE(NameSpace, User, True, cn) <= 0
        
        
End Function

Public Sub ChangePassword()

    Rptneeded = False
    Set frm = New frmpass
    
     Load frm
    With frm
    
        .UserName = fUser
        .NameSpace = FNameSpace
        Set .Connection = FConnection
        .Show vbModal
        
    End With
    
    Unload frm
    Set frm = Nothing
End Sub

Public Property Get Connection() As ADODB.Connection
On Error GoTo Handled

    Set Connection = FConnection
    
Handled:
    If Err Then
        RaiseErrorEvent (Err.Description)
        Err.Clear
    End If
End Property

Public Property Set Connection(NewVAlue As ADODB.Connection)
On Error GoTo Handled

    Set FConnection = NewVAlue
    
Handled:
    If Err Then
        RaiseErrorEvent (Err.Description)
        Err.Clear
    End If
End Property


Public Property Get UserName() As String
    UserName = fUser
End Property

Public Property Let UserName(ByVal vNewValue As String)
    fUser = vNewValue
    CurrentUser = vNewValue
End Property

Public Property Get NameSpace() As String
    NameSpace = FNameSpace
End Property

Public Property Let NameSpace(ByVal vNewValue As String)
    FNameSpace = vNewValue
End Property

Public Sub AddUser(ByVal User As String)

    
    Set frm = New frmUserProfile
    
    'Load frm 'Hidden by Juan
    
    With frm
        Call .SetNameSpace(FNameSpace)
        Call .SetConnection(FConnection)
        CurrentUser = IIf(User = "", fUser, User)
    
        Call .SetUser(CurrentUser)
        If .OpenRecordset Then Call .Show(vbModal)

    End With
    
    Unload frm
    Set frm = Nothing
End Sub

Public Sub AssignTempOwnerPassWord(ByVal User As String, FirstTimeUser As Boolean)

    Rptneeded = False
    Set frm = New frmTempPass

    With frm
    
        .FirstTimeUser = FirstTimeUser
        .User = IIf(User = "", fUser, User)
        
        Set .cn = FConnection
        .NameSpace = FNameSpace
        
        'This means we are popping up the initial password form.
        If FirstTimeUser = True Then
        
         .Caption = " Assign initial password " & .Tag
         
        Else
        
          .Caption = " Assign temporary password " & .Tag
         
        End If
        
        .Show vbModal
    End With
    
    Unload frm
    Set frm = Nothing
End Sub

Private Sub Class_Initialize()
    Rptneeded = True
End Sub

Private Sub Class_Terminate()
On Error Resume Next

Dim wh As imsutilsx.imsWinhelp

    Set wh = New imsutilsx.imsWinhelp
    Set frm = Nothing
    Call wh.HelpQuit(0, App.HelpFile)
    
    Set wh = Nothing
    
    'Hided by Juan because it was an obsolete process for the connection to the server
'    If Rptneeded Then
'        Call frm.CrystalReport1.LogOffServer(LoginKey(0), True)
'        Call frm.CrystalReport1.LogOffServer(LoginKey(1), True)
'    End If
    
    Unload frmReport
    Set frmReport = Nothing
    
    If Err Then _
        Call LogErr("ImsSecMod::Class_Terminate", Err.Description, Err, True)
        
    If Err Then Err.Clear
End Sub

Private Sub frm_Load()
On Error Resume Next
    If Rptneeded Then LogonOnReport
End Sub

Private Sub frm_Unload(Cancel As Integer)
On Error Resume Next
    frm.Hide
    
End Sub

Public Sub ShowMenuOptions(mf As MenuForms, Optional ByVal sNameSpace As String, Optional cn As ADODB.Connection)
On Error Resume Next
    Select Case mf
        Case mfLevel
            Set frm = New frmMenuLevel
            
        Case mfOption
            Set frm = New frmMenuOption
             
        Case mfUser
            Set frm = New frmUserMenu
            
        Case mfTemplate
            Set frm = New frmMenuTemp
            
        Case Else
            Call MsgBox(LoadResString(1003))
            Exit Sub
    End Select
    
    sNameSpace = Trim$(sNameSpace)
    
    If sNameSpace = "" And FNameSpace = "" Then
        Err.Raise 1000, "ImsSec", LoadResString(1000)
        
    ElseIf sNameSpace <> "" Then
        Call frm.SetNameSpace(sNameSpace)
        
    ElseIf FNameSpace <> "" Then
        Call frm.SetNameSpace(FNameSpace)
        
    End If
    
        
    If Not imsutilsx.IsNothing(cn) Then
        Call frm.SetConnection(cn)
        
    ElseIf Not imsutilsx.IsNothing(FConnection) Then
        Set FConnection = cn
        Call frm.SetConnection(cn)
        
    Else
        Err.Raise 1001, "ImsSec", LoadResString(1001)

    End If
        
    On Error GoTo 0
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
    
    Call frm.Show(vbModal)
        
    Unload frm
    Set frm = Nothing
End Sub

Public Sub ShowBuyers(Optional ByVal sNameSpace As String, Optional cn As ADODB.Connection)
 
    Set frm = New frmBuyerRight
    sNameSpace = Trim$(sNameSpace)
    
    If sNameSpace = "" And FNameSpace = "" Then
        Err.Raise 1000, "ImsSec", LoadResString(1000)
        
    ElseIf sNameSpace <> "" Then
        Call frm.SetNameSpace(sNameSpace)
       
    ElseIf FNameSpace <> "" Then
        Call frm.SetNameSpace(FNameSpace)
        
    End If
    
        
    
    If Not imsutilsx.IsNothing(cn) Then
        Call frm.SetConnection(cn)
        
    ElseIf Not imsutilsx.IsNothing(FConnection) Then
        Set FConnection = cn
        Call frm.SetConnection(cn)
        
    Else
        Err.Raise 1001, "ImsSec", LoadResString(1001)

    End If
        
    On Error GoTo 0
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
    
    Call frm.Show(vbModal)
         
    Unload frm
    Set frm = Nothing
End Sub

Public Property Let ReportFilePath(ByVal vNewValue As String)
    ReportPath = imsutilsx.FixDir(vNewValue)
End Property


Private Sub LogonOnReport()
On Error Resume Next

    If Not Rptneeded Then Exit Sub
    
    'Load frmReport
    
'''    If Err Then Call RaiseErrorEvent(Err.Description): Err.Clear
'''    LoginKey(0) = frm.CrystalReport1.LogOnServer("pdssql.dll", "ims", "SAKHALIN", "sa", "2r2m9k3")
'''
'''
'''    If Err Then Call RaiseErrorEvent(Err.Description): Err.Clear
'''    LoginKey(1) = frm.CrystalReport1.LogOnServer("pdsodbc.dll", "imsO", "SAKHALIN", "sa", "2r2m9k3")
'''
'''    If Err Then Call LogErr("imssec::Initialize", Err.Description, Err)
End Sub

Public Property Get dsnNAME() As String
dsnNAME = m_dsnname

End Property

Public Property Let dsnNAME(ByRef vNewValue As String)
m_dsnname = vNewValue
End Property




