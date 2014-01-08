VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)

Dim FNameSpace As Collection
Private fUserName As String
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Public Enum UserType
    utRoot
    utAdmin
    utOwner
    utPublic
End Enum
'local variable(s) to hold property value(s)
Private fUserID As String
Private fCreatedBy As String
Private fCreatedOn As Date
Private fExpireDate As Date
Private fPreviousPasswordToKeep As Integer
Private fUserStatus As String

Public Property Let UserStatus(ByVal vData As String)
    fUserStatus = vData
End Property


Public Property Get UserStatus() As String
    UserStatus = fUserStatus
End Property

Public Property Let PreviousPasswordToKeep(ByVal vData As Integer)
    fPreviousPasswordToKeep = vData
End Property


Public Property Get PreviousPasswordToKeep() As Integer
    PreviousPasswordToKeep = fPreviousPasswordToKeep
End Property

Public Property Let ExpireDate(ByVal vData As Date)
    fExpireDate = vData
End Property

Public Property Get ExpireDate() As Date
    ExpireDate = fExpireDate
End Property

Public Property Let CreatedOn(ByVal vData As Date)
    fCreatedOn = vData
End Property


Public Property Get CreatedOn() As Date
    CreatedOn = fCreatedOn
End Property

Public Property Let CreatedBy(ByVal vData As String)
    fCreatedBy = vData
End Property


Public Property Get CreatedBy() As String
    CreatedBy = fCreatedBy
End Property

Public Property Let UserID(ByVal vData As String)
    fUserID = vData
End Property

Public Property Get UserID() As String
    UserID = fUserID
End Property

Public Sub Move(ByVal NumRecords As Integer, Optional Start As Integer)
    Call rs.Move(NumRecords)
End Sub

Public Sub MovePrevious()
    rs.MovePrevious
End Sub

Public Sub MoveLast()
    rs.MoveLast
End Sub

Public Sub MoveNext()
    rs.MoveNext
End Sub

Public Sub MoveFirst()
    rs.MoveFirst
End Sub

Public Property Let UserName(ByVal vData As String)
    fUserName = vData
End Property

Public Property Get UserName() As String
    UserName = fUserName
End Property

Public Function NameSpace() As Collection
    Set NameSpace = FNameSpace
End Function

Public Function RecordCount() As Integer

End Function
