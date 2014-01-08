VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStockHeader As StockHeader

Private mManufacturer As Manufacturer

Private mOpenConnection As Boolean

Private mNamespace As String * 5

Private mConnection As ADODB.Connection

Private mLooKups As LookUps

Public Sub Configure(NameSpace As String, Cn As ADODB.Connection)

mNamespace = NameSpace

Set mConnection = Cn

End Sub


Public Property Get Manufacturer() As Manufacturer
  
  If mManufacturer Is Nothing Then
    
    Set mManufacturer = New Manufacturer
  
    mManufacturer.NameSpace = mNamespace
    mManufacturer.Connection = mConnection
    
  End If
    
  Set Manufacturer = mManufacturer
  
End Property

Public Property Get StockHeader() As StockHeader
  
   If mStockHeader Is Nothing Then
   
        Set mStockHeader = New StockHeader
   
        mStockHeader.NameSpace = mNamespace
        mStockHeader.Connection = mConnection
            
    End If
        
    Set StockHeader = mStockHeader
End Property

Private Property Get Connection() As ADODB.Connection
  
  If mConnection Is Nothing Then
     
     If mOpenConnection = True Then
     'We can add code to Open a new connection HERE.
     End If
     
     
  End If
     
     Set Connection = mConnection
       
End Property

Public Property Get OpenConnection() As Boolean
OpenConnection = mOpenConnection
End Property

Public Property Let OpenConnection(ByVal vNewValue As Boolean)
mOpenConnection = OpenConnection
End Property

Public Property Get LookUps() As LookUps

If mLooKups Is Nothing Then

    Set mLooKups = New LookUps
    mLooKups.NameSpace = mNamespace
    mLooKups.Connection = mConnection
    
 End If

    Set LookUps = mLooKups

End Property


Public Function Save() As String

Dim Errdesc As String

mConnection.Errors.Clear

mConnection.BeginTrans

If Not mStockHeader Is Nothing Then Errdesc = mStockHeader.Update

If Not mManufacturer Is Nothing And Len(Trim(Errdesc)) = 0 Then Errdesc = mManufacturer.Update
 
If Len(Trim(Errdesc)) > 0 Or mConnection.Errors.Count > 0 Then

    mConnection.RollbackTrans
    
Else

    mConnection.CommitTrans
    
End If

Set mManufacturer = Nothing

Save = Errdesc

End Function
