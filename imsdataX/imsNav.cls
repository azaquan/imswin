VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38955A7A00EE"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=38955A7A0170
Public Enum TupdateMode
    '##ModelId=38955A7A0184
    umDelayed
    '##ModelId=38955A7A018E
    umImediate
End Enum

'##ModelId=38955A7A01AC
Public Property Get UpdateMode() As TupdateMode
End Property

'Cancel the changes made to the recordset
'##ModelId=38955A7A01B6
Public Sub CancelUpdate()
Attribute CancelUpdate.VB_Description = "Cancel the changes made to the recordset"
End Sub

'Updates the database with any changes made
'##ModelId=38955A7A01C0
Public Sub Update()
Attribute Update.VB_Description = "Updates the database with any changes made"
End Sub

'Add a new record
'##ModelId=38955A7A01D4
Public Sub AddNew()
Attribute AddNew.VB_Description = "Add a new record"
End Sub

'Move to the previous record
'##ModelId=38955A7A01DE
Public Sub MovePrevious()
Attribute MovePrevious.VB_Description = "Move to the previous record"
End Sub

'Moves to the next record
'##ModelId=38955A7A01E8
Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Moves to the next record"
End Sub

'Moves to the last record
'##ModelId=38955A7A01FC
Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Moves to the last record"
End Sub

'Move to the first record
'##ModelId=38955A7A0206
Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Move to the first record"
End Sub

'##ModelId=38955A7A0211
Public Function Validate() As Boolean
End Function

'##ModelId=38955A7A021B
Public Sub Requery()
End Sub

'##ModelId=38955A7A022F
Public Sub UpdateBatch()
End Sub

'##ModelId=38955A7A0239
Public Property Get Editting() As Boolean
End Property

'##ModelId=38955A7A024D
Public Property Get AbsolutePosition() As Long
End Property

'##ModelId=38955A7A0257
Public Property Get Eor() As Boolean
End Property

'##ModelId=38955A7A026B
Public Property Get Bor() As Boolean
End Property

'##ModelId=38955A7A0275
Public Function Find(Criteria As String) As Boolean
End Function
