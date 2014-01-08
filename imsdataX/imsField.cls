VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum imsFieldType
    ftDate
    ftNumeric
    ftString
End Enum

Dim FFormat As String
Dim FValue As Variant
Dim FLenght As Integer
Dim FReadOnly As Boolean
Dim fFieldType As imsFieldType
Dim FOriginalValue As Variant

Public Property Get Value() As Variant

    If Len(FFormat) Then
        Value = VBA.Format(FValue, FFormat)
    Else
        Value = FValue
    End If
End Property

Public Property Let Value(ByVal vNewValue As Variant)
On Error GoTo Handler

    If FLenght <> 0 Then
    
        If fFieldType = ftString Then _
            If Len(vNewValue) > FLenght Then _
                Err.Raise 1000, , "Value is too large"
    End If
    
    If fFieldType = ftDate Then
    
        If Not IsDate(Trim(vNewValue)) Then _
            Err.Raise 1000, , "Not a valid date"
    
    ElseIf fFieldType = ftNumeric Then
    
        If Not IsNumeric(vNewValue) Then _
            Err.Raise 1000, , "Not a valid number"
    End If
    
    FValue = vNewValue

Handler:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Get OriginalValue() As Variant
    OriginalValue = FOriginalValue
End Property

Public Property Let OriginalValue(ByVal vNewValue As Variant)
    FOriginalValue = vNewValue
End Property

Public Property Get Length() As Integer
     Length = FLenght
End Property

Public Property Let Length(ByVal vNewValue As Integer)
    FLenght = vNewValue
End Property


Public Property Get FieldType() As imsFieldType
    FieldType = fFieldType
End Property

Public Property Let FieldType(ByVal vNewValue As imsFieldType)
    If vNewValue < ftDate Or vNewValue > ftString Then
        Err.Raise 1000, "Invalid Value"
    Else
        fFieldType = vNewValue
    End If
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = FReadOnly
End Property

Public Property Let ReadOnly(ByVal vNewValue As Boolean)
    FReadOnly = vNewValue
End Property

Public Property Get Format() As String
    Format = FFormat
End Property

Public Property Let Format(ByVal vNewValue As String)
    FFormat = vNewValue
End Property
