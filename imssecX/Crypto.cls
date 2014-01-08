VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsCryptoClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'encrypt text function

Public Function Encrypttext(Text As String, Key As String) As String
Dim i As Long
Dim textChar As String * 1
Dim keyChar As String * 1
Dim encryptedChar As Integer

    If Text = "" Then
        Encrypttext = ""
        Exit Function
    End If
    If Key = "" Then
        Encrypttext = ""
        Exit Function
    End If
        
    Encrypttext = ""
    For i = 1 To Len(Text)
        textChar = Mid(Text, i, 1)
        keyChar = Mid(Key, (i Mod Len(Key)) + 1)
        encryptedChar = Asc(textChar) Xor Asc(keyChar)
        Encrypttext = Encrypttext & Chr(encryptedChar)
    Next
    'Encrypttext = 1
End Function

'decypt text function

Public Function Decypttext(Text As String, Key As String) As String
    Decypttext = Encrypttext(Text, Key)
End Function


