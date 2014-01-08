VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblProgress 
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape shpIncrement 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   360
      Index           =   1
      Left            =   130
      Top             =   2175
      Width           =   840
   End
   Begin VB.Label lblProgress 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Shape shpIncrement 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   360
      Index           =   0
      Left            =   130
      Top             =   1215
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpBackground 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   390
      Index           =   1
      Left            =   120
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Shape shpBackground 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   390
      Index           =   0
      Left            =   120
      Top             =   1200
      Width           =   5655
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaxWidth As Long
Dim Percent As Integer
Dim PercentIncrement As Long

'load form set form size

Private Sub Form_Load()

    MaxWidth = ((shpIncrement(0).Left - shpBackground(0).Left) * 2)
    
    MaxWidth = shpBackground(0).Width - MaxWidth
    PercentIncrement = MaxWidth / 100
    
    shpIncrement(0).Width = 0
    shpIncrement(1).Width = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If open_forms <= 5 Then ShowNavigator
End Sub

'set form size

Public Function IncrementProgress(Optional Percentage As Integer = 1) As Integer
Dim i As Long


    With shpIncrement(0)
        Percent = (100 / (MaxWidth / .Width))
        
        Percent = Percent + Percentage
        If Percent > 100 Then Percent = 100
    
        IncrementProgress = Percent
        .Width = Percent * PercentIncrement
        .Refresh
    End With
    
End Function

'set form progess increase values

Public Function IncrementTotalProgress(Optional Percentage As Integer = 1) As Integer
Dim i As Long


    With shpIncrement(1)
        Percent = GePercent(1)
        
        Percent = Percent + Percentage
        
        If Percent > 100 Then
            Percent = 100
        ElseIf Percent < 0 Then
            Percent = 0
        End If
    
        If Percent > 90 Then Stop
        IncrementTotalProgress = Percent
        .Width = Percent * PercentIncrement
    End With
    
    Refresh
End Function

'set progress back ground values

Public Property Get ShowTotalProgress() As Boolean
    ShowTotalProgress = Height > shpBackground(1).Top + shpBackground(1).Height
End Property

'calculate progress values

Public Property Let ShowTotalProgress(ByVal vNewValue As Boolean)
    If vNewValue Then
        Height = shpBackground(1).Top + shpBackground(1).Height + 240
    Else
        Height = lblProgress(1).Top - 60
    End If
End Property

'set back ground

Public Property Get ShowProgress() As Boolean
    ShowProgress = Height > shpBackground(0).Top + shpBackground(0).Height
End Property

'calculate values

Public Property Let ShowProgress(ByVal vNewValue As Boolean)

    If Not ShowTotalProgress Then
    
        If vNewValue Then
            Height = shpBackground(0).Top + shpBackground(0).Height + 60
        Else
            Height = lblProgress(0).Top - 60
        End If
    End If
End Property

'calculate percent values

Private Function GePercent(Index As Integer) As Integer
        Percent = (100 / (MaxWidth / shpIncrement(Index).Width))
End Function
