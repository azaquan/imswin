VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#7.0#0"; "LRNAVIGATORS.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   5145
   Begin LRNavigators.NavBar NavBar1 
      Height          =   435
      Left            =   360
      TabIndex        =   2
      Top             =   4140
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   767
      ButtonHeight    =   329.953
      ButtonWidth     =   345.26
      MouseIcon       =   "Grid.frx":0000
      DeleteVisible   =   -1  'True
      EditVisible     =   -1  'True
      CloseToolTipText=   ""
      PrintToolTipText=   ""
      EmailToolTipText=   ""
      NewToolTipText  =   ""
      SaveToolTipText =   ""
      CancelToolTipText=   ""
      NextToolTipText =   ""
      LastToolTipText =   ""
      FirstToolTipText=   ""
      PreviousToolTipText=   ""
      DeleteToolTipText=   ""
      EditToolTipText =   ""
      EmailEnabled    =   -1  'True
      DeleteEnabled   =   -1  'True
      EditEnabled     =   -1  'True
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBGGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4875
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ";"
      stylesets.count =   2
      stylesets(0).Name=   "RowFont"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "Grid.frx":001C
      stylesets(0).AlignmentText=   0
      stylesets(1).Name=   "ColHeader"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "Grid.frx":0038
      stylesets(1).AlignmentText=   1
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1508
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).DataField=   "ctry_code"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HeadStyleSet=   "ColHeader"
      Columns(0).StyleSet=   "RowFont"
      Columns(1).Width=   6562
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "ctry_name"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).HeadStyleSet=   "ColHeader"
      Columns(1).StyleSet=   "RowFont"
      Columns(2).Width=   5292
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NameSpace"
      Columns(2).Name =   "NameSpace"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   8599
      _ExtentY        =   5318
      _StockProps     =   79
      DataMember      =   "COUNTRY"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get sLabel() As String
    sLabel = Label1
End Property

Public Property Let sLabel(ByVal vNewValue As String)
    Label1 = vNewValue
    Caption = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
If open_forms <= 5 Then frmNavigator.Visible = True
End Sub

Private Sub NavBar1_OnCancelClick()
    SSDBGGrid.CancelUpdate
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
    Set frmGrid = Nothing
End Sub

Private Sub NavBar1_OnFirstClick()
    SSDBGGrid.MoveFirst
End Sub

Private Sub NavBar1_OnLastClick()
    SSDBGGrid.MoveLast
End Sub

Private Sub NavBar1_OnNewClick()
    SSDBGGrid.AddNew
End Sub

Private Sub NavBar1_OnNextClick()
    SSDBGGrid.MoveNext
End Sub

Private Sub NavBar1_OnPreviousClick()
    SSDBGGrid.MovePrevious
End Sub

Private Sub NavBar1_OnSaveClick()
    SSDBGGrid.Update
End Sub

Public Property Get Description() As SSDataWidgets_B_OLEDB.Column
    Set Description = SSDBGGrid.Columns("Description")
End Property


Public Property Get Code() As SSDataWidgets_B_OLEDB.Column
    Set Code = SSDBGGrid.Columns("Code")
End Property

Public Property Get DataMember() As String
    DataMember = SSDBGGrid.DataMember
End Property

Public Property Let DataMember(ByVal vNewValue As String)
Dim ds As Object
    
    Set ds = SSDBGGrid.DataSource
    Set SSDBGGrid.DataSource = Nothing
    
    SSDBGGrid.DataMember = vNewValue
    Set SSDBGGrid.DataSource = ds
End Property

Public Property Get DataSource() As Variant
    Set DataSource = SSDBGGrid.DataSource
End Property

Public Property Set DataSource(ByVal vNewValue As Object)
      
    If IsObject(vNewValue) Then
        Set SSDBGGrid.DataSource = vNewValue
    Else
        Err.Raise 999, "Grid form", "invalid data source property"
    End If
    
    
End Property
