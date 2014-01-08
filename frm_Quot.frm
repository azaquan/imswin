VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frm_Quot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select from a pending quotation?"
   ClientHeight    =   3180
   ClientLeft      =   30
   ClientTop       =   3330
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSDBQuot 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6640
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      FieldSeparator  =   ";"
      Col.Count       =   4
      HeadFont3D      =   4
      DefColWidth     =   5292
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   8388608
      BackColorEven   =   16771818
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   4
      Columns(0).Width=   2328
      Columns(0).Caption=   "Quotation #"
      Columns(0).Name =   "Quotation #"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1217
      Columns(1).Caption=   "LI #"
      Columns(1).Name =   "LI #"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1720
      Columns(2).Caption=   "Quantity"
      Columns(2).Name =   "Quantity"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   5715
      Columns(3).Caption=   "Description"
      Columns(3).Name =   "Description"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11730
      _ExtentY        =   5212
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox VisM1 
      Height          =   480
      Left            =   195
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   4170
      Width           =   1200
   End
End
Attribute VB_Name = "frm_Quot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim li_x As Integer
    
    'Added by Juan (9/14/2000) for Multilingual
    Call translator.Translate_Forms("frm_Quot")
    '------------------------------------------
    
'    Me.BackColor = frm_Color.txt_WBackground.BackColor
'    For li_x = 0 To (Controls.Count - 1)
'        If Not (TypeOf Controls(li_x) Is Toc) Then Call gsb_fade_to_black(Controls(li_x))
'    Next li_x

    frm_Quot.Caption = frm_Quot.Caption + " - " + frm_Quot.Tag
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Hide
    If open_forms <= 5 Then ShowNavigator
End Sub

Private Sub SSDBQuot_DblClick()
'    Dim ls_record As String
'    Dim ls_temp As String
'    Dim li_increment As Integer
'    Dim ls_Cost As String
'    Dim ls_QuotNum As String
'    Dim li_LineItem As Integer
'
'    ls_QuotNum = SSDBQuot.Columns(0).Text
'    li_LineItem = SSDBQuot.Columns(1).Value
'
'    VisM1.P4 = li_LineItem
'    VisM1.P3 = DoubleQuote() & "MARC" & DoubleQuote()
'    VisM1.P2 = "^PE" & gs_UT & "PO"
'    VisM1.P1 = DoubleQuote() & ls_QuotNum & DoubleQuote()
'    VisM1.P0 = "N/A"
'    VisM1.code = "d ^loadglob3"
'    VisM1.ExecFlag = 1
'
'    ls_record = VisM1.P0
'
'    frm_Purchase.txt_Quotfake.Text = ls_QuotNum
'    frm_Purchase.txt_Quotation2.Text = "TEST"
'    frm_Purchase.txt_LineItem2.Text = li_LineItem
'    ls_temp = piece1(ls_record, "~")
'
'    If Len(ls_temp) < 5 Then
'        frm_Purchase.txt_Requested = Mid$("00000", 1, 5 - Len(ls_temp)) & ls_temp
'    Else:
'        frm_Purchase.txt_Requested = ls_temp
'    End If
'
'    frm_Purchase.txt_Descript.Text = piece1(ls_record, "~")
'    Call piece1(ls_record, "~")
'    ls_Cost = piece1(ls_record, "~")
'
'    If Len(ls_Cost) < 9 Then
'        frm_Purchase.txt_Price = Mid$("000000.00", 1, 9 - Len(ls_Cost)) & ls_Cost
'    Else:
'        frm_Purchase.txt_Price = ls_Cost
'    End If
'
'    ls_Cost = piece1(ls_record, "~")
'    frm_Purchase.txt_Total = ls_Cost
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    frm_Purchase.txt_Commodity = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'
'    If gfn_cbo_finditem(frm_Purchase.cbo_Unit, ls_temp) > -1 Then frm_Purchase.cbo_Unit.Text = ls_temp
'
'    frm_Purchase.txt_Requisition = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    frm_Purchase.txt_SerialNum = piece1(ls_record, "~")
'    frm_Purchase.cbo_PartNum = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~") '     24
'    VisM1.P1 = gs_Months(gs_Language)
'    VisM1.P2 = ls_temp
'    VisM1.code = "s P0=$$^eurodate(P1,P2)"
'    VisM1.ExecFlag = 1
'
'    If VisM1.P0 <> "N/A" Then
'        ls_temp = VisM1.P0
'        If Len(ls_temp) = 10 Then 'year 2000
'            frm_Purchase.txt_RequDate2.Text = Mid$(ls_temp, 1, 6) & Mid$(ls_temp, 9, 2)
'        Else:
'            frm_Purchase.txt_RequDate2.Text = ls_temp
'        End If
'    End If
'
'    VisM1.P1 = gs_Months(gs_Language)
'    VisM1.P2 = frm_Purchase.txt_RequDate2.Text
'    VisM1.code = "s P0=$$^eurodate(P1,P2)"
'    VisM1.ExecFlag = 1
'
'    If VisM1.P0 <> "N/A" Then frm_Purchase.txt_RequDate2 = VisM1.P0
'    frm_Purchase.txt_Released.Text = piece1(ls_record, "~")
'    VisM1.P1 = gs_Months(gs_Language)
'    VisM1.P2 = frm_Purchase.txt_Released.Text
'    VisM1.code = "s P0=$$^eurodate(P1,P2)"
'    VisM1.ExecFlag = 1
'
'    If VisM1.P0 <> "N/A" Then frm_Purchase.txt_Released = VisM1.P0
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    frm_Purchase.txt_StatItem = piece1(ls_record, "~")
'    frm_Purchase.txt_StatDelivery = piece1(ls_record, "~")
'    frm_Purchase.txt_StatShipping = piece1(ls_record, "~")
'    frm_Purchase.txt_StatInventory = piece1(ls_record, "~")
'    frm_Purchase.txt_Currency2 = frm_Purchase.LoadDescript(piece1(ls_record, "~"), "^WH" & gs_UT & "CUR")
'    frm_Purchase.txt_AFE = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    If gfn_cbo_finditem(frm_Purchase.cbo_Custom, ls_temp) > -1 Then frm_Purchase.cbo_Custom.Text = ls_temp
'    ls_temp = piece1(ls_record, "~")
'    ls_temp = piece1(ls_record, "~")
'    'frm_Purchase.txt_Quotation2 = piece1(ls_record, "~")
'
'    'Set the values in PO Line items for the line
'    'item selected
'    Unload Me
'    frm_Purchase.txt_SerialNum.SetFocus
End Sub

