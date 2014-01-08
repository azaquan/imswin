VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~2.OCX"
Begin VB.Form FrmRequisition 
   Caption         =   "Requisition Managment"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   360
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   86573057
      CurrentDate     =   37316
   End
   Begin LRNavigators.LROleDBNavBar LROleDBNavBar1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   6480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      NewEnabled      =   -1  'True
      DeleteVisible   =   -1  'True
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDBDDStockNumber 
      Height          =   855
      Left            =   5280
      TabIndex        =   4
      Top             =   6480
      Width           =   3735
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "StockNumber"
      Columns(0).Name =   "StockNumber"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6535
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6588
      _ExtentY        =   1508
      _StockProps     =   77
      DataFieldToDisplay=   "Column 0"
   End
   Begin TabDlg.SSTab SSTabRequisitions 
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Requisition"
      TabPicture(0)   =   "FrmRequisition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSDDCompany"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSDDLocation"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSDDBuyer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSoleDbDetails"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSGridSelection"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MonthView2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Recepients"
      TabPicture(1)   =   "FrmRequisition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   360
         TabIndex        =   9
         Top             =   6120
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   86573057
         CurrentDate     =   37316
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSGridSelection 
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   8655
         ScrollBars      =   1
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   7
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   7
         Columns(0).Width=   3200
         Columns(0).Caption=   "Company"
         Columns(0).Name =   "Company"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   " Location"
         Columns(1).Name =   " Location"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Stock/Folio # Y/N"
         Columns(2).Name =   "Stock/Folio # Y/N"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Buyer"
         Columns(3).Name =   "Buyer"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1799
         Columns(4).Caption=   "Days Open"
         Columns(4).Name =   "Days Open"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "From"
         Columns(5).Name =   "From"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Caption=   "To"
         Columns(6).Name =   "To"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         _ExtentX        =   15266
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Selection Criteria"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSoleDbDetails 
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   8655
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   8
         RowHeight       =   423
         Columns.Count   =   8
         Columns(0).Width=   3200
         Columns(0).Caption=   "Requisition #"
         Columns(0).Name =   "Requisition #"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Date Created"
         Columns(1).Name =   "Date Created"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Line Item"
         Columns(2).Name =   "Line Item"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "Date Approved"
         Columns(3).Name =   "Date Approved"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Originator"
         Columns(4).Name =   "Originator"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "Day Open"
         Columns(5).Name =   "Day Open"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Caption=   "Buyer"
         Columns(6).Name =   "Buyer"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Caption=   "POs Included"
         Columns(7).Name =   "POs Included"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         _ExtentX        =   15266
         _ExtentY        =   5953
         _StockProps     =   79
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDBuyer 
         Height          =   855
         Left            =   5520
         TabIndex        =   5
         Top             =   5760
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Code"
         Columns(0).Name =   "Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Buyer Name"
         Columns(1).Name =   "Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDLocation 
         Height          =   855
         Left            =   5160
         TabIndex        =   6
         Top             =   5880
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         GroupHeaders    =   0   'False
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Location Code"
         Columns(0).Name =   "Location Code"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         DataFieldToDisplay=   "Column 1"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBDropDown SSDDCompany 
         Height          =   855
         Left            =   5640
         TabIndex        =   7
         Top             =   5640
         Width           =   3615
         DataFieldList   =   "Column 1"
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "CompanyCode"
         Columns(0).Name =   "CompanyCode"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4260
         Columns(1).Caption=   "Description"
         Columns(1).Name =   "Description"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   1508
         _StockProps     =   77
         DataFieldToDisplay=   "Column 1"
      End
   End
End
Attribute VB_Name = "FrmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SelectionCodes
    
    CompanyCode As String
    LocationCode As String
    Stocknumber As String
    Buyer As String
    FromDate As String
    ToDate As String
    OpenFor As Integer

End Type

Dim GselectionCode As SelectionCodes
Private Sub Label4_Click()

End Sub

Private Sub Form_Load()

deIms.cnIms.Open

deIms.NameSpace = "Angol"

Call PopulateBuyers
Call PopulateCompany
Call PopulateStockNumber

SSGridSelection.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9)

End Sub

Private Sub SSDBDDStockNumber_Click()
 GselectionCode.Stocknumber = Trim(SSDBDDStockNumber.Columns(0).text)
End Sub

Private Sub SSDDBuyer_Click()
 
 GselectionCode.Buyer = Trim(SSDDBuyer.Columns(0).text)
 
End Sub

Private Sub SSDDCompany_Click()
GselectionCode.CompanyCode = Trim(SSDDCompany.Columns(0).text)
End Sub

Private Sub SSDDLocation_Click()
GselectionCode.LocationCode = Trim(SSDDLocation.Columns(0).text)
End Sub

Private Sub SSGridSelection_Click()

Call LostFocusOnDatesColumns(1)

If SSGridSelection.Col = 1 Then

 Call PopulateLocation(Trim(GselectionCode.CompanyCode))

ElseIf SSGridSelection.Col = 5 Then
    
   Call SetFocusOnDatesColumns(5)
    
ElseIf SSGridSelection.Col = 6 Then
    
   Call SetFocusOnDatesColumns(6)

End If

End Sub

Private Sub SSGridSelection_InitColumnProps()

SSGridSelection.Columns(0).DropDownHwnd = SSDDCompany.HWND

SSGridSelection.Columns(1).DropDownHwnd = SSDDLocation.HWND

SSGridSelection.Columns(2).DropDownHwnd = SSDBDDStockNumber.HWND

SSGridSelection.Columns(3).DropDownHwnd = SSDDBuyer.HWND

SSGridSelection.Columns(4).DropDownHwnd = MonthView1.HWND

SSGridSelection.Columns(5).DropDownHwnd = MonthView1.HWND

SSGridSelection.Columns(6).DropDownHwnd = MonthView2.HWND

End Sub

Public Function GetDataForTheSelection(Query As String) As ADODB.Recordset

Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset

Rs.Source = Query

Rs.ActiveConnection = deIms.cnIms

Rs.Open , , adOpenKeyset, adLockOptimistic



End Function

Public Function PopulateBuyers()

Dim Rsbuyer As New ADODB.Recordset

Rsbuyer.Source = "select buy_username,usr_username , buy_npecode from buyer,xuserprofile where buy_username = usr_userid and buy_npecode = usr_npecode and usr_npecode ='" & deIms.NameSpace & "'"

Rsbuyer.ActiveConnection = deIms.cnIms

Rsbuyer.Open

SSDDBuyer.RemoveAll

  Do While Not Rsbuyer.EOF
    
    SSDDBuyer.AddItem Rsbuyer("buy_username") & Chr(9) & Rsbuyer("usr_username")

    Rsbuyer.MoveNext

   Loop

 Rsbuyer.Close

 Set Rsbuyer = Nothing


End Function

Public Function PopulateCompany()

Dim rsCOMPANY As New ADODB.Recordset

rsCOMPANY.Source = "select com_compcode,  com_name  from company where  com_npecode  ='" & deIms.NameSpace & "'"

rsCOMPANY.ActiveConnection = deIms.cnIms

rsCOMPANY.Open

SSDDCompany.RemoveAll

  Do While Not rsCOMPANY.EOF
    
   SSDDCompany.AddItem rsCOMPANY("com_compcode") & Chr(9) & rsCOMPANY("com_name")

    rsCOMPANY.MoveNext

   Loop

 rsCOMPANY.Close

 Set rsCOMPANY = Nothing


End Function

Public Function PopulateLocation(CompanyCode As String)

Dim RsLocation As New ADODB.Recordset

RsLocation.Source = "select loc_locacode , loc_name    from location where loc_npecode='" & deIms.NameSpace & "' and loc_compcode  ='" & CompanyCode & "'"

RsLocation.ActiveConnection = deIms.cnIms

RsLocation.Open

SSDDLocation.RemoveAll

  Do While Not RsLocation.EOF
    
    SSDDLocation.AddItem RsLocation("loc_locacode") & Chr(9) & RsLocation("loc_name")

    RsLocation.MoveNext

   Loop

 RsLocation.Close

 Set RsLocation = Nothing
 
End Function

Public Function PopulateStockNumber()

Dim RsLocation As New ADODB.Recordset

RsLocation.Source = "select stk_stcknumb , stk_desc from stockmaster where stk_npecode='" & deIms.NameSpace & "'"

RsLocation.ActiveConnection = deIms.cnIms

RsLocation.Open

SSDBDDStockNumber.RemoveAll

  Do While Not RsLocation.EOF
    
     SSDBDDStockNumber.AddItem RsLocation("stk_stcknumb") & Chr(9) & RsLocation("stk_desc")

    RsLocation.MoveNext

   Loop

 RsLocation.Close

 Set RsLocation = Nothing

End Function


Private Sub SSGridSelection_KeyPress(KeyAscii As Integer)


If SSGridSelection.Col = 5 Then
    
    SetFocusOnDatesColumns (5)
    
ElseIf SSGridSelection.Col = 6 Then
    
    SetFocusOnDatesColumns (6)

End If
    
    
End Sub

Private Sub SSGridSelection_LostFocus()

MonthView1.Visible = False

MonthView2.Visible = False

End Sub


Public Sub SetFocusOnDatesColumns(Column As Integer)

If Column = 5 Then
   
   
   MonthView1.Top = 1400 'SSGridSelection.Columns(5).Top
   MonthView1.Left = SSGridSelection.Columns(5).Left + 250

   MonthView1.Visible = True
   MonthView2.Visible = False
   
ElseIf Column = 6 Then

   MonthView2.Top = 1200 'SSGridSelection.Columns(6).Top
   MonthView2.Left = SSGridSelection.Columns(6).Left + 10
 
   MonthView1.Visible = False
   MonthView2.Visible = True

End If


End Sub

Public Sub LostFocusOnDatesColumns(Column As Integer)

''If Column = 5 Then
''
''   MonthView1.Visible = False
''   MonthView2.Visible = False
''
''ElseIf Column = 6 Then
''
''   MonthView1.Visible = False
''   MonthView2.Visible = False
''
''End If


   MonthView1.Visible = False
   MonthView2.Visible = False

End Sub

