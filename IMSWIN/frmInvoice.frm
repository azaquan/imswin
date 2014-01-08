VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F8D97923-5EB1-11D3-BA04-0040F6348B67}#9.1#0"; "LRNAVI~1.OCX"
Object = "{27609682-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "ImsMailVBX.ocx"
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Supplier"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   12540
   Tag             =   "02050700"
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Header"
      TabPicture(0)   =   "frmInvoice.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Shape1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "POComboList"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "InvoiceComboList"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "remark"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPicker1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cell(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cell(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cell(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cell(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cell(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cell(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cell(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cell(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cell(8)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cell(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "supplierDATA"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Picture1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Line Item List"
      TabPicture(1)   =   "frmInvoice.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "invoiceLABEL"
      Tab(1).Control(1)=   "nomLabel(0)"
      Tab(1).Control(2)=   "nomLabel(1)"
      Tab(1).Control(3)=   "currencyLABEL"
      Tab(1).Control(4)=   "POtitles"
      Tab(1).Control(5)=   "POlist"
      Tab(1).Control(6)=   "TextLINE"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "nomPicture(0)"
      Tab(1).Control(9)=   "nomPicture(1)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Recipients"
      TabPicture(2)   =   "frmInvoice.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_Recipients"
      Tab(2).Control(1)=   "Imsmail1"
      Tab(2).Control(2)=   "RecipientList"
      Tab(2).Control(3)=   "cmd_Remove"
      Tab(2).Control(4)=   "cmd_Add"
      Tab(2).ControlCount=   5
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   4680
         ScaleHeight     =   1665
         ScaleWidth      =   2865
         TabIndex        =   44
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "SAVING INVOICE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.PictureBox nomPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -66600
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   41
         Top             =   680
         Width           =   255
      End
      Begin VB.PictureBox nomPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -68280
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   39
         Top             =   680
         Width           =   255
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Add"
         Height          =   288
         Left            =   -74400
         TabIndex        =   36
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Remove 
         Caption         =   "Remove"
         Height          =   288
         Left            =   -74400
         TabIndex        =   35
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Show Only Selection"
         Height          =   375
         Left            =   -65040
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TextLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -64320
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid supplierDATA 
         Height          =   2055
         Left            =   8040
         TabIndex        =   28
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   7
         FixedRows       =   0
         RowHeightMin    =   285
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         GridLinesFixed  =   0
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   9
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   8
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   7
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   6
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   5
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox cell 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   5880
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         CustomFormat    =   "MMMM/dd/yyyy"
         Format          =   64094211
         CurrentDate     =   36867
      End
      Begin VB.TextBox remark 
         Height          =   3675
         Left            =   240
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3360
         Width           =   11775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid InvoiceComboList 
         Height          =   975
         Left            =   1920
         TabIndex        =   27
         Top             =   1350
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid POComboList 
         Height          =   975
         Left            =   1920
         TabIndex        =   6
         Top             =   990
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16776960
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid POlist 
         Height          =   5720
         Left            =   -74760
         TabIndex        =   31
         Top             =   1440
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10081
         _Version        =   393216
         Cols            =   18
         RowHeightMin    =   285
         GridColorFixed  =   0
         HighLight       =   0
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   18
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid POtitles 
         Height          =   450
         Left            =   -74760
         TabIndex        =   32
         Top             =   1140
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   794
         _Version        =   393216
         Cols            =   5
         RowHeightMin    =   285
         GridColorFixed  =   0
         HighLight       =   0
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid RecipientList 
         Height          =   2535
         Left            =   -72840
         TabIndex        =   34
         Top             =   720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4471
         _Version        =   393216
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ImsMailVB.Imsmail Imsmail1 
         Height          =   3375
         Left            =   -74520
         TabIndex        =   37
         Top             =   3480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5953
      End
      Begin VB.Label currencyLABEL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71760
         TabIndex        =   43
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label nomLabel 
         Caption         =   "Already Invoiced"
         Height          =   375
         Index           =   1
         Left            =   -66240
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label nomLabel 
         Caption         =   "Purchase Unit"
         Height          =   375
         Index           =   0
         Left            =   -67920
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_Recipients 
         Caption         =   "Recipient(s)"
         Height          =   300
         Left            =   -74400
         TabIndex        =   38
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label invoiceLABEL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   720
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   240
         Top             =   720
         Width           =   15
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   26
         Top             =   2445
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Buyer"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   24
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Requested"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   22
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Issued"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   20
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Currency"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   18
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Created"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2445
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Invoice"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2085
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Created By"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1725
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Invoice"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction #"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   765
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   4455
      End
   End
   Begin LRNavigators.LROleDBNavBar NavBar1 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   7680
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   661
      AllowCustomize  =   0   'False
      CancelEnabled   =   0   'False
      EMailEnabled    =   0   'False
      EMailVisible    =   -1  'True
      FirstVisible    =   0   'False
      LastVisible     =   0   'False
      NewEnabled      =   -1  'True
      NextVisible     =   0   'False
      PreviousVisible =   0   'False
      PrintEnabled    =   0   'False
      SaveEnabled     =   0   'False
      Mode            =   3
      CommandType     =   8
      EditEnabled     =   0   'False
      EditVisible     =   0   'False
      CancelToolTipText=   "Undo the changes made to the current Screen"
   End
   Begin VB.Label lblStatu 
      Alignment       =   1  'Right Justify
      Caption         =   "Visualization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   7560
      Width           =   4335
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Form As FormMode
Dim readyFORsave As Boolean
Dim rs As ADODB.Recordset, rsReceptList As ADODB.Recordset
Dim colorsROW(12)
Dim SaveEnabled As Boolean
Dim forceNAV As Boolean
Dim moveUP As Boolean
Dim currentROW As Integer
Dim multiMARKED As Boolean
Dim selectionSTART As Integer
Sub alphaSEARCH(ByVal cellACTIVE As textBOX, ByVal gridACTIVE As MSHFlexGrid, column)
Dim i, ii As Integer
Dim word As String
Dim found As Boolean
    If cellACTIVE <> "" Then
        With gridACTIVE
            If Not .Visible Then .Visible = True
            If .Rows = 0 Then Exit Sub
            If IsNumeric(.Tag) Then
                .row = val(.Tag)
                .Col = column
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
            End If
            .Col = column
            .Tag = ""
            found = False
            For i = 0 To .Rows - 1
                word = Trim(UCase(.TextMatrix(i, column)))
                If Trim(UCase(cellACTIVE)) = Left(word, Len(cellACTIVE)) Then
                    .row = i
                    .CellBackColor = &H800000 'Blue
                    .CellForeColor = &HFFFFFF 'White
                    .Tag = .row
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                .row = 0
                .Tag = ""
            End If
            If IsNumeric(.Tag) Then .topROW = val(.Tag)
        End With
    End If
End Sub

Sub arrowKEYS(direction As String, Index As Integer)
Dim Grid As MSHFlexGrid
    With cell(Index)
        Select Case Index
            Case 0
                Set Grid = POComboList
            Case 1
                Set Grid = InvoiceComboList
                
        End Select
        
        Select Case Index
            Case 0, 1
                If IsNumeric(Grid.Tag) Then
                    Grid.row = val(Grid.Tag)
                    Grid.CellBackColor = &HFFFF00   'Cyan
                    Grid.CellForeColor = &H80000008 'Default Window Text
                End If
                Select Case direction
                Case "down"
                    If Grid.row < (Grid.Rows - 1) Then
                        If Grid.row = 0 And .text = "" Then
                            .text = Grid.text
                        Else
                            Grid.row = Grid.row + 1
                        End If
                    Else
                        Grid.row = Grid.Rows - 1
                    End If
                Case "up"
                    If Grid.row > 0 Then
                        Grid.row = Grid.row - 1
                    Else
                        Grid.row = 1
                    End If
            End Select
            If Not Grid.Visible Then
                Grid.Visible = True
            End If
            Grid.ZOrder
            Grid.topROW = Grid.row
            Grid.SetFocus
        End Select
    End With
End Sub

Sub BeforePrint()
    With MDI_IMS.CrystalReport1
        .Reset
        msg1 = translator.Trans("L00176")
        .WindowTitle = IIf(msg1 = "", "Invoice", msg1)
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        If cell(1) = "" Then
            .ReportFileName = FixDir(App.Path) + "CRreports\InvoiceGlobal.rpt"
            .ParameterFields(1) = "ponumb;" + cell(0) + ";TRUE"
            Call translator.Translate_Reports("invoiceGlobal.rpt")
        Else
            .ReportFileName = FixDir(App.Path) + "CRreports\Invoice.rpt"
            .ParameterFields(1) = "invnumb;" + cell(1) + ";TRUE"
            .ParameterFields(2) = "ponumb;" + cell(0) + ";TRUE"
            Call translator.Translate_Reports("invoice.rpt")
            Call translator.Translate_SubReports
        End If
    End With
End Sub

Sub begining()
Dim i
    With supplierDATA
        .ColWidth(0) = 900
        .ColWidth(1) = 3000
        .ColAlignmentFixed(0) = 6
        .ColAlignment(1) = 1
        .TextMatrix(0, 0) = "Supplier"
        .TextMatrix(1, 0) = "Address"
        .TextMatrix(3, 0) = "City"
        .TextMatrix(4, 0) = "State"
        .TextMatrix(5, 0) = "Country"
        .TextMatrix(6, 0) = "Zip"
    End With
End Sub


Sub checkNEXT(i, outFOR)
'    If POlist.TextMatrix(i, 0) <> "" Then
'        If IsNumeric(POlist.TextMatrix(i, 1)) Then
'            forceNAV = True
'            POlist.row = i
'            POlist.Col = 8
'            outFOR = True
'            Call POlist_Click
'        End If
'    End If
End Sub
Sub colorCOLS(Optional previous As Boolean)
Dim i As Integer
    With POlist
        .row = POlist.Rows - 1
        .Col = 3
        .CellBackColor = &HE0E0E0
        .Col = 7
        .CellBackColor = &HE0E0E0
        .Col = 11
        .CellBackColor = &HE0E0E0
        For i = 8 To 10
            .Col = i
            If val(.TextMatrix(.row, 17)) = 0 Then
                If previous Then
                    .CellBackColor = &HC0FFFF 'Very Light Yellow
                Else
                    .CellBackColor = &HFFFFC0 'Very Light cyan
                End If
            Else
                .CellBackColor = &HFFFFC0 'Very Light Green
            End If
        Next
    End With
End Sub

Sub differences(row As Integer)
Dim d1, d2 As Double
Dim s1, s2 As String
Dim Col As Integer
    s1 = POlist.TextMatrix(row, 6)
    s2 = POlist.TextMatrix(row, 10)
    
    Select Case s1
        Case Is = "", 0
            Exit Sub
            d1 = 0
        Case Else
            If IsNull(s1) Then
                d1 = 0
            Else
                d1 = CDbl(s1)
            End If
    End Select
    
    Select Case s2
        Case "", 0
            d2 = 0
        Case Else
            If IsNull(s2) Then
                d2 = 0
            Else
                d2 = CDbl(s2)
            End If
    End Select
    
    If IsNumeric(s1) And IsNumeric(s2) Then
        POlist.TextMatrix(row, 12) = FormatNumber((d2 - d1), 2)
        Col = POlist.Col
        POlist.Col = 12
        currentROW = POlist.row
        POlist.row = row
        If (d2 - d1) >= 0 Then
            POlist.CellForeColor = vbBlack
        Else
            POlist.CellForeColor = vbRed
        End If
        POlist.Col = Col
        POlist.row = currentROW
    End If
End Sub

Sub drawLINEcol(ByVal Grid As MSHFlexGrid, Col As Integer)
    With Grid
        .ColWidth(Col) = 50 'Line
        .Col = Col
        .CellBackColor = &H808080
    End With
End Sub

Sub fixPOtitles(diff As Integer)
Dim i, w As Integer
    If diff = 0 Then diff = 1
    With POtitles
        If diff <= 2 Then
            w = POlist.ColWidth(0)
        Else
            w = 0
        End If
        For i = diff To 13
            w = w + POlist.ColWidth(i)
            Select Case i
                Case 2
                    .ColWidth(0) = w
                    w = 0
                Case 6
                    .ColWidth(2) = w - IIf(w >= 50, 50, 0)
                    w = 0
                Case 13
                    .ColWidth(4) = w - IIf(w >= 50, 50, 0)
            End Select
        Next
        .TextMatrix(0, 0) = "General"
        .TextMatrix(0, 2) = "Transaction"
        .TextMatrix(0, 4) = "Supplier Invoice"
        If diff > 2 Then
            .ColWidth(0) = POlist.ColWidth(0)
            .TextMatrix(0, 0) = ""
        End If
        If diff > 6 Then
            .ColWidth(2) = POlist.ColWidth(0)
            .TextMatrix(0, 2) = ""
        End If
    End With
End Sub

Sub getCOLORSrow()
Dim i, currentCOL As Integer
    currentCOL = POlist.Col
    For i = 1 To 12
        POlist.Col = i
        colorsROW(i) = POlist.CellBackColor
    Next
    POlist.Col = currentCOL
End Sub

Sub getINVOICE(INVOICE As String)
On Error Resume Next
Dim dataINVOICE  As New ADODB.Recordset
Dim sql As String
        
    Screen.MousePointer = 11
    Call clearDOCUMENT
    
    'Header
    If Left(cell(0), 1) <> "(" And Right(cell(0), 1) <> ")" Then cell(0) = UCase(cell(0))
    If INVOICE = "*" Then
        cell(1) = ""
        sql = "SELECT * from PO_Header_for_Invoice WHERE NameSpace = '" + deIms.NameSpace + "' " _
        & "AND PO = '" + cell(0) + "'"
    Else
        If cell(1) = "" Then
            sql = "SELECT * from PO_Header_for_Invoice WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "'"
        Else
            If cell(0) = "" Or cell(0) = "(By Invoice)" Then
                sql = "SELECT * from Invoice_Header WHERE NameSpace = '" + deIms.NameSpace + "' " _
                    & "AND Invoice = '" + Trim(cell(1).text) + "'"
            Else
                sql = "SELECT * from Invoice_Header WHERE NameSpace = '" + deIms.NameSpace + "' " _
                    & "AND PO = '" + cell(0) + "' AND Invoice = '" + Trim(cell(1).text) + "'"
            End If
        End If
    End If
    Set dataINVOICE = New ADODB.Recordset
    dataINVOICE.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then Exit Sub
        
    With dataINVOICE
        If .RecordCount > 0 Then
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
            cell(0) = !po
            cell(2) = IIf(IsNull(!UserName), "", !UserName)
            cell(3) = IIf(IsNull(!InvoicedDate), "", !InvoicedDate)
            cell(4) = IIf(IsNull(!CreatedDate), "", !CreatedDate)
            cell(5) = IIf(IsNull(!Currency), "", !Currency)
            cell(6) = IIf(IsNull(!DateIssued), "", !DateIssued)
            cell(7) = IIf(IsNull(!DateRequested), "", !DateRequested)
            cell(8) = IIf(IsNull(!Buyer), "", !Buyer)
            cell(9) = IIf(IsNull(!BuyerPhone), "", !BuyerPhone)
            remark = IIf(IsNull(!remarks), "", !remarks)
                        
            supplierDATA.TextMatrix(0, 1) = IIf(IsNull(!Supplier), "", !Supplier)
            supplierDATA.TextMatrix(1, 1) = IIf(IsNull(!address1), "", !address1)
            supplierDATA.TextMatrix(2, 1) = IIf(IsNull(!address2), "", !address2)
            supplierDATA.TextMatrix(3, 1) = IIf(IsNull(!City), "", !City)
            supplierDATA.TextMatrix(4, 1) = IIf(IsNull(!State), "", !State)
            supplierDATA.TextMatrix(5, 1) = IIf(IsNull(!Country), "", !Country)
            supplierDATA.TextMatrix(6, 1) = IIf(IsNull(!Zip), "", !Zip)
            supplierDATA.TextMatrix(7, 1) = IIf(IsNull(!Telephone), "", !Telephone)
            
            'Details
            Err.Clear
            If INVOICE = "*" Then
                Call getLINEitems("*")
                cell(0).SelStart = 0
                cell(0).SelLength = Len(cell(0))
                cell(0).SetFocus
                cell(1) = ""
                POComboList.Visible = True
            Else
                Call getLINEitems(cell(1))
                cell(1).SelStart = 0
                cell(1).SelLength = Len(cell(1))
                cell(1).SetFocus
            End If
            NavBar1.NewEnabled = SaveEnabled
        Else
            NavBar1.PrintEnabled = False
            NavBar1.EMailEnabled = False
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00088")
            MsgBox IIf(msg1 = "", "Does not exist yet", msg1)
            cell(0) = ""
        End If
        Call getRECIPIENTSlist
    End With
    
    Screen.MousePointer = 0
End Sub

Sub getInvoiceComboList()
Dim sql As String
Dim dataLIST As ADODB.Recordset
    Err.Clear
    Set dataLIST = New ADODB.Recordset
    sql = "SELECT inv_invcnumb FROM INVOICE " _
        & "WHERE inv_npecode = '" + deIms.NameSpace + "'"
    If cell(0) <> "(By Invoice)" And cell(0) <> "" Then
        sql = sql + " AND inv_ponumb = '" + Trim(cell(0).text) + "' "
    End If
    sql = sql + " ORDER BY inv_invcnumb"
    dataLIST.Open sql, deIms.cnIms, adOpenForwardOnly
    
    With InvoiceComboList
        .Visible = False
        .ColWidth(0) = 1600
        .Clear
        .Rows = 0
        .ColAlignment(0) = 1
    End With
    If Err.number = 0 Then
        If dataLIST.RecordCount > 0 Then
            Do While Not dataLIST.EOF
                InvoiceComboList.AddItem " " + Trim(dataLIST!inv_invcnumb)
                dataLIST.MoveNext
            Loop
            InvoiceComboList.row = 0
            InvoiceComboList.RowHeightMin = 240
        End If
    End If
End Sub


Function isOPEN(po As String) As Boolean
Dim sql As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    isOPEN = False
    po = Trim(cell(0))
    sql = "SELECT po_ponumb, po_stas from PO WHERE po_npecode = '" + deIms.NameSpace + "' " _
        & "AND po_ponumb = '" + cell(0) + "'"
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then Exit Function
    If dataPO.RecordCount > 0 Then
        If dataPO!po_stas = "OP" Then
            isOPEN = True
        Else
            isOPEN = False
        End If
    Else
        isOPEN = False
    End If
    If Not isOPEN Then
        Screen.MousePointer = 0
        Select Case dataPO!po_stas
            Case "OH"
                MsgBox "This PO is not Approved"
            Case "CL"
                MsgBox "This PO is already Closed"
            Case "CA"
                MsgBox "This Po has been Canceled"
        End Select
        cell(0).SetFocus
    End If
End Function

Sub markROW(Optional multi As Boolean)
Dim nextROW, originalROW, purchaseUNIT As String
Dim i, itemX As Integer
Dim response
    With POlist
        If multi Then
            originalROW = .row
        Else
            originalROW = .MouseRow
        End If
        Select Case .TextMatrix(originalROW, 1)
            Case ""
                Exit Sub
            Case "§"
                nextROW = "UP"
                itemX = val(.TextMatrix(.row - 1, 1))
            Case Else
                If .row < .Rows - 1 Then
                    .row = .row + 1
                    If .TextMatrix(.row, 1) = "§" Then
                        nextROW = "DOWN"
                    Else
                        nextROW = "NO"
                    End If
                    .row = .row - 1
                Else
                    nextROW = "NO"
                End If
                itemX = val(.TextMatrix(.row, 1))
        End Select
        
        If val(.TextMatrix(.row, 17)) > 0 Then
            If .TextMatrix(.row, 0) = "" Then
                response = MsgBox("You have already invoiced this line item.  Please print a report.  Do you want to continue", vbYesNo)
                If response = vbNo Then
                    Exit Sub
                End If
            Else
                If .TextMatrix(.row, 1) <> "" Then
                    Call getLINEitems("*", itemX)
                End If
            End If
        End If
        .row = originalROW
        
        .Col = 0
        For i = 1 To 2
            .CellFontName = "Wingdings 3"
            .CellFontSize = 10
            If .text = "" Then
                .text = "Æ"
                .TextMatrix(.row, 9) = .TextMatrix(.row, 5)
                purchaseUNIT = Trim(.TextMatrix(.row, 15))
                If purchaseUNIT = "P" Or purchaseUNIT = "" Then
                    If i = 1 Then
                        .TextMatrix(.row, 10) = .TextMatrix(.row, 6)
                        .TextMatrix(.row, 12) = "00.0"
                    Else
                        .TextMatrix(.row, 8) = .TextMatrix(.row, 4)
                    End If
                Else
                    If i = 1 Then
                        .TextMatrix(.row, 8) = .TextMatrix(.row, 4)
                    Else
                        .TextMatrix(.row, 10) = .TextMatrix(.row, 6)
                        .TextMatrix(.row, 12) = "00.0"
                    End If
                End If
            Else
                .text = ""
                If val(.TextMatrix(.row, 17)) < 1 Then
                    .TextMatrix(.row, 8) = ""
                    .TextMatrix(.row, 10) = ""
                    .TextMatrix(.row, 12) = ""
                End If
            End If

            Select Case nextROW
                Case "UP"
                    .row = .row - 1
                Case "DOWN"
                    If .row < .Rows - 1 Then
                        .row = .row + 1
                    End If
                Case "NO"
                    Exit For
            End Select
        Next
        .row = originalROW
    End With
End Sub

Sub clearDOCUMENT()
Dim i As Integer
    readyFORsave = False
    For i = 2 To 9
        cell(i) = ""
        If i = 0 Or i = 1 Or i = 3 Then cell(i).BackColor = remark.BackColor
    Next
    For i = 0 To 6
        supplierDATA.TextMatrix(i, 1) = ""
    Next
    POComboList.Visible = False
    InvoiceComboList.Visible = False
    remark = ""
    nomPicture(0).Visible = False
    nomLabel(0).Visible = False
    Command1.Caption = "&Show Only Selection"
End Sub

Function controlOBJECT(controlNAME As String) As Control
Dim c As Control
    For Each c In Me.Controls
        If c.Name = controlNAME Then
            Exit For
        End If
        Set c = Nothing
    Next
    Set controlOBJECT = c
End Function

Sub datePICKER(controlNAME As String)
Dim h, i As Integer
Dim c As Control

    With DTPicker1
        .Tag = ""
        For Each c In Me.Controls
            If c.Name = controlNAME Then
                Exit For
            End If
            Set c = Nothing
        Next
        If c Is Nothing Then Exit Sub
        .Tag = controlNAME
    
        .Left = c.Left + c.ColWidth(0)
        .Height = c.RowHeight(i)
        If c.row = 0 Then
            .Top = c.Top
            .Height = .Height - 80
        Else
            h = 20
            For i = 0 To c.row - 1
                h = h + c.RowHeight(i)
            Next
            .Top = h + c.Top - 30
            .Height = .Height + 10
        End If
        .Visible = True
        .Value = IIf(IsDate(c.text), c.text, Now)
        .SetFocus
        Call DTPicker1_DropDown
    End With
End Sub

Sub getPOComboList()
On Error Resume Next
Dim sql As String
Dim datPO As New ADODB.Recordset

    Err.Clear
    With POComboList
        .Visible = False
        .ColWidth(0) = 1600
        .ColAlignment(0) = 1
    End With
    
    Set datPO = New ADODB.Recordset
        
    sql = "SELECT po_ponumb FROM PO WHERE po_npecode = '" + deIms.NameSpace + "' " _
        & "ORDER BY po_ponumb"
    
    POComboList.Rows = 0
    With datPO
        .Open sql, deIms.cnIms, adOpenForwardOnly
        If Err.number <> 0 Then Exit Sub
        If .RecordCount > 0 Then
            POComboList.AddItem "(By Invoice)"
            Do While Not .EOF
                POComboList.AddItem Trim(!PO_PONUMB)
                .MoveNext
            Loop
        End If
        POComboList.row = 0
        POComboList.RowHeightMin = 240
    End With
End Sub

Sub getLINEitems(INVOICE As String, Optional lineITEM As Integer)
Dim dataPO As New ADODB.Recordset
Dim sql, rowTEXT, stock As String
Dim Q, U, P
Dim i As Integer
Dim qty As Double

    On Error Resume Next
    Screen.MousePointer = 11
    If lineITEM < 1 Then Call makeDETAILgrid
    If INVOICE = "*" Then
        sql = "SELECT * from PO_Details_For_Invoice WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "' "
            If lineITEM > 0 Then
                sql = sql + "AND lineItem = " + Format(lineITEM) + " "
            End If
            sql = sql + "ORDER BY PO, CONVERT(integer, LineItem)"
    Else
        INVOICE = Trim(INVOICE)
        sql = "SELECT * from Invoice_Details WHERE NameSpace = '" + deIms.NameSpace + "' " _
            & "AND PO = '" + cell(0) + "' AND Invoice = '" + INVOICE + "' ORDER BY PO, CONVERT(integer, LineItem)"
    End If
    POlist.RowHeightMin = 0
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then Exit Sub
    With dataPO
        If .RecordCount > 0 Then
            
            Do While Not .EOF
                rowTEXT = "" + vbTab
                rowTEXT = rowTEXT + IIf(IsNull(!lineITEM), "", !lineITEM) + vbTab 'PO Line Item
                stock = IIf(IsNull(!StockNumber), "", Trim(!StockNumber)) + " - " + IIf(IsNull(!Description), "", !Description)
                rowTEXT = rowTEXT + stock + vbTab 'Stock Number + Description
                rowTEXT = rowTEXT + "" + vbTab 'Line
                
                'Purchase
                rowTEXT = rowTEXT + FormatNumber(!Quantity1, 2) + vbTab 'Primary Quantity
                rowTEXT = rowTEXT + IIf(IsNull(!unit1), "", Trim(!unit1)) + vbTab 'Primary Unit
                rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!unitprice1), 0, !unitprice1), 2) + vbTab 'Primary Unit Price
                
                'Invoice
                rowTEXT = rowTEXT + "" + vbTab 'Line
                If INVOICE = "*" Then
                    If IsNumeric(!SumQty1) Then
                        qty = !SumQty1
                    Else
                        qty = 0
                    End If
                    Q = IIf(qty = 0, "", FormatNumber(qty, 2))
                    U = IIf(IsNull(!unit1), "", Trim(!unit1))
                    P = IIf(IsNull(!SumUnitPrice1), "", FormatNumber(!SumUnitPrice1, 2))
                    rowTEXT = rowTEXT + Q + vbTab   'Sumary Primary Quantity
                    rowTEXT = rowTEXT + U + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + P + vbTab 'Sumary Primary Unit Price
                Else
                    If IsNumeric(!QuantityI1) Then
                        qty = !QuantityI1
                    Else
                        qty = 0
                    End If
                    Q = IIf(qty = 0, "", FormatNumber(qty, 2))
                    U = IIf(IsNull(!unit1), "", Trim(!unit1))
                    P = IIf(IsNull(!UnitPriceI1), 0, FormatNumber(!UnitPriceI1, 2))
                    rowTEXT = rowTEXT + Q + vbTab   'Primary Quantity
                    rowTEXT = rowTEXT + U + vbTab 'Primary Unit
                    rowTEXT = rowTEXT + P + vbTab 'Primary Unit Price
                End If
                
                If lineITEM = 0 Then
                    POlist.AddItem rowTEXT
                    POlist.row = POlist.Rows - 1
                Else
                    POlist.TextMatrix(POlist.row, 8) = Q 'Primary Quantity
                    POlist.TextMatrix(POlist.row, 9) = U 'Primary Unit
                    POlist.TextMatrix(POlist.row, 10) = P 'Primary Unit Price
                End If
                
                POlist.TextMatrix(POlist.row, 16) = !Unit1Code
                POlist.TextMatrix(POlist.row, 17) = IIf(IsNull(!invoices), 0, !invoices)
                If lineITEM = 0 Then Call colorCOLS(INVOICE = "*")
                Call differences(POlist.row)
                If !unit1 = !unit2 Then
                    POlist.TextMatrix(POlist.row, 15) = ""
                Else
                    POlist.TextMatrix(POlist.row, 15) = !UnitSwitch
                    nomPicture(0).Visible = True
                    nomLabel(0).Visible = True
                    POlist.RowHeight(POlist.row) = 240
                    rowTEXT = "" + vbTab + "" + vbTab + "" + vbTab
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    
                    'Purchase
                    rowTEXT = rowTEXT + FormatNumber(!Quantity2, 2) + vbTab 'Secundary Quantity
                    rowTEXT = rowTEXT + IIf(IsNull(!unit2), "", Trim(!unit2)) + vbTab 'Secundary Unit
                    rowTEXT = rowTEXT + FormatNumber(IIf(IsNull(!UnitPrice2), 0, !UnitPrice2), 2) + vbTab 'Secundary Unit Price
                    
                    'Invoice
                    rowTEXT = rowTEXT + "" + vbTab 'Line
                    If INVOICE = "*" Then
                        If IsNumeric(!SumQty2) Then
                            qty = !SumQty2
                        Else
                            qty = 0
                        End If
                        Q = IIf(qty = 0, "", FormatNumber(qty, 2))
                        U = IIf(IsNull(!unit2), "", Trim(!unit2))
                        P = IIf(IsNull(!SumUnitPrice2), "", FormatNumber(!SumUnitPrice2, 2))
                        rowTEXT = rowTEXT + Q + vbTab   'Sumary Primary Quantity
                        rowTEXT = rowTEXT + U + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + P + vbTab 'Sumary Primary Unit Price
                    Else
                        If IsNumeric(!QuantityI2) Then
                            qty = !QuantityI2
                        Else
                            qty = 0
                        End If
                        Q = IIf(qty = 0, "", FormatNumber(qty, 2))
                        U = IIf(IsNull(!unit2), "", Trim(!unit2))
                        P = IIf(IsNull(!UnitPriceI2), 0, FormatNumber(!UnitPriceI2, 2))
                        rowTEXT = rowTEXT + Q + vbTab   'Primary Quantity
                        rowTEXT = rowTEXT + U + vbTab 'Primary Unit
                        rowTEXT = rowTEXT + P + vbTab 'Primary Unit Price
                    End If
                    
                    If lineITEM = 0 Then
                        POlist.AddItem rowTEXT
                        POlist.row = POlist.Rows - 1
                    Else
                        POlist.TextMatrix(POlist.row, 8) = Q 'Primary Quantity
                        POlist.TextMatrix(POlist.row, 9) = U 'Primary Unit
                        POlist.TextMatrix(POlist.row, 10) = P 'Primary Unit Price
                    End If
                                        
                    POlist.TextMatrix(POlist.row, 15) = !UnitSwitch
                    POlist.TextMatrix(POlist.row, 16) = !Unit2Code
                    POlist.TextMatrix(POlist.row, 17) = IIf(IsNull(!invoices), 0, !invoices)
                    If lineITEM = 0 Then Call colorCOLS(INVOICE = "*")
                    POlist.Col = 1
                    If POlist = "" Then
                        POlist = "§"
                        POlist.CellFontName = "Wingdings"
                    End If
                    Call differences(POlist.row)
                    If UCase(Trim(!UnitSwitch)) = "P" Or IsNull(!UnitSwitch) Then POlist.row = POlist.Rows - 2
                    For i = 4 To 6
                        POlist.Col = i
                        POlist.CellBackColor = &HC0C0FF
                    Next
                    
                    If lineITEM = 0 Then POlist.row = POlist.Rows - 1
                End If
                
                If lineITEM = 0 Then
                    POlist.RowHeight(POlist.row) = 240
                    POlist.AddItem ""
                    POlist.row = POlist.Rows - 1
                    For i = 0 To POlist.Cols - 1
                        POlist.Col = i
                        If i = 0 Then
                            POlist.CellBackColor = &H808080
                        Else
                            POlist.CellBackColor = &HE0E0E0
                        End If
                    Next
                    POlist.RowHeight(POlist.row) = 50
                    POlist.TextMatrix(POlist.row, 13) = 50
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
            If lineITEM = 0 Then
                POlist.RemoveItem (1)
                POlist.RemoveItem (POlist.Rows - 1)
                POlist.row = 0
            End If
        End If
    End With
    Screen.MousePointer = 0
End Sub

Sub getRECIPIENTSlist()
    With RecipientList
        .ColWidth(0) = 300
        .ColWidth(1) = 9095
        .Rows = 2
        .Clear
        msg1 = translator.Trans("L00241")
        .TextMatrix(0, 1) = "Recipient List"
    End With
    If Not IsNothing(rsReceptList) Then
        With rsReceptList
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    RecipientList.AddItem "" + vbTab + .Fields(0)
                    .MoveNext
                Loop
            End If
        End With
        If RecipientList.Rows > 2 Then RecipientList.RemoveItem 1
    End If
End Sub


Sub gridLIST(ByVal mainGRID As MSHFlexGrid, ByVal childGRID As MSHFlexGrid)
Dim h, i As Integer
    
    With childGRID
        .Left = mainGRID.Left + mainGRID.ColWidth(0)
        h = 20
        For i = 0 To mainGRID.row
            h = h + mainGRID.RowHeight(i)
        Next
        .Top = h + mainGRID.Top - 30
        .Visible = True
        .SetFocus
    End With
End Sub

Sub gridONfocus(ByRef Grid As MSHFlexGrid)
Dim i, x As Integer
    With Grid
        x = .Col
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = &H800000   'Blue
            .CellForeColor = &HFFFFFF   'White
        Next
        .Col = x
        .Tag = .row
    End With
End Sub

Sub lockDOCUMENT(locked As Boolean)
Dim i As Integer
    
    If locked Then
        cell(3).locked = True
    Else
        cell(3).locked = False
    End If
    
    If locked Then
        remark.locked = True
        Imsmail1.Enabled = False
        cmd_Add.Enabled = False
        cmd_Remove.Enabled = False
    Else
        remark.locked = False
        Imsmail1.Enabled = True
        cmd_Add.Enabled = True
        cmd_Remove.Enabled = True
    End If
End Sub

Sub makeDETAILgrid()
Dim i, Col As Integer
    With POlist
        .Clear
        .Rows = 2
        For i = 0 To 12
            .ColWidth(i) = 1000
            .ColAlignment(i) = 6
            .ColAlignmentFixed(i) = 4
        Next
        
        'Col 0
        .ColAlignment(0) = 4
        .ColWidth(0) = 285
        .row = 0
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .TextMatrix(0, 0) = "®"
        
        'Section 1
        .ColAlignment(1) = 6
        .ColWidth(1) = 400
        .ColWidth(2) = 3600
        .ColAlignment(2) = 0
        .TextMatrix(0, 1) = "Line #"
        .TextMatrix(0, 2) = "Commodity Description"
        
        'Section 2
        .TextMatrix(0, 4) = "Quantity"
        .TextMatrix(0, 5) = "Unit"
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Unit Price"
        
        'Section 3
        .TextMatrix(0, 8) = "Quantity"
        .TextMatrix(0, 9) = "Unit"
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Unit Price"
        
        Call drawLINEcol(POlist, 3)
        For i = 0 To 2
            Col = i * 4
            Call drawLINEcol(POlist, 3 + Col)
        Next
        .TextMatrix(0, 12) = "Unit Price Difference"
        
        'Invisible columns
        For i = 13 To 17
            .ColWidth(i) = 0
        Next
        .TextMatrix(0, 13) = "Real Height"
        .TextMatrix(0, 14) = "Old value"
        .TextMatrix(0, 15) = "Switch"
        .TextMatrix(0, 16) = "Unit of Mesure Code"
        .TextMatrix(0, 17) = "Invoices"
        .row = 1
        .Col = 1
        .RowHeight(0) = 500
        .RowHeightMin = 240
        .WordWrap = True
        .Tag = ""
    End With
        
    With POtitles
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(4) = 4
        .row = 0
        Call drawLINEcol(POtitles, 1)
        Call drawLINEcol(POtitles, 3)
        .row = 1
        Call fixPOtitles(0)
    End With

End Sub

Function Iexists() As Boolean
Dim sql, INVOICE As String
Dim dataPO  As New ADODB.Recordset
    On Error Resume Next
    Iexists = True
    INVOICE = Trim(cell(0))
    sql = "SELECT inv_invcnumb from Invoice WHERE inv_npecode = '" + deIms.NameSpace + "' " _
        & "AND inv_ponumb = '" + cell(0) + "' AND inv_invcnumb = '" + cell(1) + "'"
    Set dataPO = New ADODB.Recordset
    dataPO.Open sql, deIms.cnIms, adOpenForwardOnly
    If Err.number <> 0 Then
        Iexists = False
        Exit Function
    End If
    If dataPO.RecordCount < 1 Then
        Iexists = False
    End If
End Function

Sub showDTPicker1(cellNUMBER As Integer)
    With cell(cellNUMBER)
        DTPicker1.Tag = cellNUMBER
        DTPicker1.Top = .Top
        DTPicker1.Height = .Height
        DTPicker1.Left = .Left
        DTPicker1.Width = .Width
        DTPicker1.ZOrder
        DTPicker1.Visible = True
        DTPicker1.SetFocus
    End With
End Sub

Sub showLIST(ByRef Grid As MSHFlexGrid)
    With Grid
        If .Rows > 0 And .text <> "" Then
            .ZOrder
            .Visible = True
        End If
    End With
End Sub

Sub showTEXTline()
Dim positionX, positionY, i, currentCOL As Integer
    With POlist
        currentCOL = .Col
        currentROW = .row
        If .TextMatrix(.row, 0) <> "" Then
            If Trim(.TextMatrix(.row, 15)) = "P" Then
                If .TextMatrix(.row, 1) = "§" Then
                    If .Col = 10 Then Exit Sub
                End If
            Else
                If .TextMatrix(.row, 1) <> "§" Then
'                    If .col = 10 Then Exit Sub
                End If
            End If
            
            positionX = .Left + 30
            For i = 0 To .Col - 1
                positionX = positionX + .ColWidth(i)
            Next
            positionY = .Top + 30 + .RowPos(currentROW)
            TextLINE.text = .text
            TextLINE.Left = positionX
            TextLINE.Width = .ColWidth(.Col) - 20
            TextLINE.Top = positionY
            TextLINE.Height = .RowHeight(.row) - 20
            TextLINE.Tag = .row
            TextLINE.SelStart = 0
            TextLINE.SelLength = Len(TextLINE.text)
            TextLINE.Visible = True
            TextLINE.SetFocus
        End If
        .Col = currentCOL
        .row = currentROW
    End With
End Sub

Sub textBOX(ByVal mainCONTROL As MSHFlexGrid, standard As Boolean)
Dim h, i As Integer
Dim box As textBOX

    With mainCONTROL
        box.Height = .RowHeight(i)
        box.Height = box.Height + 10
        If .row = 0 And .FixedRows > 0 Then
            box.Top = .Top
            box.Height = box.Height - 80
        Else
            If standard Then
                box.Left = .Left + .ColWidth(0)
                h = 20
                For i = 0 To .row - 1
                    h = h + .RowHeight(i)
                Next
                box.Top = h + .Top - 30
                box.Width = .ColWidth(1)
            Else
                box.Left = .Left
                box.Top = .Top - box.Height
                box.Width = .ColWidth(0)
            End If
        End If
        box.Visible = True
        box.text = .text
        If standard Then
            box.SetFocus
        End If
    End With
End Sub



Private Sub cell_Change(Index As Integer)
    If Me.ActiveControl.Name = "cell" Then
        With cell(Index)
            Select Case Index
                Case 0
                    If Form = mdVisualization Then
                        If cell(Index) = "" Then
                            Call clearDOCUMENT
                            NavBar1.NewEnabled = False
                        Else
                            If Me.ActiveControl.Name = "cell" Then
                                If Me.ActiveControl.Index = 0 Then Call alphaSEARCH(cell(Index), POComboList, 0)
                            End If
                        End If
                    Else
                        If cell(0) = "" Then
                        End If
                    End If
                Case 1
                    If Form <> mdVisualization Then
                        If Index = 1 Then Exit Sub
                        If cell(Index) <> "" Then Call alphaSEARCH(cell(Index), InvoiceComboList, 0)
                    End If
            End Select
        End With
    End If
End Sub

Private Sub cell_Click(Index As Integer)
    Select Case Index
        Case 0
            If Form = mdVisualization Then
                Call showLIST(POComboList)
            Else
                POComboList.Visible = False
            End If
        Case 1
            If Form = mdVisualization Then
                Call showLIST(InvoiceComboList)
            Else
                InvoiceComboList.Visible = False
            End If
    End Select
End Sub

Private Sub cell_GotFocus(Index As Integer)
    With cell(Index)
        If Not .locked Then
            .BackColor = vbYellow
            .Appearance = 1
            .Refresh
            .Tag = .text
            Select Case Index
                Case 0
                    If Form = mdVisualization Then
                        If POComboList.Visible Then
                            POComboList.Visible = False
                        Else
                            Call showLIST(POComboList)
                        End If
                    End If
                Case 1
                    If Form = mdVisualization Then
                        If InvoiceComboList.Visible Then
                            InvoiceComboList.Visible = False
                        Else
                            Call showLIST(InvoiceComboList)
                        End If
                    End If
                Case 3
                    If IsDate(cell(Index)) Then DTPicker1.Value = CDate(cell(Index))
                    If Form <> mdVisualization Then
                        If .text = "" Then
                            DTPicker1.Value = Now
                        Else
                            DTPicker1.Value = CDate(.text)
                        End If
                        Call showDTPicker1(Index)
                    End If
            End Select
        End If
    End With
End Sub

Private Sub cell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim activeARROWS As Boolean
    With cell(Index)
        If Not .locked Then
            activeARROWS = False
            If Index <= 2 And Form = mdVisualization Then activeARROWS = True
            If activeARROWS Then
                Select Case KeyCode
                    Case 40
                        Call arrowKEYS("down", Index)
                    Case 38
                        Call arrowKEYS("up", Index)
                End Select
            End If
        End If
    End With
End Sub
Private Sub cell_KeyPress(Index As Integer, KeyAscii As Integer)
    With cell(Index)
        If Not .locked Then
            Select Case KeyAscii
                Case 13
                    If cell(Index) <> "" Then
                        Select Case Index
                            Case 0
                                If KeyAscii = 13 Then
                                    Select Case Form
                                        Case mdVisualization
                                            cell(0) = POComboList
                                            POComboList.Visible = False
                                            Call getINVOICE("*")
                                            Call getInvoiceComboList
                                        Case mdCreation
                                    End Select
                                End If
                                POComboList.Visible = False
                                cell(1).SetFocus
                            Case 1
                                If KeyAscii = 13 Then
                                    Select Case Form
                                        Case mdVisualization
                                            If cell(1) <> "" Then
                                                Call getINVOICE("*")
                                            End If
                                            InvoiceComboList.Visible = False
                                            cell(1).SetFocus
                                        Case mdCreation
                                            If Iexists Then
                                                msg1 = translator.Trans("M00282")
                                                MsgBox IIf(msg1 = "", "Transaction Number is already exist", msg1)
                                                Exit Sub
                                            Else
                                                cell(3).SetFocus
                                            End If
                                    End Select
                                End If
                            Case 7
                        End Select
                    End If
                Case 27
                    .text = cell(Index).Tag
                    Select Case Index
                        Case 0
                            POComboList.Visible = False
                        Case 1
                            InvoiceComboList.Visible = False
                        Case 7

                    End Select
            End Select
        End If
    End With
End Sub

Private Sub cell_LostFocus(Index As Integer)
On Error Resume Next
    With cell(Index)
        If Not .locked Then
            .BackColor = remark.BackColor
            Select Case Index
                Case 0
                    Select Case Form
                        Case mdVisualization
                            If cell(0) = Right(invoiceLABEL, Len(cell(0))) Then Exit Sub
                        Case mdCreation
                            POComboList.Visible = False
                            Exit Sub
                    End Select
                Case 1
                    Select Case Form
                        Case mdVisualization

                            If InvoiceComboList.Visible Then
                                InvoiceComboList.Visible = False
                            End If
                        Case mdCreation
                            If Iexists Then
                                msg1 = translator.Trans("M00282")
                                MsgBox IIf(msg1 = "", "Transaction Number is already exist", msg1)
                                cell(1).SelStart = 0
                                cell(1).SelLength = Len(cell(1))
                                SSTab1.Tab = 0
                                cell(1).SetFocus
                                Exit Sub
                            Else
                                cell(3).SetFocus
                            End If
                    End Select
                Case 2, 8, 9
                    .text = .Tag
                    If Me.ActiveControl.Name <> "DTPicker1" Then
                        DTPicker1.Visible = False
                    End If
                Case 3
                    If Me.ActiveControl.Name <> "invoiceComboList" Then
                        InvoiceComboList.Visible = False
                    End If
                Case 7
                    If Me.ActiveControl.Name <> "destinationList" Then

                    End If
            End Select
        End If
    End With
End Sub



Public Sub cell_Validate(Index As Integer, Cancel As Boolean)
    If Form <> mdVisualization Then
        With cell(Index)
            If Not .locked Then
                If .text <> "" Then
                    If Form = mdCreation Then
                        Select Case Index
                            Case 0, 1
                            Case 2, 8, 9
                                If Not IsDate(.text) Then
                                    .text = ""
                                End If
                            Case 3
                                If .text <> InvoiceComboList Then
                                    .text = ""
                                End If
                            Case 4
                            Case 5
                            Case 6
                            Case 7
                        End Select
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub cmd_Add_Click()
    Imsmail1.AddCurrentRecipient
End Sub

Private Sub cmd_Remove_Click()
On Error Resume Next
    If RecipientList.row > 0 Then
        If RecipientList.TextMatrix(RecipientList.row, 1) <> "" Then
            rsReceptList.MoveFirst
            rsReceptList.Find "Recipients = '" & Trim$(RecipientList.TextMatrix(RecipientList.row, 1)), , adSearchForward
            If Not rsReceptList.EOF Then
                rsReceptList.Delete
                rsReceptList.Update
            End If
        End If
        Call getRECIPIENTSlist
    End If
    If Err Then Err.Clear
End Sub


Private Sub Command1_Click()
Dim showALL As Boolean
Dim i As Integer
    If Command1.Caption = "&Show Only Selection" Then
        Command1.Caption = "&Show All Records"
        showALL = False
    Else
        Command1.Caption = "&Show Only Selection"
        showALL = True
    End If
    
    With POlist
        .Col = 0
        If showALL Then
            .RowHeightMin = 50
            .RowHeight(-1) = 240
        Else
            For i = 1 To .Rows - 1
                If .RowHeight(i) > 240 Then
                    .TextMatrix(i, 13) = .RowHeight(i)
                End If
            Next
            .RowHeightMin = 0
            .RowHeight(-1) = 0
            For i = .Rows - 1 To 1 Step -1
                .row = i
                If .text <> "" Then
                    .RowHeight(i) = 240
                End If
            Next
        End If
        .RowHeight(0) = 500
        For i = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(i, 13)) Then
                If val(.TextMatrix(i, 13)) > 240 Then
                    .RowHeight(i) = val(.TextMatrix(i, 13))
                End If
            End If

            If showALL Then
                If val(.TextMatrix(i, 13)) = 50 Then .RowHeight(i) = 50
            Else
                If .TextMatrix(i, 0) <> "" And Not IsNumeric(.TextMatrix(i, 1)) Then
                    If .Rows > i + 1 Then .RowHeight(i + 1) = 50
                End If
            End If
        Next
    End With
End Sub



Private Sub DTPicker1_Change()
    If DTPicker1.Month = 0 Then DTPicker1.Month = Month(Now)
    If DTPicker1.Day = 0 Then DTPicker1.Month = Day(Now)
    If DTPicker1.Year = 0 Then DTPicker1.Month = Year(Now)
End Sub


Public Sub DTPicker1_DropDown()
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    With DTPicker1
        Select Case KeyCode
            Case 13
                cell(val(.Tag)).text = Format(.Value, "MMMM/dd/yyyy")
                remark.SetFocus
        End Select
    End With
End Sub

Private Sub DTPicker1_LostFocus()
Dim indexCELL As Integer
    With DTPicker1
        If IsNumeric(.Tag) Then
            cell(val(.Tag)).text = Format(.Value, "MMMM/dd/yyyy")
            indexCELL = val(.Tag)
            If Me.ActiveControl.Name = "cell" Then
                If Me.ActiveControl.Index <> val(.Tag) Then .Visible = False
                indexCELL = Me.ActiveControl.Index
            End If
            If Me.ActiveControl.Name = "cell" Then
                remark.SetFocus
            Else
                .Visible = False
            End If
        End If
        .Value = Now
    End With
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = 0
    If Form = mdVisualization Then
        NavBar1.SaveEnabled = False
        NavBar1.CancelEnabled = False
        NavBar1.NewEnabled = False
        If Iexists Then
            NavBar1.PrintEnabled = True
            NavBar1.EMailEnabled = True
        End If
    End If
    cell(0).SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rights

    SSTab1.TabVisible(2) = False
    Call translator.Translate_Forms("frmInvoice")
    Imsmail1.NameSpace = deIms.NameSpace
    Imsmail1.SetActiveConnection deIms.cnIms
    Imsmail1.Language = Language
    NavBar1.Language = Language
    Call begining
    Form = mdVisualization
    Screen.MousePointer = 11
    SSTab1.Tab = 0
    Call lockDOCUMENT(True)
    Call getPOComboList
    frmInvoice.Caption = frmInvoice.Caption + " - " + frmInvoice.Tag
    rights = Getmenuuser(deIms.NameSpace, CurrentUser, Me.Tag, deIms.cnIms)
    SaveEnabled = rights
    NavBar1.NewEnabled = SaveEnabled
    cell(0).Enabled = True
    cell(1).Enabled = True
    Screen.MousePointer = 0
    If Err Then Call LogErr(Name & "::Form_Load", Err.Description, Err.number, True)
    frmInvoice.Left = Int((MDI_IMS.Width - frmInvoice.Width) / 2)
    frmInvoice.Top = Int((MDI_IMS.Height - frmInvoice.Height) / 2) - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim closing
    If Form <> mdVisualization Then
        closing = MsgBox("Do you really want to close and lose your last record?", vbYesNo)
        If closing = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub IMSMail1_OnAddClick(ByVal address As String)
On Error Resume Next

    If IsNothing(rsReceptList) Then
        Set rsReceptList = New ADODB.Recordset
        Call rsReceptList.Fields.Append("Recipients", adVarChar, 60, adFldUpdatable)
        rsReceptList.Open
    End If
    
    If (InStr(1, address, "@") > 0) And InStr(1, UCase(address), "INTERNET!") = 0 Then
        address = "INTERNET!" & UCase(address)
    End If
    
    If Not IsInList(address, "Recipients", rsReceptList) Then
        Call rsReceptList.AddNew(Array("Recipients"), Array(address))
    End If

    Call getRECIPIENTSlist
End Sub

Private Sub NavBar1_BeforeSaveClick()
Dim wrong, wrong2 As Boolean
Dim i, ii, position, Col As Integer
On Error Resume Next
Screen.MousePointer = 11
    
    'Revision for Header
    wrong = False
    For i = 0 To 3
        If cell(i) = "" Then
            NavBar1.SaveEnabled = SaveEnabled
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00016")
            MsgBox IIf(msg1 = "", Label(i) + " Cannot be left empty", msg1)
            cell(i).SetFocus
            Exit Sub
        End If
    Next
    If wrong Then
        NavBar1.SaveEnabled = SaveEnabled
        Screen.MousePointer = 0
        msg1 = translator.Trans("M00122")
        MsgBox IIf(msg1 = "", "Invalid Value in " + Label(position), msg1)
        cell(position).SetFocus
        Exit Sub
    End If

    'Revision for Details
    wrong = True
    wrong2 = False
    position = 0
    wrong = False
    readyFORsave = True
    For i = 1 To POlist.Rows - 1
        If POlist.TextMatrix(i, 0) <> "" Then
            For ii = 0 To 1
                Col = 8 + (ii * 2)
                If IsNumeric(POlist.TextMatrix(i, Col)) Then
                    If CDbl(POlist.TextMatrix(i, Col)) > 0 Then
                    Else
                        readyFORsave = False
                        wrong = True
                        position = i
                        Exit For
                    End If
                Else
                    readyFORsave = False
                    wrong = True
                    position = i
                    Exit For
                End If
            Next
            If wrong2 Then
                readyFORsave = False
                wrong = True
                Exit For
            End If
        End If
        If wrong Then Exit For
    Next
    If wrong Then
        SSTab1.Tab = 1
        If position > 0 Then
            NavBar1.SaveEnabled = SaveEnabled
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00122")
            MsgBox IIf(msg1 = "", "Invalid Value", msg1)
            POlist.row = position
            POlist.Col = Col
            POlist.SetFocus
        Else
            NavBar1.SaveEnabled = SaveEnabled
            Screen.MousePointer = 0
            msg1 = translator.Trans("M00707")
            MsgBox IIf(msg1 = "", "You have to select at least one line item.", msg1)
        End If
    Else
        Call SAVE
        Call ChangeMode(mdVisualization)
        Call getPOComboList
        Call getInvoiceComboList
        Call getINVOICE(cell(0))
        cell(0).locked = False
        cell(0).SelLength = Len(cell(0))
        cell(0).SelStart = 0
        Picture1.Visible = False
        'msg1 = translator.Trans("M00306")
        MsgBox IIf(msg1 = "", "Insert into Supplier Invoice List is completed successfully", msg1)
        NavBar1.CancelEnabled = False
        POComboList.Visible = True
        cell(0).SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub NavBar1_OnCancelClick()
Dim response As String
    msg1 = translator.Trans("M00706")
    msg2 = translator.Trans("L00441")
    response = MsgBox(IIf(msg1 = "", "Are you sure you want to cancel changes?", msg1), vbYesNo, IIf(msg2 = "", "Cancel", msg2))
    If response = vbYes Then
        With NavBar1
            cell(0).locked = False
            Call ChangeMode(mdVisualization)
            If SSTab1.Tab > 0 Then SSTab1.Tab = 0
            Call lockDOCUMENT(True)
            Call clearDOCUMENT
            invoiceLABEL = ""
            currencyLABEL = ""
            If cell(0) <> "" Then
                .NewEnabled = SaveEnabled
                Call getINVOICE("*")
            End If
            .CancelEnabled = False
            .SaveEnabled = False
            .PrintEnabled = False
        End With
    End If
End Sub

Private Sub NavBar1_OnCloseClick()
    Unload Me
End Sub

Private Sub NavBar1_OnEMailClick()
Dim Params(1) As String
Dim rptinfo As RPTIFileInfo
Screen.MousePointer = 11
On Error Resume Next
    Call BeforePrint
    
    With rptinfo
        Params(0) = "namespace=" + deIms.NameSpace
        Params(1) = "manifestnumb=" + cell(0)
        .ReportFileName = ReportPath & "Invoice.rpt"
        Call translator.Translate_Reports("Invoice.rpt")
        .Parameters = Params
        
        
'            .ReportFileName = FixDir(App.Path) + "CRreports\Invoice.rpt"
'            .ParameterFields(1) = "invnumb;" + cell(1) + ";TRUE"
'            .ParameterFields(2) = "ponumb;" + cell(0) + ";TRUE"
'            Call translator.Translate_Reports("invoice.rpt")
'            Call translator.Translate_SubReports
        
        
        
    End With
    
    Params(0) = ""
    Call WriteRPTIFile(rptinfo, Params(0))
    Call SendEmailAndFax(rsReceptList, "Recipient", "Transaction " & cell(0), "", Params(0))
    Screen.MousePointer = 0
If Err Then Call LogErr(Name & "::NavBar1_OnEMailClick", Err.Description, Err.number, True)
End Sub

Private Sub NavBar1_OnNewClick()
Dim i As Integer
Dim sql, response As String
Dim dataUSER As ADODB.Recordset

    Screen.MousePointer = 11
    currentROW = 0
    With NavBar1
        If cell(0) = "" Then
            Screen.MousePointer = 0
            MsgBox "Invalid Transaction Number"
        Else
            If isOPEN(cell(0)) Then
                POComboList.Visible = False
                InvoiceComboList.Visible = False
                For i = 1 To 3
                    cell(i) = ""
                Next
                remark = ""
                cell(4) = Format(Now, "MMMM/dd/yyyy")
                Call ChangeMode(mdCreation)
                Call begining
                Set dataUSER = New ADODB.Recordset
                sql = "SELECT usr_username FROM XUSERPROFILE WHERE usr_npecode = '" + deIms.NameSpace + "' AND usr_userid = '" + CurrentUser + "'"
                dataUSER.Open sql, deIms.cnIms, adOpenForwardOnly
                If dataUSER.RecordCount > 0 Then
                    cell(2) = dataUSER!usr_username
                End If
                Screen.MousePointer = 0
                .NewEnabled = False
                .CancelEnabled = True
                .SaveEnabled = True
                .PrintEnabled = False
                
                Screen.MousePointer = 11
                Call getLINEitems("*")
                Call lockDOCUMENT(False)
            Else
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End With
    Screen.MousePointer = 0
    cell(1).SetFocus
End Sub

Private Function ChangeMode(FMode As FormMode) As Boolean
On Error Resume Next
    Select Case FMode
        Case mdCreation
            lblStatu.ForeColor = vbRed
            msg1 = translator.Trans("L00125")
            lblStatu.Caption = IIf(msg1 = "", "Creation", msg1)
            lblStatu.Tag = "Creation"
            ChangeMode = True
        Case mdVisualization
            lblStatu.ForeColor = vbGreen
            msg1 = translator.Trans("L00092") 'J added
            lblStatu.Caption = IIf(msg1 = "", "Visualization", msg1) 'J modified
            lblStatu.Tag = "Visualization"
            ChangeMode = True
    End Select
    Form = FMode
End Function

Private Sub NavBar1_OnPrintClick()
On Error Resume Next
Screen.MousePointer = 11
    If cell(1) = "" Then
        Screen.MousePointer = 0
        MsgBox "Please select a valid Vendor Invoice"
        cell(1).SetFocus
        Exit Sub
    End If

    With MDI_IMS.CrystalReport1
        Call BeforePrint
        'msg1 = translator.Trans("L00213")
        .WindowTitle = IIf(msg1 = "", "Invoice", msg1)
        .Action = 1
    End With
Screen.MousePointer = 0
End Sub

Sub SAVE()
Dim header As New ADODB.Recordset
Dim details As New ADODB.Recordset
Dim remarks As New ADODB.Recordset
Dim INVitem As New ADODB.Recordset
Dim i, row As Integer
Dim sql As String
Dim Q, Quantity, PRICE As Double
On Error Resume Next

    Screen.MousePointer = 11
    If readyFORsave Then
        Picture1.Visible = True
        Picture1.ZOrder
        Picture1.Refresh
        Me.Refresh
        
        'Header routine
        msg1 = translator.Trans("M00708")
        MDI_IMS.StatusBar1.Panels(1).text = IIf(msg1 = "", "Saving Header", msg1)
        deIms.cnIms.BeginTrans
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM INVOICE WHERE inv_ponumb = ''"
        header.Open sql, deIms.cnIms, adOpenDynamic, adLockPessimistic
        With header
            .AddNew
            !inv_creauser = CurrentUser
            !inv_npecode = deIms.NameSpace
            !inv_ponumb = cell(0)
            !inv_invcnumb = cell(1)
            !inv_invcdate = CDate(cell(3))
            !inv_creadate = CDate(cell(4))
            .Update
        End With
        
        'Remarks routine
        msg1 = translator.Trans("M00719")
        MDI_IMS.StatusBar1.Panels(1).text = IIf(msg1 = "", "Saving Remarks", msg1)
        Set header = New ADODB.Recordset
        sql = "SELECT * FROM INVOICEREM WHERE invr_ponumb = ''"
        remarks.Open sql, deIms.cnIms, adOpenDynamic, adLockPessimistic
        
        With remarks
            .AddNew
            !invr_creauser = CurrentUser
            !invr_npecode = deIms.NameSpace
            !invr_creadate = CDate(cell(4))
            
            !invr_ponumb = cell(0)
            !invr_invcnumb = cell(1)
            !invr_rem = remark
            !invr_linenumb = 1
            .Update
        End With
                
        'Details routine
        msg1 = translator.Trans("M00710")
        MDI_IMS.StatusBar1.Panels(1).text = IIf(msg1 = "", "Saving Details", msg1)
        Set details = New ADODB.Recordset
        sql = "SELECT * FROM INVOICEDETL WHERE invd_ponumb = ''"
        details.Open sql, deIms.cnIms, adOpenKeyset, adLockPessimistic
        With details
            For i = 1 To POlist.Rows - 1
                If POlist.TextMatrix(i, 0) <> "" Then
                    If IsNumeric(POlist.TextMatrix(i, 1)) Then
                        .AddNew
                        !invd_npecode = deIms.NameSpace
                        !invd_creauser = CurrentUser
                        !invd_creadate = CDate(cell(4))
                        !invd_ponumb = cell(0)
                        !invd_invcnumb = cell(1)
                        !invd_liitnumb = POlist.TextMatrix(i, 1)
                        Quantity = IIf(IsNumeric(POlist.TextMatrix(i, 8)), CDbl(POlist.TextMatrix(i, 8)), 0)
                        !invd_primreqdqty = Quantity
                        !invd_primuom = POlist.TextMatrix(i, 16)
                        PRICE = IIf(IsNumeric(POlist.TextMatrix(i, 10)), CDbl(POlist.TextMatrix(i, 10)), 0)
                        !invd_unitpric = PRICE
                        !invd_totapric = Quantity * PRICE
                                                
                        If Trim(POlist.TextMatrix(i, 15)) = "" Then
                            row = i
                        Else
                            row = i + 1
                        End If
                        Quantity = IIf(IsNumeric(POlist.TextMatrix(row, 8)), CDbl(POlist.TextMatrix(row, 8)), 0)
                        !invd_secoreqdqty = Quantity
                        !invd_secouom = POlist.TextMatrix(row, 16)
                        PRICE = IIf(IsNumeric(POlist.TextMatrix(row, 10)), CDbl(POlist.TextMatrix(row, 10)), 0)
                        !invd_secounitprice = PRICE
                        !invd_secototaprice = Quantity * PRICE
                    End If
                End If
            Next
            msg1 = translator.Trans("M00714")
            MDI_IMS.StatusBar1.Panels(1).text = IIf(msg1 = "", "Saving Transaction", msg1)
            .UpdateBatch
        End With
        msg1 = translator.Trans("M00715")
        MDI_IMS.StatusBar1.Panels(1).text = IIf(msg1 = "", "Commiting Transaction", msg1)
        deIms.cnIms.CommitTrans
        MDI_IMS.StatusBar1.Panels(1).text = ""
        Screen.MousePointer = 11
        Call lockDOCUMENT(True)
        Call clearDOCUMENT
        Call getPOComboList
    End If
    Screen.MousePointer = 0
End Sub

Private Sub POComboList_Click()
    Select Case Form
        Case mdVisualization
            POComboList.Tag = POComboList.row
            cell(0) = Trim(POComboList)
            If Left(cell(0), 1) = "(" And Right(cell(0), 1) = ")" Then
                Call clearDOCUMENT
                POComboList.Visible = True
                cell(0).SetFocus
            Else
                Call getINVOICE("*")
            End If
            Call getInvoiceComboList
            cell(0).SetFocus
        Case mdCreation
            cell(1).SetFocus
    End Select
End Sub

Private Sub POComboList_KeyPress(KeyAscii As Integer)
    With POComboList
        Select Case KeyAscii
            Case 13
                Select Case Form
                    Case mdVisualization
                        cell(0) = .text
                        Call getINVOICE("*")
                        Call getInvoiceComboList
                    Case mdCreation
                        cell(1).SetFocus
                End Select
            Case 27
                POComboList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(0) = Chr(KeyAscii)
                Call alphaSEARCH(cell(0), POComboList, 0)
                .Tag = ""
                cell(0).SetFocus
                cell(0).SelStart = Len(cell(0))
                cell(0).SelLength = 0
        End Select
    End With
End Sub

Public Sub POlist_Click()
Dim i, currentCOL, pointerCOL As Integer
    If Form <> mdVisualization Then
        With POlist
            If .TextMatrix(.row, 1) <> "" Then
                If .row > 0 Then
                    selectionSTART = .row
'                    If forceNAV Then
'                        pointerCOL = .Col
'                        forceNAV = False
'                    Else
                        pointerCOL = .MouseCol
'                    End If
                    Select Case pointerCOL
                        Case 0, 1
                            If multiMARKED Then
                                multiMARKED = False
                            Else
                                Call markROW
                            End If
                        Case 8, 10
                            Call showTEXTline
                    End Select
                End If
            End If
        End With
    End If
End Sub

Private Sub POComboList_EnterCell()
    With POComboList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
    End With
End Sub

Private Sub POComboList_GotFocus()
    Call gridONfocus(POComboList)
End Sub

Private Sub POComboList_LeaveCell()
    With POComboList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub


Private Sub POComboList_LostFocus()
    With POComboList
        cell(0).text = Trim(.text)
    End With
End Sub

Public Sub POComboList_Validate(Cancel As Boolean)
    cell(0) = Trim(POComboList)
End Sub

Private Sub POlist_EnterCell()
Dim changeCOLORS As Boolean
    If Form <> mdVisualization Then
        Dim i, currentCOL As Integer
        With POlist
            currentCOL = .Col
            currentROW = .row
            If IsNumeric(.Tag) Then
                If val(.Tag) = .row Then
                    changeCOLORS = False
                Else
                    If TextLINE.Visible Then
'                        currentROW = val(TextLINE.Tag)
                        'TextLINE.Visible = False
                    Else
                        currentROW = .row
                    End If
                    .row = val(.Tag)
                    If colorsROW(1) <> "" Then
                        For i = 1 To 12
                            .Col = i
                            .CellBackColor = colorsROW(i)
                        Next
                        .Col = currentCOL
                    End If
                    .row = currentROW
                    .Tag = currentROW
                    Call getCOLORSrow
                    changeCOLORS = True
                End If
            Else
                POlist.Tag = .row
                Call getCOLORSrow
                changeCOLORS = True
            End If
            
            If .TextMatrix(.row, 1) <> "" Then
                currentCOL = .Col
                If changeCOLORS Then
                    For i = 1 To 12
                        .Col = i
                        Select Case .CellBackColor
                            Case &HC0FFFF 'Very Light Yellow
                                .CellBackColor = &HFFC0C0 'Very Light Blue
                            Case &HC0C0FF 'Very Light Red
                                .CellBackColor = &HFFC0FF 'Very Light Magenta
                            Case &HE0E0E0 'Very Light Gray
                            Case Else
                                .CellBackColor = &HFFC0C0 'Very Light Blue
                        End Select
                    Next
                    Select Case .Col
                        Case 8, 10
                            Call showTEXTline
                    End Select
                End If
            End If
            .Col = currentCOL
        End With
    End If
End Sub

Private Sub POlist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i, n

    With POlist
        If .TextMatrix(.MouseRow, 1) = "" Then
            If IsNumeric(.TextMatrix(.MouseRow, 13)) Then
                .RowHeight(.MouseRow) = val(.TextMatrix(.MouseRow, 13))
            End If
        End If
        If Shift = 1 Then
            multiMARKED = True
            n = 0
            For i = selectionSTART To .MouseRow
                If .TextMatrix(i, 0) = "" And .RowHeight(i) > 200 Then Exit For
                n = n + 1
            Next
            If selectionSTART > 0 Then
                If .MouseRow >= (selectionSTART + n) Then
                    For i = selectionSTART + n To .MouseRow
                        .row = i
                        .Col = 0
                        If .TextMatrix(i, 0) = "" Then
                            If .RowHeight(i) > 200 Then
                                Call markROW(True)
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
    End With
    
End Sub

Private Sub POlist_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim row, Col As Integer
    With POlist
        row = .MouseRow
        Col = .MouseCol
        If Col = 0 Then
            If .TextMatrix(row, 1) = "" Then
                If IsNumeric(.TextMatrix(row, 13)) Then
                    .RowHeight(row) = val(.TextMatrix(row, 13))
                Else
                    .RowHeight(row) = 240
                End If
            End If
        End If
    End With
End Sub

Private Sub POlist_Scroll()
    If Form <> mdVisualization Then TextLINE.Visible = False
    If POlist.leftCOL > 0 Then
        Call fixPOtitles(POlist.leftCOL)
    End If
End Sub

Private Sub POlist_SelChange()
    With POlist
        If Form <> mdVisualization Then
            If .TextMatrix(.row, 1) <> "" Then
                If .RowHeight(POlist.row) > 240 Then
                    .TextMatrix(POlist.row, 13) = .RowHeight(POlist.row)
                End If
            End If
        End If
    End With
End Sub

Private Sub InvoiceComboList_Click()
    Select Case Form
        Case mdVisualization
            InvoiceComboList.Tag = InvoiceComboList.row
            cell(1) = Trim(InvoiceComboList)
            If Left(cell(0), 1) = "(" And Right(cell(0), 1) = ")" Then
                Call getINVOICE(cell(1))
                cell(1).SetFocus
            Else
                Call getINVOICE(cell(1))
            End If
        Case mdCreation
            cell(2).SetFocus
    End Select
End Sub

Private Sub InvoiceComboList_EnterCell()
    With InvoiceComboList
        .CellBackColor = &H800000 'Blue
        .CellForeColor = &HFFFFFF 'White
        If Me.ActiveControl.Name = .Name Then cell(1) = .text
    End With
End Sub


Private Sub InvoiceComboList_GotFocus()
    Call gridONfocus(InvoiceComboList)
End Sub

Private Sub InvoiceComboList_KeyPress(KeyAscii As Integer)
    With InvoiceComboList
        Select Case KeyAscii
            Case 13
                cell(2).SetFocus
            Case 27
                InvoiceComboList.Visible = False
            Case Else
                .CellBackColor = &HFFFF00   'Cyan
                .CellForeColor = &H80000008 'Default Window Text
                .Tag = ""
                cell(1) = Chr(KeyAscii)
                Call alphaSEARCH(cell(1), InvoiceComboList, 0)
                .Tag = ""
                cell(1).SetFocus
                cell(1).SelStart = Len(cell(1))
                cell(1).SelLength = 0
        End Select
    End With
End Sub

Private Sub InvoiceComboList_LeaveCell()
    With InvoiceComboList
        .CellBackColor = &HFFFF00   'Cyan
        .CellForeColor = &H80000008 'Default Window Text
    End With
End Sub

Private Sub InvoiceComboList_LostFocus()
    With InvoiceComboList
        cell(1).text = Trim(.text)
        cell(1).SetFocus
        cell(1).SelStart = Len(cell(1))
        cell(1).SelLength = 0
    End With
End Sub

Private Sub InvoiceComboList_Validate(Cancel As Boolean)
    cell(1) = InvoiceComboList
    InvoiceComboList.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case PreviousTab
        Case 0
            If cell(0) = "" Then
                SSTab1.Tab = 0
            Else
                If NavBar1.CancelEnabled Or Form = mdVisualization Then
                    invoiceLABEL = "Transaction # " + cell(0)
                    currencyLABEL = "Currency: " + cell(5)
                    If Form = mdVisualization Then
                        Command1.Enabled = False
                    Else
                        Command1.Enabled = True
                    End If
                Else
                    SSTab1.Tab = 0
                End If
            End If
        Case 1
    End Select
    With NavBar1
        If SSTab1.Tab = 0 Then
            If Form = mdVisualization Then
                .NewEnabled = SaveEnabled
            Else
                .SaveEnabled = True
            End If
        Else
            .NewEnabled = False
            .SaveEnabled = False
        End If
    End With
End Sub

Private Sub TextLINE_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyUp
'            If POlist.Col = 8 Then
'                moveUP = True
'                POlist.SetFocus
'            End If
'        Case vbKeyRight
'        Case vbKeyDown
'            If POlist.Col = 8 Then
'                moveUP = False
'                POlist.SetFocus
'            End If
'        Case vbKeyLeft
'    End Select
End Sub

Private Sub TextLINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call TextLINE_Validate(True)
        Case 27
            TextLINE.Visible = False
    End Select
End Sub


Private Sub TextLINE_LostFocus()
Dim h, i, n As Integer
Dim outFOR As Boolean
    With TextLINE
        If .Visible Then
            .Visible = False
            Call TextLINE_Validate(True)
'            If Command1.Caption = "&Show All Records" Then
'                If moveUP Then
'                    If POlist.row > 1 Then
'                        n = POlist.row - 1
'                    Else
'                        n = POlist.Rows - 1
'                    End If
'                Else
'                    If POlist.row < POlist.Rows - 2 Then
'                        n = POlist.row + 1
'                    Else
'                        n = 1
'                    End If
'                End If
'                outFOR = False
'                For h = 1 To 2
'                    If moveUP Then
'                        For i = n To 1 Step -1
'                            Call checkNEXT(i, outFOR)
'                            If outFOR Then Exit For
'                        Next
'                    Else
'                        For i = n To POlist.Rows - 1
'                            Call checkNEXT(i, outFOR)
'                            If outFOR Then Exit For
'                        Next
'                    End If
'                    If outFOR Then
'                        Exit For
'                    Else
'                        n = 1
'                    End If
'                Next
'            End If
        End If
'        moveUP = False 'TO DISABLE IT
    End With
End Sub

Public Sub TextLINE_Validate(Cancel As Boolean)
Dim i, Col, row As Integer
Dim qty, switch, sql, t As String
Dim newPRICE, qty1, qty2, uPRICE1, uPRICE2, sumQTY, sumPRICE As Double
Dim newPRICEok As Boolean
Dim answer
Dim dataLINE As New ADODB.Recordset

    With TextLINE
        currentROW = val(.Tag)
        If POlist.Col = 8 Or POlist.Col = 10 Then
            Col = POlist.Col
            If IsNumeric(.text) Then
                If val(.text) > 0 Then
                    t = POlist.TextMatrix(currentROW, 1)
                    If Not IsNumeric(t) Then
                        t = POlist.TextMatrix(currentROW - 1, 1)
                        If Not IsNumeric(t) Then
                            MsgBox "Invalid value"
                            POlist.TextMatrix(currentROW, 1) = ""
                            Exit Sub
                        End If
                    End If
                    sql = "SELECT * from PO_Details_For_Invoice WHERE NameSpace = '" + deIms.NameSpace + "' " _
                        & "AND PO = '" + cell(0) + "' AND lineItem = " + t
                    Set dataLINE = New ADODB.Recordset
                    dataLINE.Open sql, deIms.cnIms, adOpenForwardOnly
                    If dataLINE.RecordCount > 0 Then
                        If IsNumeric(POlist.TextMatrix(currentROW, 1)) Then
                            sumQTY = CDbl(.text) + IIf(IsNull(dataLINE!SumQty1), 0, dataLINE!SumQty1)
                            sumPRICE = CDbl(.text) + IIf(IsNull(dataLINE!SumUnitPrice1), 0, dataLINE!SumUnitPrice1)
                        Else
                            sumQTY = CDbl(.text) + IIf(IsNull(dataLINE!SumQty2), 0, dataLINE!SumQty2)
                            sumPRICE = CDbl(.text) + IIf(IsNull(dataLINE!SumUnitPrice2), 0, dataLINE!SumUnitPrice2)
                        End If
                    Else
                        sumQTY = CDbl(.text)
                        sumPRICE = CDbl(.text)
                    End If
                    Select Case POlist.Col
                        Case 8
                            If IsNumeric(CDbl(POlist.TextMatrix(currentROW, 4))) Then
                                If sumQTY > CDbl(POlist.TextMatrix(currentROW, 4)) Then
                                    answer = MsgBox("This line item is being over invoice.  Do you want to continue?", vbYesNo)
                                    If answer = vbNo Then
                                        .text = FormatNumber(CDbl(POlist.TextMatrix(currentROW, 4)) - (sumQTY - CDbl(.text)), 2)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Case 10
                            If IsNumeric(CDbl(POlist.TextMatrix(currentROW, 6))) Then
                                If sumPRICE > CDbl(POlist.TextMatrix(currentROW, 6)) Then
                                    answer = MsgBox("This line item is over price.  Do you want to continue?", vbYesNo)
                                    If answer = vbNo Then
                                        .text = FormatNumber(POlist.TextMatrix(currentROW, 6) - (sumPRICE - CDbl(.text)), 2)
                                        Exit Sub
                                    End If
                                End If
                            End If
                    End Select
                    POlist.TextMatrix(val(.Tag), Col) = FormatNumber(.text, 2)
                    switch = Trim(POlist.TextMatrix(val(.Tag), 15))
                    Select Case switch
                        Case ""
                            Call differences(currentROW)
                        Case "P", "S"
                            If POlist.TextMatrix(currentROW, 1) = "§" Then
                                row = currentROW - 1
                            Else
                                row = currentROW
                            End If
                            newPRICEok = True
                            If IsNumeric(POlist.TextMatrix(row, 8)) Then
                                qty1 = CDbl(POlist.TextMatrix(row, 8))
                            Else
                                qty1 = 0
                                newPRICEok = False
                            End If
                            If IsNumeric(POlist.TextMatrix(row + 1, 8)) Then
                                qty2 = CDbl(POlist.TextMatrix(row + 1, 8))
                            Else
                                qty2 = 0
                                newPRICEok = False
                            End If
                            If switch = "P" Then
                                If IsNumeric(POlist.TextMatrix(row, 10)) Then
                                    uPRICE1 = CDbl(POlist.TextMatrix(row, 10))
                                Else
                                    uPRICE1 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE2 = (qty1 * uPRICE1) / qty2
                                    POlist.TextMatrix(row + 1, 10) = FormatNumber(uPRICE2, 2)
                                End If
                            Else
                                If IsNumeric(POlist.TextMatrix(row + 1, 10)) Then
                                    uPRICE2 = CDbl(POlist.TextMatrix(row + 1, 10))
                                Else
                                    uPRICE2 = 0
                                    newPRICEok = False
                                End If
                                If newPRICEok Then
                                    uPRICE1 = (qty2 * uPRICE2) / qty1
                                    POlist.TextMatrix(row, 10) = FormatNumber(uPRICE1, 2)
                                End If
                            End If
                            Call differences(row)
                            Call differences(row + 1)
                    End Select
                    
                    .Tag = ""
                    .text = ""
                    .Visible = False
                    Exit Sub
                End If
            End If
            If .text <> "" Then
                msg1 = translator.Trans("M00122")
                MsgBox IIf(msg1 = "", "Invalid Value", msg1)
                TextLINE = ""
            End If
        End If
    End With
End Sub


