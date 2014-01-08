VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_StockonHand 
   Caption         =   "Stock on Hand"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Tag             =   "03030400"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6450
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   11377
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Query"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Company"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Warehouse"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_filelen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_category"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_logwhse"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Subloc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_Currency"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_Descript"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbo_Warehouse"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frm_DirList"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbo_Company"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbo_Logwhse"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbo_subloc"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbo_Category"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_Company"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_Warehouse"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_Category"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_logwhse"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_subloc"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lst_query"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_selected"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_Currency"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "frm_Descript"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_Print"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_Send"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmd_Fax"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Recipients"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_recipients"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmd_Fax 
         Caption         =   "Fax"
         Height          =   288
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2880
         Width           =   972
      End
      Begin VB.CommandButton cmd_Send 
         Caption         =   "Send"
         Height          =   288
         Left            =   1200
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2880
         Width           =   972
      End
      Begin VB.CommandButton cmd_Print 
         Caption         =   "Print"
         Height          =   288
         Left            =   2280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2880
         Width           =   972
      End
      Begin VB.Frame frm_Descript 
         Height          =   600
         Left            =   5850
         TabIndex        =   39
         Top             =   2655
         Width           =   2310
         Begin VB.OptionButton opt_Short 
            Caption         =   "Short"
            Height          =   375
            Left            =   1215
            TabIndex        =   41
            Top             =   135
            Width           =   960
         End
         Begin VB.OptionButton opt_Long 
            Caption         =   "Long"
            Height          =   375
            Left            =   45
            TabIndex        =   40
            Top             =   135
            Width           =   735
         End
      End
      Begin VB.TextBox txt_Currency 
         Height          =   288
         Left            =   4080
         TabIndex        =   37
         Top             =   2835
         Width           =   705
      End
      Begin VB.TextBox txt_selected 
         Height          =   288
         Left            =   450
         TabIndex        =   35
         Top             =   3285
         Width           =   8715
      End
      Begin VB.ListBox lst_query 
         Height          =   1620
         Left            =   4680
         TabIndex        =   34
         Top             =   765
         Width           =   3480
      End
      Begin VB.TextBox txt_subloc 
         Height          =   288
         Left            =   4185
         MaxLength       =   1
         TabIndex        =   33
         Top             =   2340
         Width           =   300
      End
      Begin VB.TextBox txt_logwhse 
         Height          =   288
         Left            =   4185
         MaxLength       =   1
         TabIndex        =   32
         Top             =   1935
         Width           =   300
      End
      Begin VB.TextBox txt_Category 
         Height          =   288
         Left            =   4185
         MaxLength       =   1
         TabIndex        =   31
         Top             =   1530
         Width           =   300
      End
      Begin VB.TextBox txt_Warehouse 
         Height          =   288
         Left            =   4185
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1125
         Width           =   300
      End
      Begin VB.TextBox txt_Company 
         Height          =   288
         Left            =   4185
         MaxLength       =   1
         TabIndex        =   29
         Top             =   720
         Width           =   300
      End
      Begin VB.ComboBox cbo_Category 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1530
         Width           =   2175
      End
      Begin VB.ComboBox cbo_subloc 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2340
         Width           =   2175
      End
      Begin VB.ComboBox cbo_Logwhse 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1935
         Width           =   2175
      End
      Begin VB.ComboBox cbo_Company 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.Frame frm_DirList 
         Caption         =   "Filename"
         Height          =   2535
         Left            =   450
         TabIndex        =   16
         Top             =   3600
         Width           =   8700
         Begin VB.ComboBox cbo_pattern 
            Height          =   315
            Left            =   4140
            TabIndex        =   21
            Top             =   2115
            Width           =   4425
         End
         Begin VB.FileListBox fle_Stock 
            Height          =   1650
            Left            =   4095
            TabIndex        =   19
            Top             =   225
            Width           =   4515
         End
         Begin VB.TextBox txt_filename 
            Height          =   330
            Left            =   90
            TabIndex        =   18
            Top             =   2115
            Width           =   3990
         End
         Begin VB.DirListBox dir_stockdir 
            Height          =   1890
            Left            =   90
            TabIndex        =   17
            Top             =   180
            Width           =   3975
         End
      End
      Begin VB.ComboBox cbo_Warehouse 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1125
         Width           =   2175
      End
      Begin VB.Frame fra_recipients 
         Height          =   5640
         Left            =   -74730
         TabIndex        =   1
         Top             =   540
         Width           =   9240
         Begin VB.TextBox txt_Recipient 
            Height          =   288
            Left            =   1845
            TabIndex        =   10
            Top             =   3150
            Width           =   6108
         End
         Begin VB.CommandButton cmd_Remove 
            Caption         =   "Remove"
            Height          =   288
            Left            =   765
            TabIndex        =   9
            Top             =   2505
            Width           =   972
         End
         Begin VB.CommandButton cmd_Add 
            Caption         =   "Add"
            Height          =   288
            Left            =   765
            TabIndex        =   8
            Top             =   2130
            Width           =   972
         End
         Begin VB.ListBox lst_Destination 
            Height          =   2205
            Left            =   1845
            TabIndex        =   7
            Top             =   450
            Width           =   6684
         End
         Begin VB.Frame fra_FaxSelect 
            Height          =   1644
            Left            =   720
            TabIndex        =   3
            Top             =   3750
            Width           =   1356
            Begin VB.OptionButton opt_SupFax 
               Caption         =   "Supplier's"
               Height          =   288
               Left            =   96
               TabIndex        =   6
               Top             =   336
               Width           =   1020
            End
            Begin VB.OptionButton opt_FaxNum 
               Caption         =   "Fax Numbers"
               Height          =   330
               Left            =   96
               TabIndex        =   5
               Top             =   768
               Width           =   1080
            End
            Begin VB.OptionButton opt_Email 
               Caption         =   "Email"
               Height          =   288
               Left            =   96
               TabIndex        =   4
               Top             =   1260
               Width           =   684
            End
         End
         Begin VB.ListBox lst_Phonebook 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1635
            Left            =   2160
            TabIndex        =   2
            Top             =   3570
            Width           =   6396
         End
         Begin VB.Label lbl_Recipients 
            Caption         =   "Recipient(s)"
            Height          =   300
            Left            =   720
            TabIndex        =   12
            Top             =   450
            Width           =   1020
         End
         Begin VB.Label lbl_New 
            Caption         =   "New"
            Height          =   300
            Left            =   1350
            TabIndex        =   11
            Top             =   3225
            Width           =   390
         End
         Begin VB.Line Line1 
            X1              =   720
            X2              =   8736
            Y1              =   3030
            Y2              =   3030
         End
      End
      Begin VB.Label lbl_Descript 
         Caption         =   "Description"
         Height          =   225
         Left            =   4860
         TabIndex        =   38
         Top             =   2880
         Width           =   825
      End
      Begin VB.Label lbl_Currency 
         Caption         =   "Currency"
         Height          =   228
         Left            =   3360
         TabIndex        =   36
         Top             =   2880
         Width           =   732
      End
      Begin VB.Label lbl_Subloc 
         Caption         =   "Sub-Location"
         Height          =   225
         Left            =   495
         TabIndex        =   28
         Top             =   2385
         Width           =   960
      End
      Begin VB.Label lbl_logwhse 
         Caption         =   "Logical Warehouse"
         Height          =   225
         Left            =   90
         TabIndex        =   27
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lbl_category 
         Caption         =   "Category"
         Height          =   225
         Left            =   450
         TabIndex        =   26
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label lbl_filelen 
         Height          =   330
         Left            =   4500
         TabIndex        =   20
         Top             =   1755
         Width           =   2400
      End
      Begin VB.Label lbl_Warehouse 
         Caption         =   "Warehouse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   450
         TabIndex        =   15
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label lbl_Company 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   495
         TabIndex        =   13
         Top             =   765
         Width           =   870
      End
   End
End
Attribute VB_Name = "frm_StockonHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Public Keyvalid1 As Variant
'Public Keyvalid2 As Variant
'Public Keyvalid3 As Variant
'Public Keyvalid4 As Variant
'Public Keyvalid5 As Variant
'
'Public Counter As Integer
'
'Public Company As String
'Public Warehouse As String
'Public Category As String
'Public Logwhouse As String
'Public Subloc As String
'Public datachanged As Boolean
'Public ListNumber As Integer
'
'Public Function LoadList(a_cbo As ComboBox, a_tablename As String)
'    Dim ls_stockrecord As String
'    Dim ls_combo As String
'
'    VisM1.P0 = a_tablename
'    VisM1.Code = "d ^listkey"
'    VisM1.ExecFlag = 1
'
'    ls_combo = ""
'    ls_stockrecord = VisM1.P0
'    Do While ls_stockrecord <> ""
'        ls_combo = piece1(ls_stockrecord, ";")
'        a_cbo.AddItem (ls_combo)
'    Loop
'End Function
'
'Public Function LoadListbox(a_lst As ListBox, a_tablename As String)
'    Dim ls_stockrecord As String
'    Dim ls_combo As String
'
'    VisM1.P0 = a_tablename
'    VisM1.Code = "d ^listkey"
'    VisM1.ExecFlag = 1
'
'    ls_combo = ""
'    ls_stockrecord = VisM1.P0
'    Do While ls_stockrecord <> ""
'        ls_combo = piece1(ls_stockrecord, ";")
'        a_lst.AddItem (ls_combo)
'    Loop
'End Function
'
'  This function returns the form files found in the
'  current directory.  It also reads each file to get
'  the report description from the file
'
'Private Sub GetFormFiles(wildcard As String)
'    Dim Filename As String
'    Dim iFile As Integer
'    Dim FormHdr As TypeFormHdr
'
'    Filename = Dir(wildcard, 0)
'    iFile = 1                ' file number
'
'    Do While Filename <> ""
'
'       ' oprn and read the file header
'       Open Filename For Binary Access Read As #iFile
'
'       ErrorCode = 0
'       On Error GoTo ErrorGetFormFiles
'       Get #iFile, 1, FormHdr
'       On Error GoTo 0
'
'       If ErrorCode = 0 And TotalForms < MAX_FORMS Then   ' add to the list
'          If FormHdr.FormSign = FORM_SIGN Then
'            FormName(TotalForms) = FormHdr.name
'            FormFile(TotalForms) = Filename
'            TotalForms = TotalForms + 1
'          End If
'       End If
'
'       Close iFile
'
'       ' get the next file
'       Filename = Dir        ' get the next file
'
'    Loop
'
'    Exit Sub
'
'ErrorGetFormFiles:
'    ErrorCode = Err
'    Resume Next
'
'End Sub
'
'  This routine displays the report template available in
'  the current directory and lets use select a file to
'  edit.
'
'Private Sub GetFormSelection(Filename As String, FormEdit As Integer)
'
'    Filename = ""        ' initialize the file name
'    TotalForms = 0       ' total form files
'
'    Call GetFormFiles("*.FPC") ' accumulate the files with the .FPC extension
'
'    ' add new report option to form selection
'    If FormEdit Then
'       NewReport = TotalForms ' index of the new report option
'       TotalForms = TotalForms + 1
'       FormName(NewReport) = "New Report Form"
'       FormFile(NewReport) = ""
'    End If
'
'    FrmEdit.Show 1
'
'    If DlgResult >= 0 Then Filename = FormFile(DlgResult)
'
'End Sub
'
'  This routine reads each record from the sorted data file
'  (HEADER.SRT) and fills the values for the report fields
'  and calls the RvbRec function pass the data to the
'  report executor.
'
'Private Sub PrintRecords(RepParm As TypeRep)
'   Dim I As Integer
'   Dim X As Integer
'   Dim zeros As Integer
'   Dim CurLen As Integer
'   Dim iFile As Integer
'   Dim FileNo As Integer
'   Dim FieldNo As Integer
'   Dim RecoNo  As Integer
'   Dim DataSetName As String
'   Dim DataValue As Variant
'   Dim Filename As String
'   Dim temp As Integer
'   Dim field As TypeField
'   Dim MaxFiles As Integer
'   Dim FieldSource As Long
'   Dim FieldType As Long
'   Dim FileId As Long
'   Dim FieldId As Long
'   Dim linum As Integer
'   Dim maxlinum As Integer
'    Dim ls_record As String
'    Dim ls_code As String
'    Dim ls_data As String
'    Dim RTF As String
'    Dim SortOrder As String
'    Dim node1 As String
'    Dim node2 As String
'    Dim node3 As String
'    Dim node4 As String
'    Dim node5 As String
'    Dim GLOB As String
'
'   ' open the sorted data file
''   FileName = "temp" + Str(RepParm.ReportId) + ".SRT"
''
''   iFile = 1
''   Open FileName For Input Access Read As #iFile
''
''   ' Check if SALES data file is used
''   If SalesFileUsed Then MaxFiles = MAX_FILES Else
'   MaxFiles = 2
'    linum = 1
'
'   ' Retrieve the information about the fields used in this report
'   For I = 0 To RepParm.TotalFields - 1
'      temp = RvbGetDataField(Rep1.pCtl, I, field)
'      RepFieldSource(I) = field.source
'      RepFieldType(I) = field.type
'      RepFieldId(I) = field.FieldId
'      RepFileId(I) = field.FileId
'   Next I
'
''    If txt_selected.Text = "" Then txt_selected.Text = "~~~~~"
''    If Warehouse = "" Then
'''        VisM1.P0 = ""
'''        VisM1.P3 = "BASE"
'''        VisM1.P4 = 13
'''        VisM1.P2 = "WH" & gs_UT & "LOC"
'''        VisM1.Code = "S P0=$$^ffindbase(P2,P3,P4)"
'''        VisM1.ExecFlag = 1
'''        Warehouse = VisM1.P0
'''
'''        VisM1.P0 = ""
'''        VisM1.P3 = "SITE"
'''        VisM1.P4 = 13
'''        VisM1.P2 = "WH" & gs_UT & "LOC"
'''        VisM1.Code = "S P0=$$^ffindbase(P2,P3,P4)"
'''        VisM1.ExecFlag = 1
'''
'''        Warehouse = Warehouse & VisM1.P0
'''        txt_selected.Text = Company & "~" & Warehouse & "~" & Category & "~" & Logwhouse & "~" & Subloc & "~"
''    End If
''    VisM1.P3 = txt_selected.Text
''    VisM1.P4 = Trim(txt_Company.Text) & Trim(txt_Warehouse.Text) & Trim(txt_Category.Text) & Trim(txt_logwhse.Text) & Trim(txt_subloc.Text)
'''    VisM1.P7 = "USD"
''    VisM1.P7 = txt_Currency.Text
''
''    ls_code = "d ^onhnd(P3,P4)"
'''    ls_code = "d ^onhndml(P3,P4)"
''    VisM1.P0 = ""
'''    Debug.Print "ls_code = " & ls_code
'''    Debug.Print "string1 = " & txt_selected.text
'''    Debug.Print "string2 = " & VisM1.P4
''    VisM1.Code = ls_code
''    VisM1.ExecFlag = 1
''
''    ls_code = "S P0=$Q(^ONHAND)"
''    VisM1.P0 = ""
'''    Debug.Print "ls_code = " & ls_code
''    VisM1.Code = ls_code
''    VisM1.ExecFlag = 1
''    ls_record = VisM1.P0
''    GLOB = ls_record
''    Debug.Print "ls_record = " & ls_record
'
'    zeros = 0
'    If txt_Company = "0" Then zeros = zeros + 1
'    If txt_Warehouse = "0" Then zeros = zeros + 1
'    If txt_Category = "0" Then zeros = zeros + 1
'    If txt_logwhse = "0" Then zeros = zeros + 1
'    If txt_subloc = "0" Then zeros = zeros + 1
'
'    SortOrder = ""
'    For X = 1 To (5 - zeros)
'        If txt_Company = X Then SortOrder = SortOrder & "Company      "
'        If txt_Warehouse = X Then SortOrder = SortOrder & "Warehouse    "
'        If txt_Category = X Then SortOrder = SortOrder & "Category     "
'        If txt_logwhse = X Then SortOrder = SortOrder & "Logical W.H. "
'        If txt_subloc = X Then SortOrder = SortOrder & "Sub-Location "
'    Next X
'    SortOrder = Mid$(SortOrder, 1, Len(SortOrder) - 1)
'
''    Select Case zeros
''        Case 0
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node4 = piece1(ls_data, DoubleQuote())
''            node4 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node5 = piece1(ls_data, DoubleQuote())
''            node5 = piece1(ls_data, DoubleQuote())
''        Case 1
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node4 = piece1(ls_data, DoubleQuote())
''            node4 = piece1(ls_data, DoubleQuote())
''            node5 = ""
''        Case 2
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            node4 = ""
''            node5 = ""
''        Case 3
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            node3 = ""
''            node4 = ""
''            node5 = ""
''        Case 4
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            node2 = ""
''            node3 = ""
''            node4 = ""
''            node5 = ""
''        Case 5
''        Case Else
''    End Select
'''    Debug.Print "nodes = " & node1, node2, node3, node4, node5
'
'READ_RECORD:
'   ' Initialize the field structure
'   For I = 0 To RepParm.TotalFields - 1
'      If RepFieldSource(I) = SRC_APPL Then
'        If RepFieldType(I) = TYPE_TEXT Then temp = RvbSetTextField(Rep1.pCtl, I, "", 0)
'        If RepFieldType(I) = TYPE_NUM Then temp = RvbSetNumField(Rep1.pCtl, I, (0))
'        If RepFieldType(I) = TYPE_DBL Then temp = RvbSetNumField(Rep1.pCtl, I, (0))
'      End If
'   Next I
'
''    'Call function to load up Header info into a temp
''    'string. depiece the thing in READ RECORD
''    VisM1.P0 = ""
'''    VisM1.P1 = cbo_Purchase.text
''    VisM1.Code = "d ^loadupstockhand"
''    VisM1.ExecFlag = 1
''    ls_record = VisM1.P0
''
''    'Call function to load up item info into a temp
''    'string. depiece the thing in READ RECORD
''    VisM1.P0 = ""
''    ls_code = "S P0=" & GLOB
''    VisM1.P0 = ""
''    VisM1.Code = ls_code
'''    Debug.Print "Vism1.code = " & VisM1.Code
''    VisM1.ExecFlag = 1
''    ls_code = VisM1.P0
'''    Debug.Print "ls_code = " & ls_code
''
''   ' read fields values for the next record
''   For FileNo = 0 To MaxFiles - 1
'''        Debug.Print "total fields = " & DataFile(FileNo).TotalFields
''     For FieldNo = 0 To DataFile(FileNo).TotalFields - 1
''       ErrorCode = 0
''       On Error GoTo PrintRecordsHandler
'''       Input #iFile, DataValue 'get information from screens
''        Select Case FileNo
''            Case 0
''                Select Case FieldNo
''                    Case 1
''                        DataValue = txt_Currency.Text
''                    Case 2
''                        DataValue = SortOrder
''                    Case 3
''                        DataValue = node1
''                    Case 4
''                        DataValue = node2
''                    Case 5
''                        DataValue = node3
''                    Case 6
''                        DataValue = node4
''                    Case 7
''                        DataValue = node5
''                    Case Else
''                        DataValue = piece1(ls_record, "~")
''                End Select
''            Case 1
''                Select Case FieldNo
''                    Case 0
''                        DataValue = piece(ls_code, "~", 2)
''                    Case 1
''                        If opt_Long.Value = True Then DataValue = piece(ls_code, "~", 8) Else DataValue = piece(ls_code, "~", 7)
''                    Case 2
''                        DataValue = piece(ls_code, "~", 8)
''                    Case 3
''                        DataValue = piece(ls_code, "~", 1)
''                    Case 4
''                        DataValue = piece(ls_code, "~", 3)
''                    Case 5
''                        DataValue = piece(ls_code, "~", 5)
''                    Case 6
''                        DataValue = piece(ls_code, "~", 4)
''                    Case 7
''                        DataValue = piece(ls_code, "~", 6)
''                    Case 8
''                        If Company = "" Then DataValue = "Company           : " & "ALL" Else DataValue = "Company           : " & Company
''                    Case 9
''                        If Warehouse = "" Then DataValue = "Warehouse         : " & "ALL" Else DataValue = "Warehouse         : " & Warehouse
''                    Case 10
''                        If Category = "" Then DataValue = "Category          : " & "ALL" Else DataValue = "Category          : " & Category
''                    Case 11
''                        If Logwhouse = "" Then DataValue = "Logical Warehouse : " & "ALL" Else DataValue = "Logical Warehouse : " & Logwhouse
''                    Case 12
''                        If Subloc = "" Then DataValue = "Sub-Location      : " & "ALL" Else DataValue = "Sub-Location      : " & Subloc
''                    Case 13
''                        DataValue = 0
''                    Case Else
''                        DataValue = ""
''                End Select
''        End Select
''       On Error GoTo 0
''       If ErrorCode Then
''          Close iFile
''          Exit Sub
''       End If
''
''       ' Pass this data field to the VBX
''       For i = 0 To RepParm.TotalFields - 1
''          If RepFieldSource(i) = SRC_APPL And RepFileId(i) = FileNo And RepFieldId(i) = FieldNo Then
''             If RepFieldType(i) = TYPE_TEXT Then temp = RvbSetTextField(Rep1.pCtl, i, DataValue, Len(DataValue))
''             If RepFieldType(i) = TYPE_NUM Then temp = RvbSetNumField(Rep1.pCtl, i, DataValue)
''             If RepFieldType(i) = TYPE_DBL Then temp = RvbSetDoubleField(Rep1.pCtl, i, DataValue)
''             If RepFieldType(i) = TYPE_DATE Then temp = RvbSetNumField(Rep1.pCtl, i, DataValue)
''             If RepFieldType(i) = TYPE_PICT Then
''                temp = RvbSetNumField(Rep1.pCtl, i, DataValue)  ' pass picture id
''                'Debug.Print "g_pict = " & g_Pict
''             End If
''             If RepFieldType(i) = TYPE_LOGICAL Then
''                If DataValue = "Y" Or DataValue = "y" Then DataValue = 1 Else DataValue = 0
''                temp = RvbSetNumField(Rep1.pCtl, i, DataValue)
''             End If
''          End If
''       Next i
''     Next FieldNo
''   Next FileNo
'
'   ' Print the record
'   temp = RvbRec(Rep1.pCtl)
'   If temp <> 0 Then     ' user aborted the report
'     Close iFile
'     Exit Sub
'   End If
'''    LiNum = LiNum + 1
'''    linum = VisM1.P2
'''    Debug.Print "linum = " & linum
''    'counter for item number
''    'then exit sub
'''    If LiNum = 1 Then Exit Sub
''
''    ls_code = "S P0=$Q(" & GLOB & ")"
''    VisM1.P0 = ""
'''    Debug.Print "ls_code = " & ls_code
''    VisM1.Code = ls_code
''    VisM1.ExecFlag = 1
''    ls_record = VisM1.P0
''    GLOB = ls_record
''
''    Select Case zeros
''        Case 0
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node4 = piece1(ls_data, DoubleQuote())
''            node4 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node5 = piece1(ls_data, DoubleQuote())
''            node5 = piece1(ls_data, DoubleQuote())
''        Case 1
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node4 = piece1(ls_data, DoubleQuote())
''            node4 = piece1(ls_data, DoubleQuote())
''            node5 = ""
''        Case 2
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node3 = piece1(ls_data, DoubleQuote())
''            node3 = piece1(ls_data, DoubleQuote())
''            node4 = ""
''            node5 = ""
''        Case 3
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            ls_data = piece1(ls_record, ",")
''            node2 = piece1(ls_data, DoubleQuote())
''            node2 = piece1(ls_data, DoubleQuote())
''            node3 = ""
''            node4 = ""
''            node5 = ""
''        Case 4
''            ls_data = piece1(ls_record, ",")
''            node1 = piece1(ls_data, DoubleQuote())
''            node1 = piece1(ls_data, DoubleQuote())
''            node2 = ""
''            node3 = ""
''            node4 = ""
''            node5 = ""
''        Case 5
''        Case Else
''    End Select
''
''    If Trim(GLOB) = "" Then
''        txt_selected.Text = ""
''        Company = ""
''        Warehouse = ""
''        Category = ""
''        Logwhouse = ""
''        Subloc = ""
''        Exit Sub
''    End If
'''    Exit Sub
''
''   GoTo READ_RECORD
'
'PrintRecordsHandler:
'   ErrorCode = Err
'   Resume Next
'
'End Sub
'
''  This routine reads the field definition file for the
''  specified file and extracts the individual field names
''  and field properties.
''
'Private Sub ReadFields(FileNo As Integer)
'   Dim Filename As String
'   Dim TextLine As String
'   Dim CharReturn As String
'   Dim CurLen As Integer
'   Dim FieldNo As Integer
'   Dim LineIdx As Integer
'   Dim FieldType As String
'   Dim iFile As Integer
'
'
'   FieldNo = 0
'   DataFile(FileNo).TotalFields = 0
'
'   ' open the field definition file
''   Filename = App.Path & "\" & DataFile(FileNo).name + ".DF"  'field definition file added app.path ML 2/17/99
'''   Filename = DataFile(FileNo).name + ".DF"  'field definition file added app.path ML 2/17/99
''   iFile = 1
'''   Debug.Print "filename = " & Filename
''   Open Filename For Input Access Read As #iFile
'
'READ_LINE:
'
'    ErrorCode = 0
'    On Error GoTo ErrorHandler
'    Input #iFile, DataField(FileNo, FieldNo).ShortName, DataField(FileNo, FieldNo).width, FieldType, DataField(FileNo, FieldNo).DecPlaces
'    On Error GoTo 0
'    If ErrorCode Then
'       Close iFile
'       Exit Sub
'    End If
'
'    ' construct other fields
'    DataField(FileNo, FieldNo).FullName = DataFile(FileNo).name + "->" + DataField(FileNo, FieldNo).ShortName
'
'    If FieldType = "T" Then DataField(FileNo, FieldNo).type = TYPE_TEXT
'    If FieldType = "N" Then DataField(FileNo, FieldNo).type = TYPE_NUM
'    If FieldType = "F" Then DataField(FileNo, FieldNo).type = TYPE_DBL
'    If FieldType = "D" Then DataField(FileNo, FieldNo).type = TYPE_DATE
'    If FieldType = "L" Then DataField(FileNo, FieldNo).type = TYPE_LOGICAL
'    If FieldType = "P" Then DataField(FileNo, FieldNo).type = TYPE_PICT
'
'    ' increment the field no
'    FieldNo = FieldNo + 1
'    DataFile(FileNo).TotalFields = FieldNo
'
'    GoTo READ_LINE
'
'ErrorHandler:
'    ErrorCode = Err
'    Resume Next
'
'End Sub
'
'Private Sub cbo_Category_Change()
'    If cbo_Category.Text = "ONE" Or cbo_Category.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "MUD"
'        lst_query.AddItem "TUB"
'        ListNumber = 3
'    End If
'End Sub
'
'Private Sub cbo_Category_Click()
'    If cbo_Category.Text = "ONE" Or cbo_Category.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "MUD"
'        lst_query.AddItem "TUB"
'        ListNumber = 3
'    End If
'End Sub
'
'Private Sub cbo_Company_Change()
'    If cbo_Company.Text = "ONE" Or cbo_Company.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "EXXONEFTE"
'        ListNumber = 1
'    End If
'End Sub
'
'Private Sub cbo_Company_Click()
'    If cbo_Company.Text = "ONE" Or cbo_Company.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "EXXONEFTE"
'        ListNumber = 1
'    End If
'End Sub
'
'Private Sub cbo_Logwhse_Change()
'    If cbo_Logwhse.Text = "ONE" Or cbo_Logwhse.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "GENERAL"
'        lst_query.AddItem "STOCK"
'        ListNumber = 4
'    End If
'End Sub
'
'Private Sub cbo_Logwhse_Click()
'    If cbo_Logwhse.Text = "ONE" Or cbo_Logwhse.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "GENERAL"
'        lst_query.AddItem "STOCK"
'        ListNumber = 4
'    End If
'End Sub
'
'Private Sub cbo_pattern_KeyPress(KeyAscii As Integer)
'    If (KeyAscii = vbKeyReturn) Then
'        fle_Stock.Pattern = cbo_pattern.Text
'    End If
'End Sub
'
'Private Sub cbo_subloc_Change()
'    If cbo_subloc.Text = "ONE" Or cbo_subloc.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "WHSE 5"
'        lst_query.AddItem "WHSE 6"
'        ListNumber = 5
'    End If
'End Sub
'
'Private Sub cbo_subloc_Click()
'    If cbo_subloc.Text = "ONE" Or cbo_subloc.Text = "SEVERAL" Then
'        lst_query.Clear
'        lst_query.AddItem "WHSE 5"
'        lst_query.AddItem "WHSE 6"
'        ListNumber = 5
'    End If
'End Sub
'
'Private Sub cbo_warehouse_Change()
'    Dim ls_data As String
'    Dim ls_combo As String
'
'    If cbo_Warehouse.Text = "ONE" Or cbo_Warehouse.Text = "SEVERAL" Then
'        lst_query.Clear
'        Call LoadListbox(lst_query, "WHPECOM")
'        jjkjlkjl
'        lst_query.AddItem "DAGI13"
'        lst_query.AddItem "DAGI15"
'        ListNumber = 2
'    End If
'
'End Sub
'
'Private Sub cbo_Warehouse_Click()
'    Dim ls_data As String
'    Dim ls_combo As String
'    If cbo_Warehouse.Text = "ONE" Or cbo_Warehouse.Text = "SEVERAL" Then
'        lst_query.Clear
'        Call LoadListbox(lst_query, "WHPECOM")
'        jjkjlkjl
'        lst_query.AddItem "DAGI13"
'        lst_query.AddItem "DAGI15"
'        ListNumber = 2
'    End If
'
'End Sub
'
'Private Sub cmd_Send_Click()
'    Dim message As String
'    Dim file_err As Integer
'   Dim temp As Integer
'   Dim Filename As String
'   Dim EditingForm As Integer
'
'   ' check if ReportEase already active, only one session per control allowed
'   If FormParm.open Then
'      MsgBox "A Form Editor Window Already Open!"
'      Exit Sub
'   End If
'   FormParm.open = True
'
'   EditingForm = False
''   Call GetFormSelection(Filename, EditingForm) 'select a report template to edit
''
''   If DlgResult < 0 Then Exit Sub     ' a form not selected
'
'    Filename = FixDir(App.Path) & "REreport\STOCKHAND1.FPC"
'   ' fill the report parameter structure
'   RepParm.file = Filename + Chr(0)  ' specify the file name
''   RepParm.device = "P"              ' output device, P = Printer
''   RepParm.device = "S"              ' output device, S = Screen
''   RepParm.device = "T"              ' output device, T = Text
'   RepParm.device = "R"              ' output device, R = Rich text file
'   RepParm.X = FormParm.X            ' output window location
'   RepParm.y = FormParm.y
'   RepParm.width = FormParm.width    ' output window width
'   RepParm.height = FormParm.height  ' output window height
'   RepParm.OutFile = "c:\attmsg\out\mikeTest.rtf" + Chr(0)  'for email
'
'   ' Initialize the report executor
'   temp = RvbInit(Rep1.pCtl, RepParm)          ' initialize the report
'   If temp <> 0 Then Exit Sub
'
'   ' prepare the datafile for output
''   Call PrepareFile(RepParm)
'
'   ' Feed the records to the report executor
'   Call PrintRecords(RepParm)
'
'   ' Close the report executor
'   temp = RvbExit(Rep1.pCtl)
'
'   FormParm.open = False
'
'    message = "RTF"
''    file_err = send_mail("C:\ATTMSG\OUT\test.msq", lst_Destination, "TEST", "T")
'    'file_err = send_mail("C:\ATTMSG\OUT\test.msq", lst_Destination, "Stock on Hand", message)
'    'If file_err = 0 Then MsgBox (gs_Message(55))
'End Sub
'
'Private Sub cmd_Fax_Click()
'   Dim message As String
'   Dim file_err As Integer
'   Dim temp As Integer
'   Dim Filename As String
'   Dim EditingForm As Integer
'
'   ' check if ReportEase already active, only one session per control allowed
'   If FormParm.open Then
'      MsgBox "A Form Editor Window Already Open!"
'      Exit Sub
'   End If
'   FormParm.open = True
'
'   EditingForm = False
''   Call GetFormSelection(Filename, EditingForm) 'select a report template to edit
''
''   If DlgResult < 0 Then Exit Sub     ' a form not selected
'
'    Filename = FixDir(App.Path) & "REreport\STOCKHAND4.FPC"
'   ' fill the report parameter structure
'   RepParm.file = Filename + Chr(0)  ' specify the file name
'   RepParm.device = "P"              ' output device, P = Printer
''   RepParm.device = "S"              ' output device, S = Screen
''   RepParm.device = "T"              ' output device, T = Text
''   RepParm.device = "R"              ' output device, R = Rich text file
'   RepParm.X = FormParm.X            ' output window location
'   RepParm.y = FormParm.y
'   RepParm.width = FormParm.width    ' output window width
'   RepParm.height = FormParm.height  ' output window height
'   RepParm.OutFile = "c:\attmsg\out\mikeTest.rtf" + Chr(0)  'for email
'
'   ' Initialize the report executor
'   temp = RvbInit(Rep1.pCtl, RepParm)          ' initialize the report
'   If temp <> 0 Then Exit Sub
'
'   ' prepare the datafile for output
''   Call PrepareFile(RepParm)
'
'   ' Feed the records to the report executor
'   Call PrintRecords(RepParm)
'
'   ' Close the report executor
'   temp = RvbExit(Rep1.pCtl)
'
'   FormParm.open = False
'
''    message = "RTF"
'''    file_err = send_mail("C:\ATTMSG\OUT\test.msq", lst_Destination, "TEST", "T")
''    file_err = send_mail("C:\ATTMSG\OUT\test.msq", lst_Destination, "Stock on Hand", message)
''    If file_err = 0 Then MsgBox (gs_Message(55))
'End Sub
'
'Private Sub dir_stockdir_Click()
'    fle_Stock.Path = dir_stockdir.List(dir_stockdir.ListIndex)
'    txt_filename.Text = dir_stockdir.List(dir_stockdir.ListIndex)
'End Sub
'
'Private Sub fle_Stock_Click()
'    If dir_stockdir.List(dir_stockdir.ListIndex) <> "C:\" Then
'        txt_filename.Text = dir_stockdir.List(dir_stockdir.ListIndex) & "\" & fle_Stock.List(fle_Stock.ListIndex)
'    Else:
'        txt_filename.Text = dir_stockdir.List(dir_stockdir.ListIndex) & fle_Stock.List(fle_Stock.ListIndex)
'    End If
'End Sub
'
'Private Sub fle_Stock_DblClick()
'        If Len(Dir(txt_filename.Text)) <= 0 Then lbl_filelen.Caption = "file does not exist"
'        If Len(Dir(txt_filename.Text)) > 0 Then Load (frm_Filedisplay)
'End Sub
'
'this subroutine makes an ascii delimited text file
'for the stock on hand report
'Private Sub DOStxt()
'    Dim ls_code As String
'    Dim ls_record As String
'    Dim ls_data As String
'    Dim GLOB As String
'    Dim zeros As String
'    Dim node1 As String
'    Dim node2 As String
'    Dim node3 As String
'    Dim node4 As String
'    Dim node5 As String
'    Dim lastnode1 As String
'    Dim lastnode2 As String
'    Dim lastnode3 As String
'    Dim lastnode4 As String
'    Dim lastnode5 As String
'    Dim total1 As Long
'    Dim total2 As Long
'    Dim total3 As Long
'    Dim total4 As Long
'    Dim total5 As Long
'
'    Open txt_filename.Text For Output As #1
'
'    zeros = 0
'    If txt_Company = "0" Then zeros = zeros + 1
'    If txt_Warehouse = "0" Then zeros = zeros + 1
'    If txt_Category = "0" Then zeros = zeros + 1
'    If txt_logwhse = "0" Then zeros = zeros + 1
'    If txt_subloc = "0" Then zeros = zeros + 1
'
'    ls_code = "S P0=$Q(^ONHAND)"
'    VisM1.P0 = ""
''    Debug.Print "ls_code = " & ls_code
'    VisM1.Code = ls_code
'    VisM1.ExecFlag = 1
'    ls_record = VisM1.P0
'    GLOB = ls_record
''    Debug.Print "ls_record = " & ls_record
'    ls_code = "S P0=" & ls_record
'    VisM1.P0 = ""
'    VisM1.Code = ls_code
''    Debug.Print "Vism1.code = " & VisM1.Code
'    VisM1.ExecFlag = 1
'    ls_code = VisM1.P0
''    Debug.Print "ls_code = " & ls_code
'    lastnode1 = ""
'    lastnode2 = ""
'    lastnode3 = ""
'    lastnode4 = ""
'    lastnode5 = ""
'    total1 = 0
'    total2 = 0
'    total3 = 0
'    total4 = 0
'    total5 = 0
'
'    Select Case zeros
'        Case 0
'            ls_data = piece1(ls_record, ",")
'            node1 = piece1(ls_data, DoubleQuote())
'            node1 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node2 = piece1(ls_data, DoubleQuote())
'            node2 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node3 = piece1(ls_data, DoubleQuote())
'            node3 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node4 = piece1(ls_data, DoubleQuote())
'            node4 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node5 = piece1(ls_data, DoubleQuote())
'            node5 = piece1(ls_data, DoubleQuote())
'        Case 1
'            ls_data = piece1(ls_record, ",")
'            node1 = piece1(ls_data, DoubleQuote())
'            node1 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node2 = piece1(ls_data, DoubleQuote())
'            node2 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node3 = piece1(ls_data, DoubleQuote())
'            node3 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node4 = piece1(ls_data, DoubleQuote())
'            node4 = piece1(ls_data, DoubleQuote())
'            node5 = ""
'        Case 2
'            ls_data = piece1(ls_record, ",")
'            node1 = piece1(ls_data, DoubleQuote())
'            node1 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node2 = piece1(ls_data, DoubleQuote())
'            node2 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node3 = piece1(ls_data, DoubleQuote())
'            node3 = piece1(ls_data, DoubleQuote())
'            node4 = ""
'            node5 = ""
'        Case 3
'            ls_data = piece1(ls_record, ",")
'            node1 = piece1(ls_data, DoubleQuote())
'            node1 = piece1(ls_data, DoubleQuote())
'            ls_data = piece1(ls_record, ",")
'            node2 = piece1(ls_data, DoubleQuote())
'            node2 = piece1(ls_data, DoubleQuote())
'            node3 = ""
'            node4 = ""
'            node5 = ""
'        Case 4
'            ls_data = piece1(ls_record, ",")
'            node1 = piece1(ls_data, DoubleQuote())
'            node1 = piece1(ls_data, DoubleQuote())
'            node2 = ""
'            node3 = ""
'            node4 = ""
'            node5 = ""
'        Case 5
'        Case Else
'    End Select
'
'    lastnode1 = node1
'    lastnode2 = node2
'    lastnode3 = node3
'    lastnode4 = node4
'    lastnode5 = node5
'
'    While ls_record <> ""
'        Print #1, node1 & "," & node2 & "," & node3 & "," & node4 & "," & node5 & "~" & ls_code
'        ls_code = "S P0=$Q(" & GLOB & ")"
'        VisM1.P0 = ""
'    '    Debug.Print "ls_code = " & ls_code
'        VisM1.Code = ls_code
'        VisM1.ExecFlag = 1
'        ls_record = VisM1.P0
'        GLOB = ls_record
'    '    Debug.Print "ls_record = " & ls_record
'        VisM1.P0 = ""
'        ls_code = "S P0=" & ls_record
'        VisM1.P0 = ""
'        VisM1.Code = ls_code
'    '    Debug.Print "Vism1.code = " & VisM1.Code
'        VisM1.ExecFlag = 1
'        ls_code = VisM1.P0
'    '    Debug.Print "ls_code = " & ls_code
'        Select Case zeros
'            Case 0
'                ls_data = piece1(ls_record, ",")
'                node1 = piece1(ls_data, DoubleQuote())
'                node1 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node2 = piece1(ls_data, DoubleQuote())
'                node2 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node3 = piece1(ls_data, DoubleQuote())
'                node3 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node4 = piece1(ls_data, DoubleQuote())
'                node4 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node5 = piece1(ls_data, DoubleQuote())
'                node5 = piece1(ls_data, DoubleQuote())
'            Case 1
'                ls_data = piece1(ls_record, ",")
'                node1 = piece1(ls_data, DoubleQuote())
'                node1 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node2 = piece1(ls_data, DoubleQuote())
'                node2 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node3 = piece1(ls_data, DoubleQuote())
'                node3 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node4 = piece1(ls_data, DoubleQuote())
'                node4 = piece1(ls_data, DoubleQuote())
'                node5 = ""
'            Case 2
'                ls_data = piece1(ls_record, ",")
'                node1 = piece1(ls_data, DoubleQuote())
'                node1 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node2 = piece1(ls_data, DoubleQuote())
'                node2 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node3 = piece1(ls_data, DoubleQuote())
'                node3 = piece1(ls_data, DoubleQuote())
'                node4 = ""
'                node5 = ""
'            Case 3
'                ls_data = piece1(ls_record, ",")
'                node1 = piece1(ls_data, DoubleQuote())
'                node1 = piece1(ls_data, DoubleQuote())
'                ls_data = piece1(ls_record, ",")
'                node2 = piece1(ls_data, DoubleQuote())
'                node2 = piece1(ls_data, DoubleQuote())
'                node3 = ""
'                node4 = ""
'                node5 = ""
'            Case 4
'                ls_data = piece1(ls_record, ",")
'                node1 = piece1(ls_data, DoubleQuote())
'                node1 = piece1(ls_data, DoubleQuote())
'                node2 = ""
'                node3 = ""
'                node4 = ""
'                node5 = ""
'            Case 5
'            Case Else
'        End Select
'        If lastnode1 <> node1 Then
'            Print #1, "TOTAL " & lastnode1 & "," & node2 & "," & node3 & "," & node4 & "," & node5 & " = $ " & Format(total1, "###,###,###,##0.00")
'            total1 = 0
'        End If
'        If lastnode2 <> node2 Then
'            Print #1, "TOTAL " & node1 & "," & lastnode2 & "," & node3 & "," & node4 & "," & node5 & " = $ " & Format(total2, "###,###,###,##0.00")
'            total2 = 0
'        End If
'        If lastnode3 <> node3 Then
'            Print #1, "TOTAL " & node1 & "," & node2 & "," & lastnode3 & "," & node4 & "," & node5 & " = $ " & Format(total3, "###,###,###,##0.00")
'            total3 = 0
'        End If
'        If lastnode4 <> node4 Then
'            Print #1, "TOTAL " & node1 & "," & node2 & "," & node3 & "," & lastnode4 & "," & node5 & " = $ " & Format(total4, "###,###,###,##0.00")
'            total4 = 0
'        End If
'        If lastnode5 <> node5 Then
'            Print #1, "TOTAL " & node1 & "," & node2 & "," & node3 & "," & node4 & "," & lastnode5 & " = $ " & Format(total5, "###,###,###,##0.00")
'            total5 = 0
'        End If
'        lastnode1 = node1
'        lastnode2 = node2
'        lastnode3 = node3
'        lastnode4 = node4
'        lastnode5 = node5
'        total1 = total1 + Val(piece(ls_code, "~", 9))
'        total2 = total2 + Val(piece(ls_code, "~", 9))
'        total3 = total3 + Val(piece(ls_code, "~", 9))
'        total4 = total4 + Val(piece(ls_code, "~", 9))
'        total5 = total5 + Val(piece(ls_code, "~", 9))
'    Wend
'
'    Close #1
'End Sub
'
'Private Sub cmd_Print_Click()
'   Dim temp As Integer
'   Dim I As Integer
'   Dim Filename As String
'   Dim EditingForm As Integer
'   Dim oPrinter As Printer
'
'    If Trim(txt_filename) <> "" Then
'        Call DOStxt
'        Exit Sub
'    End If
'
'   ' check if ReportEase already active, only one session per control allowed
'   If FormParm.open Then
'      MsgBox "A Form Editor Window Already Open!"
'      Exit Sub
'   End If
'   FormParm.open = True
'
'   EditingForm = False
''   Call GetFormSelection(Filename, EditingForm) 'select a report template to edit
'
''   If DlgResult < 0 Then Exit Sub     ' a form not selected
'
'   ' fill the report parameter structure
'   Filename = FixDir(App.Path) & "REreport\STOCKHAND5.FPC"
'   RepParm.file = Filename + Chr(0)  ' specify the file name
''   RepParm.device = "P"              ' output device, P = Printer
''   RepParm.device = "R"              ' output device, R = Rich text file
'   RepParm.device = "S"              ' output device, S = Screen
''   RepParm.device = "T"              ' output device, T = Text
'   RepParm.X = FormParm.X            ' output window location
'   RepParm.y = FormParm.y
''   RepParm.width = FormParm.width    ' output window width
''   RepParm.height = FormParm.height  ' output window height
'   RepParm.width = 800    ' output window width
'   RepParm.height = 600  ' output window height
''   RepParm.OutFile = "mikeTest.rtf" + Chr(0)  'for email
'
'   ' Initialize the report executor
'   temp = RvbInit(Rep1.pCtl, RepParm)          ' initialize the report
'   If temp <> 0 Then Exit Sub
'
'   ' prepare the datafile for output
''   Call PrepareFile(RepParm)
'
'   ' Feed the records to the report executor
'   Call PrintRecords(RepParm)
'
''    Debug.Print Printer.DriverName
''    Debug.Print Printer.hDC
'
''    For Each oPrinter In Printers
''         If oPrinters.DeviceName = "AT&T Fax Sender" Then Set Printer = oPrinter
''    Next
'
'
'   ' Close the report executor
'   temp = RvbExit(Rep1.pCtl)
'   FormParm.open = False
'
'End Sub
'
'Private Sub Form_Load()
'    Dim I As Integer
'    Dim li_x As Integer
'    Dim ls_data As String
'    Dim ls_combo As String
'    VisM1.NameSpace = gs_Namespace
'
'    Company = ""
'    Warehouse = ""
'    Category = ""
'    Logwhouse = ""
'    Subloc = ""
'
'    fle_Stock.Path = App.Path
'    dir_stockdir.Path = App.Path
'    cbo_pattern.AddItem " "
'    cbo_pattern.AddItem "*.txt"
'    cbo_pattern.AddItem "*.rpt"
'    cbo_pattern.AddItem "*.xls"
'
'    cbo_Company.AddItem "ONE"
'    cbo_Company.AddItem "SEVERAL"
'    cbo_Company.AddItem "ALL"
'
'    cbo_Warehouse.AddItem "ONE"
'    cbo_Warehouse.AddItem "SEVERAL"
'    cbo_Warehouse.AddItem "ALL"
'
'    cbo_Category.AddItem "ONE"
'    cbo_Category.AddItem "SEVERAL"
'    cbo_Category.AddItem "ALL"
'
'    cbo_Logwhse.AddItem "ONE"
'    cbo_Logwhse.AddItem "SEVERAL"
'    cbo_Logwhse.AddItem "ALL"
'
'    cbo_subloc.AddItem "ONE"
'    cbo_subloc.AddItem "SEVERAL"
'    cbo_subloc.AddItem "ALL"
'
'    cbo_Company.ListIndex = 2
'    cbo_Warehouse.ListIndex = 2
'    cbo_Category.ListIndex = 2
'    cbo_Logwhse.ListIndex = 2
'    cbo_subloc.ListIndex = 2
'
'    txt_Company.Text = "1"
'    txt_Warehouse.Text = "2"
'    txt_Category.Text = "3"
'    txt_logwhse.Text = "4"
'    txt_subloc.Text = "5"
'
'    Keyvalid1 = Array("0", "1", "2", "3", "4", "5")
'    Keyvalid2 = Array("0", "1", "2", "3", "4", "5")
'    Keyvalid3 = Array("0", "1", "2", "3", "4", "5")
'    Keyvalid4 = Array("0", "1", "2", "3", "4", "5")
'    Keyvalid5 = Array("0", "1", "2", "3", "4", "5")
'
'    opt_Long.Value = True 'set the long description as default
'    txt_Currency.Text = "USD" 'set the currency as US Dollars as default
'
'     Initialize the report input parameter structure
'    RepParm.hInst = 0
'    RepParm.hPrevInst = 0
'    RepParm.hParentWnd = hWnd           ' window handle for the form which contains the roc control
'    RepParm.style = WS_OVERLAPPEDWINDOW ' window style
'    RepParm.SwapDir = Chr(0)            ' default swap directory
'    RepParm.SuppressPrintMessages = False ' show print messages
'    RepParm.UseCurrentPrinter = False ' use the printer specified in the form
'    RepParm.OutFile = Chr(0)          ' use default
'
'     Initialize the report viewer parameter structure
'    ViewParm.file = Chr(0)              ' let the API do the file selection
'    ViewParm.device = "S"
'    ViewParm.X = FormParm.X            ' output window location
'    ViewParm.y = FormParm.y
'    ViewParm.width = FormParm.width    ' output window width
'    ViewParm.height = FormParm.height  ' output window height
'    ViewParm.hInst = 0
'    ViewParm.hPrevInst = 0
'    ViewParm.hParentWnd = hWnd         ' window handle for the form which contains the roc control
'    ViewParm.style = WS_OVERLAPPEDWINDOW ' window style
'    ViewParm.SwapDir = Chr(0)            ' default swap directory
'    ViewParm.SuppressPrintMessages = False ' show print messages
'
'     set the demo data file names and read the field names
'    DataFile(0).name = "STOCKHAND1"
'    DataFile(1).name = "STOCKITEM1"
'
'    For I = 0 To MAX_FILES - 1
'      ReadFields (I)
'    Next I
'
'     load the HEADER logo bitmap
'    Picture1.Picture = LoadPicture(App.Path & "\exxon.BMP") 'added ML 2/17/99
'
''    color the controls and form backcolor
'    Me.BackColor = frm_Color.txt_WBackground.BackColor
'    For li_x = 0 To (Controls.count - 1)
''        Debug.Print Controls(li_x).name
'        'If TOC then deny it a call to gsb_fade_to_black
'        If Not (TypeOf Controls(li_x) Is Toc) Then Call gsb_fade_to_black(Controls(li_x))
'    Next li_x
'End Sub
'
'Private Sub lst_query_DblClick()
'    Select Case ListNumber
'        Case 1
'            If Company = "" Then
'                Company = lst_query.Text
'            ElseIf cbo_Company.Text = "SEVERAL" Then
'                Company = Company & ";" & lst_query.Text
'            End If
'        Case 2
'            If Warehouse = "" Then
'                Warehouse = lst_query.Text
'            ElseIf cbo_Warehouse.Text = "SEVERAL" Then
'                Warehouse = Warehouse & ";" & lst_query.Text
'            End If
'        Case 3
'            If Category = "" Then
'                Category = lst_query.Text
'            ElseIf cbo_Category.Text = "SEVERAL" Then
'                Category = Category & ";" & lst_query.Text
'            End If
'        Case 4
'            If Logwhouse = "" Then
'                Logwhouse = lst_query.Text
'            ElseIf cbo_Logwhse.Text = "SEVERAL" Then
'                Logwhouse = Logwhouse & ";" & lst_query.Text
'            End If
'        Case 5
'            If Subloc = "" Then
'                Subloc = lst_query.Text
'            ElseIf cbo_subloc.Text = "SEVERAL" Then
'                Subloc = Subloc & ";" & lst_query.Text
'            End If
'        Case Else
'    End Select
'    txt_selected.Text = Company & "~" & Warehouse & "~" & Category & "~" & Logwhouse & "~" & Subloc & "~"
'End Sub
'
'Private Sub txt_Category_Click()
'    txt_Category.SelStart = 0
'    txt_Category.SelLength = 1
'End Sub
'
'Private Sub cmd_Add_Click()
'    If (opt_SupFax Or opt_FaxNum) And (txt_Recipient.Text = "") Then
'        Call gs_AddFax_Click(Counter, Selected, Me)
'    Else:
'        Call gs_Add_Click(Counter, Selected, Me)
'    End If
'    datachanged = True
'End Sub
'
'Private Sub cmd_Remove_Click()
'    Call gs_Rmv_Click(Counter, Me)
'End Sub
'
'Private Sub opt_SupFax_Click()
'    'load from the Supplier's Fax
'    VisM1.P2 = "^PM" & gs_UT & "SUP"
'    'Set the variable numbers = correct numbers
'    VisM1.P1 = "2;1;9"
'    Call fsb_load_fax
'End Sub
'
'Private Sub opt_FaxNum_Click()
'    'load from other fax
'    VisM1.P2 = "^PM" & gs_UT & "CALP"
'    'Set the variable numbers = correct numbers
'    VisM1.P1 = "2;1;9"
'    Call fsb_load_fax
'End Sub
'
'Private Sub opt_Email_Click()
'    'load from email
'    VisM1.P2 = "^PM" & gs_UT & "CALP"
'    'Set the variable numbers = correct numbers
'    VisM1.P1 = "2;1;11"
'    Call fsb_load_fax
'End Sub
'
'Public Sub fsb_load_fax()
'    Temporary Variable for Code
'    Dim ls_code As String
'    Dim ls_data As String
'    Dim la_length As Variant
'
'    'Call the Routine for Code list
'    VisM1.Code = "d ^loadkeyandfield"
'    VisM1.ExecFlag = 1
'
'    'Load the list bo Phonebook with the
'    'Codes
'
'    la_length = Array(15, 20, 50)
'
'    Clear the phonebook & selected item
'    lst_Phonebook.Clear
'    Selected = -1
'
'    Set initial value of temp string
'    ls_data = VisM1.P0
'    ls_code = piece1(ls_data, ";")
'    'Get all the codes
'    Do While ls_data <> ""
'        ls_code = piece1(ls_data, ";")
'        ls_code = gfn_str_align(ls_code, 3, la_length)
'        If ls_code <> "" Then
'          lst_Phonebook.AddItem (ls_code)
'        End If
'    Loop
'End Sub
'
'Private Sub lst_Phonebook_Click()
'    Call gs_Phone_Click(Selected, Me)
'End Sub
'
'Private Sub lst_Phonebook_DblClick()
'    If (opt_SupFax Or opt_FaxNum) Then
'        Call gs_AddFax_Click(Counter, Selected, Me)
'    Else:
'        Call gs_Add_Click(Counter, Selected, Me)
'    End If
'    datachanged = True
'End Sub
'
'Private Sub txt_Category_KeyPress(KeyAscii As Integer)
'    Dim X As Integer
'    Dim valid As Boolean
'    Dim isZero As Boolean
'
'    valid = False 'assume that the key is invalid
'    If txt_Category.Text = "0" Then isZero = True 'use this for numeric or zero
'    For X = LBound(Keyvalid2) To UBound(Keyvalid1)
'        If Chr(KeyAscii) = Keyvalid2(X) Then valid = True 'if key is valid make it so
'    Next X
'
'    If valid = False Then
'        KeyAscii = 0 'make the key user pressed not appear
'    Else:
'        check what he wants to change the value to next.
'        if going from a number to 0 then all higher
'        numbers must be set to +1
'        if going from 0 to a number then all higher
'        numbers must be set to -1
'        If isZero Then
'            current Value Is zero
'            If KeyAscii <> 48 Then 'is new value a number?
'                yes then all higher numbers go +1
'                If Val(txt_Warehouse.Text) >= Val(Chr(KeyAscii)) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) + 1
'                If Val(txt_Company.Text) >= Val(Chr(KeyAscii)) Then txt_Company.Text = (Val(txt_Company.Text)) + 1
'                If Val(txt_logwhse.Text) >= Val(Chr(KeyAscii)) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) + 1
'                If Val(txt_subloc.Text) >= Val(Chr(KeyAscii)) Then txt_subloc.Text = (Val(txt_subloc.Text)) + 1
'            End If
'        ElseIf Not (isZero) Then
'            current value is a number
'            If KeyAscii = 48 Then 'is new value zero?
'                yes then all higher numbers go -1
'                If Val(txt_Warehouse.Text) >= Val(txt_Category.Text) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) - 1
'                If Val(txt_Company.Text) >= Val(txt_Category.Text) Then txt_Company.Text = (Val(txt_Company.Text)) - 1
'                If Val(txt_logwhse.Text) >= Val(txt_Category.Text) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) - 1
'                If Val(txt_subloc.Text) >= Val(txt_Category.Text) Then txt_subloc.Text = (Val(txt_subloc.Text)) - 1
'                Keyvalid1 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid2 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid3 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid4 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid5 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'            Else: 'if current value is a number and new value is a number
'                switch the two numbers
'                If Val(txt_Warehouse.Text) = Val(Chr(KeyAscii)) Then txt_Warehouse.Text = txt_Category.Text
'                If Val(txt_Company.Text) = Val(Chr(KeyAscii)) Then txt_Company.Text = txt_Category.Text
'                If Val(txt_logwhse.Text) = Val(Chr(KeyAscii)) Then txt_logwhse.Text = txt_Category.Text
'                If Val(txt_subloc.Text) = Val(Chr(KeyAscii)) Then txt_subloc.Text = txt_Category.Text
'            End If
'        End If
'    End If
'    txt_Category.SelStart = 0
'    txt_Category.SelLength = 1
'End Sub
'
'Private Sub txt_Company_Click()
'    txt_Company.SelStart = 0
'    txt_Company.SelLength = 1
'End Sub
'
'Private Sub txt_Company_KeyPress(KeyAscii As Integer)
'    Dim X As Integer
'    Dim valid As Boolean
'    Dim isZero As Boolean
'
'    valid = False 'assume that the key is invalid
'    If txt_Company.Text = "0" Then isZero = True 'use this for numeric or zero
'    For X = LBound(Keyvalid2) To UBound(Keyvalid1)
'        If Chr(KeyAscii) = Keyvalid2(X) Then valid = True 'if key is valid make it so
'    Next X
'
'    If valid = False Then
'        KeyAscii = 0 'make the key user pressed not appear
'    Else:
'        check what he wants to change the value to next.
'        if going from a number to 0 then all higher
'        numbers must be set to +1
'        if going from 0 to a number then all higher
'        numbers must be set to -1
'        If isZero Then
'            current Value Is zero
'            If KeyAscii <> 48 Then 'is new value a number?
'                yes then all higher numbers go +1
'                If Val(txt_Warehouse.Text) >= Val(Chr(KeyAscii)) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) + 1
'                If Val(txt_Category.Text) >= Val(Chr(KeyAscii)) Then txt_Category.Text = (Val(txt_Category.Text)) + 1
'                If Val(txt_logwhse.Text) >= Val(Chr(KeyAscii)) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) + 1
'                If Val(txt_subloc.Text) >= Val(Chr(KeyAscii)) Then txt_subloc.Text = (Val(txt_subloc.Text)) + 1
'            End If
'        ElseIf Not (isZero) Then
'            current value is a number
'            If KeyAscii = 48 Then 'is new value zero?
'                yes then all higher numbers go -1
'                If Val(txt_Warehouse.Text) >= Val(txt_Company.Text) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) - 1
'                If Val(txt_Category.Text) >= Val(txt_Company.Text) Then txt_Category.Text = (Val(txt_Category.Text)) - 1
'                If Val(txt_logwhse.Text) >= Val(txt_Company.Text) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) - 1
'                If Val(txt_subloc.Text) >= Val(txt_Company.Text) Then txt_subloc.Text = (Val(txt_subloc.Text)) - 1
'                Keyvalid1 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid2 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid3 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid4 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid5 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'            Else: 'if current value is a number and new value is a number
'                switch the two numbers
'                If Val(txt_Warehouse.Text) = Val(Chr(KeyAscii)) Then txt_Warehouse.Text = txt_Company.Text
'                If Val(txt_Category.Text) = Val(Chr(KeyAscii)) Then txt_Category.Text = txt_Company.Text
'                If Val(txt_logwhse.Text) = Val(Chr(KeyAscii)) Then txt_logwhse.Text = txt_Company.Text
'                If Val(txt_subloc.Text) = Val(Chr(KeyAscii)) Then txt_subloc.Text = txt_Company.Text
'            End If
'        End If
'    End If
'    txt_Company.SelStart = 0
'    txt_Company.SelLength = 1
'End Sub
'
'Private Sub txt_filename_KeyPress(KeyAscii As Integer)
'    If (KeyAscii = vbKeyReturn) Then
'        If Len(Dir(txt_filename.Text)) <= 0 Then lbl_filelen.Caption = "file does not exist"
'        If Len(Dir(txt_filename.Text)) > 0 Then lbl_filelen.Caption = "file exits, length " & (FileLen(txt_filename.Text))
'        If (FileLen(txt_filename.Text) = 0) Then MsgBox ("File does not exist")
'    End If
'End Sub
'
'Private Sub txt_logwhse_Click()
'    txt_logwhse.SelStart = 0
'    txt_logwhse.SelLength = 1
'End Sub
'
'Private Sub txt_logwhse_KeyPress(KeyAscii As Integer)
'    Dim X As Integer
'    Dim valid As Boolean
'    Dim isZero As Boolean
'
'    valid = False 'assume that the key is invalid
'    If txt_logwhse.Text = "0" Then isZero = True 'use this for numeric or zero
'    For X = LBound(Keyvalid2) To UBound(Keyvalid1)
'        If Chr(KeyAscii) = Keyvalid2(X) Then valid = True 'if key is valid make it so
'    Next X
'
'    If valid = False Then
'        KeyAscii = 0 'make the key user pressed not appear
'    Else:
'        check what he wants to change the value to next.
'        if going from a number to 0 then all higher
'        numbers must be set to +1
'        if going from 0 to a number then all higher
'        numbers must be set to -1
'        If isZero Then
'            current Value Is zero
'            If KeyAscii <> 48 Then 'is new value a number?
'                yes then all higher numbers go +1
'                If Val(txt_Warehouse.Text) >= Val(Chr(KeyAscii)) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) + 1
'                If Val(txt_Category.Text) >= Val(Chr(KeyAscii)) Then txt_Category.Text = (Val(txt_Category.Text)) + 1
'                If Val(txt_Company.Text) >= Val(Chr(KeyAscii)) Then txt_Company.Text = (Val(txt_Company.Text)) + 1
'                If Val(txt_subloc.Text) >= Val(Chr(KeyAscii)) Then txt_subloc.Text = (Val(txt_subloc.Text)) + 1
'            End If
'        ElseIf Not (isZero) Then
'            current value is a number
'            If KeyAscii = 48 Then 'is new value zero?
'                yes then all higher numbers go -1
'                If Val(txt_Warehouse.Text) >= Val(txt_logwhse.Text) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) - 1
'                If Val(txt_Category.Text) >= Val(txt_logwhse.Text) Then txt_Category.Text = (Val(txt_Category.Text)) - 1
'                If Val(txt_Company.Text) >= Val(txt_logwhse.Text) Then txt_Company.Text = (Val(txt_Company.Text)) - 1
'                If Val(txt_subloc.Text) >= Val(txt_logwhse.Text) Then txt_subloc.Text = (Val(txt_subloc.Text)) - 1
'                Keyvalid1 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid2 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid3 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid4 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid5 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'            Else: 'if current value is a number and new value is a number
'                switch the two numbers
'                If Val(txt_Warehouse.Text) = Val(Chr(KeyAscii)) Then txt_Warehouse.Text = txt_logwhse.Text
'                If Val(txt_Category.Text) = Val(Chr(KeyAscii)) Then txt_Category.Text = txt_logwhse.Text
'                If Val(txt_Company.Text) = Val(Chr(KeyAscii)) Then txt_Company.Text = txt_logwhse.Text
'                If Val(txt_subloc.Text) = Val(Chr(KeyAscii)) Then txt_subloc.Text = txt_logwhse.Text
'            End If
'        End If
'    End If
'    txt_logwhse.SelStart = 0
'    txt_logwhse.SelLength = 1
'End Sub
'
'Private Sub txt_subloc_Click()
'    txt_subloc.SelStart = 0
'    txt_subloc.SelLength = 1
'End Sub
'
'Private Sub txt_subloc_KeyPress(KeyAscii As Integer)
'    Dim X As Integer
'    Dim valid As Boolean
'    Dim isZero As Boolean
'
'    valid = False 'assume that the key is invalid
'    If txt_subloc.Text = "0" Then isZero = True 'use this for numeric or zero
'    For X = LBound(Keyvalid2) To UBound(Keyvalid1)
'        If Chr(KeyAscii) = Keyvalid2(X) Then valid = True 'if key is valid make it so
'    Next X
'
'    If valid = False Then
'        KeyAscii = 0 'make the key user pressed not appear
'    Else:
'        check what he wants to change the value to next.
'        if going from a number to 0 then all higher
'        numbers must be set to +1
'        if going from 0 to a number then all higher
'        numbers must be set to -1
'        If isZero Then
'            current Value Is zero
'            If KeyAscii <> 48 Then 'is new value a number?
'                yes then all higher numbers go +1
'                If Val(txt_Warehouse.Text) >= Val(Chr(KeyAscii)) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) + 1
'                If Val(txt_Category.Text) >= Val(Chr(KeyAscii)) Then txt_Category.Text = (Val(txt_Category.Text)) + 1
'                If Val(txt_Company.Text) >= Val(Chr(KeyAscii)) Then txt_Company.Text = (Val(txt_Company.Text)) + 1
'                If Val(txt_logwhse.Text) >= Val(Chr(KeyAscii)) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) + 1
'            End If
'        ElseIf Not (isZero) Then
'            current value is a number
'            If KeyAscii = 48 Then 'is new value zero?
'                yes then all higher numbers go -1
'                If Val(txt_Warehouse.Text) >= Val(txt_subloc.Text) Then txt_Warehouse.Text = (Val(txt_Warehouse.Text)) - 1
'                If Val(txt_Category.Text) >= Val(txt_subloc.Text) Then txt_Category.Text = (Val(txt_Category.Text)) - 1
'                If Val(txt_Company.Text) >= Val(txt_subloc.Text) Then txt_Company.Text = (Val(txt_Company.Text)) - 1
'                If Val(txt_logwhse.Text) >= Val(txt_subloc.Text) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) - 1
'                Keyvalid1 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid2 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid3 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid4 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid5 = Array(txt_Warehouse.Text, txt_Company.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'            Else: 'if current value is a number and new value is a number
'                switch the two numbers
'                If Val(txt_Warehouse.Text) = Val(Chr(KeyAscii)) Then txt_Warehouse.Text = txt_subloc.Text
'                If Val(txt_Category.Text) = Val(Chr(KeyAscii)) Then txt_Category.Text = txt_subloc.Text
'                If Val(txt_Company.Text) = Val(Chr(KeyAscii)) Then txt_Company.Text = txt_subloc.Text
'                If Val(txt_logwhse.Text) = Val(Chr(KeyAscii)) Then txt_logwhse.Text = txt_subloc.Text
'            End If
'        End If
'    End If
'    txt_subloc.SelStart = 0
'    txt_subloc.SelLength = 1
'End Sub
'
'Private Sub txt_Warehouse_Click()
'    txt_Warehouse.SelStart = 0
'    txt_Warehouse.SelLength = 1
'End Sub
'
'Private Sub txt_Warehouse_KeyPress(KeyAscii As Integer)
'    Dim X As Integer
'    Dim valid As Boolean
'    Dim isZero As Boolean
'
'    valid = False 'assume that the key is invalid
'    If txt_Warehouse.Text = "0" Then isZero = True 'use this for numeric or zero
'    For X = LBound(Keyvalid2) To UBound(Keyvalid1)
'        If Chr(KeyAscii) = Keyvalid2(X) Then valid = True 'if key is valid make it so
'    Next X
'
'    If valid = False Then
'        KeyAscii = 0 'make the key user pressed not appear
'    Else:
'        check what he wants to change the value to next.
'        if going from a number to 0 then all higher
'        numbers must be set to +1
'        if going from 0 to a number then all higher
'        numbers must be set to -1
'        If isZero Then
'            current Value Is zero
'            If KeyAscii <> 48 Then 'is new value a number?
'                yes then all higher numbers go +1
'                If Val(txt_Company.Text) >= Val(Chr(KeyAscii)) Then txt_Company.Text = (Val(txt_Company.Text)) + 1
'                If Val(txt_Category.Text) >= Val(Chr(KeyAscii)) Then txt_Category.Text = (Val(txt_Category.Text)) + 1
'                If Val(txt_logwhse.Text) >= Val(Chr(KeyAscii)) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) + 1
'                If Val(txt_subloc.Text) >= Val(Chr(KeyAscii)) Then txt_subloc.Text = (Val(txt_subloc.Text)) + 1
'            End If
'        ElseIf Not (isZero) Then
'            current value is a number
'            If KeyAscii = 48 Then 'is new value zero?
'                yes then all higher numbers go -1
'                If Val(txt_Company.Text) >= Val(txt_Warehouse.Text) Then txt_Company.Text = (Val(txt_Company.Text)) - 1
'                If Val(txt_Category.Text) >= Val(txt_Warehouse.Text) Then txt_Category.Text = (Val(txt_Category.Text)) - 1
'                If Val(txt_logwhse.Text) >= Val(txt_Warehouse.Text) Then txt_logwhse.Text = (Val(txt_logwhse.Text)) - 1
'                If Val(txt_subloc.Text) >= Val(txt_Warehouse.Text) Then txt_subloc.Text = (Val(txt_subloc.Text)) - 1
'                Keyvalid1 = Array(txt_Company.Text, txt_Warehouse.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid2 = Array(txt_Company.Text, txt_Warehouse.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid3 = Array(txt_Company.Text, txt_Warehouse.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid4 = Array(txt_Company.Text, txt_Warehouse.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'                Keyvalid5 = Array(txt_Company.Text, txt_Warehouse.Text, txt_Category.Text, txt_logwhse.Text, txt_subloc.Text, "0")
'            Else: 'if current value is a number and new value is a number
'                switch the two numbers
'                If Val(txt_Company.Text) = Val(Chr(KeyAscii)) Then txt_Company.Text = txt_Warehouse.Text
'                If Val(txt_Category.Text) = Val(Chr(KeyAscii)) Then txt_Category.Text = txt_Warehouse.Text
'                If Val(txt_logwhse.Text) = Val(Chr(KeyAscii)) Then txt_logwhse.Text = txt_Warehouse.Text
'                If Val(txt_subloc.Text) = Val(Chr(KeyAscii)) Then txt_subloc.Text = txt_Warehouse.Text
'            End If
'        End If
'    End If
'    txt_Warehouse.SelStart = 0
'    txt_Warehouse.SelLength = 1
'End Sub
Private Sub cbo_subloc_Change()

End Sub

Private Sub cbo_Warehouse_DropDown()
cbo_Warehouse.locked = False
End Sub

Private Sub cbo_Warehouse_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If open_forms <= 5 Then frmNavigator.Visible = True
End Sub
