VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   -360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Declare the application object used to open the rpt file.
Dim crxApplication As New CRAXDDRT.Application
'Declare the report object
Public Report As CRAXDDRT.Report


Private Sub runpo()
'Declare a DatabaseTable Object
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
'Declare a Report object to set to the subeport
Dim crxSubreport As CRAXDDRT.Report
'Open the report
Set Report = crxApplication.OpenReport("C:\projects\vb\ims\crreports\2005poGood.rpt", 1)

'Use a For Each loop to change the location of each
'DatabaseTable in the Reports DatabaseTable Collection

Call FixDB(Report.Database.Tables)

'For Each crxDatabaseTable In Report.Database.Tables
''crxDatabaseTable.ConnectionProperties("Database Name") = App.Path & "\xtreme.mdb"
'    crxDatabaseTable.SetLogOnInfo "imsO", "Pecten", "sa", "scms"
'    crxDatabaseTable.Location = crxDatabaseTable.Name
'Next crxDatabaseTable



'Pass the Parameter value to the first parameter field in the
'ParameterFields collection of the Report.
Report.ParameterFields.Item(1).AddCurrentValue "PECT"
Report.ParameterFields.Item(2).AddCurrentValue "PH-82179"



'‘Set crxSubreport to the subreport ‘Sub1’ of the main report. The subreport name needs to be known to use this ‘method.
Set crxSubreport = Report.OpenSubreport("porem.rpt")
Call FixDB(crxSubreport.Database.Tables)


'‘Use a For Each loop to change the location of each
'‘DatabaseTable in the Subreport Database Table Collection

'For Each crxDatabaseTable In crxSubreport.Database.Tables
'crxDatabaseTable.SetLogOnInfo "imsO", "Pecten", "sa", "scms"
'Next crxDatabaseTable


Set crxSubreport = Report.OpenSubreport("poclause.rpt")
Call FixDB(crxSubreport.Database.Tables)
'
'For Each crxDatabaseTable In crxSubreport.Database.Tables
'crxDatabaseTable.SetLogOnInfo "imsO", "Pecten", "sa", "scms"
'Next crxDatabaseTable


''‘Pass the formula’s text to the first formula field
'‘in the FormulaFields collection of the subreport.
'crxSubreport.FormulaFields.Item(1).Text = "‘Subreport Formula’ "

''‘Set the Report source for the Report Viewer to the Report
CRViewer1.ReportSource = Form1.Report
''‘View the Report
CRViewer1.ViewReport
'Command3_Click
End Sub

Public Function FixDB(crxDatabaseTableS As CRAXDDRT.DatabaseTables)

Dim crxDatabaseTable As CRAXDDRT.DatabaseTable

For Each crxDatabaseTable In crxDatabaseTableS
crxDatabaseTable.SetLogOnInfo "imsO", "Pecten", "sa", "scms"
    crxDatabaseTable.Location = crxDatabaseTable.Name
Next crxDatabaseTable

End Function

Private Sub Command1_Click()
'‘Call Form2 to preview the Report
Form2.Show
End Sub
'Private Sub Command2_Click()
''‘Select the printer for the report passing the
''‘Printer Driver, Printer Name and Printer Port.
'Report.SelectPrinter “HPPCL5MS.DRV”, “HP LaserJet 4m Plus”,
'“\\Vanprt\v1-1mpls-ts”
'Crystal Reports Migrating from the OCX Control to the Report Designer Component (RDC) v9
'11/1/2002 Copyright ? 2002 Crystal Decisions, Inc. All Rights Reserved. Page 10
'‘Print the Report without prompting user
'Report.PrintOut False
'End Sub





Private Sub Command3_Click()
'‘Set the report to be exported to Rich Text Format
Report.ExportOptions.FormatType = crEFTPortableDocFormat
'‘Set the destination type to disk
Report.ExportOptions.DestinationType = crEDTDiskFile
'‘Set the path and name of the exported document
Report.ExportOptions.DiskFileName = App.Path & "\" & 1 & ".pdf"
'‘export the report without prompting the user
Report.Export False
End Sub

Private Sub runbillto()

'Declare a DatabaseTable Object
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
'Declare a Report object to set to the subeport
Dim crxSubreport As CRAXDDRT.Report
'Open the report
Set Report = crxApplication.OpenReport("C:\projects\vb\ims\crreports\billto.rpt", 1)

'Use a For Each loop to change the location of each
'DatabaseTable in the Reports DatabaseTable Collection

For Each crxDatabaseTable In Report.Database.Tables
'crxDatabaseTable.ConnectionProperties("Database Name") = App.Path & "\xtreme.mdb"
 
 crxDatabaseTable.Location = crxDatabaseTable.Name
    crxDatabaseTable.SetLogOnInfo "imsO", "Pecten", "sa", "scms"
    

Next crxDatabaseTable

Report.ParameterFields.Item(1).AddCurrentValue "PECT"
Report.ParameterFields.Item(2).AddCurrentValue "PECTEN"

''‘Set the Report source for the Report Viewer to the Report
'CRViewer1.ReportSource = Form1.Report
''‘View the Report
'CRViewer1.ViewReport
Command3_Click
End Sub


Private Sub Command2_Click()
Command3_Click
End Sub

Sub main()
Form_Load
End Sub

Private Sub Form_Load()
'runbillto
'runpo
Dim x As clsExport
Set x = New clsExport
'x.po = "PH-82179"
x.po = "muz-1"
x.Namespace = "Pect"
x.ExportFilePath = "C:\Projects\VB\ims\NewReports\asdfasfd.pdf"
MsgBox x.GeneratePdfForPO()

End Sub

'
'Private Sub Form_Load()
''‘Set the Report source for the Report Viewer to the Report
'CRViewer1.ReportSource = Form1.Report
''‘View the Report
'CRViewer1.ViewReport
'End Sub
Private Sub Form_Resize()
'‘This code resizes the Report Viewer control to Form2’s

'CRViewer1.Top = 0
'CRViewer1.Left = 0
'CRViewer1.Height = ScaleHeight
'CRViewer1.Width = ScaleWidth
'CRViewer1.EnableExportButton = True
End Sub

