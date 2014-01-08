VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class
Public ExportFilePath As String
Public po As String
Public Namespace As String
        
        Public Function GeneratePdfForPO() As String
        
            
            Dim Report As CRAXDDRT.Report
            
            'Declare a DatabaseTable Object
            Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
            'Declare a Report object to set to the subeport
            Dim crxSubreport As CRAXDDRT.Report
            
            On Error GoTo ErrHandler
             Set crxApplication = New CRAXDDRT.Application
            'Open the report
            
            Set Report = crxApplication.OpenReport("C:\projects\vb\ims\crreports\2005poGood.rpt", 1)
            'Set Report = crxApplication.OpenReport("C:\projects\vb\ims\crreports\poCr8Good.rpt", 1)
            
             
            
            Call FixDB(Report.Database.Tables)
            
            '‘Set crxSubreport to the subreport ‘Sub1’ of the main report. The subreport name needs to be known to use this ‘method.
            Set crxSubreport = Report.OpenSubreport("porem.rpt")
            
            
            Call FixDB(crxSubreport.Database.Tables)
            
            Set crxSubreport = Report.OpenSubreport("poclause.rpt")
            Call FixDB(crxSubreport.Database.Tables)
            
            
            Report.ParameterFields.Item(1).AddCurrentValue Namespace
            Report.ParameterFields.Item(2).AddCurrentValue po
            
            
            Call Export(Report)
            
            Exit Function
            
ErrHandler:
            
            GeneratePdfForPO = "Errors Occurred while trying to generate a PDF, please try again." + Err.Description
            Err.Clear
            
        End Function
        
        Private Function FixDB(crxDatabaseTableS As CRAXDDRT.DatabaseTables)

            Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
            
            For Each crxDatabaseTable In crxDatabaseTableS
            
            crxDatabaseTable.SetLogOnInfo "imsO", "pecten_dev", "sa", "scms"
                'crxDatabaseTable.Location = "pecten_dev.dbo." & crxDatabaseTable.Name
                crxDatabaseTable.Location = crxDatabaseTable.Name
                
                'Debug.Print crxDatabaseTable.Name
                'Debug.Print crxDatabaseTable.Location
                
            Next crxDatabaseTable
            
            
            
            'For Each crxDatabaseTable In crxDatabaseTableS
            
                
            '    Debug.Print crxDatabaseTable.Name
            '    Debug.Print crxDatabaseTable.Location
                
            'Next crxDatabaseTable

        End Function
        
        Private Sub Export(Report As CRAXDDRT.Report)
        
            'Dim Report As CRAXDDRT.Report
            
            '‘Set the report to be exported to Rich Text Format
            Report.ExportOptions.FormatType = crEFTPortableDocFormat
            '‘Set the destination type to disk
            Report.ExportOptions.DestinationType = crEDTDiskFile
            '‘Set the path and name of the exported document
            Report.ExportOptions.DiskFileName = ExportFilePath
            
            'report.ExportOptions.pd
            '‘export the report without prompting the user
            Report.Export False
             
        End Sub


'End
