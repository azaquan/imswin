VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class
Public ExportFilePath As String
Public PO As String
Public NameSpace As String
'Public ParamsForCrystalReport() As String
Public ReportName As String
Public ReporttypesCr11 As Integer

        Public Function GeneratePdf(ParamsForCrystalReport() As String) As String
        
            
            Dim Report As CRAXDDRT.Report
            
            'Declare a DatabaseTable Object
            Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
            'Declare a Report object to set to the subeport
            Dim crxSubreport As CRAXDDRT.Report
            Dim Param As CRAXDDRT.ParameterFieldDefinition
            Dim arrparam() As String
            On Error GoTo ErrHandler
             Set crxApplication = New CRAXDDRT.Application
            'Open the report
            
            Set Report = crxApplication.OpenReport(reportPath + ReportName, 1)
            
'        For I = 0 To UBound(ParamsForCrystalReport)
        
            
'            ArrParam = Split(ParamsForCrystalReport(I), ";")
'             Report.ParameterFields.Item(1).AddCurrentValue ArrParam(1)
'
'        Next I
  
            
           Set Report = InitializeReport(Report, ParamsForCrystalReport())
            
            
'           For i = 0 To UBound(ParamsForCrystalReport)
'
'                arrparam = Split(ParamsForCrystalReport(i), ";")
'
'                Set Param = Report.ParameterFields.GetItemByName(arrparam(0))
'
'                Param.AddCurrentValue arrparam(1)
'
'           Next
            
            'Call translator.Translate_Reports(reportPath)

            'Call translator.Translate_SubReports
            'Call FixDB(Report.Database.Tables)
            Call Export(Report)
            
            Exit Function
            
ErrHandler:
            
            GeneratePdfForPO = "Errors Occurred while trying to generate a PDF, please try again." + Err.Description
            Err.Clear
            
        End Function
        
        Public Function InitializeReport(Report As CRAXDDRT.Report, ParamsForCrystalReport() As String) As CRAXDDRT.Report
        
        Dim crxSubreport As CRAXDDRT.Report
        Dim arrparam() As String
        
        On Error GoTo ErrHand
        
        '//PO Report
        If ReportName = Report_EmailFax_PO_name Then
        
                     Call FixDB(Report.Database.Tables)
                    
                    '‘Set crxSubreport to the subreport ‘Sub1’ of the main report. The subreport name needs to be known to use this ‘method.
                    Set crxSubreport = Report.OpenSubreport("porem.rpt")
                    
                    
                    Call FixDB(crxSubreport.Database.Tables)
                    
                    Set crxSubreport = Report.OpenSubreport("poclause.rpt")
                    Call FixDB(crxSubreport.Database.Tables)
                    
                     arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.Item(1).AddCurrentValue deIms.NameSpace
                    Report.ParameterFields.Item(2).AddCurrentValue arrparam(1)
        
        
        '//Freight Receipt
        ElseIf ReportName = Report_EmailFax_FreightReceipt_name Then
        
        
                    Call FixDB(Report.Database.Tables)
                    
                    arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.GetItemByName("recnumb").AddCurrentValue (arrparam(1))
                    Report.ParameterFields.GetItemByName("namespace").AddCurrentValue (deIms.NameSpace)
        
        ElseIf ReportName = Report_EmailFax_TrackingPo Then
        
        
                    Call FixDB(Report.Database.Tables)
                    
                    Report.ParameterFields.GetItemByName("namespace").AddCurrentValue (deIms.NameSpace)
                    
                    arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.GetItemByName("mesgnumb").AddCurrentValue (arrparam(1))
                    
                    arrparam = Split(ParamsForCrystalReport(2), ";")
                    Report.ParameterFields.GetItemByName("ponumb").AddCurrentValue (arrparam(1))
                    
        ElseIf ReportName = Report_EmailFax_Supplier_name Then
        
                     Call FixDB(Report.Database.Tables)
                    
                    Report.ParameterFields.GetItemByName("namespace").AddCurrentValue (deIms.NameSpace)
                    
                    arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.GetItemByName("suppcode").AddCurrentValue (arrparam(1))
                    
                    arrparam = Split(ParamsForCrystalReport(2), ";")
                    Report.ParameterFields.GetItemByName("Intloc").AddCurrentValue (arrparam(1))
                    
        ElseIf ReportName = Report_EmailFax_PackingManifest_name Then
        
                     Call FixDB(Report.Database.Tables)
                    
                    Report.ParameterFields.GetItemByName("namespace").AddCurrentValue (deIms.NameSpace)
                    
                    arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.GetItemByName("manifestnumb").AddCurrentValue (arrparam(1))
                            
        
        ElseIf ReportName = Report_EmailFax_Stockmaster_name Then
        
                    Call FixDB(Report.Database.Tables)
                    
                    
                    Set crxSubreport = Report.OpenSubreport("sbhistory.rpt")
                    
                    
                    Call FixDB(crxSubreport.Database.Tables)
                    
                    Report.ParameterFields.GetItemByName("namespace").AddCurrentValue (deIms.NameSpace)
                    
                    arrparam = Split(ParamsForCrystalReport(1), ";")
                    Report.ParameterFields.GetItemByName("stcknumb").AddCurrentValue (arrparam(1))
        
        End If
        
        Set InitializeReport = Report
        
        Exit Function
ErrHand:
        
        Call LogErr("InitializeReport ", Err.Description, Err.number)
        MsgBox "InitializeReport function : " + Err.Description
        Err.Clear
        
        End Function
        
        Private Function FixDB(crxDatabaseTableS As CRAXDDRT.DatabaseTables)

            Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
            
            For Each crxDatabaseTable In crxDatabaseTableS
            
            crxDatabaseTable.SetLogOnInfo ConnInfo.Dsnname, ConnInfo.InitCatalog, ConnInfo.UId, ConnInfo.Pwd    ' "imsO", "pecten_dev", "sa", "scms"
                'crxDatabaseTable.Location = Trim(ConnInfo.InitCatalog) & ".dbo." & crxDatabaseTable.Name
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
            '‘export the report without prompting the user
            Report.Export False
             
        End Sub


'End

