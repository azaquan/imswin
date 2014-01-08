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
Public reportName As String
Public ReporttypesCr11 As Integer

        Public Function GeneratePdf(ParamsForCrystalReport() As String, path As String) As String
        On Error GoTo ErrHandler
            Dim Report As CRAXDRT.Report
            'Declare a DatabaseTable Object
            Dim crxDatabaseTable As CRAXDRT.DatabaseTable
            'Declare a Report object to set to the subeport
            'Dim crxSubreport As CRAXDRT.Report
            'Dim Param As CRAXDRT.ParameterFieldDefinition
            'Dim arrparam() As String
            Set crxApplication = New CRAXDRT.Application
            
            'Open the report
            Set Report = crxApplication.OpenReport(reportName, 1)
            reportPATH = path
            Set Report = InitializeReport(Report, ParamsForCrystalReport())
            
            Call Export(Report)
            Exit Function
            
ErrHandler:
            GeneratePdf = "Errors Occurred while trying to generate a PDF, please try again." + Err.description
            Err.Clear
        End Function
        
        Public Function InitializeReport(Report As CRAXDRT.Report, ParamsForCrystalReport() As String) As CRAXDRT.Report
        Dim crxSubreport As CRAXDRT.Report
        Dim arrparam() As String
        Dim subReportName As String
        On Error GoTo ErrHand
            Call FixDB(Report.Database.Tables)
            'subreport refresh
            subReportName = ""
            Select Case frmWarehouse.tag
                Case "02040400" 'ReturnFromRepair
                Case "02050200" 'AdjustmentEntry
                Case "02040200" 'Issue
                    subReportName = "sbwareIrem.rpt"
                Case "02040500" 'WellToWell
                Case "02040700" 'InternalTransfer
                Case "02050300" 'AdjustmentIssue
                Case "02040600" 'To
                Case "02040100" 'Receipt
                Case "02050400" 'Sales
                Case "02040300" 'Return from Well
            End Select
            If subReportName <> "" Then
                Set crxSubreport = Report.OpenSubreport(subReportName)
                Call FixDB(crxSubreport.Database.Tables)
            End If
            Report.ParameterFields.Item(1).AddCurrentValue ParamsForCrystalReport(0)
            Report.ParameterFields.Item(2).AddCurrentValue ParamsForCrystalReport(1)
            Set InitializeReport = Report
        Exit Function
ErrHand:
        
        Call LogErr("InitializeReport ", Err.description, Err.Number)
        MsgBox "InitializeReport function : " + Err.description
        Err.Clear
        
        End Function
        
Private Function FixDB(crxDatabaseTableS As CRAXDRT.DatabaseTables)
    Dim crxDatabaseTable As CRAXDRT.DatabaseTable
    cnSTRING = Split(cn.ConnectionString, ";")
    For Each text In cnSTRING
        Select Case Left(UCase(text), InStr(text, "="))
            Case "PASSWORD="
                dsnPWD = Mid(text, InStr(text, "=") + 1)
            Case "USER ID="
                dsnUID = Mid(text, InStr(text, "=") + 1)
            Case "INITIAL CATALOG="
                dsnDSQ = Mid(text, InStr(text, "=") + 1)
         End Select
    Next text
    For Each crxDatabaseTable In crxDatabaseTableS
        crxDatabaseTable.SetLogOnInfo dsnF, dsnDSQ, dsnUID, dsnPWD
        crxDatabaseTable.Location = crxDatabaseTable.name
    Next crxDatabaseTable
End Function
        
Private Sub Export(Report As CRAXDRT.Report)
    On Error GoTo errorHandler
    With Report
        '‘Set the report to be exported to Rich Text Format
        .ExportOptions.FormatType = crEFTPortableDocFormat
        '‘Set the destination type to disk
        .ExportOptions.DestinationType = crEDTDiskFile
        '‘Set the path and name of the exported document
        .ExportOptions.DiskFileName = ExportFilePath
        .VerifyOnEveryPrint = False
        .EnableSelectDistinctRecords = False
        .PerformGroupingOnServer = True
        .UseIndexForSpeed = True

        '‘export the report without prompting the user
        .Export False
    End With
errorHandler:
    If Err.Number > 0 Then
        MsgBox Err.description
    End If
End Sub




