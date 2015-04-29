VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imsTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TR_LANGUAGE As String
Public mainREPORT As Boolean
Public thisREPO As Crystal.CrystalReport
Public Sub changeLANGUAGE()
'Procedure for start and open de database resources for translation
    Dim sql
    deIms.Recordsets("TR_CONTROLS").Source = deIms.Recordsets("TR_CONTROLS").Source + " WHERE TRMESSAGE.msg_lang = '" + TR_LANGUAGE + "'"
    deIms.Recordsets("TR_CONTROLS").Open
    sql = deIms.Recordsets("TR_MESSAGES").Source
    deIms.Recordsets("TR_MESSAGES").Source = Left(sql, InStr(sql, "ORDER") - 1) + " WHERE TRLANGUAGE.lan_code  = '" + TR_LANGUAGE + "' " + Mid(sql, InStr(sql, "ORDER"))
    deIms.Recordsets("TR_MESSAGES").Open
End Sub


Public Function Trans(MessageCode) As String
'Function for retrieve direct texts for translation
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        With deIms.Recordsets("TR_MESSAGES")
            .MoveFirst
            .Find "msg_numb = '" + MessageCode + "'"
            If .EOF Then
                Trans = ""
            Else
                Trans = !msg_text
            End If
        End With
        Err.Clear
    End If
End Function

Public Sub Translate_Forms(Form_name As String)
'Procedure for captions translations in every form
    Dim i, j, k, indexARRAY, indexTAB, indexCOL As Integer
    Dim originalFILTER, nameCONTROLs, nameCONTROLs2  As String
    Dim withARRAY, withTAB As Boolean
    Dim lilControl As NavBar
    
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        On Error Resume Next
        With deIms.Recordsets("TR_CONTROLS")
            originalFILTER = .Filter
            For i = 0 To VB.Forms.Count - 1
                If VB.Forms(i).Name = Form_name Then
                    .Filter = "trs_enty = '" + Form_name + "'"
                    If .RecordCount > 0 Then
                        .Find "trs_obj = '" + Form_name + "'"
                        If Not .EOF Then VB.Forms(i).Caption = !msg_text
                        
                            For j = 0 To VB.Forms(i).Controls.Count - 1
                                nameCONTROLs = VB.Forms(i).Controls(j).Name
                                If TypeOf VB.Forms(i).Controls(j) Is lrnavigators.NavBar Or TypeOf VB.Forms(i).Controls(j) Is lrnavigators.LROleDBNavBar Then
                                    Set VB.Forms(i).Controls(j).ActiveConnection = deIms.cnIms
                                    VB.Forms(i).Controls(j).Language = translator.TR_LANGUAGE
                                End If
                                indexARRAY = -1
                                indexARRAY = VB.Forms(i).Controls(j).Index
                                If indexARRAY >= 0 Then
                                    nameCONTROLs = nameCONTROLs + "(" + Format(indexARRAY) + ")"
                                    .MoveFirst
                                    .Find "trs_obj = '" + nameCONTROLs + "'"
                                    If Not .EOF Then VB.Forms(i).Controls(j).Caption = !msg_text
                                Else
                                    indexTAB = -1
                                    indexTAB = VB.Forms(i).Controls(j).Tabs
                                    If indexTAB >= 0 Then
                                        For k = 0 To indexTAB - 1
                                            nameCONTROLs2 = nameCONTROLs + ".Tab(" + Format(k) + ")"
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                            If Not .EOF Then VB.Forms(i).Controls(j).TabCaption(k) = !msg_text
                                        Next
                                    Else
                                        indexCOL = -1
                                        indexCOL = VB.Forms(i).Controls(j).Columns.Count
                                        If indexCOL >= 0 Then
                                            For k = 0 To indexCOL - 1
                                                nameCONTROLs2 = nameCONTROLs + "." + VB.Forms(i).Controls(j).Columns(k).Caption
                                                .MoveFirst
                                                .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                                If Not .EOF Then
                                                    VB.Forms(i).Controls(j).Columns(k).Caption = !msg_text
                                                Else
                                                    nameCONTROLs2 = nameCONTROLs + ".Columns(" + Format(k) + ")"
                                                    .MoveFirst
                                                    .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                                    If Not .EOF Then VB.Forms(i).Controls(j).Columns(k).Caption = !msg_text
                                                End If
                                            Next
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs + "'"
                                            If Not .EOF Then VB.Forms(i).Controls(j).Caption = !msg_text
                                        Else
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs + "'"
                                            If Not .EOF Then VB.Forms(i).Controls(j).Caption = !msg_text
                                        End If
                                    End If
                                End If
                            Next
                            
                        End If
                    Exit For
                End If
            Next
        End With
        Err.Clear
    End If
End Sub


Public Sub Translate_Reports(reportNAME As String)
'Procedure for labels translations in every report
    Dim n, x, xx As Integer
    Dim tableNAME As String
    On Error GoTo errSTOP
    
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        With deIms.Recordsets("TR_CONTROLS")
            .Filter = ""
            .Filter = "trs_enty = '" + reportNAME + "'"
            If .RecordCount > 0 Then
                n = 0
                Do While Not .EOF
                    thisREPO.Formulas(n) = !trs_obj + " = '" + !msg_text + "'"
                    n = n + 1
                    .MoveNext
                Loop
            End If
            Err.Clear
        End With
    End If
            
    If translator.mainREPORT Then
        x = thisREPO.RetrieveLogonInfo - 1
        x = thisREPO.RetrieveDataFiles - 1
    Else
        Dim subreportQUERY As New ADODB.Recordset
        Dim sql, mainREP, subREP
        mainREP = thisREPO.ReportFileName
        mainREP = Mid(mainREP, InStrRev(mainREP, "\") + 1)
        subREP = reportNAME
        tableNAME = ""
        sql = "SELECT * FROM REPORTS WHERE report = '" + mainREP + "' and subreport = '" + subREP + "'"
        Set subreportQUERY = New ADODB.Recordset
        With subreportQUERY
            .Open sql, deIms.cnIms, adOpenForwardOnly
            x = .RecordCount
            ReDim alltables(x) As String
            x = 0
            If .RecordCount > 0 Then
                Do While Not .EOF
                    alltables(x) = !tableNAME
                    x = x + 1
                    .MoveNext
                Loop
                x = x - 1
            End If
            .Close
        End With
    End If
    For n = 0 To x
        thisREPO.LogonInfo(n) = "dsn=" + ConnInfo.Dsnname + ";dsq=" _
        & ConnInfo.InitCatalog + ";uid=" + ConnInfo.UId + ";pwd=" + ConnInfo.Pwd
        tableNAME = thisREPO.DataFiles(n)
        
        If tableNAME = "" Then
            If Not translator.mainREPORT Then
                If IsNull(alltables(n)) Or alltables(n) = "" Then
                Else
                    xx = InStr(alltables(n), ".")
                    tableNAME = Mid(alltables(n), IIf(xx = 0, 1, xx))
                End If
            End If
        Else
            xx = InStr(tableNAME, ".")
            tableNAME = Mid(tableNAME, IIf(xx = 0, 1, xx))
        End If
        thisREPO.DataFiles(n) = ConnInfo.InitCatalog + tableNAME
    Next
    Exit Sub
    
errSTOP:
    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub Translate_SubReports()
'Procedure for process sub reports
    Dim repo As String
    Dim i As Integer
    translator.mainREPORT = False
    With thisREPO
        For i = 0 To .GetNSubreports - 1
            repo = .GetNthSubreportName(i)
            .SubreportToChange = repo
            Call translator.Translate_Reports(repo)
        Next
        translator.mainREPORT = True
        .SubreportToChange = ""
    End With
End Sub


