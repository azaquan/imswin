Attribute VB_Name = "translator"
Option Explicit
Global thisrepo As CrystalReport
Global TR_LANGUAGE
Global mainREPORT

Public Function Trans(MessageCode) As String
'Function for retrieve direct texts for translation
Dim data As New ADODB.Recordset
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translationMESSAGE", Array(TR_LANGUAGE, MessageCode))
        If data.RecordCount > 0 Then
            Trans = data!msg_text
        Else
            Trans = ""
        End If
        Err.Clear
    End If
End Function
Public Function getIt(objectId) As String
'Function to retrieve direct texts for translation
Dim data As New ADODB.Recordset
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translationGetIt", Array(TR_LANGUAGE, "frmWarehouse", objectId))
        If data.RecordCount > 0 Then
            getIt = data!msg_text
        Else
            getIt = ""
        End If
        Err.Clear
    End If
End Function
Public Sub Translate_Forms(Form_name As String)
'Procedure for captions translations in every form
    Dim data As New ADODB.Recordset
    Dim i, j, k, indexARRAY, indexTAB, indexCOL As Integer
    Dim originalFILTER, nameCONTROLs, nameCONTROLs2  As String
    Dim withARRAY, withTAB As Boolean
    
    On Error Resume Next
    
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translationCONTROLS", Array(TR_LANGUAGE, Form_name))
        With data
            originalFILTER = .Filter
            For i = 0 To VB.Forms.Count - 1
                If VB.Forms(i).name = Form_name Then
                    If .RecordCount > 0 Then
                        .Find "trs_obj = '" + Form_name + "'"
                        'If Not .EOF Then VB.Forms(i).Caption = !msg_text
                        For j = 0 To VB.Forms(i).Controls.Count - 1
                            nameCONTROLs = VB.Forms(i).Controls(j).name
                            If TypeOf VB.Forms(i).Controls(j) Is LRNavigators.NavBar Or TypeOf VB.Forms(i).Controls(j) Is LRNavigators.LROleDBNavBar Then
                                Set VB.Forms(i).Controls(j).ActiveConnection = cn
                                VB.Forms(i).Controls(j).language = TR_LANGUAGE
                            End If
                            indexARRAY = -1
                            indexARRAY = VB.Forms(i).Controls(j).Index
                            If indexARRAY >= 0 Then
                                nameCONTROLs = nameCONTROLs + "(" + Format(indexARRAY) + ")"
                                .MoveFirst
                                .Find "trs_obj = '" + nameCONTROLs + "'"
                                If Not .EOF Then VB.Forms(i).Controls(j) = !msg_text
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
                                    indexCOL = VB.Forms(i).Controls(j).columns.Count
                                    If indexCOL >= 0 Then
                                        For k = 0 To indexCOL - 1
                                            nameCONTROLs2 = nameCONTROLs + "." + VB.Forms(i).Controls(j).columns(k).Caption
                                            .MoveFirst
                                            .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                            If Not .EOF Then
                                                VB.Forms(i).Controls(j).columns(k).Caption = !msg_text
                                            Else
                                                nameCONTROLs2 = nameCONTROLs + ".Columns(" + Format(k) + ")"
                                                .MoveFirst
                                                .Find "trs_obj = '" + nameCONTROLs2 + "'"
                                                If Not .EOF Then VB.Forms(i).Controls(j).columns(k).Caption = !msg_text
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


Public Sub Translate_Reports(repo As String)
'Procedure for labels translations in every report
    Dim data As New ADODB.Recordset
    Dim i, j, n, x, xx As Integer
    Dim tableNAME As String
    Dim formNAME, controlNAME
    Dim subreportQUERY As New ADODB.Recordset
    Dim sql, mainREP, subREP
    mainREP = thisrepo.ReportFileName
    mainREP = Mid(mainREP, InStrRev(mainREP, "\") + 1)
    On Error GoTo errSTOP
    
    
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translatorCONTROLS", Array(TR_LANGUAGE, repo))
        With data
            If .RecordCount > 0 Then
                n = 0
                Do While Not .EOF
                    thisrepo.Formulas(n) = !trs_obj + " = '" + !msg_text + "'"
                    n = n + 1
                    .MoveNext
                Loop
            End If
            Err.Clear
        End With
    End If
            
    If mainREPORT Then
        x = thisrepo.RetrieveLogonInfo - 1
        x = thisrepo.RetrieveDataFiles - 1
    Else
        subREP = repo
        subREP = Mid(subREP, InStrRev(subREP, "\") + 1)
        tableNAME = ""
        Set subreportQUERY = getDATA("subREPORTS", Array(mainREP, subREP))
        With subreportQUERY
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
        thisrepo.LogonInfo(n) = "dsn=" + dsnF + ";dsq=" + dsnDSQ + ";uid=" + dsnUID + ";pwd=" + dsnPWD
        tableNAME = thisrepo.DataFiles(n)
        
        If tableNAME = "" Then
            If Not mainREPORT Then
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
        thisrepo.DataFiles(n) = dsnDSQ + tableNAME
    Next
    Exit Sub
    
errSTOP:
    MsgBox Err.description
    Exit Sub
End Sub

Public Sub Translate_SubReports()
'Procedure for process sub reports
    Dim repo As String
    Dim i As Integer
    mainREPORT = False
    With thisrepo
        For i = 0 To .GetNSubreports - 1
            repo = .GetNthSubreportName(i)
            .SubreportToChange = repo
            Call Translate_Reports(repo)
        Next
        mainREPORT = True
        .SubreportToChange = ""
    End With
End Sub


