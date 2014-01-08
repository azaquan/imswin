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
Attribute VB_Exposed = True
Option Explicit

Public Function Trans(MessageCode) As String
'Function for retrieve direct texts for translation
Dim data As New ADODB.Recordset
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translationMESSAGES", Array(TR_LANGUAGE, MessageCode))
        If data.RecordCount > 0 Then
            Trans = data!msg_text
        Else
            Trans = ""
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
                        If Not .EOF Then VB.Forms(i).Caption = !msg_text
                            For j = 0 To VB.Forms(i).Controls.Count - 1
                                nameCONTROLs = VB.Forms(i).Controls(j).name
                                If TypeOf VB.Forms(i).Controls(j) Is LRNavigators.NavBar Or TypeOf VB.Forms(i).Controls(j) Is LRNavigators.LROleDBNavBar Then
                                    Set VB.Forms(i).Controls(j).ActiveConnection = cn
                                    VB.Forms(i).Controls(j).language = TR_LANGUAGE
                                End If
                                indexARRAY = -1
                                indexARRAY = VB.Forms(i).Controls(j).index
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


Public Sub Translate_Reports(repo As CrystalReport)
'Procedure for labels translations in every report
    Dim data As New ADODB.Recordset
    Dim i, j, n, x, xx As Integer
    Dim tableNAME As String
    Dim formNAME, controlNAME
    On Error GoTo errSTOP
    Set thisREPO = repo
    For i = 0 To VB.Forms.Count - 1
        If VB.Forms(i).name = "MDI_IMS" Then
            For j = 0 To VB.Forms(i).Controls.Count - 1
                If VB.Forms(i).Controls(j).name = "CrystalReport1" Then
                    Set thisREPO = VB.Forms(i).Controls(j)
                End If
            Next
        End If
    Next
    
    If TR_LANGUAGE <> "*" And TR_LANGUAGE <> "" Then
        Set data = getDATA("translatorCONTROLS", Array(TR_LANGUAGE, repoNAME))
        With data
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
            
    If mainREPORT Then
        x = thisREPO.RetrieveLogonInfo - 1
        x = thisREPO.RetrieveDataFiles - 1
    Else
        Dim subreportQUERY As New ADODB.Recordset
        Dim sql, mainREP, subREP
        mainREP = repoNAME
        mainREP = Mid(mainREP, InStrRev(mainREP, "\") + 1)
        subREP = repoNAME
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
        thisREPO.LogonInfo(n) = "dsn=" + connINFO.Dsnname + ";dsq=" _
        & connINFO.InitCatalog + ";uid=" + connINFO.UId + ";pwd=" + connINFO.Pwd
        tableNAME = thisREPO.DataFiles(n)
        
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
        thisREPO.DataFiles(n) = connINFO.InitCatalog + tableNAME
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
    With thisREPO
        For i = 0 To .GetNSubreports - 1
            repo = .GetNthSubreportName(i)
            .SubreportToChange = repo
            Call Translate_Reports(repo)
        Next
        mainREPORT = True
        .SubreportToChange = ""
    End With
End Sub
Public Function getDATA(Access, Parameters) As ADODB.Recordset
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = Access
        Set getDATA = .Execute(1, Parameters)
    End With
End Function


