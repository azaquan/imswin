VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connstring As String

Private Sub Form_Load()
'DataEnvironment1.cn.Open
Dim rsReceptList As New ADODB.Recordset

connstring = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Initial Catalog=sugeko;Data Source=imssql002;pwd=0egxpx4"

rsReceptList.Source = "select dis_mail from distribution where dis_npecode='PECT' and dis_id ='LD'"
rsReceptList.ActiveConnection = connstring
rsReceptList.Open , , adOpenStatic

Call sendOutlookEmailandFax("Later Delivery Report", "lateDeliveryReport", rsReceptList, "Late Delivery Report", "", "PROStream", "dis_mail")
Unload Me

End Sub
Public Function sendOutlookEmailandFax(reportNAME As String, ReportCaption As String, rsReceptList As ADODB.Recordset, Subject As String, attention As String, Optional sender As String, Optional FieldName As String)

Dim Params(1) As String

Dim I As Integer

Dim Attachments() As String

Dim Recepients() As String

Dim str As String

On Error GoTo errMESSAGE
     
     If rsReceptList.RecordCount > 0 Then
        
        attention = "Attention Please "
        
        Attachments = generateattachments(reportNAME, ReportCaption)
     
     If Len(Trim(FieldName)) = 0 Then
     
        Recepients = ToArrayFromRecordset(rsReceptList) ' This is just to keep with the old compatiblity. It thinks
                                                        ' the First field in the recordset are the email Addresses.
        
     Else
        
        Recepients = ToArray(rsReceptList, FieldName, I, str)
        
     End If
     
        Call WriteParameterFiles(Recepients, sender, Attachments, Subject, attention)
            
    Else
    
         MsgBox "No Recipients to Send", , "Imswin"
     
    End If
    
errMESSAGE:
    
    If Err.Number <> 0 Then
        
        MsgBox Err.Description
    
    End If

End Function
Public Function WriteParameterFiles(Recepients() As String, sender As String, Attachments() As String, Subject As String, attention As String)
 
 Dim l
 Dim x
 Dim y
 Dim I
 Dim Email() As String
 Dim fax() As String
 Dim rs As New ADODB.Recordset
 
 If Len(Trim(sender)) = 0 Then
 
    rs.Source = "select com_name from company where com_compcode = ( select psys_compcode from pesys where psys_npecode ='" & deIms.NameSpace & "')"
    rs.ActiveConnection = connstring
    rs.Open
    
    If rs.RecordCount > 0 Then
        If Len(rs("com_name") & "") > 0 Then sender = rs("com_name")
    End If
    rs.Close
    
    
End If
 
On Error GoTo errMESSAGE
  
'Splitting the address into Emails and Faxes.
 l = UBound(Recepients)
 
 
 
    x = 0
    y = 0
 
 
 For I = 0 To l
 
     If InStr(Recepients(I), "@") > 0 Then
       
       ReDim Preserve Email(x)
       Email(x) = Recepients(I)
       x = x + 1
       
    Else
      
       ReDim Preserve fax(y)
       fax(y) = Recepients(I)
       y = y + 1
       
    End If
       
       
       
 Next I

    If IsArrayLoaded(Email) Then 'M 02/23/02
    
        If Not (UBound(Email) = 0 And Email(0) = "") Then
            If UBound(Email) >= 0 Then Call WriteParameterFileEmail(Attachments, Email, Subject, sender, attention)
        End If
        
    End If                      'M 02/23/02
    
''    If IsArrayLoaded(fax) Then 'M 02/23/02
''
''        If Not (UBound(fax) = 0 And fax(0) = "") Then
''
''            If UBound(fax) >= 0 Then Call WriteParameterFileFax(Attachments, fax, Subject, sender, attention)
''
''        End If
''
''    End If 'M 02/23/02

errMESSAGE:
    
    If Err.Number <> 0 And Err.Number <> 9 Then
        
        MsgBox Err.Description
    
    Else
        
        Err.Clear
    
    End If

End Function

Public Function WriteParameterFileEmail(Attachments() As String, Recipients() As String, Subject As String, sender As String, attention As String) As Integer
On Error GoTo errMESSAGE
     Dim Filename As String
     Dim FileNumb As Integer
     Dim I As Integer, l As Integer
     Dim reports As String
     Dim recepientSTR As String

     Filename = "Email" & "-" & "sogeco" & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".txt"
     FileNumb = FreeFile

     Filename = "c:\IMSRequests\IMSRequests\" & Filename

     'Filename = ConnInfo.EmailParameterFolder & Filename

    For I = 0 To UBound(Recipients)
            recepientSTR = recepientSTR & Trim$(Recipients(I) & ";")
    Next

      I = 0

    For I = 0 To UBound(Attachments)
            reports = reports & Trim$(Attachments(I) & ";")
    Next

    Open Filename For Output As FileNumb

        Print #FileNumb, "[Email]"
        Print #FileNumb, "Recipients=" & recepientSTR
        Print #FileNumb, "Reports=" & reports
        Print #FileNumb, "Subject=" & Subject
        Print #FileNumb, "Sender=" & sender
        Print #FileNumb, "Attention=" & Trim$(attention)

    Close #FileNumb
    
WriteParameterFileEmail = 1

Exit Function
errMESSAGE:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function
Public Function generateattachments(reportNAME As String, ReportCaption As String) As String()
  
  Dim Attachments(0) As String
  
  Dim Filename As String
  
  Dim I As Integer
  
  Dim rs As New ADODB.Recordset
  Dim textstring As String
  Dim sa As New Scripting.FileSystemObject
  Dim t As Scripting.TextStream
  
  rs.Source = " select sup_name supplier, po_ponumb Ponumb, poi_liitnumb 'LineNo', poi_primreqdqty Qty from poitem "
rs.Source = rs.Source & " inner join po on po_ponumb =poi_ponumb and po_npecode =poi_npecode"
rs.Source = rs.Source & " inner join supplier on sup_code =po_suppcode and sup_npecode =po_npecode"
rs.Source = rs.Source & " where getdate() > poi_liitreqddate and poi_stasliit ='OP' and poi_stasdlvy <> 'RC'"

  rs.ActiveConnection = connstring
  rs.Open , , adOpenKeyset
  
  textstring = reportNAME & vbCrLf
  'textstring = textstring & "supplier" & "  ," & "po" & "   ," & "line#" & "    ," & "qty" & vbCrLf
  
  Do While Not rs.EOF
  
      textstring = textstring & rs.AbsolutePosition & "." & "   Supplier :" & rs("supplier")
      textstring = textstring & "   , PO :" & rs("ponumb")
      textstring = textstring & "   , Line# :" & rs("LineNo")
      textstring = textstring & "   , Qty :" & rs("qty") & vbCrLf
  
      rs.MoveNext
      
  Loop


On Error GoTo errMESSAGE
  
    Attachments(0) = "Report-" & ReportCaption & "-" & "SOGECO" & "-" & Replace(Replace(Replace(Now(), "/", "_"), " ", "-"), ":", "_") & ".rtf"
     
     Filename = "c:\IMSRequests\IMSRequests\OUT\" & Attachments(0)
     
     'Filename = ConnInfo.EmailOutFolder & Attachments(0)
     
     If sa.FileExists(Filename) Then sa.DeleteFile (Filename)
        
     If Not sa.FileExists(Filename) Then sa.CreateTextFile Filename
     Set t = sa.OpenTextFile(Filename, ForWriting)
     t.Write textstring
     
       
     generateattachments = Attachments
    
errMESSAGE:

    If Err.Number <> 0 Then
    
        MsgBox Err.Description
        
    End If

End Function
Public Function ToArrayFromRecordset(rs As ADODB.Recordset) As String()
Dim str() As String
Dim UpperBound As Integer

On Error GoTo Errhandler
    ReDim str(0)
    UpperBound = -1
    If rs Is Nothing Then Exit Function
        
    rs.MoveFirst
    Do While Not rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve str(UpperBound)
        If InStr(UCase(rs(0)), "INTERNET!") > 0 Then
            str(UpperBound) = Mid(rs(0), 10)
        Else
            str(UpperBound) = rs(0)
        End If
        rs.MoveNext
    Loop
    ToArrayFromRecordset = str
    Exit Function
    
Errhandler:
    Err.Raise Err.Number, Err.Description
    Err.Clear
End Function

Public Function ToArray(rs As ADODB.Recordset, ByVal FieldName As String, Optional UpperBound As Integer, Optional ByVal Filter As String) As String()
Dim BK As Variant
Dim str() As String
Dim OldFilter As Variant

On Error GoTo Errhandler
    ReDim str(0)
    UpperBound = -1
    If rs Is Nothing Then Exit Function
    
    BK = rs.Bookmark
    
    
    If Len(Filter) Then
        OldFilter = rs.Filter
        rs.Filter = adFilterNone
        rs.Filter = Filter
    End If
    
    rs.MoveFirst
    Do While Not rs.EOF
        UpperBound = UpperBound + 1
        ReDim Preserve str(UpperBound)
        str(UpperBound) = rs(FieldName)
        rs.MoveNext
    Loop
    
    ToArray = str
    
    If Len(Filter) Then rs.Filter = OldFilter
    rs.Bookmark = BK
    Exit Function
    
Errhandler:


    Err.Clear
End Function
Public Function IsArrayLoaded(ArrayToTest() As String) As Boolean

Dim x As Integer

On Error GoTo Errhandler

IsArrayLoaded = False

x = UBound(ArrayToTest)

IsArrayLoaded = True

Exit Function

Errhandler:

Err.Clear

End Function
