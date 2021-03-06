VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mRsHeader As ADODB.Recordset
Private MNamespace As String
Private MPonumb As String
Private miutility As Iutility
Private mCheckLoad As Boolean
Private mRecordCount As Integer
Private cancel As Boolean
Private mValueChanged As Boolean
Private mAutoGeneratedNo As String 'AM

Private MStatus As Status
Private mErrors As Errors
Private mError As ErrorDetl
Private mCount As Integer

Public Event BeforeAddNEW(cancel As Boolean)
Public Event AfterAddNEW()
Public Event BeforeDeleteHeader(cancel As Boolean)
Public Event AfterDeleteHeader()


''Public Enum Status
''  ADSTATUSUnknow = 0
''  ADSTATUSMODIFIED = 1
''  ADSTATUSINSERT = 2
''End Enum
Public Property Get POnumb() As String
  POnumb = mRsHeader!PO_ponumb
End Property

Public Property Let POnumb(ByVal indata As String)
If Not mRsHeader!PO_ponumb = Trim$(indata) Then mValueChanged = True
  mRsHeader!PO_ponumb = Trim$(indata)
End Property

Public Property Get NpeCode() As String
  NpeCode = mRsHeader!po_Npecode
End Property

Public Property Let NpeCode(ByVal indata As String)
  
If Not mRsHeader!po_Npecode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_Npecode = Trim$(indata)
End Property

Public Property Get buyr() As String
  buyr = mRsHeader!po_buyr
End Property

Public Property Let buyr(ByVal indata As String)
If Not mRsHeader!po_buyr = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_buyr = Trim$(indata)
End Property

Public Property Get Createdate() As String
  Createdate = IIf(IsNull(mRsHeader!po_date), "", mRsHeader!po_date)
End Property

Public Property Let Createdate(ByVal indata As String)
  indata = Trim$(indata)
  If Not mRsHeader!po_date = Trim$(indata) Then mValueChanged = True
    mRsHeader!po_date = IIf(IsEmpty(indata), Null, CDate(indata))
End Property

Public Property Get apprby() As String
  apprby = IIf(IsNull(mRsHeader!po_apprby), "", mRsHeader!po_apprby)
End Property

Public Property Let apprby(ByVal indata As String)
  If Not mRsHeader!po_apprby = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_apprby = Trim$(indata)
End Property

Public Property Get totacost() As Double
  totacost = mRsHeader!po_totacost
End Property

Public Property Let totacost(ByVal indata As Double)
  If Not mRsHeader!po_totacost = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_totacost = Trim$(indata)
End Property

Public Property Get tbuf() As String
   tbuf = IIf(IsNull(mRsHeader!po_tbuf), "", mRsHeader!po_tbuf)
End Property

Public Property Let tbuf(ByVal indata As String)
  If Not mRsHeader!po_tbuf = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_tbuf = Trim$(indata)
End Property

Public Property Get suppcode() As String
  suppcode = mRsHeader!po_suppcode
End Property

Public Property Let suppcode(ByVal indata As String)
  If Not mRsHeader!po_suppcode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_suppcode = Trim$(indata)
End Property

Public Property Get docutype() As String
  
  docutype = mRsHeader!po_docutype
End Property

Public Property Let docutype(ByVal indata As String)
  If Not mRsHeader!po_docutype = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_docutype = Trim$(indata)
End Property

Public Property Get priocode() As String
  priocode = mRsHeader!po_priocode
End Property

Public Property Let priocode(ByVal indata As String)
  If Not mRsHeader!po_priocode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_priocode = Trim$(indata)
End Property

Public Property Get Currcode() As String
  Currcode = mRsHeader!po_Currcode
End Property

Public Property Let Currcode(ByVal indata As String)
  If Not mRsHeader!po_Currcode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_Currcode = Trim$(indata)
End Property


Public Property Get reqddelvdate() As Date
  reqddelvdate = mRsHeader!po_reqddelvdate
End Property

Public Property Let reqddelvdate(ByVal indata As Date)
   If Not mRsHeader!po_reqddelvdate = Trim$(indata) Then mValueChanged = True
   mRsHeader!po_reqddelvdate = indata
End Property

Public Property Get shipcode() As String
  shipcode = IIf(IsNull(mRsHeader!po_shipcode), "", mRsHeader!po_shipcode)
End Property

Public Property Let shipcode(ByVal indata As String)
  
  If Not mRsHeader!po_shipcode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_shipcode = Trim$(indata)
End Property

Public Property Get datesent() As String
   datesent = IIf(IsNull(mRsHeader!po_datesent), "", mRsHeader!po_datesent)
  
  'datesent = IIf(IsNull(mRsHeader!po_datesent), "", mRsHeader!po_datesent)
  
End Property

Public Property Let datesent(ByVal indata As String)

 indata = Trim$(indata)
  If Not mRsHeader!po_datesent = Trim$(indata) Then mValueChanged = True
   If Len(Trim$(indata)) = 0 Then
       mRsHeader!po_datesent = Null
   Else
      mRsHeader!po_datesent = CDate(indata)
   End If
    'mRsHeader!po_datesent = IIf(IsEmpty(indata), Null, CDate(indata))

 ' If Not mRsHeader!po_datesent = (Trim$(indata)) Then mValueChanged = True
 '   mRsHeader!po_datesent = Trim(indata)
  'mRsHeader!po_datesent = IIf(Len(Trim$(indata)) = 0, Null, CDate(Trim$(indata)))
End Property

Public Property Get stas() As String
  stas = mRsHeader!po_stas
End Property

Public Property Let stas(ByVal indata As String)
  If Not mRsHeader!po_stas = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_stas = Trim$(indata)
End Property

Public Property Get orig() As String
  orig = IIf(IsNull(mRsHeader!po_orig), "", mRsHeader!po_orig)
End Property

Public Property Let orig(ByVal indata As String)
  If Not mRsHeader!po_orig = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_orig = Trim$(indata)
End Property

Public Property Get site() As String
  site = mRsHeader!po_site
End Property

Public Property Let site(ByVal indata As String)
  If Not mRsHeader!po_site = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_site = Trim$(indata)
End Property

Public Property Get chrgto() As String
  chrgto = IIf(IsNull(mRsHeader!po_chrgto), "", mRsHeader!po_chrgto)
End Property

Public Property Let chrgto(ByVal indata As String)
  If Not mRsHeader!po_chrgto = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_chrgto = Trim$(indata)
End Property

Public Property Get sendby() As String
  sendby = mRsHeader!po_sendby
End Property

Public Property Let sendby(ByVal indata As String)
  If Not mRsHeader!po_sendby = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_sendby = Trim$(indata)
End Property
Public Property Get SuppContactName() As String
SuppContactName = IIf(IsNull(mRsHeader!po_suppconta), "", mRsHeader!po_suppconta)
End Property

Public Property Let SuppContactName(ByVal vNewValue As String)
mRsHeader!po_suppconta = IIf(Len(vNewValue) = 0, Null, vNewValue)
End Property

Public Property Get SuppContaPH() As String
SuppContaPH = IIf(IsNull(mRsHeader!po_SuppContaPH), "", mRsHeader!po_SuppContaPH)
End Property

Public Property Let SuppContaPH(ByVal vNewValue As String)
mRsHeader!po_SuppContaPH = IIf(Len(vNewValue) = 0, Null, vNewValue)
End Property

Public Property Get confordr() As Boolean
  confordr = mRsHeader!po_confordr
End Property

Public Property Let confordr(ByVal indata As Boolean)
  If Not mRsHeader!po_confordr = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_confordr = indata
End Property

Public Property Get quotnumb() As String
  quotnumb = mRsHeader!po_quotnumb
End Property

Public Property Let quotnumb(ByVal indata As String)
  If Not mRsHeader!po_quotnumb = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_quotnumb = Trim$(indata)
End Property

Public Property Get forwr() As Boolean
  forwr = mRsHeader!po_forwr
End Property

Public Property Let forwr(ByVal indata As Boolean)
  If Not mRsHeader!po_forwr = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_forwr = indata
End Property

Public Property Get catecode() As String
  catecode = IIf(IsNull(mRsHeader!po_catecode), "", mRsHeader!po_catecode)
End Property

Public Property Let catecode(ByVal indata As String)
  If Not mRsHeader!po_catecode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_catecode = Trim$(indata)
End Property

Public Property Get shipto() As String
  
  shipto = IIf(IsNull(mRsHeader!po_shipto), "", mRsHeader!po_shipto)
End Property

Public Property Let shipto(ByVal indata As String)
  If Not mRsHeader!po_shipto = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_shipto = Trim$(indata)
End Property

Public Property Get stasdelv() As String
  stasdelv = mRsHeader!po_stasdelv
End Property

Public Property Let stasdelv(ByVal indata As String)
  If Not mRsHeader!po_stasdelv = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_stasdelv = Trim$(indata)
End Property

Public Property Get Stasship() As String
  Stasship = mRsHeader!po_stasship
End Property

Public Property Let Stasship(ByVal indata As String)
  If Not mRsHeader!po_stasship = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_stasship = Trim$(indata)
End Property

Public Property Get Stasinvt() As String
  Stasinvt = mRsHeader!po_stasinvt
End Property

Public Property Let Stasinvt(ByVal indata As String)
  If Not mRsHeader!po_stasinvt = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_stasinvt = Trim$(indata)
End Property

Public Property Get revinumb() As Integer
  revinumb = mRsHeader!po_revinumb
End Property

Public Property Let revinumb(ByVal indata As Integer)
  If Not mRsHeader!po_revinumb = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_revinumb = Trim$(indata)
End Property
Public Property Get Originalrevinumb() As Integer
  Originalrevinumb = mRsHeader("po_revinumb").OriginalValue
End Property
Public Property Get reqddelvflag() As Boolean
  reqddelvflag = mRsHeader!po_reqddelvflag
End Property

Public Property Let reqddelvflag(ByVal indata As Boolean)
  If Not mRsHeader!po_reqddelvflag = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_reqddelvflag = indata
End Property

Public Property Get srvccode() As String
 srvccode = IIf(IsNull(mRsHeader!po_srvccode), "", mRsHeader!po_srvccode)
End Property

Public Property Let srvccode(ByVal indata As String)
  If Not mRsHeader!po_srvccode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_srvccode = Trim$(indata)
End Property

Public Property Get invloca() As String
  invloca = IIf(IsNull(mRsHeader!po_invloca), "", mRsHeader!po_invloca)
End Property

Public Property Let invloca(ByVal indata As String)
  If Not mRsHeader!po_invloca = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_invloca = Trim$(indata)
End Property

Public Property Get fromstckmast() As Boolean
  fromstckmast = mRsHeader!po_fromstckmast
End Property

Public Property Let fromstckmast(ByVal indata As Boolean)
  If Not mRsHeader!po_fromstckmast = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_fromstckmast = indata
End Property

Public Property Get daterevi() As Variant
  daterevi = IIf(IsNull(mRsHeader!po_daterevi), "", mRsHeader!po_daterevi)
End Property

Public Property Let daterevi(ByVal indata As Variant)
  If Not mRsHeader!po_daterevi = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_daterevi = Trim(indata)
End Property

Public Property Get taccode() As String
  taccode = mRsHeader!po_taccode
End Property

Public Property Let taccode(ByVal indata As String)
  If Not mRsHeader!po_taccode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_taccode = Trim$(indata)
End Property

Public Property Get freigforwr() As Boolean
freigforwr = mRsHeader!PO_freigforwr
End Property

Public Property Let freigforwr(ByVal vNewValue As Boolean)
mRsHeader!PO_freigforwr = vNewValue
End Property

Public Property Get termcode() As String
  termcode = mRsHeader!po_termcode
End Property

Public Property Let termcode(ByVal indata As String)
  If Not mRsHeader!po_termcode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_termcode = Trim$(indata)
End Property

Public Property Get creauser() As String
  creauser = mRsHeader!po_creauser
End Property

Public Property Let creauser(ByVal indata As String)
  If Not mRsHeader!po_creauser = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_creauser = Trim$(indata)
End Property

Public Property Get modiuser() As String
  modiuser = mRsHeader!po_modiuser
End Property

Public Property Let modiuser(ByVal indata As String)
  If Not mRsHeader!po_modiuser = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_modiuser = Trim$(indata)
End Property

Public Property Get compcode() As String
  compcode = mRsHeader!po_compcode
End Property

Public Property Let compcode(ByVal indata As String)
  If Not mRsHeader!po_compcode = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_compcode = Trim$(indata)
End Property

Public Property Get MyStatus() As Status
  MyStatus = MStatus
End Property

Public Property Let MyStatus(ByVal indata As Status)
   MStatus = indata
End Property
Public Property Get USExport() As Boolean
  USExport = IIf(IsNull(mRsHeader!po_usexport), 0, mRsHeader!po_usexport)
End Property

Public Property Let USExport(ByVal indata As Boolean)
  If Not mRsHeader!po_usexport = Trim$(indata) Then mValueChanged = True
  mRsHeader!po_usexport = indata
End Property
Public Property Get Count() As Integer
If mCheckLoad = False Then mCheckLoad = LoadCollection

If mCheckLoad = True Then
Count = mRsHeader.CursorType = adOpenDynamic
End If
End Property
'===========================================================
Public Function Configure(Obj As Iutility) As Boolean
  Set miutility = Obj
End Function

Public Function Save() As Boolean
  Dim IsItANewPO As Integer
 On Error GoTo handler
 Save = False
  
  If mValueChanged Then
  
  IsItANewPO = mRsHeader.EditMode 'AM
  
     mRsHeader.Update
     
     'This means that this is the creation mode and the user is using the AutoGenerated number.
    'In such a case the AUTO number has to be updated in the Chrono table

    'If IsItANewPO = 2 And Ponumb = mAutoGeneratedNo Then UpdateAutoNumberValue 'AM
    
    If IsItANewPO = 2 Then UpdateAutoNumberValue 'AM
    
    If IsItANewPO = 2 Then 'AM
    
        If UCase(Trim(docutype)) = "R" Then buyr = ""
        
    End If

    Call mRsHeader.Find("Po_ponumb='" & POnumb & "'", , adSearchForward)
     
    If mRsHeader.Eof = False Then
     
    End If
  
  End If
     
     Save = True
     
     Exit Function
     
handler:
   
    'Filling The Collection with The Error
    
    Err.Clear
    
End Function
Public Function Move(POnumb As String) As Boolean
On Error GoTo handler
Move = False

If mCheckLoad = False Then mCheckLoad = LoadCollection

If mCheckLoad = True Then
   mRsHeader.MoveFirst
   Call mRsHeader.Find("Po_ponumb='" & POnumb & "'", , adSearchForward)
   
   If mRsHeader.Eof = False Then
   Move = True
   'Call FillMeWithTheCurrentRecord(mRsHeader)
   End If
End If
   
   
   Exit Function
   
handler:

   Err.Clear
End Function



Public Function AddNew() As Boolean
 
  AddNew = False
  
  On Error GoTo handler
  
  RaiseEvent BeforeAddNEW(cancel)
    
           mRsHeader.CancelUpdate
           
           mRsHeader.AddNew
           
          If SetTheNewHeader = False Then 'AM
            
                mRsHeader.CancelUpdate 'AM
                
                Exit Function 'AM
                
          End If 'AM
          
  cancel = True
  
  AddNew = True
  
  Exit Function

handler:
     

    Err.Clear
End Function

Private Function LoadCollection() As Boolean

Dim CmdHead As ADODB.Command
'Dim RsHead As ADODB.Recordset
Dim RecordCount As Integer

LoadCollection = False
On Error GoTo handler
Set CmdHead = New ADODB.Command

    With CmdHead
       .CommandType = adCmdStoredProc
       .CommandText = "GetAllPO"
    
       Set .ActiveConnection = miutility.Connection
       .Parameters.Append .CreateParameter("", adVarChar, adParamInput, 5, miutility.Namespace)
       .Parameters.Append .CreateParameter("", adVarChar, adParamOutput, 3, RecordCount)
    End With
    
   Set mRsHeader = CmdHead.Execute
       mRsHeader.Close
       mRsHeader.CursorLocation = adUseClient
       mRsHeader.Open , , adOpenKeyset, adLockOptimistic
   'LoadCollection = FillMeWithTheCurrentRecord(RsHead)

   LoadCollection = True
   mCheckLoad = True
   Exit Function
   
handler:
   Err.Raise Err.Number, "IMSPODLL.POHEADERS.Loadcollection", Err.Description
   Err.Clear
End Function

Private Function FillmeWithTheCurrentRecord(RsHead As ADODB.Recordset) As Boolean
Dim mPoheader As PoHeader

FillmeWithTheCurrentRecord = False
On Error GoTo handler

If Not RsHead.Eof Then
     
     mPoheader.apprby = IIf(IsNull(RsHead!po_apprby), "", RsHead!po_apprby)
     mPoheader.buyr = RsHead!po_buyr
     mPoheader.catecode = IIf(IsNull(RsHead!po_catecode), "", RsHead!po_catecode)
     mPoheader.chrgto = IIf(IsNull(RsHead!po_chrgto), "", RsHead!po_chrgto)
     mPoheader.compcode = RsHead!po_compcode
     mPoheader.confordr = RsHead!po_confordr
     mPoheader.Createdate = IIf(IsNull(RsHead!po_date), Now(), RsHead!po_date)
     mPoheader.creauser = RsHead!po_creauser
     mPoheader.Currcode = RsHead!po_Currcode
     mPoheader.daterevi = IIf(IsNull(RsHead!po_daterevi), "", RsHead!po_daterevi)
     mPoheader.datesent = IIf(IsNull(RsHead!po_datesent), "", RsHead!po_datesent)
     mPoheader.docutype = RsHead!po_docutype
     mPoheader.forwr = RsHead!po_forwr
     mPoheader.fromstckmast = RsHead!po_fromstckmast
     mPoheader.invloca = IIf(IsNull(RsHead!po_invloca), "", RsHead!po_invloca)
     mPoheader.modiuser = RsHead!po_modiuser
     mPoheader.NpeCode = RsHead!po_Npecode
     mPoheader.orig = IIf(IsNull(RsHead!po_orig), "", RsHead!po_orig)
     mPoheader.POnumb = RsHead!PO_ponumb
     mPoheader.priocode = RsHead!po_priocode
     mPoheader.quotnumb = IIf(IsNull(RsHead!po_quotnumb), "", RsHead!po_catecode)
     mPoheader.reqddelvdate = RsHead!po_reqddelvdate
     mPoheader.reqddelvflag = RsHead!po_reqddelvflag
     mPoheader.revinumb = RsHead!po_revinumb
     mPoheader.sendby = IIf(IsNull(RsHead!po_sendby), "", RsHead!po_sendby)
     mPoheader.shipcode = RsHead!po_shipcode
     mPoheader.shipto = IIf(IsNull(RsHead!po_shipto), "", RsHead!po_shipto)
     mPoheader.site = RsHead!po_site
     mPoheader.srvccode = IIf(IsNull(RsHead!po_srvccode), "", RsHead!po_srvccode)
     mPoheader.stas = RsHead!po_stas
     mPoheader.stasdelv = RsHead!po_stasdelv
     mPoheader.Stasinvt = RsHead!po_stasinvt
     mPoheader.Stasship = RsHead!po_stasship
     mPoheader.suppcode = RsHead!po_suppcode
     mPoheader.taccode = RsHead!po_taccode
     mPoheader.tbuf = IIf(IsNull(RsHead!po_tbuf), "", RsHead!po_tbuf)
     mPoheader.termcode = RsHead!po_termcode
     mPoheader.totacost = IIf(IsNull(RsHead!po_totacost), 0, RsHead!po_totacost)
     mPoheader.MyStatus = ADSTATUSMODIFIED
    
    

     FillmeWithTheCurrentRecord = True
End If
 Exit Function

handler:
  
  Err.Clear
End Function


Public Function MoveFirst() As Boolean
On Error GoTo handler
MoveFirst = False
  If Not mRsHeader.Bof Then
    mRsHeader.CancelUpdate
     mRsHeader.MoveFirst
   '  If Not mRsHeader.Bof Then FillmeWithTheCurrentRecord (mRsHeader)
  MoveFirst = True
  End If
  
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
On Error GoTo handler
MoveNext = False
  If Not mRsHeader.Eof Then
     mRsHeader.CancelUpdate
     mRsHeader.MoveNext
     If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Eof = True And mRsHeader.Bof <> True Then mRsHeader.MoveLast
     'If Not mRsHeader.Eof Then Call FillmeWithTheCurrentRecord(mRsHeader)
   If Not mRsHeader.Eof Then MoveNext = True
 
  End If
  
 Exit Function
handler:
  
End Function

Public Function MovePrevious() As Boolean
On Error GoTo handler
MovePrevious = False
  If Not mRsHeader.Bof Then
    mRsHeader.CancelUpdate
     mRsHeader.MovePrevious
      If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
     If Not mRsHeader.Bof Then MovePrevious = True
   
  End If

Exit Function
handler:
  Err.Clear
End Function

Public Function MoveLast() As Boolean
On Error GoTo handler
MoveLast = False
  If Not mRsHeader.Eof Then
    mRsHeader.CancelUpdate
     mRsHeader.MoveLast
    ' If Not mRsHeader.Eof Then FillmeWithTheCurrentRecord (mRsHeader)
     MoveLast = True
  End If

  Exit Function
handler:
  
End Function

''Private Sub SetTheNewHeader()
''
''   MStatus = ADSTATUSINSERT
''   cancel = False
''    mRsHeader!po_stas = "OH"
''           mRsHeader!po_stasdelv = "NR"
''           mRsHeader!po_stasship = "NS"
''           mRsHeader!po_stasinvt = "NI"
''End Sub

Private Function SetTheNewHeader() As Boolean 'AM

On Error GoTo Errhandler 'AM

SetTheNewHeader = False 'AM
            
           mAutoGeneratedNo = "" 'AM
            
           MStatus = ADSTATUSINSERT
           cancel = False
           mRsHeader!po_stas = "OH"
           mRsHeader!po_stasdelv = "NR"
           mRsHeader!po_stasship = "NS"
           mRsHeader!po_stasinvt = "NI"
                  
           mAutoGeneratedNo = GenerateATempPONumb ' GenerateAutoNumber 'AM
           
           If Len(Trim$(mAutoGeneratedNo)) > 0 Then 'AM
                  
                mRsHeader!PO_ponumb = mAutoGeneratedNo 'AM
           
           Else 'AM
           
                Exit Function 'AM
                  
           End If 'AM
           
SetTheNewHeader = True 'AM

Exit Function 'AM

Errhandler: 'AM

'n
End Function

Public Property Get Errors() As Errors
Set Errors = mErrors
End Property



Private Sub HandleErrors(ErrNumber As Long, ErrDesc As String, ErrLocation As String)
 
 If mErrors Is Nothing Then Set mErrors = New Errors
   
       Set mError = New ErrorDetl
    
       mError.ErrorDesc = ErrDesc
       mError.ErrorLocation = ErrLocation
       mError.ErrorNumber = ErrNumber
       
  mErrors.Add mError
       
End Sub

Public Property Get Eof() As Boolean
If mCheckLoad = False Then mCheckLoad = LoadCollection
If mCheckLoad = True Then
  Eof = mRsHeader.Eof
End If
End Property

Public Property Get Bof() As Boolean
If mCheckLoad = False Then mCheckLoad = LoadCollection
Bof = mRsHeader.Bof
End Property
''''Public Property Get Counta() As Integer
''''If mCheckLoad = False Then mCheckLoad = LoadCollection
''''
''''If mCheckLoad = True Then
''''Counta = mCount
''''End If
''''End Property

Public Function CancelUpdate() As Boolean
On Error GoTo handler
If mRsHeader.EditMode = adEditAdd Then
   mRsHeader.CancelUpdate
   mAutoGeneratedNo = "" 'AM
End If
Exit Function
   
handler:
   Err.Clear
End Function


Public Function Requery()

 On Error GoTo handler
 
  mRsHeader.Requery
  
  Exit Function
  
handler:

  Err.Raise Err.Number, , Err.Description
  
End Function


Public Function Delete(POnumb As String) As Boolean
Dim cancel As Boolean
On Error GoTo handler
Dim Cmd As ADODB.Command
 Delete = False
RaiseEvent BeforeDeleteHeader(cancel)
If cancel = False Then
   
   Set Cmd = New ADODB.Command
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "deletepo"
    With Cmd
     .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, miutility.Namespace)
     .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, POnumb)
    End With
        
    
    Cmd.ActiveConnection = miutility.Connection
    Cmd.Execute
    If Not Err.Number > 0 Then Exit Function
        RaiseEvent AfterDeleteHeader
        
Else

   cancel = False
           
End If
Delete = True
Exit Function
handler:
MsgBox "Error Occurred During Deleting the PO Header.Err.desc -- " & Err.Description
Err.Clear
End Function

Public Function UpdateAutoNumberValue() As Boolean 'AM

    Dim Cmd As New ADODB.Command 'AM
    
    Cmd.CommandType = adCmdStoredProc 'AM
    Cmd.CommandText = "UpdateAutoNumberForPO" 'AM
    Cmd.ActiveConnection = miutility.Connection 'AM
   
   With Cmd 'AM
        
        .Parameters.Append .CreateParameter("@ChronoCode", adVarChar, adParamInput, 5, GChronoCodeForPO)  'AM
        
        .Parameters.Append .CreateParameter("@DOCUTYPE", adVarChar, adParamInput, 5, docutype) 'AM
        
        .Parameters.Append .CreateParameter("@NPECODE", adVarChar, adParamInput, 5, miutility.Namespace) 'AM
        
        .Parameters.Append .CreateParameter("@AUTOGENERATEDNUMBER", adVarChar, adParamInput, 15, POnumb)  'AM
        
        .Parameters.Append .CreateParameter("ServiceCode", adVarChar, adParamInput, 15, mRsHeader("po_srvccode") & "") 'AM
   
        .Execute 'AM
   
   End With 'AM
   
End Function 'AM
'This Function generates a unique number.
Public Function GenerateATempPONumb() As String
Dim str As String
str = "_" & Replace(Replace(Replace(Now, "/", ""), " ", ""), ":", "")
GenerateATempPONumb = Mid(str, 1, Len(str) - 2)

End Function

Public Function LoadFromRequsition(RequsitionNo As String, TemporaryPoNumb As String, CurrentUser As String, ErrorIfany As String) As Boolean
Dim RsRequisition As ADODB.Recordset
Dim field As ADODB.field
On Error GoTo Errhandler
LoadFromRequsition = True

If Len(Trim(TemporaryPoNumb)) = 0 Then Exit Function

Set RsRequisition = mRsHeader.Clone

RsRequisition.MoveFirst

Call RsRequisition.Find("Po_ponumb='" & RequsitionNo & "'", , adSearchForward)

If RsRequisition.Eof = False Then

    
    TemporaryPoNumb = mRsHeader("po_ponumb")
    
    For Each field In RsRequisition.Fields
    
            mRsHeader(field.Name) = field.Value
            Debug.Print field.Name
    Next
   
    mRsHeader("po_ponumb") = TemporaryPoNumb
    mRsHeader("po_revinumb") = 0
    mRsHeader("po_daterevi") = Null
    mRsHeader("po_buyr") = CurrentUser
    mRsHeader("po_apprby") = Null
    mRsHeader("po_datesent") = Null
    mRsHeader("po_date") = Now
    mRsHeader("po_stas") = "OH"
    mRsHeader("po_stasdelv") = "NR"
    mRsHeader("po_stasship") = "NS"
    mRsHeader("po_stasinvt") = "NI"
    mRsHeader("po_reqddelvdate") = DateAdd("d", 1, Now())
    'mRsHeader("po_AssignreqDate") = Null
    
    
 End If
   
Set RsRequisition = Nothing

LoadFromRequsition = False

Exit Function
Errhandler:

ErrorIfany = Err.Description
Err.Clear
End Function
