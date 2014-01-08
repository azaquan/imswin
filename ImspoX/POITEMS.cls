VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "POITEMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains a Collection of POITEM"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"POITEM"
Attribute VB_Ext_KEY = "Member0" ,"POITEM"
Option Explicit

Private mIsItemLoaded As Boolean

'''Public Enum Status
'''  ADSTATUSUnknow = 0
'''  ADSTATUSMODIFIED = 1
'''  ADSTATUSINSERT = 2
'''End Enum

Private mRsPOITEM As ADODB.Recordset

Private MStatus  As Status

Private MNamespace As String

Private MPonumb As String

Private mCount As Integer

Private mRsHeader As ADODB.Recordset

Private mValueChanged As Boolean
Private miutility As Iutility
Public mCancel As Boolean
Private mErrors As Errors
Private mError As ErrorDetl

Public Event BeforeAddNEW(cancel As Boolean)
Public Event BeforeDeleteALLPOITEM(cancel As Boolean)
Public Event AfterDeletePOITEMS()
Public Event AfterDeleteCurrentLI()
Public Event AfterCancelUpdate()


'Private POITEM As POITEM ' Created from: Private POITEM As POITEM

'''''''Private mPonumb As String ' Created from: Private Ponumb
'''''''
'''''''Private mNpecode As String ' Created from: Private Npecode
'''''''
'''''''Private mlineNumb As Integer ' Created from: Private lineNumb
'''''''
'''''''Private mUnitOfPurch As String ' Created from: Private UnitOfPurch
'''''''
'''''''Private mDescription As String ' Created from: Private Description
'''''''
'''''''Private mPrimreqdqty As Double ' Created from: Private Primreqdqty
'''''''
'''''''Private mPrimuom As String ' Created from: Private Primuom
'''''''
'''''''Private mSecoreqdqty As Double ' Created from: Private Secoreqdqty
'''''''
'''''''Private mSecouom As String ' Created from: Private Secouom
'''''''
'''''''Private mPrimUnitprice As Double ' Created from: Private PrimUnitprice
'''''''
'''''''Private mSecUnitPrice As Double ' Created from: Private SecUnitPrice
'''''''
'''''''Private mPriTotaprice As Double ' Created from: Private PriTotaprice
'''''''
'''''''Private mSecTotaprice As Double ' Created from: Private SecTotaprice
'''''''
'''''''Private mPriQtydlvd As Double ' Created from: Private PriQtydlvd
'''''''
'''''''Private mSecqtydlvd As Double ' Created from: Private Secqtydlvd
'''''''
'''''''Private mPriQtyship As Double ' Created from: Private PriQtyship
'''''''
'''''''Private mSecQtyship As Double ' Created from: Private SecQtyship
'''''''
'''''''Private mPriQtyinvt As Double ' Created from: Private PriQtyinvt
'''''''
'''''''Private mSecqtyinvt As Double ' Created from: Private Secqtyinvt
'''''''
'''''''Private mComm As String ' Created from: Private Comm
'''''''
'''''''Private mRequnumb As String ' Created from: Private Requnumb
'''''''
'''''''Private mRequliitnumb As String ' Created from: Private Requliitnumb
'''''''
'''''''Private mQuotnum As String ' Created from: Private Quotnum
'''''''
'''''''Private mQuotliitnumb As String ' Created from: Private Quotliitnumb
'''''''
'''''''Private mLocatax As Double ' Created from: Private Locatax
'''''''
'''''''Private mRemk As String ' Created from: Private Remk
'''''''
'''''''Private mSerlnumb As String ' Created from: Private Serlnumb
'''''''
'''''''Private mManupartnumb As String ' Created from: Private Manupartnumb
'''''''
'''''''Private mLiitreqddate As Date ' Created from: Private Liitreqddate
'''''''
'''''''Private mLiitrelsdate As Date ' Created from: Private Liitrelsdate
'''''''
'''''''Private mStarrendate As Date ' Created from: Private Starrendate
'''''''
'''''''Private mEndrentdate As Date ' Created from: Private Endrentdate
'''''''
'''''''Private mStasliit As String ' Created from: Private Stasliit
'''''''
'''''''Private mStasdlvy As String ' Created from: Private Stasdlvy
'''''''
'''''''Private mStasship As String ' Created from: Private Stasship
'''''''
'''''''Private mStasinvt As String ' Created from: Private Stasinvt
'''''''
'''''''Private mCurrcode As String ' Created from: Private Currcode
'''''''
'''''''Private mAfe As String ' Created from: Private Afe
'''''''
'''''''Private mCustcate As String ' Created from: Private Custcate
'''''''
'''''''Private mLastinvcnumb As String ' Created from: Private Lastinvcnumb
'''''''
'''''''Private mQtytobedlvd As Double ' Created from: Private Qtytobedlvd



Public Sub Class_Initialize()
  
   MStatus = ADSTATUSINSERT
   End Sub


Public Property Get MyStatus() As Status
 MyStatus = MStatus
End Property

Public Property Let MyStatus(ByVal indata As Status)
MStatus = indata
End Property

Public Property Get POnumb() As String
POnumb = MPonumb
End Property

Public Property Let POnumb(ByVal indata As String)
'If Not mRsHeader!poi_ponumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_ponumb = (Trim$(indata))
MPonumb = mRsHeader!poi_ponumb 'AM
End Property

Public Property Get Namespace() As String
'Namespace = MNamespace
Namespace = miutility.Namespace
End Property
Public Property Let Namespace(ByVal indata As String)
'Namespace = MNamespace
 mRsHeader!poi_npecode = indata
End Property

Public Property Get lineNumb() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
lineNumb = mRsHeader!poi_liitnumb
End Property

Public Property Let lineNumb(ByVal indata As Integer)
'If Not mRsHeader!poi_liitnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_liitnumb = (Trim$(indata))
End Property

Public Property Get UnitOfPurch() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
UnitOfPurch = IIf(IsNull(mRsHeader!poi_UnitofP), "", mRsHeader!poi_UnitofP)
End Property

Public Property Get USExport() As Boolean
  USExport = mRsHeader!poi_usexport
End Property

Public Property Let USExport(ByVal indata As Boolean)
  If Not mRsHeader!poi_usexport = Trim$(indata) Then mValueChanged = True
  mRsHeader!poi_usexport = indata
End Property

Public Property Let Eccnlicsreq(ByVal indata As Boolean)
  If Not mRsHeader!poi_Eccnlicsreq = Trim$(indata) Then mValueChanged = True
  mRsHeader!poi_Eccnlicsreq = indata
End Property

Public Property Get Eccnlicsreq() As Boolean
  Eccnlicsreq = mRsHeader!poi_Eccnlicsreq
End Property

Public Property Let Eccnid(ByVal indata As String)
  If Not mRsHeader!poi_Eccnid = Trim$(indata) Then mValueChanged = True
  mRsHeader!poi_Eccnid = indata
End Property

Public Property Get Sourceofinfoid() As String
  Sourceofinfoid = mRsHeader!poi_sourceid & ""
End Property
Public Property Let Sourceofinfoid(ByVal indata As String)
  If Not mRsHeader!poi_sourceid = Trim$(indata) Then mValueChanged = True
  mRsHeader!poi_sourceid = indata
End Property

Public Property Get Eccnid() As String
  Eccnid = mRsHeader!poi_Eccnid & ""
End Property
''Public Property Let Eccnno(ByVal indata As String)
''  If Not mRsHeader!Eccn_no = Trim$(indata) Then mValueChanged = True
''  mRsHeader!Eccn_no = indata
''End Property
''
''Public Property Get Eccnno() As String
''  Eccnno = mRsHeader!Eccn_no & ""
''End Property

Public Property Let UnitOfPurch(ByVal indata As String)
'If Not mRsHeader!poi_UnitofP = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_UnitofP = (Trim$(indata))
End Property
Public Property Get Description() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Description = IIf(IsNull(mRsHeader!poi_desc), "", mRsHeader!poi_desc)
End Property

Public Property Let Description(ByVal indata As String)
'If Not mRsHeader!poi_desc = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_desc = (Trim$(indata))
End Property

Public Property Get Primreqdqty() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Primreqdqty = mRsHeader!poi_primreqdqty
End Property

Public Property Let Primreqdqty(ByVal indata As Double)
'If Not mRsHeader!poi_primreqdqty = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_primreqdqty = (Trim$(indata))
End Property

Public Property Get Primuom() As String
 If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
 Primuom = IIf(IsNull(mRsHeader!poi_primuom), "", mRsHeader!poi_primuom)
End Property

Public Property Let Primuom(ByVal indata As String)
'If Not mRsHeader!poi_primuom = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_primuom = (Trim$(indata))
End Property

Public Property Get Secoreqdqty() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Secoreqdqty = mRsHeader!poi_secoreqdqty
End Property

Public Property Let Secoreqdqty(ByVal indata As Double)
'If Not mRsHeader!poi_secoreqdqty = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_secoreqdqty = (Trim$(indata))
End Property

Public Property Get Secouom() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Secouom = mRsHeader!poi_secouom
End Property

Public Property Let Secouom(ByVal indata As String)
'If Not mRsHeader!poi_secouom = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_secouom = (Trim$(indata))
End Property

Public Property Get PrimUnitprice() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PrimUnitprice = mRsHeader!poi_unitprice
End Property

Public Property Let PrimUnitprice(ByVal indata As Double)
'If Not mRsHeader!poi_unitprice = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_unitprice = (Trim$(indata))
End Property

Public Property Get SecUnitPrice() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
SecUnitPrice = mRsHeader!poi_SecOUnitPrice
End Property

Public Property Let SecUnitPrice(ByVal indata As Double)
'If Not mRsHeader!poi_SecUnitPrice = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_SecOUnitPrice = (Trim$(indata))
End Property

Public Property Get PriTotaprice() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PriTotaprice = mRsHeader!poi_totaprice
End Property

Public Property Let PriTotaprice(ByVal indata As Double)
'If Not mRsHeader!poi_totaprice = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_totaprice = (Trim$(indata))
End Property

Public Property Get SecTotaprice() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
SecTotaprice = mRsHeader!poi_SecOTotaprice
End Property

Public Property Let SecTotaprice(ByVal indata As Double)
'If Not mRsHeader!poi_SecTotaprice = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_SecOTotaprice = (Trim$(indata))
End Property

Public Property Get PriQtydlvd() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PriQtydlvd = mRsHeader!poi_qtydlvd
End Property

Public Property Let PriQtydlvd(ByVal indata As Double)
'If Not mRsHeader!poi_qtydlvd = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_qtydlvd = (Trim$(indata))
End Property

'''Public Property Get Secqtydlvd() As Double
'''If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
'''Secqtydlvd = mRsHeader!poi_Secqtydlvd
'''End Property
'''
'''Public Property Let Secqtydlvd(ByVal indata As Double)
''''If Not mRsHeader!poi_Secqtydlvd = (Trim$(indata)) Then mValueChanged = True
'''mRsHeader!poi_Secqtydlvd = (Trim$(indata))
'''End Property

Public Property Get PriQtyship() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PriQtyship = mRsHeader!poi_qtyship
End Property

Public Property Let PriQtyship(ByVal indata As Double)
'If Not mRsHeader!poi_qtyship = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_qtyship = (Trim$(indata))
End Property

'''Public Property Get SecQtyship() As Double
'''If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
'''SecQtyship = mRsHeader!poi_Secqtyship
'''End Property
'''
'''Public Property Let SecQtyship(ByVal indata As Double)
''''If Not mRsHeader!poi_Secqtyship = (Trim$(indata)) Then mValueChanged = True
'''mRsHeader!poi_Secqtyship = (Trim$(indata))
'''End Property

Public Property Get PriQtyinvt() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PriQtyinvt = mRsHeader!poi_qtyinvt
End Property

Public Property Let PriQtyinvt(ByVal indata As Double)
'If Not mRsHeader!poi_qtyinvt = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_qtyinvt = (Trim$(indata))
End Property

''''Public Property Get Secqtyinvt() As Double
''''If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
''''Secqtyinvt = mRsHeader!poi_Secqtyinvt
''''End Property
''''
''''Public Property Let Secqtyinvt(ByVal indata As Double)
'''''If Not mRsHeader!poi_Secqtyinvt = (Trim$(indata)) Then mValueChanged = True
''''mRsHeader!poi_Secqtyinvt = (Trim$(indata))
''''End Property

Public Property Get Comm() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Comm = IIf(IsNull(mRsHeader!poi_comm), "", mRsHeader!poi_comm)
End Property

Public Property Let Comm(ByVal indata As String)
'If Not mRsHeader!poi_comm = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_comm = (Trim$(indata))
    
 End Property

Public Property Get Requnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Requnumb = IIf(IsNull(mRsHeader!poi_requnumb), "", mRsHeader!poi_requnumb)
End Property

Public Property Let Requnumb(ByVal indata As String)
'If Not mRsHeader!poi_requnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_requnumb = IIf(Len(Trim$(indata)) = 0, Null, Trim$(indata))



End Property

Public Property Get Requliitnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Requliitnumb = IIf(IsNull(mRsHeader!poi_requliitnumb), "", mRsHeader!poi_requliitnumb)
End Property

Public Property Let Requliitnumb(ByVal indata As String)
'If Not mRsHeader!poi_requliitnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_requliitnumb = IIf(Len(Trim$(indata)) = 0, Null, Trim$(indata))

End Property

Public Property Get Quotnum() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Quotnum = mRsHeader!poi_quotnum
End Property

Public Property Let Quotnum(ByVal indata As String)
'If Not mRsHeader!poi_quotnum = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_quotnum = (Trim$(indata))
End Property

Public Property Get Quotliitnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Quotliitnumb = mRsHeader!poi_quotliitnumb
End Property

Public Property Let Quotliitnumb(ByVal indata As String)
'If Not mRsHeader!poi_quotliitnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_quotliitnumb = (Trim$(indata))
End Property

Public Property Get Locatax() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Locatax = mRsHeader!poi_locatax
End Property

Public Property Let Locatax(ByVal indata As Double)
'If Not mRsHeader!poi_locatax = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_locatax = (Trim$(indata))
End Property

Public Property Get Remk() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Remk = IIf(IsNull(mRsHeader!poi_remk), "", mRsHeader!poi_remk)
End Property

Public Property Let Remk(ByVal indata As String)
'If Not mRsHeader!poi_remk = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_remk = (Trim$(indata))
End Property

Public Property Get Serlnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Serlnumb = IIf(IsNull(mRsHeader!poi_serlnumb), "", mRsHeader!poi_serlnumb)
End Property

Public Property Let Serlnumb(ByVal indata As String)
'If Not mRsHeader!poi_serlnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_serlnumb = (Trim$(indata))
End Property

Public Property Get Manupartnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Manupartnumb = IIf(IsNull(mRsHeader!poi_manupartnumb), "", mRsHeader!poi_manupartnumb)
End Property

Public Property Let Manupartnumb(ByVal indata As String)
''If IsNull(mRsHeader!poi_manupartnumb) Then
''mValueChanged = True
''
''Else
''
''If Not mRsHeader!poi_manupartnumb = (Trim$(indata)) Then
''mValueChanged = True
''End If
''End If

mRsHeader!poi_manupartnumb = Trim$(indata)


End Property



Public Property Get Liitreqddate() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Liitreqddate = IIf(IsNull(mRsHeader!poi_liitreqddate), "", mRsHeader!poi_liitreqddate)
End Property

Public Property Let Liitreqddate(ByVal indata As String)
'If Not mRsHeader!poi_liitreqddate = (Trim$(indata)) Then mValueChanged = True


If Len(Trim$(indata)) = 0 Then
  
  mRsHeader!poi_liitreqddate = Null
  
Else

  mRsHeader!poi_liitreqddate = CDate(indata)

End If


End Property

Public Property Get Liitrelsdate() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Liitrelsdate = mRsHeader!poi_liitrelsdate
End Property

Public Property Let Liitrelsdate(ByVal indata As String)
'If Not mRsHeader!poi_liitrelsdate = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_liitrelsdate = (Trim$(indata))
End Property

Public Property Get Starrendate() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Starrendate = mRsHeader!poi_starrendate
End Property

Public Property Let Starrendate(ByVal indata As String)
'If Not mRsHeader!poi_starrendate = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_starrendate = (Trim$(indata))
End Property

Public Property Get Endrentdate() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Endrentdate = mRsHeader!poi_endrentdate
End Property

Public Property Let Endrentdate(ByVal indata As String)
'If Not mRsHeader!poi_endrentdate = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_endrentdate = (Trim$(indata))
End Property

Public Property Get Stasliit() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Stasliit = mRsHeader!poi_stasliit
End Property

Public Property Let Stasliit(ByVal indata As String)
'If Not mRsHeader!poi_stasliit = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_stasliit = (Trim$(indata))
End Property

Public Property Get Stasdlvy() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Stasdlvy = mRsHeader!poi_stasdlvy
End Property

Public Property Let Stasdlvy(ByVal indata As String)
'If Not mRsHeader!poi_stasdlvy = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_stasdlvy = (Trim$(indata))
End Property

Public Property Get Stasship() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Stasship = mRsHeader!poi_stasship
End Property

Public Property Let Stasship(ByVal indata As String)
'If Not mRsHeader!poi_stasship = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_stasship = (Trim$(indata))
End Property

Public Property Get Stasinvt() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Stasinvt = mRsHeader!poi_stasinvt
End Property

Public Property Let Stasinvt(ByVal indata As String)
'If Not mRsHeader!poi_stasinvt = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_stasinvt = (Trim$(indata))
End Property

Public Property Get Currcode() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Currcode = mRsHeader!poi_currcode
End Property

Public Property Let Currcode(ByVal indata As String)
'If Not mRsHeader!poi_currcode = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_currcode = (Trim$(indata))
End Property

Public Property Get Afe() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Afe = IIf(IsNull(mRsHeader!poi_afe), "", mRsHeader!poi_afe)
End Property

Public Property Let Afe(ByVal indata As String)
'if isnull(mrsheader!poi_afe) and trim$(indata)
'If Not mRsHeader!poi_afe = (Trim$(indata)) Then mValueChanged = True

mRsHeader!poi_afe = (Trim$(indata))
End Property

Public Property Get Custcate() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Custcate = IIf(IsNull(mRsHeader!poi_custcate), "", mRsHeader!poi_custcate)
End Property

Public Property Let Custcate(ByVal indata As String)
'If Not mRsHeader!poi_custcate = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_custcate = (Trim$(indata))
End Property

Public Property Get Lastinvcnumb() As String
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Lastinvcnumb = mRsHeader!poi_lastinvcnumb
End Property

Public Property Let Lastinvcnumb(ByVal indata As String)
'If Not mRsHeader!poi_lastinvcnumb = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_lastinvcnumb = (Trim$(indata))
End Property

Public Property Get PriQtytobedlvd() As Double
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
PriQtytobedlvd = mRsHeader!poi_Qtytobedlvd
End Property

Public Property Let PriQtytobedlvd(ByVal indata As Double)
'If Not mRsHeader!po_Qtytobedlvd = (Trim$(indata)) Then mValueChanged = True
mRsHeader!poi_Qtytobedlvd = (Trim$(indata))
End Property


'------------------------------------------------------
Public Function AddNew() As Boolean
Dim cancel As Boolean
AddNew = False
  On Error Resume Next
  
  RaiseEvent BeforeAddNEW(cancel)
       If cancel = False Then
           
    'If mRsHeader.State = 1 And (mRsHeader.EditMode = 1 Or mRsHeader.EditMode = 2) Then
    
    '       mRsHeader!poi_npecode = mIutility.Namespace
     'End If
           mRsHeader.AddNew
           If Err.Number = "-2147217864" Then Err.Clear
    On Error GoTo handler

           SetTheNewHeader
   End If
  cancel = True
  AddNew = True
  
  Exit Function

handler:
     

    Err.Clear
End Function

Public Function DeleteCurrentLI() As Boolean
Dim varBookmark  As Variant
On Error GoTo handler
DeleteCurrentLI = False


If Not mRsHeader.AbsolutePosition = adPosEOF Then
     
     mRsHeader.Delete adAffectCurrent
  If mRsHeader.RecordCount > 0 Then
            mRsHeader.MoveNext
            
            If mRsHeader.AbsolutePosition = adPosEOF Then
               mRsHeader.MoveFirst
            End If
        
            varBookmark = mRsHeader.Bookmark
            mRsHeader.MoveFirst
            
            Do While Not mRsHeader.Eof
               mRsHeader!poi_liitnumb = mRsHeader.AbsolutePosition
               mRsHeader.MoveNext
            Loop
            mRsHeader.Bookmark = varBookmark
  End If
  
  RaiseEvent AfterDeleteCurrentLI
  
End If
DeleteCurrentLI = True
Exit Function
handler:
   MsgBox "Error while deleting the line item." & Err.Description
   Err.Clear
   
End Function
Public Function DeleteAllPOITEM(PO_ponumb As String) As Boolean
Dim cancel As Boolean
DeleteAllPOITEM = False
On Error GoTo handler

RaiseEvent BeforeDeleteALLPOITEM(cancel)

If cancel = False Then
    mRsHeader.Filter = "poi_ponumb='" & PO_ponumb & "'"
    mRsHeader.Delete adAffectGroup
    mRsHeader.Filter = ""
    
Else
    cancel = False
End If

RaiseEvent AfterDeletePOITEMS

DeleteAllPOITEM = True
Exit Function
handler:
MsgBox "Error Deleting POITEMS.Err.desc -- " & Err.Description
Err.Clear
End Function
Public Function Update() As Boolean
Dim TPrice As Double


 On Error GoTo handler
 Update = False
  
 
   'If SavePoitemRequisition = False Then Exit Function
      
      TPrice = CalculateTotal
     If UpdateTotalPriceOfPO(TPrice) = False Then Exit Function
     If AnyChangesBeforeSaving = False Then Exit Function
     mRsHeader.UpdateBatch
RESUMENEXT:
     If AnyChangesAfterSaving(MPonumb) = False Then Exit Function
     
     'mRsHeader.Requery
     'Call mRsHeader.Find("Po_ponumb='" & Ponumb & "'", , adSearchForward)
     If mRsHeader.Eof = False Then ' FillMeWithTheCurrentRecord (mRsHeader)
     End If
 
  
     Update = True
     Exit Function
handler:
    'Filling The Collection with The Error
    If Err.Number = "-2147217864" Then
        Update = True
        Err.Clear
        GoTo RESUMENEXT
    Else
        MsgBox "Errors Occured in Saving POITEMS  & " & Err.Description
        Err.Clear
    End If
End Function

Public Function Configure(Iutilty As Iutility)
Set miutility = Iutilty
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
       .CommandText = "GETPOITEM"
    
       Set .ActiveConnection = miutility.Connection
       
       .Parameters.Append .CreateParameter("@PONUMB", adVarChar, adParamInput, 15, MPonumb)
       .Parameters.Append .CreateParameter("@NAMESPACE", adVarChar, adParamInput, 5, miutility.Namespace)
    
    End With
    
   Set mRsHeader = CmdHead.Execute
       mRsHeader.Close
       mRsHeader.CursorLocation = adUseClient
       mRsHeader.LockType = adLockBatchOptimistic
       mRsHeader.Open , , adOpenKeyset, adLockBatchOptimistic

   LoadCollection = True
   mIsItemLoaded = True
   Exit Function
   
handler:
   Err.Raise Err.Number, "IMSPODLL.POITEMS.Loadcollection", Err.Description
   Err.Clear
End Function


Public Function Move(POnumb As String) As Boolean
On Error GoTo handler
Move = False
If (POnumb) <> MPonumb Then mIsItemLoaded = False
MPonumb = (POnumb)

If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection

If mIsItemLoaded = True Then
   mRsHeader.MoveFirst
'   Call mRsHeader.Find("Po_ponumb='" & Ponumb & "'", , adSearchForward)
   
   If mRsHeader.Eof = False Then
   Move = True
   'Call FillMeWithTheCurrentRecord(mRsHeader)
   End If
End If
   
   
   Exit Function
   
handler:

   Err.Clear
End Function

Public Property Get Count() As Integer
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Count = mRsHeader.RecordCount
End Property
Public Function MoveFirst() As Boolean
MoveFirst = False
On Error GoTo handler
  If Not mRsHeader.Bof Then
   
     mRsHeader.MoveFirst
     If mRsHeader.Bof = True Then 'FillmeWithTheCurrentRecord (mRsHeader)
       MoveFirst = False
     Else
       MoveFirst = True
      End If
  End If
  Exit Function
handler:
  
End Function



Public Function MoveNext() As Boolean
MoveNext = False
On Error GoTo handler
  If Not mRsHeader.Eof Then
   
     mRsHeader.MoveNext
    ' If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Eof = True And mRsHeader.Bof <> True Then mRsHeader.MoveLast 'AM
     
     If mRsHeader.AbsolutePosition = adPosEOF Then mRsHeader.MoveLast 'AM
     If Not mRsHeader.Eof Then
        MoveNext = True
     End If
  End If
  Exit Function
handler:
  
End Function

Public Function MovePrevious() As Boolean
MovePrevious = False
On Error GoTo handler
  If Not mRsHeader.Bof Then
   
     mRsHeader.MovePrevious
     If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
     If mRsHeader.Bof Then  'FillmeWithTheCurrentRecord (mRsHeader)
      MovePrevious = False
     Else
      MovePrevious = True
     End If
  End If
 Exit Function
handler:
  
End Function

Public Function MoveLast() As Boolean

MoveLast = False
On Error GoTo handler
  If Not mRsHeader.Eof Then
   
     mRsHeader.MoveLast
     If mRsHeader.Eof = True Then 'FillmeWithTheCurrentRecord (mRsHeader)
       MoveLast = False
     Else
       MoveLast = True
     End If
  End If
  Exit Function
handler:
  Err.Clear
End Function

Private Sub SetTheNewHeader()

   MStatus = ADSTATUSINSERT
   mCancel = False
   
End Sub

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
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
If mIsItemLoaded = True Then
  Eof = mRsHeader.Eof
End If
End Property

Public Property Get Bof() As Boolean
If mIsItemLoaded = False Then mIsItemLoaded = LoadCollection
Bof = mRsHeader.Bof
End Property


Public Function CancelUpdate() As Boolean
Dim LineNo As Integer
On Error GoTo handler
If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Bof <> True And mRsHeader.Eof = True Then mRsHeader.MoveLast
If mRsHeader.AbsolutePosition = adPosUnknown And mRsHeader.RecordCount > 0 Then mRsHeader.MoveLast

If mRsHeader.EditMode = adEditAdd Then
   
'   LineNo = mRsHeader("poi_liitnumb")
   mRsHeader.CancelUpdate

 ' If mRsHeader.RecordCount > 0 Then
           
  '          mRsHeader.MoveFirst
            
  '          Do While Not mRsHeader.Eof
  '             mRsHeader!poi_liitnumb = mRsHeader.AbsolutePosition
  '             mRsHeader.MoveNext
  '          Loop
  '          mRsHeader.MoveLast
            
  'End If
  
End If

If mRsHeader.AbsolutePosition = adPosBOF And mRsHeader.Eof <> True And mRsHeader.Bof = True Then mRsHeader.MoveFirst
If mRsHeader.AbsolutePosition = adPosEOF And mRsHeader.Bof <> True And mRsHeader.Eof = True Then mRsHeader.MoveLast
If mRsHeader.AbsolutePosition = adPosUnknown And mRsHeader.RecordCount > 0 Then mRsHeader.MoveLast

RaiseEvent AfterCancelUpdate

Exit Function
   
handler:
   Err.Clear
End Function

Public Function CancelBatch()
On Error GoTo handler
If mRsHeader.EditMode = adEditAdd Then
   mRsHeader.CancelBatch
End If
Exit Function
   
handler:
   Err.Clear
End Function


Public Property Get EditMode() As String
If mRsHeader.State = 1 And mRsHeader.RecordCount > 0 Then
  EditMode = mRsHeader.EditMode
Else
   EditMode = "-1"
End If
End Property


Public Function Requery()
On Error GoTo handler
   mRsHeader.Requery
   Exit Function
   
handler:

  MsgBox "Error When Poitem tried to Requery Database"
  Err.Clear
End Function

Private Function SavePoitemRequisition() As Boolean
On Error GoTo handler
SavePoitemRequisition = False

Dim mCmdReq As ADODB.Command
  mRsHeader.MoveFirst
    Do While Not mRsHeader.Eof
      If mRsHeader.EditMode = 2 Then
        If Len(mRsHeader!poi_requnumb & "") > 0 And Len(mRsHeader!poi_requliitnumb & "") > 0 Then
                Set mCmdReq = New ADODB.Command
                mCmdReq.ActiveConnection = miutility.Connection
                mCmdReq.CommandType = adCmdText
                mCmdReq.CommandText = "update poitem set poi_primreqdqty='0.0',poi_secoreqdqty='0.0',poi_stasliit='CL' where poi_ponumb='" & mRsHeader!poi_requnumb & "' and poi_liitnumb='" & mRsHeader!poi_requliitnumb & "'"
                mCmdReq.Execute , , adExecuteNoRecords
                Set mCmdReq = Nothing
        End If
       End If
        mRsHeader.MoveNext
    Loop
    SavePoitemRequisition = True
    Exit Function
handler:
    Err.Raise Err.Number, , "Could not Save the Requisition Line Items." & Err.Description
    Err.Clear
End Function

Public Function UpdateTotalPriceOfPO(totalpo As Double) As Boolean
Dim CmdPo As New ADODB.Command
UpdateTotalPriceOfPO = False
On Error GoTo handler
   CmdPo.CommandText = "uPDATE PO SET po_totacost =" & totalpo & " where ( ( po_ponumb ='" & MPonumb & "') AND  ( po_npecode ='" & miutility.Namespace & "' )  )"
   CmdPo.ActiveConnection = miutility.Connection
   CmdPo.Execute , adExecuteNoRecords
   UpdateTotalPriceOfPO = True
   Exit Function
handler:
   Err.Raise Err.Number, , "Error in Updating PO Total." & Err.Description
   Err.Clear
End Function

Public Function CalculateTotal() As Double
  
  CalculateTotal = 0
  If mRsHeader.RecordCount > 0 Then
    mRsHeader.MoveFirst
    Do While Not mRsHeader.Eof
        CalculateTotal = CalculateTotal + mRsHeader!poi_totaprice
        mRsHeader.MoveNext
    Loop
     CalculateTotal = FormatNumber(CalculateTotal, 4)
  End If
End Function

Public Function AnyChangesBeforeSaving() As Boolean
On Error GoTo handler

AnyChangesBeforeSaving = False
If mRsHeader.RecordCount > 0 Then
   mRsHeader.MoveFirst
Else
   AnyChangesBeforeSaving = True
   Exit Function
End If
Do While Not mRsHeader.Eof
   If mRsHeader!poi_primreqdqty = 0 And Trim$(UCase(mRsHeader!poi_stasliit & "")) <> "CL" Then mRsHeader!poi_stasliit = "CA"
   mRsHeader.MoveNext
Loop

AnyChangesBeforeSaving = True
Exit Function
handler:
   Err.Raise Err.Number, , "Error in setting the Status of a POitem to Cancelled when Quantity is 0."
   Err.Clear
End Function
Public Function AnyChangesAfterSaving(POnumb As String) As Boolean
Dim Cmd As New ADODB.Command
On Error GoTo handler

If mRsHeader.RecordCount = 0 Then AnyChangesAfterSaving = True: Exit Function

With Cmd

    .CommandText = "UPDATESTOCKWITHECCN"
    .CommandType = adCmdStoredProc
    .ActiveConnection = miutility.Connection
    .Parameters.Append .CreateParameter("@npecode", adVarChar, adParamInput, 5, miutility.Namespace)
    .Parameters.Append .CreateParameter("@po", adVarChar, adParamInput, 15, POnumb)
    .Execute
    
End With


AnyChangesAfterSaving = True
Exit Function
handler:
   Err.Raise Err.Number, , "Error occurred in saving the Eccn# to the stockmaster."
   Err.Clear
End Function

Public Function Replace(PropertyName As String, Value As String) As Boolean

Dim AfterSlash As String

On Error GoTo Errhandler

Replace = False

PropertyName = Trim$(PropertyName)

If mRsHeader.RecordCount > 0 Then

    mRsHeader.MoveFirst
    
Else
    
    Replace = True
    
    Exit Function

End If

Select Case PropertyName

    Case "PONUMB"
          
          mRsHeader.MoveFirst
          
          Do While Not mRsHeader.Eof
          
                mRsHeader!poi_ponumb = Value
                
                MPonumb = Value
                
                If (InStr(1, LTrim(mRsHeader!poi_comm), "_") = 1) Then
                
                        AfterSlash = Mid(mRsHeader!poi_comm, InStr(mRsHeader!poi_comm, "/"), Len(mRsHeader!poi_comm))
                        
                        mRsHeader!poi_comm = RTrim(MPonumb) & AfterSlash
                
                 End If
                 
                mRsHeader.MoveNext
          Loop
      
End Select

Replace = True

Exit Function

Errhandler:

'n

End Function


Public Function LoadFromRequsition(RequsitionNo As String, TemporaryPoNumb As String, Fromstocknumber As Boolean, ErrorIfany As String) As Boolean
Dim RsRequisition As ADODB.Recordset
Dim field As ADODB.field
On Error GoTo Errhandler
LoadFromRequsition = True

Set RsRequisition = Nothing

If Move(RequsitionNo) = True Then

Set RsRequisition = mRsHeader.Clone

If RsRequisition.RecordCount > 0 Then

    Set mRsHeader = Nothing
    Move TemporaryPoNumb
    
    RsRequisition.MoveFirst

    Do While Not RsRequisition.Eof
        
        mRsHeader.AddNew
        
        For Each field In RsRequisition.Fields
        
                mRsHeader(field.Name) = field.Value
        
        Next
       
       mRsHeader("poi_ponumb") = TemporaryPoNumb
       
       mRsHeader("poi_stasliit") = "OH"
       mRsHeader("poi_stasdlvy") = "NR"
       mRsHeader("poi_stasship") = "NS"
       mRsHeader("poi_stasinvt") = "NI"
       mRsHeader("poi_liitreqddate") = DateAdd("d", 1, Now)
       
        'Juan 2010-9-25 necessary to reset lineitem
        mRsHeader("poi_qtydlvd") = 0
        mRsHeader("poi_qtytobedlvd") = mRsHeader("poi_primreqdqty")
        mRsHeader("poi_qtyship") = 0
        mRsHeader("poi_qtyinvt") = 0
        '-------------------
       
       If Fromstocknumber = False Then
       
        mRsHeader("poi_comm") = ""
       
       End If

        
       RsRequisition.MoveNext
       
     Loop
       
     End If
       
    Set RsRequisition = Nothing

End If

LoadFromRequsition = False

Exit Function
Errhandler:
ErrorIfany = Err.Description

Err.Clear
End Function

