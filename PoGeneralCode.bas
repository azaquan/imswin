Attribute VB_Name = "PoGeneral"
Option Explicit
''''Public Enum FormMode
''''    mdNa = 0
''''    mdCreation
''''    mdModified
''''    mdModification
''''    mdVisualization
''''End Enum
Global mDidUserOpenStkMasterForm As Boolean
Public Enum LoadMode
   LoadingPOheader = 0
   loadingPoItem
   loadingPoRemark
   loadingPoClause
   NoLoadInProgress = -1
End Enum
Public Type StockDesc
    No As String
    PriUnit As String
    SecUnit As String
    CompFactor As Double
End Type



Public Sub ToggleNavButtons(FMode As FormMode)


   

End Sub

Public Sub InsertPoRevision(Ponumb As String)
Dim cmd As ADODB.Command
  On Error GoTo handler
   
    Set cmd = New ADODB.Command
    
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = deIms.cnIms
    cmd.Prepared = True
    cmd.CommandText = "InsertPoRevision"
    cmd.Execute , Array(deIms.NameSpace, Ponumb)
    Exit Sub
    
handler:
  MsgBox "Errors occurred while inserting a PO revision.Error description  " & Err.Description
  Err.Clear
  
End Sub





