VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Rs As New ADODB.Recordset

On Error GoTo Errhandler
De.Cn.Open
Rs.Source = "Update xuserprofile set usr_stas ='A' where usr_userid ='imsusa'"
Rs.Open , De.Cn
MsgBox "User Status for IMSUSA activated successfully.", vbInformation, "Ims"
Unload Me
Exit Sub

Errhandler:
MsgBox "Errors Occurrred while trying to activate IMSUSA. Err Desc " & Err.Description, vbCritical, "Ims"
Unload Me
End Sub
