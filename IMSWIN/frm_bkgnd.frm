VERSION 5.00
Begin VB.Form frm_bkgnd 
   BorderStyle     =   0  'None
   ClientHeight    =   7485
   ClientLeft      =   6015
   ClientTop       =   1485
   ClientWidth     =   3495
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HasDC           =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frm_bkgnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  '    color the controls and form backcolor
    Call Move(0, 0)
    'Me.BackColor = frm_Color.txt_Background.BackColor
End Sub
