VERSION 5.00
Object = "{27609697-380F-11D5-99AB-00D0B74311D4}#1.0#0"; "LRLabelsX.ocx"
Begin VB.Form frmNavigator 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Navigator"
   ClientHeight    =   8295
   ClientLeft      =   3840
   ClientTop       =   0
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOwner 
      BorderStyle     =   0  'None
      Height          =   8400
      Left            =   0
      ScaleHeight     =   560
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   787
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   0
      Width           =   11805
      Begin VB.PictureBox picReport 
         BorderStyle     =   0  'None
         Height          =   9030
         Left            =   3000
         ScaleHeight     =   9030
         ScaleWidth      =   8835
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   8835
         Begin VB.Frame fraHolder 
            Height          =   1695
            Index           =   14
            Left            =   3180
            TabIndex        =   149
            Top             =   6480
            Width           =   5475
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Menu User History"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   18
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   163
               Tag             =   "03050700"
               Top             =   1380
               Width           =   1575
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Access Level"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   17
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   155
               Tag             =   "03050100"
               Top             =   165
               Width           =   5115
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Application User Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   16
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   154
               Tag             =   "03050200"
               Top             =   360
               Width           =   5130
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Login/Logoff"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   15
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   153
               Tag             =   "03050300"
               Top             =   570
               Width           =   5175
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Security Changes Log"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   14
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   152
               Tag             =   "03050400"
               Top             =   765
               Width           =   5115
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Access Level + Buyer + User"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   13
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   151
               Tag             =   "03050600"
               Top             =   1170
               Width           =   5115
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Individual User Profile"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   11
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   150
               Tag             =   "03050500"
               Top             =   960
               Width           =   5130
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   2295
            Index           =   6
            Left            =   3180
            TabIndex        =   36
            Top             =   960
            Width           =   5475
            Begin VB.Label lblnewSupplier 
               AutoSize        =   -1  'True
               Caption         =   "Supplier by Date created"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2760
               MousePointer    =   99  'Custom
               TabIndex        =   164
               Tag             =   "03020104"
               Top             =   585
               Width           =   2130
            End
            Begin VB.Label lblrequisitionstatus 
               AutoSize        =   -1  'True
               Caption         =   "Requisition Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2760
               MousePointer    =   99  'Custom
               TabIndex        =   162
               Tag             =   "03020102"
               Top             =   375
               Width           =   2640
            End
            Begin VB.Label lblreorder 
               AutoSize        =   -1  'True
               Caption         =   "Optima"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   10
               Left            =   2760
               MousePointer    =   99  'Custom
               TabIndex        =   161
               Tag             =   "03020101"
               Top             =   180
               Width           =   600
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Report Writer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   9
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   125
               Tag             =   "03020900"
               Top             =   1785
               Width           =   1155
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "General Status Report(By Req)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   124
               Tag             =   "03021000"
               Top             =   1980
               Width           =   5160
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Late Shipping Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   123
               Tag             =   "03020800"
               Top             =   1575
               Width           =   5175
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Late Delivery Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   122
               Tag             =   "03020700"
               Top             =   1380
               Width           =   5130
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Order Delivery Schedule"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   43
               Tag             =   "03020600"
               Top             =   1185
               Width           =   5205
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Order Tracking Record"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Tag             =   "03020500"
               Top             =   975
               Width           =   5205
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Order Activity Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   41
               Tag             =   "03020400"
               Top             =   780
               Width           =   5160
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Open Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   40
               Tag             =   "03020300"
               Top             =   585
               Width           =   5175
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Stock Number History"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   39
               Tag             =   "03020200"
               Top             =   375
               Width           =   2340
            End
            Begin VB.Label lblRepPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Print Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   37
               Tag             =   "03020100"
               Top             =   180
               Width           =   2250
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   2115
            Index           =   7
            Left            =   3180
            TabIndex        =   45
            Top             =   3240
            Width           =   5475
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Stock On Hand per Stock Number"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   126
               Tag             =   "03030900"
               Top             =   1770
               Width           =   5175
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Stock On Hand across all locations"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   53
               Tag             =   "03030800"
               Top             =   1560
               Width           =   5175
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "New Stock On Hand"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   52
               Tag             =   "03030500"
               Top             =   960
               Width           =   1755
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Historical Stock Movement"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   51
               Tag             =   "03030700"
               Top             =   1365
               Width           =   5175
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Slow Moving Inventory"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   50
               Tag             =   "03030600"
               Top             =   1170
               Width           =   5175
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Stock On-Hand"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   49
               Tag             =   "03030400"
               Top             =   765
               Width           =   5160
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Transactions per Date Range"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   48
               Tag             =   "03030300"
               Top             =   570
               Width           =   5160
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Inventory per Stock Number"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   47
               Tag             =   "03030200"
               Top             =   360
               Width           =   5160
            End
            Begin VB.Label lblRepWharehouse 
               AutoSize        =   -1  'True
               Caption         =   "Orders to be Received"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   46
               Tag             =   "03030100"
               Top             =   165
               Width           =   5175
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   885
            Index           =   8
            Left            =   3180
            TabIndex        =   54
            Top             =   5340
            Width           =   5475
            Begin VB.Label lblRepAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Upload Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   2880
               MousePointer    =   99  'Custom
               TabIndex        =   168
               Tag             =   "03040500"
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label lblRepAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Price Control Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   2880
               MousePointer    =   99  'Custom
               TabIndex        =   167
               Tag             =   "03040400"
               Top             =   165
               Width           =   1740
            End
            Begin VB.Label lblRepAccounting 
               AutoSize        =   -1  'True
               Caption         =   "SAP Analysis Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   57
               Tag             =   "03040300"
               Top             =   570
               Width           =   2715
            End
            Begin VB.Label lblRepAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Transaction Valuation Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   56
               Tag             =   "03040200"
               Top             =   360
               Width           =   2625
            End
            Begin VB.Label lblRepAccounting 
               AutoSize        =   -1  'True
               Caption         =   "SAP Valuation Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   55
               Tag             =   "03040100"
               Top             =   165
               Width           =   2460
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   960
            Index           =   5
            Left            =   3180
            TabIndex        =   33
            Top             =   0
            Width           =   5475
            Begin VB.Label lblRepCataloging 
               BackStyle       =   0  'Transparent
               Caption         =   "Export Stock Master to Excel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   121
               Tag             =   "03010300"
               Top             =   585
               Width           =   5250
            End
            Begin VB.Label lblRepCataloging 
               BackStyle       =   0  'Transparent
               Caption         =   "Mfr/Stock# Xref"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   35
               Tag             =   "03010200"
               Top             =   375
               Width           =   5250
            End
            Begin VB.Label lblRepCataloging 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Master"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   34
               Tag             =   "03010100"
               Top             =   180
               Width           =   5220
            End
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   58
            X1              =   2760
            X2              =   3180
            Y1              =   7800
            Y2              =   7800
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   600
            Index           =   0
            Left            =   480
            Top             =   120
            Width           =   2160
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   25
            X1              =   2760
            X2              =   3180
            Y1              =   700
            Y2              =   700
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   90
            X1              =   2640
            X2              =   3240
            Y1              =   6780
            Y2              =   6780
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   30
            X1              =   2760
            X2              =   2760
            Y1              =   6780
            Y2              =   8040
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   88
            X1              =   2760
            X2              =   3180
            Y1              =   7575
            Y2              =   7575
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   84
            X1              =   2760
            X2              =   3180
            Y1              =   7380
            Y2              =   7380
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   81
            X1              =   2760
            X2              =   3180
            Y1              =   7185
            Y2              =   7185
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   80
            X1              =   2760
            X2              =   3180
            Y1              =   8040
            Y2              =   8040
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   79
            X1              =   2760
            X2              =   3180
            Y1              =   6975
            Y2              =   6975
         End
         Begin VB.Label lblReportMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Security"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   600
            TabIndex        =   148
            Tag             =   "03050000"
            Top             =   6540
            Width           =   1905
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   26
            X1              =   0
            X2              =   480
            Y1              =   6600
            Y2              =   6600
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   600
            Index           =   5
            Left            =   480
            Top             =   6360
            Width           =   2160
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   66
            X1              =   2760
            X2              =   3180
            Y1              =   3060
            Y2              =   3060
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   65
            X1              =   2760
            X2              =   3180
            Y1              =   2860
            Y2              =   2860
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   64
            X1              =   2760
            X2              =   3180
            Y1              =   2660
            Y2              =   2660
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   4
            X1              =   2760
            X2              =   3180
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   24
            X1              =   2760
            X2              =   3180
            Y1              =   6045
            Y2              =   6045
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   600
            Index           =   3
            Left            =   480
            Top             =   5460
            Width           =   2160
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   23
            X1              =   0
            X2              =   480
            Y1              =   5700
            Y2              =   5700
         End
         Begin VB.Label lblReportMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Financial Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   58
            Tag             =   "03040000"
            Top             =   5640
            Width           =   1920
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   19
            X1              =   2640
            X2              =   3240
            Y1              =   5640
            Y2              =   5640
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   18
            X1              =   2760
            X2              =   3180
            Y1              =   5835
            Y2              =   5835
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   51
            X1              =   2760
            X2              =   2760
            Y1              =   5640
            Y2              =   6030
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   22
            X1              =   2760
            X2              =   3180
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   21
            X1              =   2760
            X2              =   3180
            Y1              =   4940
            Y2              =   4940
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   20
            X1              =   2760
            X2              =   3180
            Y1              =   5140
            Y2              =   5140
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   17
            X1              =   2760
            X2              =   3180
            Y1              =   3740
            Y2              =   3740
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   16
            X1              =   2760
            X2              =   3180
            Y1              =   4540
            Y2              =   4540
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   15
            X1              =   2760
            X2              =   3180
            Y1              =   3940
            Y2              =   3940
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   14
            X1              =   2760
            X2              =   3180
            Y1              =   4140
            Y2              =   4140
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   13
            X1              =   2760
            X2              =   3180
            Y1              =   4340
            Y2              =   4340
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   50
            X1              =   2760
            X2              =   2760
            Y1              =   3540
            Y2              =   5140
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   12
            X1              =   2640
            X2              =   3240
            Y1              =   3540
            Y2              =   3540
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   11
            X1              =   2760
            X2              =   3180
            Y1              =   2260
            Y2              =   2260
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   10
            X1              =   2760
            X2              =   3180
            Y1              =   2060
            Y2              =   2060
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   9
            X1              =   2760
            X2              =   3180
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   8
            X1              =   2760
            X2              =   3180
            Y1              =   1660
            Y2              =   1660
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   7
            X1              =   2760
            X2              =   3180
            Y1              =   1460
            Y2              =   1460
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   49
            X1              =   2760
            X2              =   2760
            Y1              =   1260
            Y2              =   3060
         End
         Begin VB.Label lblReportMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Inventory Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   44
            Tag             =   "03030000"
            Top             =   3420
            Width           =   1950
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   6
            X1              =   0
            X2              =   480
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   600
            Index           =   2
            Left            =   480
            Top             =   3240
            Width           =   2160
         End
         Begin VB.Label lblReportMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Purchasing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   38
            Tag             =   "03020000"
            Top             =   1260
            Width           =   1920
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   1
            X1              =   0
            X2              =   480
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   5
            X1              =   2640
            X2              =   3195
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   3
            X1              =   2760
            X2              =   3180
            Y1              =   500
            Y2              =   500
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   46
            X1              =   2760
            X2              =   2760
            Y1              =   300
            Y2              =   700
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   2
            X1              =   2640
            X2              =   3195
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Label lblReportMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Cataloging"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   32
            Tag             =   "03010000"
            Top             =   280
            Width           =   1875
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   0
            X1              =   0
            X2              =   480
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   15
            X2              =   15
            Y1              =   240
            Y2              =   6600
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   600
            Index           =   1
            Left            =   480
            Top             =   1080
            Width           =   2160
         End
      End
      Begin VB.PictureBox picActivities 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   9030
         Left            =   3000
         ScaleHeight     =   9030
         ScaleWidth      =   8715
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   8715
         Begin VB.Frame fraHolder 
            Height          =   975
            Index           =   1
            Left            =   3180
            TabIndex        =   12
            Top             =   2340
            Width           =   5475
            Begin VB.Label lblManifestPOD 
               AutoSize        =   -1  'True
               Caption         =   "Manifest POD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2880
               MousePointer    =   99  'Custom
               TabIndex        =   165
               Tag             =   "02030400"
               Top             =   405
               Width           =   1305
            End
            Begin VB.Label lblSubLogisitic 
               AutoSize        =   -1  'True
               Caption         =   "Create Shipping Manifest Tracking Message"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   15
               Tag             =   "02030300"
               Top             =   645
               Width           =   5205
            End
            Begin VB.Label lblSubLogisitic 
               Caption         =   "Create Shipping Manifest"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   14
               Tag             =   "02030200"
               Top             =   405
               Width           =   2460
            End
            Begin VB.Label lblSubLogisitic 
               AutoSize        =   -1  'True
               Caption         =   "Receive Freight"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   13
               Tag             =   "02030100"
               Top             =   165
               Width           =   2085
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   1935
            Index           =   3
            Left            =   3180
            TabIndex        =   22
            Top             =   3300
            Width           =   5475
            Begin VB.Label lblSubWharehousing 
               AutoSize        =   -1  'True
               Caption         =   "Logical Warehouse-Sub Location Movement"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   29
               Tag             =   "02040700"
               Top             =   1605
               Width           =   5205
            End
            Begin VB.Label lblSubWharehousing 
               Caption         =   "Warehouse to Warehouse Transfer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   28
               Tag             =   "02040600"
               Top             =   1365
               Width           =   5130
            End
            Begin VB.Label lblSubWharehousing 
               Caption         =   "Well to Well Transfer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   27
               Tag             =   "02040500"
               Top             =   1125
               Width           =   5145
            End
            Begin VB.Label lblSubWharehousing 
               Caption         =   "Return from Repair"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   26
               Tag             =   "02040400"
               Top             =   885
               Width           =   5145
            End
            Begin VB.Label lblSubWharehousing 
               AutoSize        =   -1  'True
               Caption         =   "Return from Well Site"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   25
               Tag             =   "02040300"
               Top             =   645
               Width           =   5175
            End
            Begin VB.Label lblSubWharehousing 
               Caption         =   "Issue"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   24
               Tag             =   "02040200"
               Top             =   405
               Width           =   5130
            End
            Begin VB.Label lblSubWharehousing 
               AutoSize        =   -1  'True
               Caption         =   "Order Receipt"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   23
               Tag             =   "02040100"
               Top             =   165
               Width           =   5160
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   2895
            Index           =   4
            Left            =   3180
            TabIndex        =   30
            Top             =   5220
            Width           =   5475
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Global Transfer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   11
               Left            =   2640
               MousePointer    =   99  'Custom
               TabIndex        =   166
               Tag             =   "02050801"
               Top             =   180
               Width           =   1320
            End
            Begin VB.Label lblModifyFQA 
               AutoSize        =   -1  'True
               Caption         =   "Modify FQA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   158
               Tag             =   "02051200"
               Top             =   2580
               Width           =   1815
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Audit SAP Valuation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   10
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   120
               Tag             =   "02051100"
               Top             =   2340
               Width           =   5115
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "SAP Analysis Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   9
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   119
               Tag             =   "02051000"
               Top             =   2100
               Width           =   5115
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Transaction Valuation Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   118
               Tag             =   "02050800"
               Top             =   1620
               Width           =   5145
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Supplier Invoice Input"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   117
               Tag             =   "02050700"
               Top             =   1380
               Width           =   5130
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Condition Code Valuation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   116
               Tag             =   "02050600"
               Top             =   1140
               Width           =   5160
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "SAP inquiry"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   115
               Tag             =   "02050500"
               Top             =   900
               Width           =   5175
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Sale"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   114
               Tag             =   "02050400"
               Top             =   660
               Width           =   5175
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Inventory Write On"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   2640
               MousePointer    =   99  'Custom
               TabIndex        =   113
               Tag             =   "02050300"
               Top             =   420
               Width           =   2385
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Inventory Write Off"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   112
               Tag             =   "02050200"
               Top             =   420
               Width           =   2235
            End
            Begin VB.Label lblSubAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Inventory Initial Load"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   111
               Tag             =   "02050100"
               Top             =   180
               Width           =   2175
            End
            Begin VB.Label lblSubAccounting 
               Caption         =   "Sap Adjustment"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   110
               Tag             =   "02050900"
               Top             =   1860
               Width           =   5145
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   735
            Index           =   0
            Left            =   3180
            TabIndex        =   9
            Top             =   0
            Width           =   5475
            Begin VB.Label lblsubCatalog 
               Caption         =   "Create Modify Stock Record"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   11
               Tag             =   "02010100"
               Top             =   165
               Width           =   5220
            End
            Begin VB.Label lblsubCatalog 
               Caption         =   "Search on Stock Records"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   10
               Tag             =   "02010200"
               Top             =   405
               Width           =   5160
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   1635
            Index           =   2
            Left            =   3180
            TabIndex        =   16
            Top             =   720
            Width           =   5475
            Begin VB.Label lblSubPurchasing 
               Caption         =   "General Status Report(By Transaction)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   127
               Tag             =   "02020600"
               Top             =   1365
               Width           =   5220
            End
            Begin VB.Label lblSubPurchasing 
               Caption         =   "Approve && Send Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   21
               Tag             =   "02020500"
               Top             =   1125
               Width           =   5220
            End
            Begin VB.Label lblSubPurchasing 
               Caption         =   "Print Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   20
               Tag             =   "02020400"
               Top             =   885
               Width           =   5220
            End
            Begin VB.Label lblSubPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Close/Cancel Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   19
               Tag             =   "02020300"
               Top             =   645
               Width           =   5160
            End
            Begin VB.Label lblSubPurchasing 
               Caption         =   "Create Order Tracking Message"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   18
               Tag             =   "02020200"
               Top             =   405
               Width           =   5145
            End
            Begin VB.Label lblSubPurchasing 
               Caption         =   "Create/Revise Order"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Tag             =   "02020100"
               Top             =   165
               Width           =   5175
            End
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   32
            X1              =   2760
            X2              =   3180
            Y1              =   2220
            Y2              =   2220
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   31
            X1              =   2760
            X2              =   3180
            Y1              =   7680
            Y2              =   7680
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   9
            X1              =   2760
            X2              =   3180
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   6
            X1              =   2760
            X2              =   3180
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   28
            X1              =   2760
            X2              =   3180
            Y1              =   7440
            Y2              =   7440
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   47
            X1              =   2775
            X2              =   3195
            Y1              =   7920
            Y2              =   7920
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   44
            X1              =   2775
            X2              =   3195
            Y1              =   7200
            Y2              =   7200
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   43
            X1              =   2775
            X2              =   3195
            Y1              =   6720
            Y2              =   6720
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   41
            X1              =   2775
            X2              =   3195
            Y1              =   6960
            Y2              =   6960
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   40
            X1              =   2760
            X2              =   3180
            Y1              =   6480
            Y2              =   6480
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   39
            X1              =   2760
            X2              =   3180
            Y1              =   6240
            Y2              =   6240
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   38
            X1              =   2760
            X2              =   3180
            Y1              =   5760
            Y2              =   5760
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   37
            X1              =   2760
            X2              =   2760
            Y1              =   5520
            Y2              =   7920
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   36
            X1              =   2760
            X2              =   3180
            Y1              =   6000
            Y2              =   6000
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   35
            X1              =   2550
            X2              =   3180
            Y1              =   5520
            Y2              =   5520
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   29
            X1              =   2760
            X2              =   3180
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   27
            X1              =   2760
            X2              =   3180
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   26
            X1              =   2775
            X2              =   3195
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   25
            X1              =   2760
            X2              =   3180
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   24
            X1              =   2760
            X2              =   3180
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   23
            X1              =   2760
            X2              =   2760
            Y1              =   3600
            Y2              =   5040
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   22
            X1              =   2775
            X2              =   3195
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   21
            X1              =   2550
            X2              =   3195
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   20
            X1              =   2550
            X2              =   3195
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   19
            X1              =   2775
            X2              =   3195
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   18
            X1              =   2760
            X2              =   2760
            Y1              =   2640
            Y2              =   3120
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   17
            X1              =   2775
            X2              =   3195
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   14
            X1              =   2760
            X2              =   3180
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   13
            X1              =   2760
            X2              =   3180
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   12
            X1              =   2760
            X2              =   2760
            Y1              =   1020
            Y2              =   2220
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   11
            X1              =   2550
            X2              =   3360
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   8
            X1              =   2760
            X2              =   3240
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   7
            X1              =   2760
            X2              =   2760
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   5
            X1              =   2550
            X2              =   3195
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financial Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   420
            TabIndex        =   8
            Tag             =   "02050000"
            Top             =   5520
            Width           =   2055
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchasing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   6
            Tag             =   "02020000"
            Top             =   1020
            Width           =   2055
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logistics"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   5
            Tag             =   "02030000"
            Top             =   2640
            Width           =   1995
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cataloging"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   4
            Tag             =   "02010000"
            Top             =   195
            Width           =   2025
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   4
            X1              =   0
            X2              =   360
            Y1              =   5520
            Y2              =   5520
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   540
            Index           =   3
            Left            =   360
            Top             =   3300
            Width           =   2220
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   3
            X1              =   0
            X2              =   360
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   2
            X1              =   0
            X2              =   360
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   600
            Index           =   2
            Left            =   360
            Top             =   2460
            Width           =   2220
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   600
            Index           =   1
            Left            =   360
            Top             =   840
            Width           =   2220
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   600
            Index           =   0
            Left            =   360
            Top             =   0
            Width           =   2220
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   1
            X1              =   0
            X2              =   360
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   0
            X1              =   0
            X2              =   630
            Y1              =   250
            Y2              =   250
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   15
            X2              =   15
            Y1              =   240
            Y2              =   5520
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   480
            TabIndex        =   7
            Tag             =   "02040000"
            Top             =   3420
            Width           =   2025
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   600
            Index           =   4
            Left            =   360
            Top             =   5340
            Width           =   2220
         End
      End
      Begin VB.PictureBox picTables 
         BorderStyle     =   0  'None
         Height          =   12000
         Left            =   3000
         ScaleHeight     =   600
         ScaleMode       =   2  'Point
         ScaleWidth      =   438.75
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   8775
         Begin VB.Frame fraHolder 
            Height          =   3555
            Index           =   9
            Left            =   3180
            TabIndex        =   60
            Top             =   -60
            Width           =   5475
            Begin VB.Label lbleccnsource 
               AutoSize        =   -1  'True
               Caption         =   "Eccn Source"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   160
               Tag             =   "01010107"
               Top             =   1190
               Width           =   2550
            End
            Begin VB.Label lbleccn 
               AutoSize        =   -1  'True
               Caption         =   "Eccn"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   27
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   159
               Tag             =   "01010106"
               Top             =   1380
               Width           =   2490
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Group"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   26
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   139
               Tag             =   "01011800"
               Top             =   2985
               Width           =   2565
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Charge Account"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   25
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   138
               Tag             =   "01012000"
               Top             =   3180
               Width           =   2580
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Manufacturer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   24
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   137
               Tag             =   "01011700"
               Top             =   2775
               Width           =   2580
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Stock Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Index           =   23
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   136
               Tag             =   "01012100"
               Top             =   3180
               Width           =   2550
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Service Code Category"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   21
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   134
               Tag             =   "01011900"
               Top             =   2985
               Width           =   2565
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Category"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   19
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   132
               Tag             =   "01011600"
               Top             =   2775
               Width           =   2445
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   18
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   131
               Tag             =   "01011500"
               Top             =   2580
               Width           =   2520
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Phone Directory"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   17
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   130
               Tag             =   "01011400"
               Top             =   2385
               Width           =   2460
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "To Be Used For Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   16
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   129
               Tag             =   "01011300"
               Top             =   2580
               Width           =   2490
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Forwarder"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   15
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   128
               Tag             =   "01011200"
               Top             =   2385
               Width           =   2535
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Terms & Conditions"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   14
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   75
               Tag             =   "01011100"
               Top             =   2175
               Width           =   2490
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Terms of Delivery"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   13
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   74
               Tag             =   "01011000"
               Top             =   2175
               Width           =   2460
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Custom Category"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   12
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   73
               Tag             =   "01010900"
               Top             =   1980
               Width           =   2520
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Service Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   11
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   72
               Tag             =   "01010800"
               Top             =   1980
               Width           =   2520
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Document Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   10
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   71
               Tag             =   "01010700"
               Top             =   1785
               Width           =   2430
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Shipping Mode"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   9
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   70
               Tag             =   "01010600"
               Top             =   1785
               Width           =   2475
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Currency"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   2805
               MousePointer    =   99  'Custom
               TabIndex        =   69
               Tag             =   "01010500"
               Top             =   1575
               Width           =   2445
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Originator"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   68
               Tag             =   "01010400"
               Top             =   1575
               Width           =   2520
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Shipment Terms & Conditions"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   67
               Tag             =   "01010300"
               Top             =   1380
               Width           =   2610
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Shipper Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   66
               Tag             =   "01010200"
               Top             =   1185
               Width           =   2400
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Local Supplier Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   65
               Tag             =   "01010105"
               Top             =   975
               Width           =   3090
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "List of Supplier Code used in TO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   64
               Tag             =   "01010104"
               Top             =   780
               Width           =   5175
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Print Local Supplier Records"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   63
               Tag             =   "01010103"
               Top             =   585
               Width           =   5205
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "International Supplier Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   62
               Tag             =   "01010101"
               Top             =   180
               Width           =   2850
            End
            Begin VB.Label lblTblPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Print International Supplier Records"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   61
               Tag             =   "01010102"
               Top             =   375
               Width           =   5205
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   1035
            Index           =   10
            Left            =   3180
            TabIndex        =   79
            Top             =   3480
            Width           =   5475
            Begin VB.Label lblTblLogistics 
               AutoSize        =   -1  'True
               Caption         =   "Destination"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   83
               Tag             =   "01020400"
               Top             =   780
               Width           =   5175
            End
            Begin VB.Label lblTblLogistics 
               AutoSize        =   -1  'True
               Caption         =   "Sold to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   82
               Tag             =   "01020300"
               Top             =   585
               Width           =   5175
            End
            Begin VB.Label lblTblLogistics 
               AutoSize        =   -1  'True
               Caption         =   "Ship to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   81
               Tag             =   "01020200"
               Top             =   375
               Width           =   5175
            End
            Begin VB.Label lblTblLogistics 
               AutoSize        =   -1  'True
               Caption         =   "Bill to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   80
               Tag             =   "01020100"
               Top             =   180
               Width           =   5175
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   1875
            Index           =   11
            Left            =   3180
            TabIndex        =   84
            Top             =   4500
            Width           =   5475
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Location/SITE Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   2400
               MousePointer    =   99  'Custom
               TabIndex        =   157
               Tag             =   "01030900"
               Top             =   780
               Width           =   1680
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Condition"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   142
               Tag             =   "01030700"
               Top             =   1380
               Width           =   5175
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Company"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   141
               Tag             =   "01030800"
               Top             =   1575
               Width           =   5175
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Sub-Location Table"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   140
               Tag             =   "01030600"
               Top             =   1185
               Width           =   5175
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Logical Warehouse Table"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   90
               Tag             =   "01030500"
               Top             =   975
               Width           =   5175
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Location Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   89
               Tag             =   "01030400"
               Top             =   780
               Width           =   1440
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Country Utility"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   88
               Tag             =   "01030300"
               Top             =   585
               Width           =   5160
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Phone Directory"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   87
               Tag             =   "01030200"
               Top             =   375
               Width           =   5130
            End
            Begin VB.Label lblTblWharehouse 
               Caption         =   "Transaction Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   86
               Tag             =   "01030100"
               Top             =   180
               Width           =   2160
            End
         End
         Begin VB.Frame fraHolder 
            Height          =   1635
            Index           =   12
            Left            =   3180
            TabIndex        =   85
            Top             =   6360
            Width           =   5475
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "System File"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   147
               Tag             =   "01040700"
               Top             =   1380
               Width           =   5175
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Electronic Distribution (User)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   146
               Tag             =   "01040600"
               Top             =   1185
               Width           =   5100
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Electronic Distribution (System)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   145
               Tag             =   "01040500"
               Top             =   975
               Width           =   5175
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Auto-Numbering"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   144
               Tag             =   "01040400"
               Top             =   780
               Width           =   5085
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Site Consolidation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   143
               Tag             =   "01040300"
               Top             =   585
               Width           =   5145
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Site"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   92
               Tag             =   "01040200"
               Top             =   375
               Width           =   5145
            End
            Begin VB.Label lblTblAccounting 
               AutoSize        =   -1  'True
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   91
               Tag             =   "01040100"
               Top             =   180
               Width           =   5115
            End
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   75
            X1              =   135
            X2              =   159
            Y1              =   393
            Y2              =   393
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   74
            X1              =   135
            X2              =   159
            Y1              =   383
            Y2              =   383
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   73
            X1              =   135
            X2              =   159
            Y1              =   373
            Y2              =   373
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   89
            X1              =   138
            X2              =   159
            Y1              =   310.5
            Y2              =   310.5
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   87
            X1              =   135
            X2              =   159
            Y1              =   363
            Y2              =   363
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   86
            X1              =   138
            X2              =   159
            Y1              =   290.25
            Y2              =   290.25
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   85
            X1              =   138
            X2              =   159
            Y1              =   280.5
            Y2              =   280.5
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   83
            X1              =   135
            X2              =   159
            Y1              =   354
            Y2              =   354
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   82
            X1              =   138
            X2              =   159
            Y1              =   300.75
            Y2              =   300.75
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   76
            X1              =   138
            X2              =   159
            Y1              =   162
            Y2              =   162
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   600
            Index           =   6
            Left            =   390
            Top             =   6480
            Width           =   2145
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   59
            X1              =   0
            X2              =   18
            Y1              =   336
            Y2              =   336
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "System"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   94
            Tag             =   "01040000"
            Top             =   6660
            Width           =   1965
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   58
            X1              =   126
            X2              =   161.25
            Y1              =   333
            Y2              =   333
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   57
            X1              =   135
            X2              =   135
            Y1              =   333
            Y2              =   393
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   53
            X1              =   135
            X2              =   159
            Y1              =   343
            Y2              =   343
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   56
            X1              =   138
            X2              =   138
            Y1              =   240
            Y2              =   310.6
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   50
            X1              =   138
            X2              =   162
            Y1              =   249.75
            Y2              =   249.75
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   49
            X1              =   138
            X2              =   159
            Y1              =   260.25
            Y2              =   260.25
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   48
            X1              =   138
            X2              =   159
            Y1              =   270
            Y2              =   270
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   55
            X1              =   0
            X2              =   18
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Shape ShpSubActivities 
            BorderWidth     =   2
            FillColor       =   &H8000000F&
            Height          =   600
            Index           =   5
            Left            =   360
            Top             =   4620
            Width           =   2160
         End
         Begin VB.Label lblSubActivities 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   420
            TabIndex        =   93
            Tag             =   "01030000"
            Top             =   4800
            Width           =   1965
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   54
            X1              =   126
            X2              =   161.25
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   46
            X1              =   138
            X2              =   159
            Y1              =   219.75
            Y2              =   219.75
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   45
            X1              =   138
            X2              =   159
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   47
            X1              =   138
            X2              =   162
            Y1              =   200.25
            Y2              =   200.25
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   53
            X1              =   138
            X2              =   138
            Y1              =   189
            Y2              =   220.05
         End
         Begin VB.Label lblTblSubCat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Logistics"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   77
            Tag             =   "01020000"
            Top             =   3780
            Width           =   1965
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   600
            Index           =   1
            Left            =   360
            Top             =   3600
            Width           =   2160
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   44
            X1              =   0
            X2              =   18
            Y1              =   189
            Y2              =   189
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   39
            X1              =   126
            X2              =   159.75
            Y1              =   189
            Y2              =   189
         End
         Begin VB.Label lblTblSubCat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Purchasing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   76
            Tag             =   "01010000"
            Top             =   210
            Width           =   1920
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   43
            X1              =   138
            X2              =   159
            Y1              =   122
            Y2              =   122
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   42
            X1              =   138
            X2              =   159
            Y1              =   132
            Y2              =   132
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   41
            X1              =   138
            X2              =   159
            Y1              =   142
            Y2              =   142
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   40
            X1              =   138
            X2              =   159
            Y1              =   152
            Y2              =   152
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   38
            X1              =   138
            X2              =   159
            Y1              =   72
            Y2              =   72
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   37
            X1              =   138
            X2              =   159
            Y1              =   82
            Y2              =   82
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   36
            X1              =   138
            X2              =   159
            Y1              =   92
            Y2              =   92
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   35
            X1              =   138
            X2              =   159
            Y1              =   102
            Y2              =   102
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   34
            X1              =   138
            X2              =   159
            Y1              =   112
            Y2              =   112
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   33
            X1              =   138
            X2              =   159
            Y1              =   22
            Y2              =   22
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   32
            X1              =   138
            X2              =   159
            Y1              =   32
            Y2              =   32
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   31
            X1              =   138
            X2              =   159
            Y1              =   42
            Y2              =   42
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   30
            X1              =   138
            X2              =   159
            Y1              =   52
            Y2              =   52
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   29
            X1              =   138
            X2              =   159
            Y1              =   62
            Y2              =   62
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   52
            X1              =   138
            X2              =   138
            Y1              =   12
            Y2              =   162
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   28
            X1              =   126
            X2              =   159.75
            Y1              =   12
            Y2              =   12
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   27
            X1              =   0
            X2              =   18
            Y1              =   13
            Y2              =   13
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   600
            Index           =   0
            Left            =   360
            Top             =   30
            Width           =   2160
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            Index           =   0
            X1              =   0.75
            X2              =   0.75
            Y1              =   12.5
            Y2              =   336
         End
      End
      Begin VB.PictureBox PicSystem 
         BorderStyle     =   0  'None
         Height          =   9030
         Left            =   3000
         ScaleHeight     =   9030
         ScaleWidth      =   8835
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   8835
         Begin VB.Frame fraHolder 
            Height          =   2055
            Index           =   13
            Left            =   2880
            TabIndex        =   96
            Top             =   5280
            Width           =   5775
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Menu Option"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   156
               Tag             =   "04010600"
               Top             =   1185
               Width           =   5415
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "User Access Level"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   104
               Tag             =   "04010900"
               Top             =   1785
               Width           =   5445
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Menu Template"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   103
               Tag             =   "04010800"
               Top             =   1590
               Width           =   5400
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Menu Level"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   102
               Tag             =   "04010700"
               Top             =   1380
               Width           =   5445
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Temporary Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   101
               Tag             =   "04010500"
               Top             =   975
               Width           =   5475
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Change Personal Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   100
               Tag             =   "04010400"
               Top             =   780
               Width           =   5445
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Initial user Password settings"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   99
               Tag             =   "04010300"
               Top             =   585
               Width           =   5355
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "User Profile"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   98
               Tag             =   "04010200"
               Top             =   375
               Width           =   5445
            End
            Begin VB.Label lblSysPurchasing 
               AutoSize        =   -1  'True
               Caption         =   "Buyer table utility/User application rights"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   180
               MousePointer    =   99  'Custom
               TabIndex        =   97
               Tag             =   "04010100"
               Top             =   180
               Width           =   5400
            End
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   600
            Index           =   2
            Left            =   240
            Top             =   5910
            Width           =   2040
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   54
            X1              =   2520
            X2              =   2940
            Y1              =   7175
            Y2              =   7175
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   63
            X1              =   2520
            X2              =   2940
            Y1              =   6585
            Y2              =   6585
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   62
            X1              =   2520
            X2              =   2940
            Y1              =   6375
            Y2              =   6375
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   61
            X1              =   2520
            X2              =   2940
            Y1              =   5580
            Y2              =   5580
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   60
            X1              =   2520
            X2              =   2940
            Y1              =   5985
            Y2              =   5985
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   59
            X1              =   2520
            X2              =   2940
            Y1              =   5775
            Y2              =   5775
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   57
            X1              =   2520
            X2              =   2940
            Y1              =   6975
            Y2              =   6975
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   56
            X1              =   2520
            X2              =   2940
            Y1              =   6780
            Y2              =   6780
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   60
            X1              =   2520
            X2              =   2520
            Y1              =   5580
            Y2              =   7175
         End
         Begin VB.Label lblTblSubCat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Security Utility"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   105
            Tag             =   "04010000"
            Top             =   6090
            Width           =   1845
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   52
            X1              =   2270
            X2              =   3080
            Y1              =   6180
            Y2              =   6180
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   51
            X1              =   -120
            X2              =   240
            Y1              =   6195
            Y2              =   6195
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   3915
         Left            =   3240
         ScaleHeight     =   3915
         ScaleWidth      =   5160
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5160
      End
      Begin LRLabels.LRHyperLabel lrhReports 
         Height          =   480
         Left            =   120
         TabIndex        =   0
         Tag             =   "03000000"
         Top             =   2565
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   847
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Reports"
         MousePointer    =   99
         HyperLinkColor  =   12582912
         BeginProperty HyperLinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   16
         X1              =   304
         X2              =   332
         Y1              =   524
         Y2              =   524
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   15
         X1              =   252
         X2              =   280
         Y1              =   516
         Y2              =   516
      End
      Begin LRLabels.LRHyperLabel lrhActivities 
         Height          =   465
         Left            =   120
         TabIndex        =   109
         Tag             =   "02000000"
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   820
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Activities"
         MousePointer    =   99
         HyperLinkColor  =   12582912
         BeginProperty HyperLinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         Visible         =   0   'False
         X1              =   216
         X2              =   72
         Y1              =   193
         Y2              =   193
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         Visible         =   0   'False
         X1              =   80
         X2              =   216
         Y1              =   413
         Y2              =   413
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         Visible         =   0   'False
         X1              =   216
         X2              =   80
         Y1              =   272
         Y2              =   271
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   208
         X2              =   60
         Y1              =   21
         Y2              =   21
      End
      Begin LRLabels.LRHyperLabel lrhSystem 
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Tag             =   "04000000"
         Top             =   5865
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   741
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "System"
         MousePointer    =   99
         HyperLinkColor  =   12582912
         BeginProperty HyperLinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LRLabels.LRHyperLabel lrhTables 
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Tag             =   "01000000"
         Top             =   3735
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   741
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tables"
         MousePointer    =   99
         HyperLinkColor  =   12582912
         BeginProperty HyperLinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   72
      X1              =   0
      X2              =   28
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTblPurchasing 
      AutoSize        =   -1  'True
      Caption         =   "To Be Used For Utility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   22
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   135
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lblTblPurchasing 
      AutoSize        =   -1  'True
      Caption         =   "To Be Used For Utility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   20
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   133
      Top             =   0
      Width           =   1890
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   10
      X1              =   0
      X2              =   28
      Y1              =   0
      Y2              =   0
   End
   Begin LRLabels.LRHyperLabel LRHyperLabel1 
      Height          =   480
      Left            =   0
      TabIndex        =   108
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reports"
      MousePointer    =   99
      HyperLinkColor  =   12582912
      BeginProperty HyperLinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRepCataloging 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   107
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WH As imsWarehouse.WareHouse
Public SC As ImsSecX.imsSecMod
'Dim TableLocked As Boolean

'JCG 2009-01-13
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1
'----------------------

Private Sub Form_Activate()
    WindowState = vbNormal
End Sub

'set window status

Private Sub Form_GotFocus()
    Me.Visible = True
    WindowState = vbNormal
End Sub

'load icon and cursor from resuor file and return it to approapriate controls

Private Sub Form_Load()

'Added by Juan (8/29/2000) for Multilingual
Call translator.Translate_Forms("frmNavigator")
'------------------------------------------

On Error Resume Next

Dim ctl As Control
    For Each ctl In Controls
        If ((TypeOf ctl Is Label) And (ctl.ForeColor = &HC00000)) Then
            ctl.MousePointer = 99
            ctl.MouseIcon = LoadResPicture(101, vbResCursor)
        End If
        
        If Err Then Err.Clear
    Next ctl
    
    lrhActivities_OnHyperLinkEnter
    Call Move(3840, 0)
End Sub

Private Sub lblLocationSup_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

If LogOff.Visible = False Then

    Cancel = True
    Load LogOff
    LogOff.Show
    
Else

    Cancel = False
    
End If
    
''If MsgBox("Are you sure you want to Exit?", vbCritical + vbYesNo, "Imswin") = vbYes Then
''Cancel = False
''
''Else
''Cancel = True
''End If

End Sub

Private Sub lbleccn_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
         Load frmEccn
         frmEccn.Show
    Screen.MousePointer = vbArrow
End Sub

Private Sub lbleccnsource_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
         Load frmPicklist
         frmPicklist.Show
         
    Screen.MousePointer = vbArrow
End Sub



Private Sub lblManifestPOD_Click()

    Screen.MousePointer = vbHourglass
         Load frmmanifestpod
         frmmanifestpod.Show
    Screen.MousePointer = vbArrow

End Sub

Private Sub lblModifyFQA_Click()
    Screen.MousePointer = vbHourglass
         Load FrmModifyFQA
         FrmModifyFQA.Show
    Screen.MousePointer = vbArrow
End Sub

Private Sub lblnewSupplier_Click()
Load frm_newSupplier
 frm_newSupplier.Show
End Sub

Private Sub lblreorder_Click(Index As Integer)
'       Screen.MousePointer = vbHourglass
       
       
'JCG 2009-13-01
Dim x
'x = ShellExecute(Me.HWND, "Open", "http://localhost:81/", &O0, &O0, SW_NORMAL)
x = ShellExecute(Me.HWND, "Open", "http://localhost:8080/Exxon/html/logon.html", &O0, &O0, SW_NORMAL)
'-----------------------
       
       
''''''    MDI_IMS.CrystalReport1.Reset
''''''
''''''
''''''    'Report call
''''''    With MDI_IMS.CrystalReport1
  ''  MDI_IMS.CrystalReport1.Connect = "Data Source=imsdev003;UID=sa;PWD=scms;DSQ=PectenEccn;"
''''''        .ReportFileName = ReportPath + "reorder.rpt"
'''''' '       Call translator.Translate_Reports("reorder.rpt")
''''''       ' .ParameterFields(0) = "Namespace;" + deIms.NameSpace + ";TRUE"
''''''
'''''''.Connect = "DSN=imsO;UID=sa;PWD=scms;DSQ=pecteneccn;"
''''''
'''''''.DiscardSavedData = 1
'''''''if you have filter for selection formula
''''''.SelectionFormula = "{vwreorderreport.namespace}='" & deIms.NameSpace & "'"
''''''
''''''        .Action = 1: .Reset
''''''    End With

'JCG 1/13/2009 commented to use it for optima
'JCGFIXES ADDED 1/9/2006
'Screen.MousePointer = vbArrow
'Dim frm As Form
'Set frm = frm_reorder
'Load frm
'frm.WindowState = vbNormal

'Call frm.Move(0, 0)
'Call frm.Show(vbModeless)
'------------------------

'''JCGFIXES COMMENTED OUT 1/9/2006
'''    With MDI_IMS.CrystalReport1
'''            .Reset
'''            .ReportFileName = ReportPath & "reorder.rpt"
'''    ''            .ParameterFields(0) = "Namespace;" + deIms.NameSpace + ";TRUE"
'''    ''            .ParameterFields(0) = "prmStartingReq#;" + IIf(UCase(Trim$(combo_begpo.Text)) = "ALL", "ALL", combo_begpo.Text) + ";TRUE"
'''    ''            .ParameterFields(1) = "prmStoppingReq#;" + IIf(UCase(Trim$(combo_begpo.Text)) = "ALL", "", Trim$(combo_endpo.Text)) + ";TRUE"
'''    ''            .ParameterFields(2) = "prmStartingReqCreateDate;date(" & Year(DTbegdate.Value) & "," & Month(DTbegdate.Value) & "," & Day(DTbegdate.Value) & ");TRUE"
'''    ''            .ParameterFields(3) = "prmStoppingReqCreateDate;date(" & Year(DTenddate.Value) & "," & Month(DTenddate.Value) & "," & Day(DTenddate.Value) & ");TRUE"
'''    ''            .ParameterFields(4) = "prmOnlyOpen;" + IIf(optYes.Value = True, "Y", "N") + ";TRUE"
'''            .ParameterFields(0) = "Namespace;" + deIms.NameSpace + ";TRUE"

'''            msg1 = translator.Trans("M00278") 'J added
'''            .WindowTitle = IIf(msg1 = "", "Reorder Report", msg1) 'J modified
'''             Call translator.Translate_Reports("reorder.rpt")
            
'''            .Action = 1
'''            .Reset
            '.Action = 1: .Reset
    
End Sub

'Private Sub Form_Paint()
''    'Added by Juan (8/29/2000) for Multilingual
''    lrhActivities.Width = TextWidth(lrhActivities.Caption) * 2
''    lrhReports.Width = TextWidth(lrhReports.Caption) * 2
''    lrhTables.Width = TextWidth(lrhTables.Caption) * 2
''    lrhSystem.Width = TextWidth(lrhSystem.Caption) * 2
''    '------------------------------------------
'End Sub

'set forms categories when load accounting menu

Private Sub lblRepAccounting_Click(Index As Integer)
On Error Resume Next

Dim frm As Form
Select Case Index

Case 0
    Set frm = frm_sap_inquiry

Case 1
   Set frm = frm_tranvaluationreport

Case 2
   Set frm = FrmModifyFQA
Case 3
   Set frm = frm_sap_analysis
   
Case 4
    Set frm = priceControlReport
Case 5
    Set frm = uploadReport
End Select


    Load frm
    frm.WindowState = vbNormal
    
    'Call frm.Move(0, 0)
    Call frm.Show(vbModeless)
    If Err Then Call LogErr(Name & "::lblRepAccounting_Click", Err.Description, Err)
End Sub

'set crystal report parameter and application path

Private Sub lblRepCataloging_Click(Index As Integer)
On Error GoTo ErrHandler
       
    MDI_IMS.CrystalReport1.Reset

Select Case Index

Case 0
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\stckmaster.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("stckmaster.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
    
Case 1
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\Xcrossmanu.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("Xcrossmanu.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With

 Case 4
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\stckmasterX.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("stckmasterX.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With

 End Select

    Exit Sub
 
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Call LogErr(Name & "::lblRepCataloging_Click", Err.Description, Err)
        Err.Clear
    End If
End Sub

'load purchasing category forms and crystal report forms

Private Sub lblRepPurchasing_Click(Index As Integer)
On Error Resume Next

Dim frm As Form
Select Case Index

Case 0
     Set frm = frm_transact_order
     
Case 1
     Set frm = frm_stockhistory
     
Case 2
       On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\orderopen.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("orderopen.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
    
Case 3
    Set frm = frm_order_activity
    
Case 4
    Set frm = frm_ordertracking
    
Case 5
    Set frm = frm_orderdelivery
    
Case 6
    Set frm = frm_latedelivery
    
Case 7
    Set frm = frm_lateshipping
    

Case 8

    'Modified by Juan (8/2/2000) for Multilingual
    msg1 = translator.Trans("M00088") 'J added
    MsgBox IIf(msg1 = "", ("does not exist yet"), msg1) 'J modified
    '--------------------------------------------

Case 9
    Shell App.Path + "\imsclient.exe", vbMaximizedFocus
End Select
    
    Load frm
    frm.WindowState = vbNormal
    
    Call frm.Move(0, 0)
    Call frm.Show(vbModeless)
    Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Call LogErr(Name & "::lblRepPurchasing_Click", Err.Description, Err)
    End If
End Sub

'load warehouse catagory forms and crystal reports forms

Private Sub lblRepWharehouse_Click(Index As Integer)

MDI_IMS.CrystalReport1.Reset
Select Case Index
'Added by Juan (2007/6/25
Case 18
    Load frm_menuhistory
    frm_menuhistory.Show
    frm_menuhistory.ZOrder
'------------------------
Case 17
On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        '.Reset
        .ReportFileName = FixDir(App.Path) + "CRreports\accesslevel.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("accesslevel.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
Case 16
On Error GoTo ErrHandler1
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\userstatus.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("userstatus.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler1:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If

Case 15
    Load frm_loginlogoff
    frm_loginlogoff.Show
    frm_loginlogoff.ZOrder

Case 11
    Load frm_individualuserprofile
    frm_individualuserprofile.Show
    frm_individualuserprofile.ZOrder
Case 1
    'MsgBox "Feature does Not exist as yet"
    Load frm_inventoryperstocknu
    frm_inventoryperstocknu.Show
    frm_inventoryperstocknu.ZOrder
   
Case 2
Load frm_tranperdaterange
 frm_tranperdaterange.Show
 frm_tranperdaterange.ZOrder
 
Case 3
Load StockOnHand
    StockOnHand.ZOrder 0
    StockOnHand.Visible = True
    StockOnHand.Show
Case 4
 Load frm_slowmoving
    frm_slowmoving.Show
    frm_slowmoving.ZOrder 0
    
Case 5
   Load frm_historicalstock
   frm_historicalstock.Show
   frm_historicalstock.ZOrder 0
Case 6
'   Load frm_physicalinventory
' frm_physicalinventory.Show
' frm_physicalinventory.ZOrder
    Load StockOnHandNew
    StockOnHandNew.Show
    StockOnHandNew.ZOrder 0
Case 8
     Load frm_sohaccrosslocation
     frm_sohaccrosslocation.Show
     frm_sohaccrosslocation.ZOrder
Case 7
    Load frmStockOnHandStock
    frmStockOnHandStock.ZOrder 0
    frmStockOnHandStock.Visible = True
    frmStockOnHandStock.Show

 Case 0
   On Error GoTo ErrHandler2
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\ordertoberevcd.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("ordertoberevcd.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler2:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    
    Case 13
    
    'On Error GoTo ErrHandler3
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\menubig.rpt"
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("menubig.rpt") 'J added
        Call translator.Translate_SubReports 'J added
        '---------------------------------------------
        
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler3:
 
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
 Case 14
    Load frmSecurityChangeLog
    frmSecurityChangeLog.Show
'    On Error GoTo ErrHandler4
'    'Report call
'    With MDI_IMS.CrystalReport1
'        .ReportFileName = FixDir(App.Path) + "CRreports\securitchange.rpt"
'
'        'Modified by Juan (8/28/2000) for Multilingual
'        Call translator.Translate_Reports("securitchange.rpt") 'J added
'        '---------------------------------------------
'
'        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'        .Action = 1: .Reset
'    End With
'        Exit Sub
    
ErrHandler4:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
    End Select
End Sub

Private Sub lblrequisitionstatus_Click()
Load frm_RequisitionRptInput
 frm_RequisitionRptInput.Show
End Sub

'load accounting catagory forms and crystal report forms

Public Sub lblSubAccounting_Click(Index As Integer)
'Dim WH As New imsWarehouse.WareHouse

    Set WH = New imsWarehouse.WareHouse
    Set WH.Connection = deIms.cnIms
    WH.currUSER = CurrentUser
    WH.NameSpace = deIms.NameSpace
    WH.reportPath = FixDir(App.Path) + "CRReports"
    WH.dsnFILE = ConnInfo.Dsnname
    WH.cEmailOutFolder = ConnInfo.EmailOutFolder
    WH.cPwd = ConnInfo.Pwd
    WH.cUid = ConnInfo.UId
    WH.Language = Language
    WH.ExtendedCurrency = GExtendedCurrency
    Select Case Index
         Case 0
            Load frmWHInitialAdjustment
            Call frmWHInitialAdjustment.Move(0, 0)
            frmWHInitialAdjustment.Visible = True
         Case 1
            WH.Loading AdjustmentIssue
            'Call MDI_IMS.OpenDLLChild
         Case 2
            WH.Loading AdjustmentEntry
            'Call MDI_IMS.OpenDLLChild
         Case 3
            WH.Loading Sales
            'Call MDI_IMS.OpenDLLChild
         Case 4
            Load frm_sap_inquiry
            frm_sap_inquiry.Show
          
         Case 5
            'Load the Condition form
            Load frm_Condition
            frm_Condition.ZOrder 0
            frm_Condition.Show
            
         Case 6
            'Load frm_invoice
            Load frmInvoice
            frmInvoice.Show
            'Call frm_invoice.Move(0, 0)
            'frm_invoice.Show
            'frm_invoice.ZOrder 0
            
         Case 7
           Load frm_tranvaluationreport
           frm_tranvaluationreport.Show
           frm_tranvaluationreport.ZOrder
          
         Case 8
            frmSapAdjustment.Show
           
         Case 9
            frm_sap_analysis.Show
            frm_sap_analysis.ZOrder 0
           
         Case 10
            frm_sap_analysis.Show
            frm_sap_analysis.ZOrder
            
            
        '        Case 0
        '
        '        Case 1
        '
        '        Case 2
        '            frm_Condition.Show
        '            frm_Condition.ZOrder 0
        ' Case 3
        '
        ' Case 4
        '    Load frm_ordertracking
        '    frm_ordertracking.Show
        '    frm_ordertracking.ZOrder 0
        Case 11
            WH.Loading2 GlobalTransfer
    End Select
End Sub

'load stock master category forms

Private Sub lblSubCatalog_Click(Index As Integer)
On Error Resume Next

    Select Case Index
        Case 0
            'frm_Stock2.Show
            Screen.MousePointer = vbHourglass
            Frm_StockMaster.Show
            Frm_StockMaster.ZOrder
            Screen.MousePointer = vbArrow
            'frm_Stock2.ZOrder
        Case 1
            Load frm_StockSearch
            frm_StockSearch.ZOrder 0
            frm_StockSearch.Show
        
    End Select
    
    If Err Then Call LogErr(Name & "::lblSubCatalog_Click", Err.Description, Err)
    Err.Clear
End Sub

'load manifest category forms

Private Sub lblSubLogisitic_Click(Index As Integer)
On Error Resume Next
Select Case Index
Dim frm As Form

Case 0
    Set frm = frmReception
Case 1
    Set frm = frmPackingList
  Case 2
    Set frm = frmTrackManifest
 End Select
 
    Load frm
    frm.WindowState = vbNormal
    
    Call frm.Move(0, 0)
    Call frm.Show(vbModeless)
    If Err Then Call LogErr(Name & "::lblSubLogistic_Click", Err.Description, Err)
End Sub

'load purchasing catagory forms

Private Sub lblSubPurchasing_Click(Index As Integer)
On Error Resume Next

Dim frm As Form
Dim FNameSpace As String
Select Case Index
Case 1
     'Set frm = frmTracking
     Set frm = Frm_TrackingPONew
Case 2
    Set frm = frmClose
Case 3
    'Set frm = FrmRequisition
    Set frm = frm_transact_order
Case 4
    Set frm = frmPOApproval
Case 5
    Set frm = frm_gnrlstatustransac
 Case 0
   'Set frm = frm_Purchase
   
   Screen.MousePointer = vbHourglass
   
   Set frm = frm_NewPurchase
   
    'FNamespace = "ANGOL"
 End Select
 
 
    Load frm
    frm.WindowState = vbNormal
    
    Call frm.Move(0, 0)
    Call frm.Show(vbModeless)
     Screen.MousePointer = vbArrow
    If Index = 0 Then
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        Call deIms.ActiveStockMasterLooKUP(deIms.NameSpace)
     End If
    
    If Err Then Call LogErr(Name & "::lblSubPurchasing_Click", Err.Description, Err)
End Sub

'load warehousing catagory forms

Public Sub lblSubWharehousing_Click(Index As Integer)
On Error Resume Next
'Dim WH As New imsWarehouse.WareHouse
    Set WH = New imsWarehouse.WareHouse
    Set WH.Connection = deIms.cnIms
    WH.currUSER = CurrentUser
    WH.NameSpace = deIms.NameSpace
    WH.reportPath = FixDir(App.Path) + "CRReports"
    'WH.dsnFILE = cnInfo.Dsnname
    WH.dsnFILE = ConnInfo.Dsnname
    WH.Language = Language
    WH.ExtendedCurrency = GExtendedCurrency
    WH.cEmailOutFolder = ConnInfo.EmailOutFolder
    WH.cPwd = ConnInfo.Pwd
    WH.cUid = ConnInfo.UId

    Select Case Index
        Case 0
            Call WH.Loading(WarehouseReceipt)
            'Call MDI_IMS.OpenDLLChild(WarehouseReceipt)
        Case 2
            Call WH.Loading(WarehouseIssue)
            'Call MDI_IMS.OpenDLLChild(WarehouseIssue)
        Case 3
            Call WH.Loading(ReturnFromWell)
            'Call MDI_IMS.OpenDLLChild(ReturnFromWell)
        Case 4
            Call WH.Loading(ReturnFromRepair)
            'Call MDI_IMS.OpenDLLChild(ReturnFromRepair)
        Case 5
            Call WH.Loading(WellToWell)
            'Call MDI_IMS.OpenDLLChild(WellToWell)
        Case 6
            Call WH.Loading(WarehouseToWarehouse)
            'Call MDI_IMS.OpenDLLChild(WarehouseToWarehouse)
        Case 7
            Call WH.Loading(InternalTransfer)
            'Call MDI_IMS.OpenDLLChild(InternalTransfer)
    End Select
    If Err Then Call LogErr(Name & "::lblSubWharehousing_Click", Err.Description, Err)
End Sub

'load system purchasing catagory forms and check user status
'and user levels

Private Sub lblSysPurchasing_Click(Index As Integer)
On Error Resume Next
'Dim SC As imssecx.imsSecMod

    Set SC = New ImsSecX.imsSecMod
    
    SC.UserName = CurrentUser
    SC.NameSpace = deIms.NameSpace
    SC.ReportFilePath = FixDir(App.Path) + "CRReports"
    Set SC.Connection = deIms.cnIms
    SC.Dsnname = ConnInfo.Dsnname
    SC.languageSELECTED = Language
    Select Case Index
        Case 0
        On Error Resume Next
         
        Call SC.ShowBuyers(deIms.NameSpace, deIms.cnIms)

        Case 1
            Call SC.AddUser(CurrentUser)
        Case 2
            Call SC.AssignTempOwnerPassWord(CurrentUser, True)
    
        Case 3
            If SC.CanChangePassword(deIms.NameSpace, CurrentUser, deIms.cnIms) Then
                SC.ChangePassword
            Else
                'Modified by Juan (8/28/2000) for Translation
                msg1 = translator.Trans("M00089") 'J added
                MsgBox IIf(msg1 = "", "Your Password is not old enough", msg1) 'J modified
                '--------------------------------------------
            End If
    
        Case 4
            Call SC.AssignTempOwnerPassWord(CurrentUser, False)
            
            
        Case 5
            Call SC.ShowMenuOptions(mfLevel, deIms.NameSpace, deIms.cnIms)
        Case 6
            Call SC.ShowMenuOptions(mfTemplate, deIms.NameSpace, deIms.cnIms)
        Case 7
            Call SC.ShowMenuOptions(mfUser, deIms.NameSpace, deIms.cnIms)
        Case 8
            Call SC.ShowMenuOptions(mfOption, deIms.NameSpace, deIms.cnIms)
   End Select
   
   
    Set SC = Nothing
    If Err Then
        Call MsgBox(Err.Description, vbCritical, Err.Source)
         If Err Then Call LogErr(Name & "::lblSysPurchasing_Click", Err.Description, Err)
        If Err Then Err.Clear
    End If
End Sub

'load system catagory forms and report forms

Private Sub lblTblAccounting_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Load frmStatus
            frmStatus.Show
            frmStatus.ZOrder
        
        Case 1
            Load frm_SiteDescript
            frm_SiteDescript.Show
            frm_SiteDescript.ZOrder
        
        Case 2
            Load frmSiteConsolidation
            frmSiteConsolidation.Show
            frmSiteConsolidation.ZOrder
            
        Case 3
            Load frmChrono
            frmChrono.Show
            frmChrono.ZOrder
        
        Case 4
            Load frmElecDistribution
            frmElecDistribution.Show
            frmElecDistribution.ZOrder

        Case 5
            Load frmEUserDistribution
            frmEUserDistribution.Show
            frmEUserDistribution.ZOrder
            
        Case 6
            Load frm_systemfile
            frm_systemfile.Show
            frm_systemfile.ZOrder
    End Select
    
    If Err Then Call LogErr(Name & "lblSubWharehousing_Click", Err.Description, Err)
End Sub

'load logical catagory forms and report forms

Private Sub lblTblLogistics_Click(Index As Integer)
On Error Resume Next

    Select Case Index
        
        Case 0
      '      If TableLocked = True Then
            Load frm_Billto
            frm_Billto.Show
            frm_Billto.ZOrder
       '     End If
        Case 1
            Load frm_ShipTo
            frm_ShipTo.Show
            frm_ShipTo.ZOrder
            
        Case 2
            Load frm_SoldTo
            frm_SoldTo.Show
            frm_SoldTo.ZOrder
            
        Case 3
            Load frm_Destination
            frm_Destination.Show
            frm_Destination.ZOrder
            
    End Select

    If Err Then Call LogErr(Name & "::lblTblLogistics_Click", Err.Description, Err)

End Sub

'load purchasing catagory forms and report forms

Private Sub lblTblPurchasing_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    Case 9
        Load frm_Priority
    frm_Priority.ZOrder 0
    frm_Priority.Visible = True
    Case 10
        'Load the Document Type form
    Load frm_Document
    frm_Document.ZOrder 0
'    frm_Document.Caption = "Document Type"
    frm_Document.Show
    Case 11
        Load frm_ServiceCode
    frm_ServiceCode.ZOrder 0
'    frm_ServiceCode.Caption = "Service Codes"
    frm_ServiceCode.Show
Case 12
       'Load the Custom form
    Load frm_Custom
    frm_Custom.ZOrder 0
   ' Call Move(0, 0, width, height)
    'Set Caption for the Custom
    'form window
'    frm_Custom.Caption = "Custom Category"
    frm_Custom.Show
Case 13
    Load frmTermDelivery
    frmTermDelivery.ZOrder 0
'    frmTermDelivery.Caption = "Term of Delivery"
    frmTermDelivery.Show

Case 14
    Load frmTermCondition
    frmTermCondition.ZOrder 0
'    frmTermCondition.Caption = "Terms of Condition"
    frmTermCondition.Show
Case 15
    frmForwarder.Show
  Case 16
    Load frm_ToBe
    frm_ToBe.ZOrder 0
'    frm_ToBe.Caption = "To Be"
    frm_ToBe.Show
  Case 17
      Load frm_Phone
    frm_Phone.ZOrder 0
'    frm_Phone.Caption = "Phone Directory"
    frm_Phone.Show
  Case 18
      'Load the Unit form
    Load frm_Unit
    frm_Unit.ZOrder 0
'    frm_Unit.Caption = "Unit"
    frm_Unit.Show
  Case 19
      'Load the Category form
    Load frm_Category
    frm_Category.ZOrder 0
    'Set Caption for the Category form
'    frm_Category.Caption = "Category"
    frm_Category.Show
  Case 21
      Load frmServiceCate
    frmServiceCate.ZOrder 0
'    frmServiceCate.Caption = "Service Code Category"
    frmServiceCate.Show
  
  Case 24
  '    'Load the manufacturer form
    Load frm_Manufacturer
    frm_Manufacturer.ZOrder 0
    frm_Manufacturer.Show
    'Set Caption for the Manufacturer
'    'form window.show
  
  Case 26
      'Load the Group Code form
    Load frm_GroupCode
    frm_GroupCode.ZOrder 0
    'Set Caption for the Group Code
    'form window
'    frm_GroupCode.Caption = "Group Code"
    frm_GroupCode.Show
    Case 25
        'Load the Charge form
    Load frm_Charge
    frm_Charge.ZOrder 0
    'Set Caption for the Charge form
'    frm_Charge.Caption = "Charge"
    frm_Charge.Show
    
        Case 0
           On Error GoTo ErrHandler
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\Intsupp.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("Intsupp.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::lblTblPurchasing_Click", Err.Description, Err)

        Err.Clear
    End If
            
        
        Case 1
        
On Error Resume Next

'    Load the International Supplier form
    Screen.MousePointer = vbHourglass

    Load frm_IntSupe
    frm_IntSupe.ZOrder 0

    Call frm_IntSupe.Move(0, 0)

    frm_IntSupe.Show
    Screen.MousePointer = vbDefault
            
        Case 2
            On Error GoTo ErrHandler5
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\Locsupp.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("Locsupp.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler5:
    If Err Then
        MsgBox Err.Description
        If Err Then Call LogErr(Name & "::lblTblPurchasing_Click", Err.Description, Err)

        Err.Clear
    End If
            
        Case 3
            On Error GoTo ErrHandler6
    'Report call
    With MDI_IMS.CrystalReport1
        .ReportFileName = FixDir(App.Path) + "CRreports\Usedsupp.rpt"
        
        'Modified by Juan (8/28/2000) for Multilingual
        Call translator.Translate_Reports("Usedsupp.rpt") 'J added
        '---------------------------------------------
        
        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
        .Action = 1: .Reset
    End With
        Exit Sub
    
ErrHandler6:
    If Err Then
        MsgBox Err.Description
        Err.Clear
    End If
        Case 4
            On Error Resume Next

'    Load the International Supplier form
    Screen.MousePointer = vbHourglass

''    Load frm_IntSupe1
''    frm_IntSupe1.ZOrder 0
''
''    Call frm_IntSupe1.Move(0, 0)
''
''    frm_IntSupe1.Show

    Load frm_LocSupe
    frm_LocSupe.ZOrder 0

    Call frm_LocSupe.Move(0, 0)

    frm_LocSupe.Show
    Screen.MousePointer = vbDefault
    'Report call
'''    With MDI_IMS.CrystalReport1
'''        .ReportFileName = FixDir(App.Path) + "CRreports\Notusedsupp.rpt"
'''
'''        'Modified by Juan (8/28/2000) for Multilingual
'''        Call translator.Translate_Reports("Notusedsupp.rpt") 'J added
'''        '---------------------------------------------
'''
'''        .ParameterFields(0) = "namespace;" + deIms.NameSpace + ";TRUE"
'''        .Action = 1: .Reset
'''    End With
        Exit Sub
    
'''ErrHandler7:
'''    If Err Then
'''        MsgBox Err.Description
'''        Err.Clear
'''    End If
        Case 5
On Error Resume Next
Dim ctl As Control

    'Load the Shipper form
    Load frm_Shipper
    frm_Shipper.ZOrder 0
    'Set Caption for the Shipper
    'form window
'    frm_Shipper.Caption = "Shipper"
    frm_Shipper.Visible = True
         Case 6
            Load frm_ShiptermsEdit
    frm_ShiptermsEdit.ZOrder 0
'    frm_ShiptermsEdit.Caption = "Ship Terms & Condition"
    frm_ShiptermsEdit.Visible = True
            
         Case 7
            Load frm_Originator
            frm_Originator.Show
            frm_Originator.ZOrder
            
        Case 8
   
    Load frmCurrency
    frmCurrency.Show
''        Case 9
''            Load frm_Document
''            frm_Document.Show
''            frm_Document.ZOrder
''
''        Case 10
'''            Load frm_Service
'''            frm_Service.Show
'''            frm_Service.ZOrder
''
''        Case 11
'''            Load frm_ServiceCode
'''            frm_ServiceCode.Show
'''            frm_ServiceCode.ZOrder
''
''        Case 12
'''            Load frm_Keys
'''            frm_Keys.Show
'''            frm_Keys.ZOrder
''
''        Case 13
''            Load frm_Custom
''            frm_Custom.Show
''            frm_Custom.ZOrder
''
''        Case 14
''            Load frm_ToBe
''            frm_ToBe.Show
''            frm_ToBe.ZOrder
''

           
                  
   End Select
   
End Sub

'load warehouse catagory forms and report forms

Private Sub lblTblWharehouse_Click(Index As Integer)
    Select Case Index
        
        Case 0
            Load frmTrantype
            frmTrantype.Show
            frmTrantype.ZOrder
            
        Case 1
            Load frm_Phone
            frm_Phone.Show
            frm_Phone.ZOrder
        
        Case 2
            Load frm_Country
            frm_Country.Show
            frm_Country.ZOrder
        
        Case 3
            Load frm_Location
            frm_Location.Show
            frm_Location.ZOrder
        
        Case 4
            Load frm_logical
            frm_logical.Show
            frm_logical.ZOrder
        
        Case 5
            Load frm_SubLocation
            frm_SubLocation.Show
            frm_SubLocation.ZOrder
            
        Case 6
            Load frm_Condition
            frm_Condition.Show
            frm_Condition.ZOrder

        Case 7
            Load frm_Company
            frm_Company.Show
            frm_Company.ZOrder
        Case 8
             Load frm_LocationSITE
            frm_LocationSITE.Show
            frm_LocationSITE.ZOrder
    End Select
End Sub

'call function to set activities menu font and lines format

Private Sub lrhActivities_OnHyperLinkEnter()
    Call HideLines
    NormalizeLinks
    Line1(0).Visible = True
    picActivities.Visible = True
    picActivities.ZOrder 0
    Call StickFont(lrhActivities)
    lrhActivities.Width = TextWidth(lrhActivities.Caption) * 2
End Sub

'call function to set reports menu picture font and lines format

Private Sub lrhReports_OnHyperLinkEnter()
    Call HideLines
    NormalizeLinks
    Line1(1).Visible = True
    picReport.Visible = True
    picReport.ZOrder 0
    Call StickFont(lrhReports)
    lrhReports.Width = TextWidth(lrhReports.Caption) * 2
End Sub

'call function to set system menu font and lines format

Private Sub lrhSystem_OnHyperLinkEnter()
    Call HideLines
    NormalizeLinks
    Line1(3).Visible = True
    PicSystem.Visible = True
    PicSystem.ZOrder 0
    Call StickFont(lrhSystem)
    lrhSystem.Width = TextWidth(lrhSystem.Caption) * 2
End Sub

'call function to set table menu font and lines format

Private Sub lrhTables_OnHyperLinkEnter()
    Call HideLines
    NormalizeLinks
    Line1(2).Visible = True
    picTables.Visible = True
    picTables.ZOrder 0
    Call StickFont(lrhTables)
    lrhTables.Width = TextWidth(lrhTables.Caption) * 2 + 10
End Sub

'call function to set color and size for menu

Public Sub NormalizeLinks()
    
    lrhSystem.ForeColor = vbRed
    lrhTables.ForeColor = vbRed
    lrhReports.ForeColor = vbRed
    lrhActivities.ForeColor = vbRed
    
    
    lrhSystem.Font.size = 12
    lrhTables.Font.size = 12
    lrhReports.Font.size = 12
    lrhActivities.Font.size = 12
  '  lrhSecurity.ForeColor = vbRed
End Sub

'function set lines

Public Sub HideLines()
Dim i As Integer, x As Integer, y As Integer

    y = Line1.LBound()
    x = Line1.UBound()
    
    Picture1.Visible = False
    Picture1.Enabled = False
    
    For i = y To x
        Line1(i).Visible = False
    Next i
End Sub

'function set font

Private Sub StickFont(lrh As LRHyperLabel)
    lrh.Font.size = 14
    lrh.ForeColor = lrh.HyperLinkColor
End Sub

