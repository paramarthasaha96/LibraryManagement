VERSION 5.00
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_Iss 
      Caption         =   "ISSUE A BOOK"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   6735
   End
   Begin VB.Frame fr_Re 
      Caption         =   "RETURN A BOOK"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   7080
      TabIndex        =   3
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Frame fr_Cur 
      Caption         =   "CURRENTLY READING"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   14175
      Begin VB.VScrollBar VScroll1 
         Height          =   1695
         Left            =   13800
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txt_Cur 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   13935
      End
   End
   Begin VB.Menu mn_Fi 
      Caption         =   "File"
      Begin VB.Menu mn_E 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mn_Fea 
      Caption         =   "Features"
      Begin VB.Menu mn_Us 
         Caption         =   "Register an User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mn_B 
         Caption         =   "Register a Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mn_Ba 
         Caption         =   "View Books Available"
      End
      Begin VB.Menu mn_Ru 
         Caption         =   "View Registered Users"
      End
      Begin VB.Menu mn_Ub 
         Caption         =   "View Users Who Borrowed"
      End
      Begin VB.Menu mn_F 
         Caption         =   "View Users Whose Fines are Due"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub

Private Sub mn_E_Click()
Unload Me
End Sub

