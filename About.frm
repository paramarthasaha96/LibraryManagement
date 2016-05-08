VERSION 5.00
Begin VB.Form About 
   Caption         =   "About"
   ClientHeight    =   5610
   ClientLeft      =   6465
   ClientTop       =   3180
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   Picture         =   "About.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   6105
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   "KOLKATA"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "INSTITUTE OF ENGINEERING AND MANAGEMENT"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "C3, CSE-2ND YEAR, B.Tech"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "RAJARSHI BASU- 118"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "PARAMARTHA SAHA- 99"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "PRESENTED BY:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   $"About.frx":2C513
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "LIBRARY MANAGEMENT"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
About.Picture = LoadPicture("library-books-shelves-1366x768-54712.jpg")
End Sub
