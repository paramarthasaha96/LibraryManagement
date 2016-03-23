VERSION 5.00
Begin VB.Form Reg_book 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register a Book"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Uid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox cb_Cat 
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
      ItemData        =   "Reg_book.frx":0000
      Left            =   2160
      List            =   "Reg_book.frx":0016
      TabIndex        =   7
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txt_Aut 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txt_Name 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "UNIQUE BOOK ID GENERATED :"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lb_Aut 
      Caption         =   "AUTHOR NAME :"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lb_Cat 
      Caption         =   "CATEGORY :"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lb_Name 
      Caption         =   "NAME :"
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
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Reg_book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
