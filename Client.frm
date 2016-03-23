VERSION 5.00
Begin VB.Form Client 
   Caption         =   "Form2"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SETTINGS"
      Height          =   255
      Left            =   13200
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MY PROFILE"
      Height          =   255
      Left            =   11400
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FINES DUE"
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORROW A BOOK"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1455
      Left            =   10680
      TabIndex        =   5
      Top             =   3840
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "RECENTLY BORROWED"
      Height          =   2295
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   10695
      Begin VB.OptionButton Option4 
         Caption         =   "View Recently Borrowed"
         Height          =   195
         Left            =   8280
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "View Currently Borrowed"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ListBox List2 
         Height          =   1425
         ItemData        =   "Client.frx":0000
         Left            =   120
         List            =   "Client.frx":0002
         TabIndex        =   4
         Top             =   360
         Width           =   10455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RECENTLY READ"
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10695
      Begin VB.OptionButton Option2 
         Caption         =   "View Recently Read"
         Height          =   255
         Left            =   8520
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "View Currently Reading"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1455
         Left            =   10320
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "Client.frx":0004
         Left            =   120
         List            =   "Client.frx":0006
         TabIndex        =   2
         Top             =   360
         Width           =   10455
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

