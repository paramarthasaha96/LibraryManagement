VERSION 5.00
Begin VB.Form Client 
   Caption         =   "Form2"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "READ A BOOK"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORROW A BOOK"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "Client.frx":0000
      Left            =   240
      List            =   "Client.frx":0007
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Recently Borrowed"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Recently Read"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

