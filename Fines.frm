VERSION 5.00
Begin VB.Form Fines 
   Caption         =   "Fines Due"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_fin 
      Caption         =   "Fines Due"
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
      Top             =   240
      Width           =   9495
      Begin VB.VScrollBar VScroll1 
         Height          =   1815
         Left            =   9120
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.ListBox txt_fin 
         Height          =   1815
         ItemData        =   "Fines.frx":0000
         Left            =   120
         List            =   "Fines.frx":0002
         TabIndex        =   1
         Top             =   360
         Width           =   9255
      End
   End
End
Attribute VB_Name = "Fines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

