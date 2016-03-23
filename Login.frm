VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton co_sgn 
      Caption         =   "SIGN UP"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton co_lg 
      Caption         =   "LOGIN"
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
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame fr_lg 
      Caption         =   "LOGIN AS:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
      Begin VB.OptionButton op_us 
         Caption         =   "USER"
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
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton op_adm 
         Caption         =   "ADMIN"
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_pas 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txt_uid 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lb_pas 
      Caption         =   "PASSWORD:"
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
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lb_uid 
      Caption         =   "US-ID:"
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
      Width           =   855
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub co_lg_Click()
'Play, click on options and login to test code. Add code for UID and password later
If (op_adm.Value = True) Then
Server.Show
Unload Login
ElseIf (op_us.Value = True) Then
Client.Show
Unload Login
End If
End Sub

Private Sub co_sgn_Click()
If (op_adm.Value = True) Then
'add code for registering admins
Unload Login
ElseIf (op_us.Value = True) Then
Reg_user.Show
Unload Login
End If
End Sub
