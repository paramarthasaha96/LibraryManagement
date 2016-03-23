VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Reg_user 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register a User"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dt_Date 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   42270721
      CurrentDate     =   42452
   End
   Begin VB.TextBox txt_Uid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txt_Id 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox txt_age 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txt_Name 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton co_Register 
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
      Left            =   4080
      TabIndex        =   0
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label lb_Uid 
      Caption         =   "UNIQUE GENERATED ID (U-ID) :"
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
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lb_Date 
      Caption         =   "DATE :"
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
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lb_Id 
      Caption         =   "ID NUMBER :"
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
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lb_Age 
      Caption         =   "AGE :"
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
      Top             =   1200
      Width           =   855
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Reg_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
