VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Server 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   6735
   End
   Begin VB.Frame fr_Reg 
      Caption         =   "REGISTER USER"
      Height          =   4335
      Left            =   7080
      TabIndex        =   3
      Top             =   3360
      Width           =   7335
      Begin VB.CommandButton but_Reg 
         Caption         =   "REGISTER"
         Height          =   495
         Left            =   4800
         TabIndex        =   12
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txt_Id 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2640
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dt_Date 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110821377
         CurrentDate     =   42450
      End
      Begin VB.TextBox txt_Age 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox txt_Name 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lab_Date 
         Caption         =   "Date :"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lab_Id 
         Caption         =   "ID NUMBER :"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lab_Age 
         Caption         =   "AGE :"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lab_Name 
         Caption         =   "NAME :"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fr_Cur 
      Caption         =   "CURRENTLY READING"
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
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub
