VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Reg_user 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register a User"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=LIBDATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=LIBDATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Users"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txt_Uid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txt_Id 
      DataField       =   "EMAIL"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox txt_age 
      DataField       =   "AGE"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txt_Name 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   5
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton co_Register 
      BackColor       =   &H0000C000&
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lb_Uid 
      BackColor       =   &H0080C0FF&
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
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lb_Eid 
      BackColor       =   &H0080C0FF&
      Caption         =   "EMAIL ID :"
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
      Width           =   1215
   End
   Begin VB.Label lb_Age 
      BackColor       =   &H0080C0FF&
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
      BackColor       =   &H0080C0FF&
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

Private Sub co_Register_Click()
   Adodc1.Recordset.Update
   Adodc1.Recordset.MoveLast
   MsgBox ("THANKS FOR REGISTERING! Your ID NUMBER IS " & txt_Uid.Text)
   Adodc1.Recordset.AddNew
   
   
End Sub

Private Sub Form_Load()
Reg_user.Picture = LoadPicture("8e4a924b-4668-4630-94dc-51b0cdc6de30.jpg")
Adodc1.Recordset.AddNew
End Sub

