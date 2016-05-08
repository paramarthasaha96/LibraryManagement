VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Reg_book 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register a Book"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Qt 
      DataField       =   "QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txt_Tmp 
      DataField       =   "CATEGORY"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "Books"
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
   Begin VB.TextBox txt_Bid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ComboBox cb_Cat 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2040
      List            =   "Reg_book.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txt_Aut 
      DataField       =   "AUTHOR"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txt_Name 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   5175
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lb_Qt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "QUANTITY :"
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
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lb_Bid 
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label lb_Aut 
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lb_Cat 
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lb_Name 
      BackColor       =   &H00C0E0FF&
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
Private Sub cb_Cat_Click()
txt_Tmp.Text = cb_Cat.List(cb_Cat.ListIndex)
End Sub


Private Sub co_Register_Click()
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveLast
    MsgBox ("THANKS FOR REGISTERING! THE BOOK ID NUMBER IS " & txt_Bid.Text)
    Adodc1.Recordset.AddNew
End Sub

Private Sub Form_Load()
Reg_book.Picture = LoadPicture("8e4a924b-4668-4630-94dc-51b0cdc6de30.jpg")
Adodc1.Recordset.AddNew
cb_Cat.ListIndex = 0
txt_Tmp.Text = cb_Cat.List(0)
End Sub
