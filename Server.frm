VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Server 
   BackColor       =   &H00FAF2BA&
   Caption         =   "Server"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8880
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5160
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "Issue"
      Caption         =   "Adodc3"
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
   Begin VB.Frame fr_Iss 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ISSUE A BOOK"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   6735
      Begin VB.TextBox txt_Nbc 
         DataField       =   "BOOK"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   5400
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_Nuc 
         DataField       =   "NAME"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   5400
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3360
         Top             =   240
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
      Begin VB.CommandButton co_Issue 
         BackColor       =   &H80000004&
         Caption         =   "ISSUE"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4320
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Option2"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   4350
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   4350
         Width           =   255
      End
      Begin VB.TextBox txt_Iid 
         DataField       =   "ID"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txt_Di 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_Nb 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txt_Idb 
         DataField       =   "BID"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txt_Nu 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txt_Idu 
         DataField       =   "UID"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lb_B 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BORROW"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lb_Cr 
         BackStyle       =   0  'Transparent
         Caption         =   "READING"
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
         Left            =   2040
         TabIndex        =   21
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lb_St 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS :"
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
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lb_Iid 
         BackStyle       =   0  'Transparent
         Caption         =   "ISSUE ID GENERATED :"
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
         TabIndex        =   13
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label lb_Di 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF ISSUE :"
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
         TabIndex        =   11
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lb_Nb 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF THE BOOK :"
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
         TabIndex        =   9
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lb_Idb 
         BackStyle       =   0  'Transparent
         Caption         =   "ID OF THE BOOK :"
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
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lb_Nu 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF THE USER :"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lb_Idu 
         BackStyle       =   0  'Transparent
         Caption         =   "ID OF THE USER :"
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
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame fr_Re 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RETURN A BOOK"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   7080
      TabIndex        =   1
      Top             =   3360
      Width           =   7335
      Begin VB.TextBox txt_brd 
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   2880
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
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
      Begin VB.TextBox txt_F 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txt_Dr 
         Height          =   285
         Left            =   2760
         TabIndex        =   34
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txt_Nbr 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txt_Idbr 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   32
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txt_Nur 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   31
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_Idur 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   30
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_Rid 
         Height          =   285
         Left            =   2760
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton co_Return 
         Caption         =   "RETURN"
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
         Left            =   5400
         TabIndex        =   28
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label lb_Idbr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ID OF THE BOOK :"
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
         TabIndex        =   26
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lb_Idur 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ID OF THE USER :"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lb_F 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FINE TO BE PAID :"
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
         TabIndex        =   19
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lb_Dr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DATE OF RETURN :"
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
         TabIndex        =   18
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lb_Nbr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NAME OF THE BOOK :"
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
         TabIndex        =   17
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lb_Nur 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NAME OF THE USER :"
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
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lb_Rid 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ISSUE ID :"
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
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fr_Cur 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CURRENTLY READING"
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
      Top             =   600
      Width           =   14175
      Begin MSDataGridLib.DataGrid Dg_cr 
         Bindings        =   "Server.frx":0000
         Height          =   1815
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mn_Fi 
      Caption         =   "File"
      Begin VB.Menu mn_Fd 
         Caption         =   "Fine Details"
      End
      Begin VB.Menu mn_E 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mn_Fea 
      Caption         =   "Features"
      Begin VB.Menu mn_Us 
         Caption         =   "Register an User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mn_B 
         Caption         =   "Register a Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mn_Ba 
         Caption         =   "View Books Available"
      End
      Begin VB.Menu mn_Ru 
         Caption         =   "View Registered Users"
      End
      Begin VB.Menu mn_Ub 
         Caption         =   "View Users Who Borrowed"
      End
      Begin VB.Menu mn_F 
         Caption         =   "View Users Whose Fines are Due"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub co_Issue_Click()
    'Adodc1.RecordSource = "select * from Users where ID =" & txt_Idu.Text
    'Adodc2.RecordSource = "select * from Books where ID =" & txt_Idb.Text
    'Adodc2.Refresh
    'Adodc1.Refresh
    If txt_Idu = "" Then
    MsgBox ("Sorry No User ID Input Found!")
    ElseIf txt_Idb = "" Then
    MsgBox ("Sorry No Book ID Input Found!")
    ElseIf txt_Nu = "" Then
    MsgBox ("Sorry No User Exists with that ID!")
    ElseIf txt_Nb = "" Then
    MsgBox ("Sorry No Book Exists with that ID!")

    
    Else
    txt_Nuc = txt_Nu
    txt_Nbc = txt_Nb
    Adodc3.Recordset.Update
    Adodc3.Recordset.AddNew
    txt_Nb = ""
    txt_Nu = ""
    End If
    
    
    
End Sub

Private Sub co_Return_Click()
Adodc4.RecordSource = "select * from Issue where ID =" & txt_Rid.Text
Adodc4.Refresh
If txt_Rid = "" Then
    MsgBox ("Sorry No ID Input Found!")
    ElseIf txt_Idur = "" Then
    MsgBox ("Sorry No Such Record Found!")
    Else
    Adodc4.Recordset.Delete
    End If
    
    
txt_Rid = ""
txt_Idur = ""
txt_Idbr = ""
txt_Nur = ""
txt_Nbr = ""

'code for Currently Reading
Adodc3.RecordSource = "select * from ISSUE where STATUS = 'C' ORDER BY ID ASC"
Adodc3.Refresh
Dg_cr.Refresh
With Dg_cr
    .Columns(6).Visible = False
    .Columns(0).Width = 1000
    .Columns(1).Width = 5500
    .Columns(2).Width = 1000
    .Columns(3).Width = 4650
    .Columns(4).Width = 1200
    .Columns(5).Visible = False
End With
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select distinct ID,EMAIL,NAME,AGE from Users"
Adodc1.Refresh
txt_Nu.DataField = "NAME"
Adodc1.RecordSource = "Users"
txt_Nu = ""
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select distinct ID,NAME,AUTHOR,CATEGORY,QUANTITY from Books"
Adodc2.Refresh
txt_Nb.DataField = "NAME"
Adodc2.RecordSource = "Books"
txt_Nb = ""

'adodc code for currently reading
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select distinct ID,NAME,UID,BOOK,BID,STATUS from Issue"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select distinct ID,NAME,UID,BOOK,BID,STATUS,BORROW_DATE from Issue"
Adodc4.Refresh

txt_Idur.DataField = "UID"
txt_Idbr.DataField = "BID"
txt_Nur.DataField = "NAME"
txt_Nbr.DataField = "BOOK"
txt_brd.DataField = "BORROW_DATE"
Adodc4.RecordSource = "Issue"
Adodc4.Recordset.AddNew
txt_Idur = ""
txt_Idbr = ""
txt_Nur = ""
txt_Nbr = ""

'code for date
Dim dt As Date
dt = DateValue(Now)
txt_Di.Text = dt
txt_Dr.Text = dt

'code for Currently Reading
Adodc3.RecordSource = "select * from ISSUE where STATUS = 'C' ORDER BY ID ASC"
Adodc3.Refresh
Dg_cr.Refresh
With Dg_cr
    .Columns(6).Visible = False
    .Columns(0).Width = 1000
    .Columns(1).Width = 5500
    .Columns(2).Width = 1000
    .Columns(3).Width = 4650
    .Columns(4).Width = 1200
    .Columns(5).Visible = False
End With

End Sub

Private Sub mn_B_Click()
Reg_book.Show
End Sub

Private Sub mn_Ba_Click()
View_Books.Show
End Sub

Private Sub mn_E_Click()
Unload Me
End Sub

Private Sub mn_Fd_Click()
Fine_Details.Show
End Sub

Private Sub mn_Ru_Click()
View_User.Show
End Sub

Private Sub mn_Ub_Click()
View_Br.Show
End Sub

Private Sub mn_Us_Click()
Reg_user.Show
End Sub

Private Sub txt_Idb_LostFocus()
On Error Resume Next
If txt_Idb = "" Then
txt_Nb = ""
Else
Adodc2.RecordSource = "select * from Books where ID =" & txt_Idb.Text
Adodc2.Refresh
End If
End Sub

Private Sub txt_Idu_LostFocus()
On Error Resume Next
If txt_Idu = "" Then
txt_Nu = ""
Else
Adodc1.RecordSource = "select * from Users where ID =" & txt_Idu.Text
Adodc1.Refresh

End If

End Sub

Private Sub txt_Rid_LostFocus()
If Not (txt_Rid = "") Then
Adodc4.RecordSource = "select * from Issue where ID =" & txt_Rid.Text
Adodc4.Refresh
'Update Currently Reading list
Adodc3.Refresh
Dg_cr.Refresh
With Dg_cr
    .Columns(6).Visible = False
    .Columns(0).Width = 1000
    .Columns(1).Width = 5500
    .Columns(2).Width = 1000
    .Columns(3).Width = 4650
    .Columns(4).Width = 1200
    .Columns(5).Visible = False
End With

'fine calculation
Dim db, dr As Date
Dim dd As Integer
dr = txt_Dr.Text
db = txt_brd.Text
dd = DateDiff("d", db, dr)
If (dd <= 30) Then
dd = 0
ElseIf (dd > 30 And dd <= 60) Then
dd = (dd - 30) * 2
ElseIf (dd > 60) Then
dd = 60 + 5 * (dd - 60)
End If
txt_F = dd
End If
End Sub
