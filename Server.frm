VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   870
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
      TabIndex        =   4
      Top             =   3360
      Width           =   6735
      Begin VB.TextBox txt_Nbc 
         DataField       =   "BOOK"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   5400
         TabIndex        =   39
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_Nuc 
         DataField       =   "NAME"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   5400
         TabIndex        =   38
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
         TabIndex        =   29
         Top             =   4320
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   4350
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   4350
         Width           =   255
      End
      Begin VB.TextBox txt_Iid 
         DataField       =   "ID"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txt_Di 
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_Nb 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txt_Idb 
         DataField       =   "BID"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txt_Nu 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txt_Idu 
         DataField       =   "UID"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lb_B 
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
         TabIndex        =   24
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lb_Cr 
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
         TabIndex        =   23
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lb_St 
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
         TabIndex        =   22
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lb_Iid 
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
         TabIndex        =   15
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label lb_Di 
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
         TabIndex        =   13
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lb_Nb 
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
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lb_Idb 
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
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lb_Nu 
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
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lb_Idu 
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
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame fr_Re 
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
      TabIndex        =   3
      Top             =   3360
      Width           =   7335
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
         TabIndex        =   37
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txt_Dr 
         Height          =   285
         Left            =   2760
         TabIndex        =   36
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txt_Nbr 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txt_Idbr 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txt_Nur 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_Idur 
         DataSource      =   "Adodc4"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_Rid 
         Height          =   285
         Left            =   2760
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label lb_Idbr 
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
         TabIndex        =   28
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lb_Idur 
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
         TabIndex        =   27
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lb_F 
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
         TabIndex        =   21
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lb_Dr 
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
         TabIndex        =   20
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lb_Nbr 
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
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lb_Nur 
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
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lb_Rid 
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
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fr_Cur 
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
   Begin VB.Menu mn_Fi 
      Caption         =   "File"
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

Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select distinct ID,NAME,UID,BOOK,BID,STATUS from Issue"
Adodc4.Refresh

txt_Idur.DataField = "UID"
txt_Idbr.DataField = "BID"
txt_Nur.DataField = "NAME"
txt_Nbr.DataField = "BOOK"
Adodc4.RecordSource = "Issue"
Adodc3.Recordset.AddNew
txt_Idur = ""
txt_Idbr = ""
txt_Nur = ""
txt_Nbr = ""

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
Adodc4.RecordSource = "select * from Issue where ID =" & txt_Rid.Text
Adodc4.Refresh
End Sub
