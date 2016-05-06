VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form View_User 
   Caption         =   "Users"
   ClientHeight    =   6075
   ClientLeft      =   3585
   ClientTop       =   1890
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Britannic Bold"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1920
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
   Begin VB.Frame fr_uv 
      Caption         =   "LIST OF USERS"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9375
      Begin VB.CommandButton co_S 
         Caption         =   "SEARCH"
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
         Left            =   7680
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton op_All 
         Caption         =   "SHOW ALL"
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
         Left            =   7680
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_Id 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_Name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton op_uid 
         Caption         =   "USER-ID"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton op_name 
         Caption         =   "NAME"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid uv_dg 
         Bindings        =   "View_User.frx":0000
         Height          =   2895
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5106
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
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
            Name            =   "Britannic Bold"
            Size            =   9
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
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "VIEW BY:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "View_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub co_S_Click()
If op_name.Value = True Then
If Not (txt_Name.Text = "") Then
Adodc1.RecordSource = "select * from Users where Name = '" & txt_Name.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
uv_dg.Refresh
End If


ElseIf op_uid.Value = True Then
If Not (txt_Id.Text = "") Then
Adodc1.RecordSource = "select * from Users where ID =" & txt_Id.Text & " ORDER BY ID ASC"
Adodc1.Refresh
uv_dg.Refresh
End If

ElseIf op_All.Value = True Then
Adodc1.RecordSource = "select * from Users ORDER BY ID ASC"
Adodc1.Refresh
uv_dg.Refresh
End If
'width changing code
With uv_dg
    .Columns(0).Width = 1500
    .Columns(1).Width = 3700
    .Columns(2).Width = 2045
    .Columns(3).Width = 1350
End With
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select distinct ID,EMAIL,NAME,AGE from Users"
Adodc1.Refresh
With Adodc1.Recordset
    Do Until .EOF
    '    Combo1.AddItem ![isbn]
     '   Combo2.AddItem ![author]
      '  Combo3.AddItem ![Title]
        .MoveNext
    Loop
    End With
Adodc1.RecordSource = "Users"
op_All.Value = True
'width changing code
With uv_dg
    .Columns(0).Width = 1500
    .Columns(1).Width = 3700
    .Columns(2).Width = 2045
    .Columns(3).Width = 1350
End With
End Sub

Private Sub op_All_Click()
If op_All.Value = True Then
txt_Id.Visible = False
txt_Name.Visible = False
Adodc1.RecordSource = "select * from USERS ORDER BY ID ASC"
Adodc1.Refresh
uv_dg.Refresh
End If
'width changing code
With uv_dg
    .Columns(0).Width = 1500
    .Columns(1).Width = 3700
    .Columns(2).Width = 2045
    .Columns(3).Width = 1350
End With
End Sub

Private Sub op_Name_Click()
If op_name.Value = True Then
txt_Name.Visible = True
txt_Id.Visible = False
End If


End Sub

Private Sub op_Uid_Click()
If op_uid.Value = True Then
txt_Id.Visible = True
txt_Name.Visible = False
End If
End Sub
