VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form View_Br 
   BackColor       =   &H00FAF2BA&
   Caption         =   "Users who Borrowed"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
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
   Picture         =   "Fines.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
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
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton op_Id 
      BackColor       =   &H0080C0FF&
      Caption         =   "ISSUE ID"
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
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txt_Bid 
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
      Left            =   6600
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_B 
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_Uid 
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_N 
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
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton co_S 
      BackColor       =   &H0000C000&
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
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton op_All 
      BackColor       =   &H0080C0FF&
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
      Left            =   8760
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton op_Bid 
      BackColor       =   &H0080C0FF&
      Caption         =   "BOOK ID"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton op_B 
      BackColor       =   &H0080C0FF&
      Caption         =   "BOOK"
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
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton op_Uid 
      BackColor       =   &H0080C0FF&
      Caption         =   "USER ID"
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
      TabIndex        =   3
      Top             =   780
      Width           =   1215
   End
   Begin VB.OptionButton op_N 
      BackColor       =   &H0080C0FF&
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
      Left            =   2160
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   5760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin VB.Frame fr_fin 
      BackColor       =   &H0080C0FF&
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
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   10695
      Begin MSDataGridLib.DataGrid dg_brw 
         Bindings        =   "Fines.frx":406B8
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5318
         _Version        =   393216
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Britannic Bold"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "View_Br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub co_S_Click()
If op_N.Value = True Then
Adodc1.RecordSource = "select * from Issue where NAME = '" & txt_N.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
ElseIf op_B.Value = True Then
Adodc1.RecordSource = "select * from Issue where BOOK = '" & txt_B.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
ElseIf op_uid.Value = True Then
Adodc1.RecordSource = "select * from Issue where UID = '" & txt_Uid.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
ElseIf op_Bid.Value = True Then
Adodc1.RecordSource = "select * from Issue where BID = '" & txt_Bid.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
ElseIf op_Id.Value = True Then
If Not (txt_Id.Text = "") Then
Adodc1.RecordSource = "select * from Issue where ID = " & txt_Id.Text & " ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
End If
ElseIf op_All.Value = True Then
Adodc1.RecordSource = "select * from Issue ORDER BY ID ASC"
Adodc1.Refresh
dg_brw.Refresh
End If
'width changing code
With dg_brw
    .Columns(0).Width = 1150
    .Columns(1).Width = 3550
    .Columns(2).Width = 1200
    .Columns(3).Width = 2900
    .Columns(4).Width = 1265
End With
End Sub

Private Sub Form_Load()
op_All.Value = True
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select distinct ID,NAME,UID,BOOK,BID from Issue"
Adodc1.Refresh
With Adodc1.Recordset
    'Do Until .EOF
    '    Combo1.AddItem ![isbn]
     '   Combo2.AddItem ![author]
      '  Combo3.AddItem ![Title]
       ' .MoveNext
    'Loop
    End With
Adodc1.RecordSource = "Issue"
'width changing code
With dg_brw
    .Columns(0).Width = 1150
    .Columns(1).Width = 3550
    .Columns(2).Width = 1200
    .Columns(3).Width = 2900
    .Columns(4).Width = 1235
End With
End Sub

Private Sub op_All_Click()
If op_All.Value = True Then
txt_N.Visible = False
txt_Bid.Visible = False
txt_Uid.Visible = False
txt_B.Visible = False
txt_Id.Visible = False
Adodc1.RecordSource = "select * from Issue ORDER BY ID ASC"
'Adodc1.Refresh
dg_brw.Refresh
End If
End Sub

Private Sub op_B_Click()
If op_B.Value = True Then
txt_N.Visible = False
txt_Bid.Visible = False
txt_Uid.Visible = False
txt_B.Visible = True
txt_Id.Visible = False
End If
End Sub

Private Sub op_Bid_Click()
If op_Bid.Value = True Then
txt_N.Visible = False
txt_Bid.Visible = True
txt_Uid.Visible = False
txt_B.Visible = False
txt_Id.Visible = False
End If
End Sub

Private Sub op_Id_Click()
If op_Id.Value = True Then
txt_N.Visible = False
txt_Bid.Visible = False
txt_Uid.Visible = False
txt_B.Visible = False
txt_Id.Visible = True
End If
End Sub

Private Sub op_N_Click()
If op_N.Value = True Then
txt_N.Visible = True
txt_Bid.Visible = False
txt_Uid.Visible = False
txt_B.Visible = False
txt_Id.Visible = False
End If
End Sub

Private Sub op_Uid_Click()
If op_uid.Value = True Then
txt_N.Visible = False
txt_Bid.Visible = False
txt_Uid.Visible = True
txt_B.Visible = False
txt_Id.Visible = False
End If
End Sub
