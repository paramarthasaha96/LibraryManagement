VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form View_Books 
   BackColor       =   &H00FAF2BA&
   Caption         =   "Books"
   ClientHeight    =   5985
   ClientLeft      =   2805
   ClientTop       =   1935
   ClientWidth     =   10710
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
   Picture         =   "View_Books.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   10710
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3120
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Frame fr_bv 
      BackColor       =   &H0080C0FF&
      Caption         =   "List of Books"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9975
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
         Height          =   495
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1455
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
         Left            =   8280
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt_Aut 
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
         Left            =   6840
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_Qty 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   3600
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton op_Aut 
         BackColor       =   &H0080C0FF&
         Caption         =   "AUTHOR"
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
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton op_Qty 
         BackColor       =   &H0080C0FF&
         Caption         =   "QUANTITY"
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
         Left            =   5040
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton op_Name 
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
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   975
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
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid bv_dg 
         Bindings        =   "View_Books.frx":46114
         Height          =   3375
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lb_view 
         BackColor       =   &H0080C0FF&
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
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "View_Books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub co_S_Click()
If op_name.Value = True Then
Adodc1.RecordSource = "select * from BOOKS where Name = '" & txt_Name.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh

ElseIf op_Bid.Value = True Then
If Not (txt_Bid.Text = "") Then
Adodc1.RecordSource = "select * from BOOKS where ID =" & txt_Bid.Text & " ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh
End If

ElseIf op_Aut.Value = True Then
If Not (txt_Aut.Text = "") Then
Adodc1.RecordSource = "select * from BOOKS where AUTHOR = '" & txt_Aut.Text & "' ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh
End If

ElseIf op_Qty.Value = True Then
If Not (txt_Qty.Text = "") Then
Adodc1.RecordSource = "select * from BOOKS where QUANTITY =" & txt_Qty.Text & " ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh
End If

ElseIf op_All.Value = True Then
Adodc1.RecordSource = "select * from BOOKS ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh
End If
'width changing code
With bv_dg
    .Columns(0).Width = 1000
    .Columns(1).Width = 3050
    .Columns(2).Width = 2400
    .Columns(3).Width = 1200
    .Columns(4).Width = 1095
End With
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=LIBDATABASE.mdb;"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select distinct ID,NAME,AUTHOR,CATEGORY,QUANTITY from Books"
Adodc1.Refresh
With Adodc1.Recordset
    Do Until .EOF
    '    Combo1.AddItem ![isbn]
     '   Combo2.AddItem ![author]
      '  Combo3.AddItem ![Title]
        .MoveNext
    Loop
    End With

Adodc1.RecordSource = "Books"
op_All.Value = True
'width changing code
With bv_dg
    .Columns(0).Width = 1000
    .Columns(1).Width = 3050
    .Columns(2).Width = 2400
    .Columns(3).Width = 1200
    .Columns(4).Width = 1095
End With
End Sub

Private Sub op_All_Click()
If op_All.Value = True Then
txt_Name.Visible = False
txt_Bid.Visible = False
txt_Qty.Visible = False
txt_Aut.Visible = False
Adodc1.RecordSource = "select * from BOOKS ORDER BY ID ASC"
Adodc1.Refresh
bv_dg.Refresh
'width changing code
With bv_dg
    .Columns(0).Width = 1000
    .Columns(1).Width = 3050
    .Columns(2).Width = 2400
    .Columns(3).Width = 1200
    .Columns(4).Width = 1095
End With
End If
End Sub

Private Sub op_Aut_Click()
If op_Aut.Value = True Then
txt_Name.Visible = False
txt_Bid.Visible = False
txt_Qty.Visible = False
txt_Aut.Visible = True
End If
End Sub

Private Sub op_Bid_Click()
If op_Bid.Value = True Then
txt_Bid.Visible = True
txt_Name.Visible = False
txt_Qty.Visible = False
txt_Aut.Visible = False
End If
End Sub

Private Sub op_Name_Click()
If op_name.Value = True Then
txt_Name.Visible = True
txt_Bid.Visible = False
txt_Qty.Visible = False
txt_Aut.Visible = False
End If
End Sub

Private Sub op_Qty_Click()
If op_Qty.Value = True Then
txt_Name.Visible = False
txt_Bid.Visible = False
txt_Qty.Visible = True
txt_Aut.Visible = False
End If
End Sub
