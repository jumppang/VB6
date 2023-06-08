VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "종 료"
      Height          =   560
      Left            =   7434
      TabIndex        =   4
      Top             =   7560
      Width           =   1778
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "검 색"
      Height          =   518
      Left            =   378
      TabIndex        =   3
      Top             =   7434
      Width           =   1526
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3523
      Left            =   468
      TabIndex        =   2
      Top             =   3042
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6218
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
            LCID            =   1042
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
            LCID            =   1042
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
   Begin MSAdodcLib.Adodc adoComic 
      Height          =   481
      Left            =   468
      Top             =   1053
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   847
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
   Begin VB.TextBox txtWriter 
      DataField       =   "FIRST_NAME"
      DataSource      =   "adoComic"
      Height          =   481
      Index           =   1
      Left            =   468
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2223
      Width           =   8764
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "JOB_TITLE"
      DataSource      =   "adoComic"
      Height          =   481
      Index           =   0
      Left            =   468
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1638
      Width           =   8820
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private dao As CDao

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSearch_Click()
     frmSearch.Show
End Sub

Private Sub Form_Initialize()
    'txtWriter.Item.DataSource = "adoComic"
    'txtWriter.Item.DataField = "FIRST_NAME"
    
    Set dao = New CDao
    
End Sub

Private Sub Form_Load()
    Dim def As String
    Dim query As String
    
    Set dao = New CDao
    
    query = "Select * From EMPLOYEES"
    
    adoComic.ConnectionString = dao.LPCONSTRING
    adoComic.RecordSource = query
    
End Sub

