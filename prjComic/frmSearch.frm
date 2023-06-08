VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Form1"
   ClientHeight    =   10320
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoSearch 
      Height          =   392
      Left            =   2772
      Top             =   9576
      Width           =   5558
      _ExtentX        =   9790
      _ExtentY        =   714
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
   Begin VB.TextBox txtPublisher 
      Height          =   518
      Left            =   3528
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   8694
      Width           =   5306
   End
   Begin VB.TextBox txtWriter 
      Height          =   518
      Left            =   3654
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6426
      Width           =   5180
   End
   Begin VB.TextBox txtTitle 
      Height          =   518
      Left            =   3654
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5418
      Width           =   5180
   End
   Begin VB.OptionButton optPublisher 
      Caption         =   "출판사"
      Height          =   392
      Left            =   2646
      TabIndex        =   6
      Top             =   8820
      Width           =   896
   End
   Begin VB.OptionButton optIBSN 
      Caption         =   "ISBN"
      Height          =   392
      Left            =   2646
      TabIndex        =   5
      Top             =   7644
      Width           =   1526
   End
   Begin VB.OptionButton optWriter 
      Caption         =   "작가"
      Height          =   266
      Left            =   2646
      TabIndex        =   4
      Top             =   6594
      Width           =   1022
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "제목"
      Height          =   266
      Left            =   2646
      TabIndex        =   3
      Top             =   5544
      Width           =   896
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   518
      Left            =   504
      TabIndex        =   2
      Top             =   8694
      Width           =   1652
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "검색"
      Height          =   518
      Left            =   504
      TabIndex        =   1
      Top             =   5418
      Width           =   1652
   End
   Begin MSDataGridLib.DataGrid dgGrid 
      Height          =   4550
      Left            =   378
      TabIndex        =   0
      Top             =   252
      Width           =   8834
      _ExtentX        =   15584
      _ExtentY        =   8017
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
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dao As CDao


Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSearch_Click()
    Dim def As String
    Dim query As String
    
    def = "Select * From EMPLOYEES "
    
    If Not dao.OraConnOpen Then
        MsgBox "Con Error"
        
        Exit Sub
    End If
    
    
    adoSearch.CommandType = adCmdUnknown
    
    If optTitle.Value = True Then
        
        If txtTitle.Text = "" Then
            query = " Order By FIRST_NAME"
        Else
            query = " Where FIRST_NAME like '%" + txtTitle.Text + "%'"
        End If
        
    End If
    
    adoSearch.RecordSource = def + query
    adoSearch.Refresh
    
End Sub

Private Sub Form_Initialize()

    Set dao = New CDao
    Me.adoSearch.ConnectionString = dao.LPCONSTRING
    'Me.adoSearch.Visible = False
       
End Sub


