VERSION 5.00
Begin VB.UserControl CLabel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6916
   ScaleHeight     =   4420
   ScaleWidth      =   6916
   Begin VB.Timer Timer1 
      Left            =   117
      Top             =   1287
   End
   Begin VB.Label lblContext 
      Caption         =   "빤짝이 레이블"
      Height          =   1183
      Left            =   234
      TabIndex        =   0
      Top             =   234
      Width           =   2821
   End
End
Attribute VB_Name = "CLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************
' RunLabel.Ctl
'**************************

Option Explicit

Event goExecute()

'Win Api Shell Execute  함수
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParamenters As String, ByVal lpdirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1

Private m_URL As String
Private t_interval As Long
Private IsURL As Boolean
Private ExecuteDir As String

Private Sub lblContext_Click()
    Dim RunDmc As String
    
    RaiseEvent goExecute
    
    If URL = False Then
        RunDmc = ExecuteDir + "\" + lblContext.Caption
    Else
        RunDmc = lblContext.Caption
    End If
    
    ShellExecute hwnd, "open", RunDmc, vbNull, vbNull, SW_NORMAL
    
End Sub

Private Sub Timer1_Timer()
    Randomize
    lblContext.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub UserControl_Initialize()
    lblContext.Width = UserControl.ScaleWidth
    lblContext.Height = UserControl.ScaleHeight
    
    m_URL = "https://www.naver.com"
    IsURL = False
    t_interval = 1000
    ExecuteDir = App.Path
End Sub

Private Sub UserControl_InitProperties()
    goFile = m_URL
    URL = IsURL
    blink = t_interval
    Directory = ExecuteDir
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    goFile = PropBag.ReadProperty("goFile", m_URL)
    URL = PropBag.ReadProperty("URL", IsURL)
    blink = PropBag.ReadProperty("blink", t_interval)
    Directory = PropBag.ReadProperty("Directory", ExecuteDir)
End Sub

Private Sub UserControl_Resize()
    lblContext.Width = UserControl.ScaleWidth
    lblContext.Height = UserControl.ScaleHeight
    
End Sub

'******** Property

Public Property Get goFile() As String
    goFile = lblContext.Caption
End Property

Public Property Let goFile(ByVal LgoFile As String)
    lblContext = LgoFile
    PropertyChanged "goFile"
End Property

Public Property Get URL() As Boolean
    URL = IsURL
End Property

Public Property Let URL(ByVal LURL As Boolean)
    IsURL = LURL
    
    ' URL이면 화면이 번쩍거림
    If IsURL = True Then
        Timer1.Enabled = True
        lblContext.ForeColor = &H80000012
    Else
        Timer1.Enabled = False
    End If
    
    PropertyChanged "URL"
End Property

Public Property Get blink() As Long
    blink = Timer1.Interval
End Property

Public Property Let blink(ByVal LBlink As Long)
    Timer1.Interval = LBlink
    PropertyChanged "blink"
End Property

Public Property Get Directory() As String
    Directory = ExecuteDir
End Property

Public Property Let Directory(ByVal LDirectory As String)
    ExecuteDir = LDirectory
    PropertyChanged ("Directory")
End Property

'******** Property End
