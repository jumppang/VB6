VERSION 5.00
Object = "{639EB20C-C582-4555-A2F2-0D03933FBB5E}#1.0#0"; "ActiveXControlTest.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5434
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   4355
   ScaleWidth      =   5434
   Begin ActiveXControlTest.CLabel CLabel1 
      Height          =   1183
      Left            =   585
      TabIndex        =   0
      Top             =   1170
      Width           =   2002
      _ExtentX        =   3522
      _ExtentY        =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command1_Click()
      
    Module1.monster_eat
  
    'On Error GoTo aaa_label
    ''On Error Resume Next
    
    'Err.Clear
    'Err.Raise 6
    
    'If Err.Number <> 0 Then
    
        'MsgBox gCC
                
    'End If
    
    
    'Exit Sub
    
'aaa_label:
    
    'MsgBox "aaa"

End Sub

Private Sub CLabel1_goExecute()
    CLabel1.URL = True
End Sub

Private Sub Form_Load()
    Call Module1.SetMoster
End Sub

