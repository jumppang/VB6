VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SP With In and Out"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run SP"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2415
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dbConn As ADODB.Connection
Dim Cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim strWName As String
Dim intCount As Long

On Error GoTo ExitME
Err.Clear
Set dbConn = New ADODB.Connection
With dbConn
    .Provider = "OraOLEDB.Oracle"
    .Properties("Data Source") = "ROWELLS"
    .Properties("User Id") = "wellsadmin"
    .Properties("Password") = "roweisgood"
    .Open
End With

Set Cmd = New ADODB.Command
Set Cmd.ActiveConnection = dbConn
With Cmd
    .Parameters.Append .CreateParameter(, adVarChar, adParamOutput, 50)
    .Parameters.Append .CreateParameter(, adNumeric, adParamOutput)
End With

If dbConn.State Then
    Cmd.Properties("PLSQLRSet") = True
    Cmd.CommandType = adCmdText
    Cmd.CommandText = "{CALL WellsAdmin.WellCounting(?, ?)}"
    Set rs = Cmd.Execute()
End If
If Not rs Is Nothing Then
    Dim i As Long
    i = 0
    
    If Not rs.BOF And Not rs.EOF Then
        Do Until rs.EOF
            i = i + 1
            If i > 1 Then
                Me.MSHFlexGrid1.Rows = Me.MSHFlexGrid1.Rows + 1
            End If
            Me.MSHFlexGrid1.TextMatrix(i, 0) = rs.Fields(0).Value
            Me.MSHFlexGrid1.TextMatrix(i, 1) = rs.Fields(1).Value
            rs.MoveNext
        Loop
    End If
End If
ExitME:
If Not Cmd Is Nothing Then
    Set Cmd = Nothing
End If
If Not dbConn Is Nothing Then
    If dbConn.State Then dbConn.Close
    Set dbConn = Nothing
End If

End Sub

Private Sub Command2_Click()
Dim dbConn As ADODB.Connection
Dim Cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim lngWellID As Long

lngWellID = InputBox("WellID Number for results:", "Well ID")


On Error GoTo ExitME
Err.Clear
Set dbConn = New ADODB.Connection
With dbConn
    .Provider = "OraOLEDB.Oracle"
    .Properties("Data Source") = "ROWELLS"
    .Properties("User Id") = "wellsadmin"
    .Properties("Password") = "roweisgood"
    .Open
End With

Set Cmd = New ADODB.Command
Set Cmd.ActiveConnection = dbConn
With Cmd
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , lngWellID)
    .Parameters.Append .CreateParameter(, adVarChar, adParamOutput, 50)
    .Parameters.Append .CreateParameter(, adNumeric, adParamOutput)
End With

If dbConn.State Then
    Cmd.Properties("PLSQLRSet") = True
    Cmd.CommandType = adCmdText
    Cmd.CommandText = "{CALL WellsAdmin.OneWellCount(?,?, ?)}"
    Set rs = Cmd.Execute()
End If
If Not rs Is Nothing Then
    
    If Not rs.BOF And Not rs.EOF Then
        Do Until rs.EOF
            MsgBox "Well Name: " & rs.Fields(0).Value & vbCrLf & "Number of Results: " & rs.Fields(1).Value & vbCrLf & "Well ID Entered: " & lngWellID
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End If

ExitME:
If Not Cmd Is Nothing Then
    Set Cmd = Nothing
End If
If Not dbConn Is Nothing Then
    If dbConn.State Then dbConn.Close
    Set dbConn = Nothing
End If


End Sub

Private Sub Form_Load()

Me.MSHFlexGrid1.TextMatrix(0, 0) = "Well Name"
Me.MSHFlexGrid1.TextMatrix(0, 1) = "Number of results"

Me.MSHFlexGrid1.ColWidth(0) = Len(Me.MSHFlexGrid1.TextMatrix(0, 0)) * 100
Me.MSHFlexGrid1.ColWidth(1) = Len(Me.MSHFlexGrid1.TextMatrix(0, 1)) * 100

End Sub


