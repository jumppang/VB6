VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Const conString As String = "Provider=OraOLEDB.Oracle.1;Password=360512;Persist Security Info=True;User ID=ot;Data Source=X44_XE"
'Provider=OraOLEDB.Oracle;dbq=localhost:1521/XE;Database=myDataBase;User Id=myUsername;Password=myPassword;
Public Conn As ADODB.Connection
Public Rs As ADODB.Recordset

Public Function OraConnOpen() As Boolean

    OraConnOpen = False
    
    On Error GoTo ConnectError
    
    Set Conn = New ADODB.Connection
    
    Conn.Open "Provider=OraOLEDB.Oracle;dbq=J-X44:1521;Database=XE;User Id=ot;Password=360512;"
    
    On Error GoTo 0
    
    OraConnOpen = True
    
    Exit Function
    
ConnectError:
    OraConnOpen = False
    MsgBox Err.Description
End Function

Public Function RunSql(nSql As String) As Boolean
    
    On Error GoTo ErrSql
    Set Rs = Conn.Execute(nSql)
    
    On Error GoTo 0
    
    RunSql = True
    Exit Function
    
ErrSql:
    RunSql = False
        
End Function

Public Function RunSql1(nSql As String, nRs As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrSql
    nRs.Open nSql, Conn, adOpenDynamic, , adCmdText
    
    On Error GoTo 0
    
    RunSql1 = True
    
    Exit Function
    
ErrSql:
    RunSql1 = False
    
End Function

Public Property Get LPCONSTRING() As String
    LPCONSTRING = conString
End Property



