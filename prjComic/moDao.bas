Attribute VB_Name = "moDao"
Option Explicit


Public Const CONSTRING As String = "Provider=OraOLEDB.Oracle.1;Password=360512;Persist Security Info=True;User ID=ot;Data Source=X44_XE"
'Provider=OraOLEDB.Oracle;dbq=localhost:1521/XE;Database=myDataBase;User Id=myUsername;Password=myPassword;
Public Conn As ADODB.Connection
Public Rs As ADODB.Recordset

Public Function OraConnOpen() As Boolean

    OraConnOpen = False
    
    On Error GoTo ConnectError
    
    Set Conn = New ADODB.Connection
    
    Conn.Open CONSTRING
    
    On Error GoTo 0
    
    OraConnOpen = True
    
    Exit Function
    
ConnectError:
    OraConnOpen = False
    MsgBox Err.Description
End Function

Public Function SqlTran(nSql As String) As Boolean
    
    On Error GoTo ErrSql
    
    Set Rs = Conn.Execute(nSql)
    
    On Error GoTo 0
    
    SqlTran = True
    Exit Function
    
ErrSql:
    SqlTran = False
        
End Function

Public Function SqlSelect(nSql As String, nRs As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrSql
    nRs.Open nSql, Conn, 1, , 1
    
    On Error GoTo 0
    
    SqlSelect = True
    
    Exit Function
    
ErrSql:
    SqlSelect = False
    
End Function






