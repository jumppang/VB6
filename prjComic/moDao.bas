Attribute VB_Name = "moDao"
Option Explicit


Const conString As String = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCO=TCP)(HOST=localhost)(PORT = 521))(CONNECT_DATA=(SERVER=DEDICATED)(SID=xe)));User ID=jusu;Password=360512"

Public Conn As ADODB.Connection
Public Rs As ADODB.Recordset

Public Function OraConnOpen() As Boolean

    OraConnOpen = False
    
    On Error GoTo ConnectError
    
    Set Conn = New ADODB.Connection
    
    Conn.Open conString
    
    On Error GoTo 0
    
    OraConnOpen = True
    
    Exit Function
    
ConnectError:
    OraConnOpen = False

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

Public Function RunSql(nSql As String, nRs As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrSql
    nRs.Open nSql, Conn, adOpenDynamic, , adCmdText
    
    On Error GoTo 0
    
    RunSql = True
    
    Exit Function
    
ErrSql:
    RunSql = False
    
End Function
