Attribute VB_Name = "Module1"
Option Explicit

Private Function Selectsql(SQL As String) As ADODB.Recordset '����ADODB.Recordset����
Dim ConnStr As String
Dim Conn As ADODB.Connection
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set Conn = New ADODB.Connection

'On Error GoTo MyErr:
ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Password=1123581321;Initial Catalog=test;Data Source=localhost" '��������SQL���ݿ�����
Conn.Open ConnStr
rs.CursorLocation = adUseClient
rs.Open Trim$(SQL), Conn, adOpenDynamic, adLockOptimistic
End Function

