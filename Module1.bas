Attribute VB_Name = "Module1"
Option Explicit
Public Player1
Public ID, Money
Public Lv1 As Integer
Public Lv2 As Integer

Dim Cn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Public MyRs As New ADODB.Recordset

Public Sub CnSql(Sql As String, TP As Long)
On Error Resume Next
    Set Cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\zc.mdb;Jet OLEDB:Database password=123"
    Cn.Open
    If TP = 1 Then
        Rs.Open Sql, Cn, 1, 1
        Set MyRs = Rs
        Set Rs = Nothing
    End If
    If TP = 2 Then
        Cn.Execute Sql
    End If
    Set Cn = Nothing
End Sub

