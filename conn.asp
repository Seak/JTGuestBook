<%
Option Explicit
Response.Buffer = True
Dim Conn

Sub ConnectionDatabase
	Dim ConnStr, Db
	Db = "gb.mdb"
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Db)
	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ������������ִ���"
		Response.End
	End If
End Sub

Sub CloseDatabase
	Conn.Close
	Set Conn = Nothing
End Sub
%>