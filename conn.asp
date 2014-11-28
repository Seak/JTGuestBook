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
		Response.Write "数据库连接出错，请检查连接字串。"
		Response.End
	End If
End Sub

Sub CloseDatabase
	Conn.Close
	Set Conn = Nothing
End Sub
%>