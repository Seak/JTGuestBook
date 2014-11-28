<!--#include file="conn.asp" -->
<%
Dim Rs, SQL

If Request.Form("name") <> "" And Request.Form("mail") <> "" And Request.Form("content") <> "" And Request.Form("topic") <> "" Then

	ConnectionDatabase

	SQL = "Select ID, DateTime, Name, IP, Mail, Home, Content, Topic, Reply From [guestbook] Order By ID DESC"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL, conn, 1,3
	Rs.AddNew 
	Rs("DateTime") = Now()
	Rs("Name") = Server.HTMLEncode(Request.Form("name"))
	Rs("IP") = Request.ServerVariables("Remote_Addr")
	Rs("Mail") = Server.HTMLEncode(Request.Form("mail"))
	Rs("Home") = Server.HTMLEncode(Request.Form("home"))
	Rs("Content") = Server.HTMLEncode(Request.Form("content"))
	Rs("Topic") = Server.HTMLEncode(Request.Form("Topic"))
	Rs("Reply") = Server.HTMLEncode(Request.Form("Reply"))
	Rs.Update 
	Rs.Close
	Set Rs = Nothing

	CloseDatabase

End If

Response.Redirect "index.asp"
%>