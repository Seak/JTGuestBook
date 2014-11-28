<!--#Include File="conn.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除留言</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="images/ALL_BG.gif" oncontextmenu="return false" onselectstart="return false">
<%
Dim ID, Rs
If Request.Form("Password") = Password Then

	ID = Request("id")

	ConnectionDatabase

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "Select ID From [guestbook] Where ID = "& ID ,conn,1,3
	Rs.Delete
	Rs.Close
	Set Rs = Nothing

	CloseDatabase
	Response.Write "删除成功！"
	Response.End
End If
%>
<form name="Delete" method="post" action="del.asp?id=<%= Request("id") %>">
  <p>密码： 
    <input name="Password" type="password" size="15">

    <input name="submit" type="submit" value="确定">
    </p>
</form>
</body>
</html>