<!--#Include File="conn.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ظ�����</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="images/ALL_BG.gif" oncontextmenu="return false" onselectstart="return false">
<%
Dim ID, Rs, Reply
If Request.Form("Password") = Password Then

	ID = Request("id")
	Reply = Request("Reply")

	ConnectionDatabase

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "Select ID, Reply From [guestbook] Where ID = "& ID ,conn,1,3
	Rs("Reply") = Reply
	Rs.Update
	Rs.Close
	Set Rs = Nothing

	CloseDatabase

	Response.Write "�ظ��ɹ���"
	Response.End
End If
%>
<form name="Reply" method="post" action="reply.asp?id=<%= Request("id") %>">
  <p>���룺 
    <input name="Password" type="password" size="15">
    <br>
    �ظ��� 
    <textarea name="Reply" cols="34" rows="6"></textarea>
    <br>
    <input name="Submit" type="submit" value="ȷ��">
    <input type="reset" name="Reset" value="����">
  </p>
</form>
</body>
</html>