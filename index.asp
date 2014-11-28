<!--#include file="conn.asp" -->
<!--#include file="config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>访客留言 - <%= Title %></title>
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="images/ALL_BG.gif" oncontextmenu="return false" onselectstart="return false">
<%
Dim RecordPerPage, Rs, SQL, absPageNum, TotalPages, absRecordNum
RecordPerPage = 10
ConnectionDatabase
SQL = "Select ID, DateTime, Name, IP, Mail, Home, Content, Topic, Reply From [guestbook] Order By ID DESC"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3
Rs.CacheSize = RecordPerPage
Rs.PageSize = RecordPerPage
Rs.Open SQL, Conn, 3, 1, &H0001
absPageNum = CInt(Request.QueryString("page"))
TotalPages = Rs.PageCount
If Request.QueryString("page") = "" Or absPageNum > TotalPages Then absPageNum = 1
If TotalPages < absPageNum Then TotalPages = absPageNum
%>
<table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="9" align="center" valign="middle"><img src="images/navbar_L.gif" width="9" height="27"></td>
    <td align="center" valign="middle" background="images/navbar.gif">[<a href="<%= URL %>">返回首页</a>]　第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= Rs.RecordCount %>条</font>　 
      <% If Not absPageNum = 1 Then %>[<a href="?page=1">首页</a>][<a href="?page=<%= absPageNum - 1 %>">上一页</a>]<%
End If
If Not absPageNum = TotalPages Then
%>[<a href="?page=<%= absPageNum + 1 %>">下一页</a>][<a href="?page=<%= TotalPages %>">尾页</a>]<% End If %>
    </td>
    <td width="9" align="center" valign="middle"><img src="images/navbar_r.gif" width="9" height="27"></td>
  </tr>
</table>
<br>
<table width="510" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="5" align="center" valign="middle" background="images/mframe_t.gif"><img src="images/lframe_t_l.gif" width="5" height="23"></td>
    <td width="460" align="left" valign="middle" background="images/mframe_t.gif"><strong>访客留言</strong></td>
    <td width="45" align="center" valign="middle" background="images/mframe_t.gif"><img src="images/lframe_t_r.gif" width="45" height="23"></td>
  </tr>
  <tr> 
    <td colspan="3"><table width="510" border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER-LEFT: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid">
        <tr> 
          <td height="5" colspan="3" align="left" valign="middle"></td>
        </tr>
        <tr align="center"> 
          <form name="guestbook" method="post" action="save.asp">
            <td width="200" height="120" valign="middle"> 姓名： 
              <input name="name" type="text" size="20"> <br>
              邮箱： 
              <input name="mail" type="text" size="20"> <br>
              主页： 
              <input name="home" type="text" size="20"> <br>
              主题： 
              <input name="topic" type="text" size="20"> </td>
            <td width="10" height="120" valign="middle">&nbsp; </td>
            <td width="300" height="120" valign="middle"> <textarea name="content" cols="40" rows="5"></textarea> 
              <br> <input type="submit" name="Submit" value="提交"> <input type="reset" name="Pino" value="重置"> 
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
</table>
<br>
  <hr width="80%" size="1">
<br>
<%
If Not(Rs.EOF) Then
	Rs.AbsolutePage = absPageNum
	For absRecordNum = 1 To Rs.PageSize
%>
<table width="510" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="5" align="center" valign="middle" background="images/mframe_t.gif"><img src="images/mframe_t_l.gif" width="5" height="23"></td>
    <td width="115" align="left" valign="middle" background="images/mframe_t.gif"><strong> 
      <%= Rs("Name") %>
      </strong></td>
    <td width="170" align="left" valign="middle" background="images/mframe_t.gif"> 
      <strong>于</strong> <%= Rs("DateTime") %> <strong>说</strong></td>
    <td width="140" align="right" valign="middle" background="images/mframe_t.gif"><a href="reply.asp?id=<%= Rs("ID") %>">Reply</a> 
      <a href="del.asp?id=<%= Rs("ID") %>">Delete</a> <a href="<%= Rs("Home") %>">Home</a> <a href="mailto:<%= Rs("Mail") %>">Mail</a></td>
    <td width="80" align="center" valign="middle" background="images/mframe_t.gif"><img src="images/mframe_t_r.gif" width="80" height="23"></td>
  </tr>
  <tr> 
    <td colspan="5"><table width="510" border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER-LEFT: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid">
        <tr> 
          <td width="69" height="5" align="left" valign="middle"></td>
        </tr>
        <tr align="center"> 
            
          <td align="left" valign="top"><strong><%= Rs("Topic") %></strong><br>
            <%= Replace(Rs("Content"), chr(13),"<br>") %></td>
        </tr>
		<% If Rs("Reply") <> "" Then %>
        <tr align="center">
          <td align="left" valign="top">
<hr width="100%" size="1">
            <strong>斑竹回复：</strong><br>
            <%= Replace(Rs("Reply"), chr(13),"<br>") %></td>
        </tr>
		<% End If %>
      </table></td>
  </tr>
</table>
<br>
<%
		Rs.MoveNext
		If Rs.EOF Then
			Exit For
		End If
	Next
End If
%>
  <hr width="80%" size="1">
<br>
<table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="9" align="center" valign="middle"><img src="images/navbar_L.gif" width="9" height="27"></td>
    <td align="center" valign="middle" background="images/navbar.gif">[<a href="<%= URL %>">返回首页</a>]　第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= Rs.RecordCount %>条</font>　 
      <% If Not absPageNum = 1 Then %>[<a href="?page=1">首页</a>][<a href="?page=<%= absPageNum - 1 %>">上一页</a>]<%
End If
If Not absPageNum = TotalPages Then
%>[<a href="?page=<%= absPageNum + 1 %>">下一页</a>][<a href="?page=<%= TotalPages %>">尾页</a>]<% End If %>
    </td>
    <td width="9" align="center" valign="middle"><img src="images/navbar_r.gif" width="9" height="27"></td>
  </tr>
</table>
<br>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle">Powered by : <a href="http://www.zzut.net/products/JTGuestBook.htm">JTGuestBook</a> <a href="http://www.zzut.net/download/JTGuestBook100.rar">v1.0.0</a><br>Copyright &copy; 2005 ZZUT.NET All rights 
      reserved<br>
      E-MAIL: <a href="mailto:Webmaster@zzut.net"> Webmaster@zzut.net</a> QQ：11265943（江海客）</td>
  </tr>
</table>
<%
Rs.Close
Set Rs = Nothing
CloseDatabase
%>
</body>
</html>