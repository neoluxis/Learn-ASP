<%
'================================================================================================
'删除留言文件
'1. 如果用户输入正确的删除密码，就可以将留言删除。
'2. 注意，这里首先将获取过来的记录ID编号存放在一个隐藏文本框中，然后提交表单后就可以获取该ID， 从而删除该记录。
'================================================================================================
%>

<% Option Explicit						'强制声明变量%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>删除留言</title>
	<link rel="stylesheet" href="guest.css">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body >
	<h2 align="center">删除留言</h2>
	<!-- 注意其中将传递过来的ID存放到隐藏文本框中了  -->
	<form name="frmDelete" method="POST" action="">
		<center>请输入删除密码：
		<input type="password" name="txtPwd" size="10" value="">
		<input type="hidden" name="txtID" value="<%=Request.QueryString("ID")%>">
		<input type="submit" value="　提交　" size="20">
		</center>
	</form>
	<%
	'这里判断一下，如果密码和配置文件中的密码相等，则删除该留言
	If Request.Form("txtPwd")=conPwd Then
		Dim strSql
		strSql="Delete From tbGuest where ID=" & Request.Form("txtID") 
		conn.Execute(strSql)
		Response.Redirect("index.asp")
	End If
	%>
</body>
</html>