<%
'================================================================================================
'在线人员页面
'1. 该页只是从Application中读取在线人员名单，它是一个数组，然后用循环列出所有人员。
'2. 其中调用配置文件中的常数，从而若干秒会自动刷新一次，以显示最新在线人员。默认60秒刷新一次。
'3. 由于聊天信息是从上往下滚动的，最后的JavaScript语句就可以将窗口滚动到最下边，从而显示最新的聊天信息。
'================================================================================================
%>

<%option explicit%>
<!--#Include file="Config.asp"-->
<!--#INCLUDE FILE="function.asp"-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<meta http-equiv="refresh" content="<%=conRefreshOnline%>">
	<link rel="stylesheet" href="chat.css">
</head>
<body background="images/bg.gif">
	<p align="center">在线人员
	<br><hr>
	<%
	'从数组中获取在线人员
	Dim strUsers
	strUsers=Application("strUsers")
	'利用循环列出数组中所有人员
	Dim I
	For I=0 To Ubound(strUsers)
		Response.Write strUsers(I) & "<br>" 
	Next
	%>
</body>
</html>