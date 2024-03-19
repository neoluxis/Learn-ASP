<%
'================================================================================================
'聊天内容页面
'1. 该页只是从Application中读取聊天信息，并显示在页面中。
'2. 其中调用配置文件中的常数，从而若干秒会自动刷新一次，以显示最新聊天信息。默认5秒刷新一次。
'3. 由于聊天信息是从上往下滚动的，最后的JavaScript语句就可以将窗口滚动到最下边，从而显示最新的聊天信息。
'================================================================================================
%>

<%option explicit%>
<!--#Include File="config.asp"-->
<html>
<head>
	<meta http-equiv="refresh" content="<%=conRefresh%>">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="chat.css">
</head>
<body bgcolor="#EEEEEE" bottomMargin="50">
	<%
	'下面从Application中获取聊天信息，并输出在页面上
	Response.Write Application("strChat")
	%>
	<!--注释：因为从上到下显示，所以要利用JavaScript滚动到最下边-->
	<script language="JavaScript" for="window" event="onload">window.scroll(0,60000);</script>
</body>
</html>