<%
'================================================================================================
'聊天室主界面
'1. 这只是一个简单的框架页，分别显示不同的框架而已。
'================================================================================================
%>

<html>
<head>
	<title>聊天室主界面</title>
</head>
<frameset framespacing="0" border="0" rows="85%,*" frameborder="0">
	<frameset cols="*,14%" border="0" frameborder="0" framespacing="0">
		<frame name="main" src="f1.asp" marginwidth="0" marginheight="0" scrolling="auto" >
		<frame name="online" src="f3.asp" scrolling="auto">
	</frameset>
	<frame name="input" src="f2.asp" marginwidth="0" marginheight="0" scrolling="auto" >
</frameset>
</html>
