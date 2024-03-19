<html>
<body>
	欢迎光临我的主页
	<%
	Server.ScriptTimeOut=100							'设置脚本最长执行时间
	Response.Write "<p>6-4.asp中ScriptTimeOut=" & Server.ScriptTimeOut	
	Server.Execute "6-5.asp"
	%>
	<p>谢谢，再见
</body> 
</html>
