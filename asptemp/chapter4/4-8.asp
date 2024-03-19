<html>
<body>
	<% 
	Dim IP
	IP=Request.ServerVariables("REMOTE_ADDR") 
	Response.Write "来访者IP地址是：" & IP
	%>
</body>
</html>
