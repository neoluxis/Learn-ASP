<html>
<body>
	<%
	Dim strNames										'注意：这里声明一个普通变量
	strNames=Session("strNames")						'读取Session变量
	Response.Write strNames(0) & "您好，欢迎您<br>"	
	Response.Write strNames(1) & "您好，欢迎您<br>"  
	%>
</body> 
</html>

