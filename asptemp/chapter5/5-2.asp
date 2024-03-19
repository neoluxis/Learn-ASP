<html>
<body>
	<%
	Dim strName,intAge
	strName=Session("strName")						'读取Session变量
	intAge=Session("intAge")							'读取Session变量
	Response.Write strName & "您好，欢迎您<br>"
	Response.Write "您的年龄是" & intAge
	%>
</body> 
</html>

