<html>
<body>
	<% 
	Dim strName,intAge                                   
	strName=Request.QueryString("strName")				'获取姓名
	intAge=Request.QueryString("intAge")				'获取年龄
	Response.Write "您的姓名是：" & strName & "，您的年龄是：" & intAge
	%>
</body>
</html>

