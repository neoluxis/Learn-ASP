<html>
<head>
	<title>求a和b的平方和</title>
</head>
<body>
	<%
	Dim a,b,c										'声明三个变量
	a=3												'给变量a赋值
	b=4												'给变量b赋值
	c=a^2+b^2										'其中^表示求平方
	Response.Write "3和4的平方和为" & c				'在页面上输出结果
	%>    
</body>
</html>
