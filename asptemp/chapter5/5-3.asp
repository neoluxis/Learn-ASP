<html>
<body>
	<% 
	Dim strNames(1)							'声明一个长度为2的数组
	strNames(0)="白芸"
	strNames(1)="海霞"
	Session("strNames")=strNames			'将数组赋值给Session变量
	Response.Write "该程序仅用来存入Session数组，请自己打开5-4.asp查看结果。"   
	%>
</body> 
</html> 
