<html>
<body>
	<% 
	Dim strName,intAge
	strName="卓云"                     
	intAge=22
	Session("strName")=strName							'给Session变量赋值
	Session("intAge")=intAge							'给Session变量赋值
	Response.Write "该程序仅用来存入Session值，请自己打开5-2.asp查看结果。"   
	%>
</body> 
</html>

