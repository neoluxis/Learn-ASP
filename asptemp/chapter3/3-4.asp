<% Option Explicit								'放在程序首行，强制变量声明  %>
<html>
<body>
	<%
	Dim strA
	strA = Split("VBScript*is*good!", "*")		'以*为分界符拆分
	Response.Write strA(0) & "<P>"
	Response.Write strA(1) & "<P>"
	Response.Write strA(2)
	%>
</body> 
</html> 
