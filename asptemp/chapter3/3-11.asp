<% Option Explicit						'强制声明变量%> 
<html>
<body>
	<%
	Dim lngSum,I						'lngSum用来存放结果，I是循环计数器变量
	lngSum=0							'给lngSum赋初值0 
	For I=1 To 100						'计数器变量I从1循环到100
		lngSum=lngSum+I^2									
	Next
	Response.Write "1到100的平方和=" & lngSum           
	%>
</body> 
</html>
