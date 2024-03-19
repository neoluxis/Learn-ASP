<% Option Explicit							'强制声明变量%>                    
<html>
<body>
	<%
	Dim lngSum,I							'lngSum用来保存结果，I用来控制循环
	lngSum=0								'给lngSum赋初值 
	I=1												
	Do While I<=100							'当I小于等于100时执行循环
		lngSum=lngSum+I^2					
		I=I+1								'I的值增加1
	Loop
	Response.Write "1到100的平方和=" & lngSum           
	%>
</body> 
</html>
