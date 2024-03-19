<% Option Explicit								'强制变量声明 %>
<html>
<body>
	<%
	Dim intM,intN								'intM和intN为实际参数
	intM=3
	intN=4
	Call mySquare(intM,intN)					'调用子程序，显示结果
	'下面是子程序，用来计算两个数的平方和
	Sub mySquare(intA,intB)						'intA和intB是形式参数
		Dim lngSum
		lngSum=intA^2+intB^2  
		Response.Write "3和4的平方和是：" & lngSum    
	End Sub
	%>   
</body>
</html>

