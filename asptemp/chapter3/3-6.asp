<% Option Explicit										'强制变量声明 %>
<html>
<body>
	<%
	Dim intM,intN,lngResult								'intM和intN为实际参数
	intM=3
	intN=4
	lngResult=mySquare(intM,intN)						'调用函数，返回平方和
	Response.Write "3和4的平方和是：" & lngResult
	'下面是函数，用来显示两个数的平方和
	Function mySquare(intA,intB)						'intA和intB是形式参数
		Dim lngSum
		lngSum=intA^2+intB^2  
		mySquare=lngSum									'赋值给函数名，很重要
	End Function
	%>   
</body>
</html>
