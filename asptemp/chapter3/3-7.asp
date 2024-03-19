<% Option Explicit										'强制变量声明 %>
<!--#Include file="3-8.asp"-->
<html>
<body>
	<%
	Dim intM,intN,lngResult								'intM和intN为实际参数
	intM=3
	intN=4
	lngResult=mySquare(intM,intN)						'调用函数，返回平方和
	Response.Write "3和4的平方和是：" & lngResult
	%>   
</body>
</html>
