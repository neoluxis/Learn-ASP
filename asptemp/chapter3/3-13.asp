<% Option Explicit						'强制声明变量%>                    
<html>
<body>
	<%
	Dim strA(2),strSum,Item				'strSum用来保存结果，Item用来返回数组元素
	strA(0)="a"							'给数组元素赋值
	strA(1)="b"
	strA(2)="c"
	For Each Item in strA				'执行循环，取出每个元素
		strSum=strSum & Item			
	Next
	Response.Write "全部数组元素形成的字符串是：" & strSum
	%>
</body> 
</html>

