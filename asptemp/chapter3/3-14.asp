<% Option Explicit					'强制声明变量%>                    
<html>
<body>
	<%
	Dim intA(9,9)					'声明一个10行10列的二维数组
	Dim I,J							'I是外层循环计数器变量，J是内层循环计数器变量
	For I=0 To 9					'外层循环
		For J=0 To 9				'内层循环
			intA(I,J)=10			'给每一个元素赋初值10
		Next
	Next
	Response.Write "程序运行结束。" 
	%>
</body> 
</html>
