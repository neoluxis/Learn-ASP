<html>
<body>
	<%
	Dim intA,intB,intC
	intA=Request.Form("txtA")				'获取表单元素txtA的值
	intB=Request.Form("txtB")				'获取表单元素txtB的值
	intC=CInt(intA)+CInt(intB)				'因为获取的是字符串，所以必须转换类型
	Response.Write "两个数的和=" & intC	
	%>
</body>
</html>

