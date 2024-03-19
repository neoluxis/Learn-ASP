<html>
<body>
	<h2 align="center">利用文件超链接组件建立目录页</h2>
	<%
	Dim link													'声明一个NextLink对象实例
	Set link=Server.CreateObject("MSWC.NextLink")
	Dim I
	For I=1 To link.GetListCount("link.txt")					'用循环依次写出所有的文件链接
		Response.Write "<a href='" & link.GetNthURL("link.txt",I) & "'>" & link.GetNthDescription("link.txt",I) & "</a><p>"
	Next
	%>
</body> 
</html>



