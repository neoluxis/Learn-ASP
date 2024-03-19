<html>
<body>
	<h2 align="center">File对象的属性示例</h2>
	<%
	Dim fso									'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim fle									'声明一个File对象实例
	Set fle=fso.GetFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt")
	Response.Write "<br>文件名：" & fle.Name 
	Response.Write "<br>文件属性：" & fle.Attributes 
	Response.Write "<br>路径：" & fle.Path 
	Response.Write "<br>大小：" & fle.Size 
	Response.Write "<br>创建日期：" & fle.DateCreated 
	%>
</body> 
</html>

