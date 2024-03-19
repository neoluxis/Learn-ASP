<html>
<body>
	<h2 align="center">列出所有驱动器名称</h2>
	<%
	Dim fso									'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Response.Write "共有" & fso.Drives.Count & "个驱动器"
	Dim objItem
	For Each objItem in fso.Drives			'下面用循环列出每个驱动器的名称
		Response.Write "<br>驱动器名称：" & objItem.DriveLetter
	Next
	%>
</body> 
</html>

