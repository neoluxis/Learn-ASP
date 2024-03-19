<html>
<body>
	<h2 align="center">读取已有文本文件</h2>
	<%
	Dim fso									'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'声明一个TextStream对象实例
	Set tsm= fso.OpenTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",1,True)
	Do While Not tsm.AtEndOfStream
		Response.Write tsm.ReadLine			'逐行读取，直到文件结尾
		Response.Write "<br>"				'在页面上换行显示
	Loop
	tsm.Close								'关闭TextStream对象
	%>
</body> 
</html> 

