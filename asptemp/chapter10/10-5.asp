<html>
<body>
	<h2 align="center">在文本文件中追加内容</h2>
	<%
	Dim fso									'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'声明一个TextStream对象实例
	Set tsm= fso.OpenTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",8,True)
	tsm.WriteLine "这是第三句" 				'追加内容
	tsm.Close								'关闭TextStream对象
	Response.Write "已经成功追加，请自己打开查看。"
	%>
</body> 
</html> 
