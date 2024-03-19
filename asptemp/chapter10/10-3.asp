<html>
<body>
   <h2 align="center">新建一个文本文件</h2>
	<%
	Dim fso									'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'声明一个TextStream对象实例
	Set tsm=fso.CreateTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",True)
	tsm.WriteLine "这是第一句" 				'向文件中写一行内容
	tsm.WriteLine "这是第二句" 				'再写一行内容
	tsm.Close								'关闭TextStream对象
	Response.Write "已经成功建立文件，请自己打开查看。"
	%>
</body> 
</html>
