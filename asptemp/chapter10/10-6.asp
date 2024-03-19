<html>
<body>
   <h2 align="center">自动生成HTML文件</h2>
	<%
	Dim fso										'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm										'声明一个TextStream对象实例
	Set tsm=fso.CreateTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\temp.htm",True)
	tsm.WriteLine "<html>"						'向文件中写入内容，以下同
	tsm.WriteLine "<head>"						
	tsm.WriteLine "<title>我的主页</title>"	
	tsm.WriteLine "</head>"					
	tsm.WriteLine "<body>"
	tsm.WriteLine "<h2 align='center'>我的主页</h2>"
	tsm.WriteLine "<p align='center'>欢迎大家访问"
	tsm.WriteLine "</body>"						
	tsm.WriteLine "</html>"						
	tsm.Close									'关闭TextStream对象
	Response.Write "已经成功建立文件，请自己打开查看。"
	%>
</body> 
</html>
