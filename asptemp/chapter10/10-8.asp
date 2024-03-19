<html>
<body>
	<h2 align="center">显示指定文件夹下所有子文件夹和文件</h2>
	<%
	Dim fso										'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim fld										'声明一个Folder对象
	Set fld=fso.GetFolder("C:\Inetpub\wwwroot\asptemp\chapter10")
	Response.Write "子文件夹共" & fld.SubFolders.Count & "个<br>"
	'下面利用循环输出子文件夹集合中的所有Folder对象
	Dim objItem									'声明一个对象变量
	For Each objItem In fld.SubFolders
		Response.Write objItem.Path & "<br>"
	Next
	Response.Write "文件共" & fld.Files.Count & "个<br>"
	'下面利用循环输出文件集合中的所有File对象
	For Each objItem In fld.Files
		Response.Write objItem.Path & "<br>"
	Next
	%>
</body> 
</html> 


