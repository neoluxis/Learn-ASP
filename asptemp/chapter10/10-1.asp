<html>
<body>
	<h2 align="center">文件的复制、移动和删除</h2>
	<%
	Dim fso											'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim SourceFile,DestiFile						'声明源文件和目标文件变量
	'复制文件---将test.txt复制为test2.txt
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\test.txt"
	DestiFile="C:\Inetpub\wwwroot\asptemp\chapter10\test2.txt"
	fso.CopyFile SourceFile, DestiFile				'复制文件   
	'移动文件---将test2.txt移动到temp文件夹下，应保证temp文件夹存在
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\test2.txt"
	DestiFile="C:\Inetpub\wwwroot\asptemp\chapter10\temp\test2.txt"
	fso.MoveFile SourceFile, DestiFile				'移动文件   
	'删除文件---如果test2.txt存在，则将其删除
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\temp\test2.txt"
	IF fso.FileExists(SourceFile)=True Then
		fso.DeleteFile SourceFile					'删除文件        
	End If
	%>
</body> 
</html>
 
