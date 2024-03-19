<html>
<body>
	<h2 align="center">文件夹的新建、复制、移动和删除</h2>
	<%
	Dim fso											'声明一个FileSystemObject对象实例
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim SourceFolder,DestiFolder					'声明源文件夹和目标文件夹变量
	'新建文件夹---新建new1文件夹
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	fso.CreateFolder SourceFolder					'新建文件夹   
	'复制文件夹---将new1复制为new2文件夹
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	DestiFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new2"
	fso.CopyFolder SourceFolder, DestiFolder		'复制文件夹   
	'移动文件夹---将new2文件夹移动到new1下
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new2"
	DestiFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1\new2"
	fso.MoveFolder SourceFolder, DestiFolder		'移动文件夹   
	'删除文件夹---如存在，将new1文件夹删除
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	IF fso.FolderExists(SourceFolder)=True Then
		fso.DeleteFolder SourceFolder				'删除文件夹        
	End If
	%>
</body> 
</html>

