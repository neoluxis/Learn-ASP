<html>
<body>
	<h2 align="center">�ļ��е��½������ơ��ƶ���ɾ��</h2>
	<%
	Dim fso											'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim SourceFolder,DestiFolder					'����Դ�ļ��к�Ŀ���ļ��б���
	'�½��ļ���---�½�new1�ļ���
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	fso.CreateFolder SourceFolder					'�½��ļ���   
	'�����ļ���---��new1����Ϊnew2�ļ���
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	DestiFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new2"
	fso.CopyFolder SourceFolder, DestiFolder		'�����ļ���   
	'�ƶ��ļ���---��new2�ļ����ƶ���new1��
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new2"
	DestiFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1\new2"
	fso.MoveFolder SourceFolder, DestiFolder		'�ƶ��ļ���   
	'ɾ���ļ���---����ڣ���new1�ļ���ɾ��
	SourceFolder="C:\Inetpub\wwwroot\asptemp\chapter10\new1"
	IF fso.FolderExists(SourceFolder)=True Then
		fso.DeleteFolder SourceFolder				'ɾ���ļ���        
	End If
	%>
</body> 
</html>

