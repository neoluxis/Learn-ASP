<html>
<body>
	<h2 align="center">�ļ��ĸ��ơ��ƶ���ɾ��</h2>
	<%
	Dim fso											'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim SourceFile,DestiFile						'����Դ�ļ���Ŀ���ļ�����
	'�����ļ�---��test.txt����Ϊtest2.txt
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\test.txt"
	DestiFile="C:\Inetpub\wwwroot\asptemp\chapter10\test2.txt"
	fso.CopyFile SourceFile, DestiFile				'�����ļ�   
	'�ƶ��ļ�---��test2.txt�ƶ���temp�ļ����£�Ӧ��֤temp�ļ��д���
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\test2.txt"
	DestiFile="C:\Inetpub\wwwroot\asptemp\chapter10\temp\test2.txt"
	fso.MoveFile SourceFile, DestiFile				'�ƶ��ļ�   
	'ɾ���ļ�---���test2.txt���ڣ�����ɾ��
	SourceFile="C:\Inetpub\wwwroot\asptemp\chapter10\temp\test2.txt"
	IF fso.FileExists(SourceFile)=True Then
		fso.DeleteFile SourceFile					'ɾ���ļ�        
	End If
	%>
</body> 
</html>
 
