<html>
<body>
	<h2 align="center">���ı��ļ���׷������</h2>
	<%
	Dim fso									'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'����һ��TextStream����ʵ��
	Set tsm= fso.OpenTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",8,True)
	tsm.WriteLine "���ǵ�����" 				'׷������
	tsm.Close								'�ر�TextStream����
	Response.Write "�Ѿ��ɹ�׷�ӣ����Լ��򿪲鿴��"
	%>
</body> 
</html> 
