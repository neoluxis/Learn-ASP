<html>
<body>
   <h2 align="center">�½�һ���ı��ļ�</h2>
	<%
	Dim fso									'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'����һ��TextStream����ʵ��
	Set tsm=fso.CreateTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",True)
	tsm.WriteLine "���ǵ�һ��" 				'���ļ���дһ������
	tsm.WriteLine "���ǵڶ���" 				'��дһ������
	tsm.Close								'�ر�TextStream����
	Response.Write "�Ѿ��ɹ������ļ������Լ��򿪲鿴��"
	%>
</body> 
</html>
