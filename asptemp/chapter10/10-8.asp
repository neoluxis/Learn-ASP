<html>
<body>
	<h2 align="center">��ʾָ���ļ������������ļ��к��ļ�</h2>
	<%
	Dim fso										'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim fld										'����һ��Folder����
	Set fld=fso.GetFolder("C:\Inetpub\wwwroot\asptemp\chapter10")
	Response.Write "���ļ��й�" & fld.SubFolders.Count & "��<br>"
	'��������ѭ��������ļ��м����е�����Folder����
	Dim objItem									'����һ���������
	For Each objItem In fld.SubFolders
		Response.Write objItem.Path & "<br>"
	Next
	Response.Write "�ļ���" & fld.Files.Count & "��<br>"
	'��������ѭ������ļ������е�����File����
	For Each objItem In fld.Files
		Response.Write objItem.Path & "<br>"
	Next
	%>
</body> 
</html> 


