<html>
<body>
	<h2 align="center">��ȡ�����ı��ļ�</h2>
	<%
	Dim fso									'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm									'����һ��TextStream����ʵ��
	Set tsm= fso.OpenTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt",1,True)
	Do While Not tsm.AtEndOfStream
		Response.Write tsm.ReadLine			'���ж�ȡ��ֱ���ļ���β
		Response.Write "<br>"				'��ҳ���ϻ�����ʾ
	Loop
	tsm.Close								'�ر�TextStream����
	%>
</body> 
</html> 

