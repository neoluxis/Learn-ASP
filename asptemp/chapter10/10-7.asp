<html>
<body>
	<h2 align="center">File���������ʾ��</h2>
	<%
	Dim fso									'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim fle									'����һ��File����ʵ��
	Set fle=fso.GetFile("C:\Inetpub\wwwroot\asptemp\chapter10\test.txt")
	Response.Write "<br>�ļ�����" & fle.Name 
	Response.Write "<br>�ļ����ԣ�" & fle.Attributes 
	Response.Write "<br>·����" & fle.Path 
	Response.Write "<br>��С��" & fle.Size 
	Response.Write "<br>�������ڣ�" & fle.DateCreated 
	%>
</body> 
</html>

