<html>
<body>
	<h2 align="center">�г���������������</h2>
	<%
	Dim fso									'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Response.Write "����" & fso.Drives.Count & "��������"
	Dim objItem
	For Each objItem in fso.Drives			'������ѭ���г�ÿ��������������
		Response.Write "<br>���������ƣ�" & objItem.DriveLetter
	Next
	%>
</body> 
</html>

