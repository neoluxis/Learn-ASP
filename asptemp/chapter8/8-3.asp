<html>
<body>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"									
	'��������Connection�����Execute����ɾ����¼
	conn.Execute("Delete From tbAddress Where ID=1")
	Response.Write "�Ѿ��ɹ�ɾ����������Լ���address.mdb�鿴�����"
	%>
</body> 
</html> 
