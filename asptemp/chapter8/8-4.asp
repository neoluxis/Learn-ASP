<html>
<body>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"                                
	'��������Connection�����Execute�������¼�¼
	conn.Execute("Update tbAddress Set intAge=22,strTel='61223211' Where ID=2")
	Response.Write "�Ѿ��ɹ����£�������Լ���address.mdb�鿴�����"
	%>
</body> 
</html> 
