<html>
<body>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"									
	'��������Connection�����Execute��������¼�¼
	conn.Execute("Insert Into tbAddress(strName,strSex,intAge,strTel,strEmail,strIntro,dtmSubmit) Values('����','Ů',21,'6112211','mm@372.net','����ϵͬѧ',#2008-8-8#)")
	Response.Write "�Ѿ��ɹ���ӣ�������Լ���address.mdb�鿴�����"
	%>
</body> 
</html>

