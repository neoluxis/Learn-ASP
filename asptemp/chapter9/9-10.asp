<html>
<body>
	<h2 align="center">����Recordset�����ȡ���ݿ�</h2>
	<%
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½���Recordset����ʵ��rs
	Dim rs,strSql
	Set rs=Server.CreateObject("ADODB.Recordset")
	strSql ="Select * From tbAddress"						
	rs.Open strSql,conn,1,2				'ע���������������ʾ����ָ��Ϳ��Ը���
	'���²�ѯ��¼������ѭ�����г���Ա�����͵绰
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	'������Ӽ�¼
	rs.AddNew									'���һ���¼�¼��ָ���ָ��ü�¼
	rs("strName")="��õ"						'���¼�¼���ֶθ�ֵ
	rs("strTel")="88888888"						'ͬ��
	rs("strSex")="Ů"							'ͬ��
	rs("intAge")=22								'ͬ��
	rs("strTel")="limei@371.net"				'ͬ��
	rs("strIntro")="һ����ѧ��"					'ͬ��
	rs("dtmSubmit")=Date()						'ͬ�ϣ����ֶ�ֵΪ��������
	rs.Update									'�������ݿ�
	'���¸��¼�¼������ID=5����Ա�ĵ绰��
	rs.MoveFirst								'����¼ָ���ƶ�����1����¼
	rs.Find "ID=5"								'�ҵ�������¼��ָ���ָ��ü�¼
	If Not rs.Eof Then							'�ж��Ƿ��ҵ��˼�¼
		rs("strTel")="66666666"					'���µ�ǰ��¼���ֶ�ֵ
		rs.Update								'�������ݿ� 
	End If
	'����ɾ����¼��ɾ��ID=6����Ա��
	rs.MoveFirst								'����¼ָ���ƶ�����1����¼
	rs.Find "ID=6"								'�ҵ�������¼��ָ���ָ��ü�¼
	If Not rs.Eof Then							'�ж��Ƿ��ҵ��˼�¼
		rs.Delete								'ɾ����ǰ��¼	
		rs.Update								'�������ݿ�
	End If
	%>
</body>
</html>

