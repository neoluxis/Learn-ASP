<html>
<body>
	<h2 align="center">����Command����������ݿ�Ļ�������</h2>
	<%
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½���Command����ʵ��cmd
	Dim cmd
	Set cmd= Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn						'ָ��Connection����
	'���²�ѯ��¼
	Dim rs,strSql
	strSql="Select * From tbAddress"
	cmd.CommandType=1								'��ʾ��ѯָ��ΪSQL���
	cmd.CommandText=strSql							'ָ����ѯָ��
	Set rs=cmd.Execute								'ִ�в�ѯ������һ��Recordset����
	'��������ѭ�����г���Ա�����͵绰
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	'������Ӽ�¼
	strSql ="Insert Into tbAddress(strName,strTel,intAge) Values('��õ','88888888',21)"
	cmd.CommandText=strSql							'ָ����ѯָ��
	cmd.Execute										'ִ�в�������Ӽ�¼
	'���¸��¼�¼
	strSql ="Update tbAddress Set strTel='66666666' Where ID=2"
	cmd.CommandText=strSql							'ָ����ѯָ��
	cmd.Execute										'ִ�в��������¼�¼
	'����ɾ����¼
	strSql="Delete From tbAddress Where ID=2"
	cmd.CommandText=strSql							'ָ����ѯָ��
	cmd.Execute										'ִ�в�����ɾ����¼
	%>
</body>
</html>
