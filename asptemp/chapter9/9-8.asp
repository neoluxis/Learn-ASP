<html>
<body>
	<h2 align="center">�ǲ�����ѯʾ��</h2>
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
	'ִ�в�ѯqryList
	Dim rs
	cmd.CommandType=4							'��ʾ��ѯָ���ǲ�ѯ��
	cmd.CommandText="qryList"						'ָ����ѯ����
	Set rs=cmd.Execute								'ִ�в�ѯ������Recordset����
	'��������ѭ�����г���Ա�����͵绰
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	%>
</body>
</html>

