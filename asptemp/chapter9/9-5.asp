<html>
<body>
	<h2 align="center">�������÷�ʾ��</h2>
	<%
	On Error Resume Next								'���������������ִ����һ��
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���濪ʼ������
	conn.BeginTrans										
	'ɾ����¼
	strSql="Delete From tbAddress Where strName='��õ'"
	conn.Execute(strSql)
	'��Ӽ�¼
	strSql="Insert Into tbAddress(strName,strTel) Values('��õ','88888888')"
	conn.Execute(strSql)
	'��������Ƿ��������������Ƿ��ύ������
	If  conn.Errors.Count=0 Then							
		conn.CommitTrans								'����޴��󣬾��ύ��������
		Response.Write "�ɹ�ִ��"
	Else
		conn.RollbackTrans								'����д��󣬾�ȡ����������
		Response.Write "�д�������ȡ��������"
	End If
	%>
</body>
</html>

