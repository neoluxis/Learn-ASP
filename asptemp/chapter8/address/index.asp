<html>
<body>
	<h2 align="center">ͨ��¼</h2>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½�����¼��������һ��RecordSet����ʵ��rs
	Dim rs,strSql 
	strSql="Select * From tbAddress Order By ID DESC"			'���Զ�����ֶν�������		        
	Set rs=conn.Execute(strSql)
	'�������ñ�����ʾ��¼���еļ�¼
	%>
	<a href="insert.asp">���Ӽ�¼</a>
	<table border="1" width="100%" >
		<tr bgcolor="#E0E0E0">
			<th>����</th><th>�Ա�</th><th>����</th><th>�绰</th><th>E-mail</th>
			<th>���</th><th>��������</th><th>ɾ��</th><th>����</th>
		</tr>
		<% 
		Do While Not rs.Eof										'ֻҪ���ǽ�β��ִ��ѭ��
		%>
			<tr>
				<td><%=rs("strName")%></td>
				<td><%=rs("strSex")%></td>
				<td><%=rs("intAge")%></td>
				<td><%=rs("strTel")%></td>
				<td><a href="mailto:<%=rs("strEmail")%>"><%=rs("strEmail")%></a> </td>
				<td><%=rs("strIntro")%></td>
				<td><%=rs("dtmSubmit")%></td>
				<td><a href="delete.asp?ID=<%=rs("ID")%>">ɾ��</a></td>
				<td><a href="update_form.asp?ID=<%=rs("ID")%>">����</a></td>
			</tr>
		<% 
			rs.MoveNext											'����¼ָ���ƶ�����һ����¼
		Loop
		%>
	</table>
</body> 
</html> 