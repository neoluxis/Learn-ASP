<html>
<body>
	<h2 align="center">�ҵ�ͨ��¼</h2>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"									'��������Դ�������ݿ�
	'���½�����¼��������һ��Recordset����ʵ��rs
	Dim rs											 
	Set rs=conn.Execute("Select * From tbAddress")			'�����������ݱ�
	'�������ñ����ʾ��¼���еļ�¼
	%>
	<table border="1" width="100%" align="center">
		<tr bgcolor="#E0E0E0">
			<th>����</th><th>�Ա�</th><th>����</th><th>�绰</th><th>E-mail</th>
			<th>���</th><th>�������</th>
		</tr>
		<% 
		Do While Not rs.Eof									'ֻҪ���ǽ�β��ִ��ѭ��
		%>
			<tr>
				<td><%=rs("strName")%></td>
				<td><%=rs("strSex")%></td>
				<td><%=rs("intAge")%></td>
				<td><%=rs("strTel")%></td>
				<td><a href="mailto:<%=rs("strEmail")%>"><%=rs("strEmail")%></a> </td>
				<td><%=rs("strIntro")%></td>
				<td><%=rs("dtmSubmit")%></td>
			</tr>
		<% 
			rs.MoveNext										'����¼ָ���ƶ�����һ����¼
		Loop
		%>
	</table>
</body> 
</html> 
