<html>
<body>
	<h2 align="center">������ʾ����ʾ��</h2>
	<% 
	'���Ȼ�ȡ���ݹ����������ֶ�����
	Dim strField
	If Request.QueryString("varField")="" Then
		strField="strName"						'���Ϊ�գ���Ĭ�ϰ������ֶ�����
	Else
		strField=Request.QueryString("varField")		'�����Ϊ�գ��Ͱ���Ӧ�ֶ�����
	End If
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½�����¼��������һ��RecordSet����ʵ��rs
	Dim rs,strSql 
	strSql="Select * From tbAddress Order By " & strField			'��ָ���ֶ�����
	Set rs=conn.Execute(strSql)
	'�������ñ����ʾ��¼���еļ�¼
	%>
	<table border="1" width="100%" >
		<tr bgcolor="#E0E0E0">
			<th><a href="9-1.asp?varField=strName">����</a></th>
			<th><a href="9-1.asp?varField=strSex">�Ա�</a></th>
			<th><a href="9-1.asp?varField=intAge">����</a></th>
			<th><a href="9-1.asp?varField=strTel">�绰</a></th>
			<th><a href="9-1.asp?varField=strEmail">E-mail</a></th>
			<th><a href="9-1.asp?varField=strIntro">���</a></th>
			<th><a href="9-1.asp?varField=dtmSubmit">�������</a></th>
		</tr>
		<% 
		Do While Not rs.Eof							'ֻҪ���ǽ�β��ִ��ѭ��
		%>
			<tr>
				<td><%=rs("strName")%></td>
				<td><%=rs("strSex")%></td>
				<td><%=rs("intAge")%></td>
				<td><%=rs("strTel")%></td>
				<td><a href="mailto:<%=rs("strEmail")%>"><%=rs("strEmail")%></a></td>
				<td><%=rs("strIntro")%></td>
				<td><%=rs("dtmSubmit")%></td>
			</tr>
		<% 
			rs.MoveNext							'����¼ָ���ƶ�����һ����¼
		Loop
		%>
	</table>
</body> 
</html>
