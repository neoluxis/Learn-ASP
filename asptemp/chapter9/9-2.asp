<html>
<body>
	<h2 align="center">���Ҽ�¼ʾ��</h2>
	<form name="frmSearch" method="POST" action="">
		������Ҫ���ҵ�������<input type="text" name="txtName">
		<input type="submit" name="btnSubmit" value=" ȷ �� ">
	</form>
	<%
	If Request.Form("txtName")<>"" Then
		'�����������ݿ⣬����һ��Connection����ʵ��conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn 
		'���½���һ��RecordSet����ʵ��rs��ע��Select�����Ҫ�õ��ύ������
		Dim rs,strSql 
		strSql="Select * From tbAddress Where strName Like '%" & Request.Form("txtName") & "%'"
		Set rs=conn.Execute(strSql)
		'�������ñ����ʾ���ҵ��ļ�¼
		%>
		<table border="1" width="100%" >
			<tr bgcolor="#E0E0E0">
				<th>����</th><th>�Ա�</th><th>����</th><th>�绰</th><th>E-mail</th>
				<th>���</th><th>�������</th>
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
	<%End If%>
</body> 
</html>
