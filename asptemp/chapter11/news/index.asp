<html>
<body>
	<h2 align="center">������Ϣ</h2>
	<center><a href="insert_form.asp">��������Ϣ</a></center>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim strConn,conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("news.mdb")
	conn.Open strConn
	'���½���Recordset����ʵ��rs
	Dim strSql,rs 
	strSql="Select * From tbNews Order By dtmSubmit DESC"        
	Set rs=conn.Execute(strSql)
	'��������ѭ����ʾ���ݿ��¼
	Do while Not rs.Eof
	%>
		<table border="0" width="80%" align="center">
			<tr><td colspan=2><hr></td></tr>
			<tr><td width="20%">����</td><td><%=rs("strTitle")%></td></tr>
			<tr><td>����</td><td><%=rs("strBody")%></td></tr>
			<tr><td>����</td><td>
			<a href="upfiles/<%=rs("strFilename")%>"><%=rs("strFilename")%></a></td></tr>
			<tr><td>ʱ��</td><td><%=rs("dtmSubmit")%></td></tr>
		</table>
	<%
		rs.MoveNext 
	Loop 
	%>
	</center>
</body> 
</html>






