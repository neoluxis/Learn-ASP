<html>
<body>
	<h2 align="center">通信录</h2>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立记录集，建立一个RecordSet对象实例rs
	Dim rs,strSql 
	strSql="Select * From tbAddress Order By ID DESC"			'按自动编号字段降序排列		        
	Set rs=conn.Execute(strSql)
	'以下利用表格显示记录集中的记录
	%>
	<table border="1" width="100%" >
		<tr bgcolor="#E0E0E0">
			<th>姓名</th><th>电话</th><th>E-mail</th><th>详细</th>
		</tr>
		<% 
		Do While Not rs.Eof										'只要不是结尾就执行循环
		%>
			<tr>
				<td><%=rs("strName")%></td>
				<td><%=rs("strTel")%></td>
				<td><a href="mailto:<%=rs("strEmail")%>"><%=rs("strEmail")%></a></td>
				<td><a href="9-4.asp?ID=<%=rs("ID")%>" target="_blank">详细</a></td>
			</tr>
		<% 
			rs.MoveNext											'将记录指针移动到下一条记录
		Loop
		%>
	</table>
</body> 
</html> 