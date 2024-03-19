<html>
<body>
	<h2 align="center">我的通信录</h2>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"									'利用数据源连接数据库
	'以下建立记录集，建立一个Recordset对象实例rs
	Dim rs											 
	Set rs=conn.Execute("Select * From tbAddress")			'返回整个数据表
	'以下利用表格显示记录集中的记录
	%>
	<table border="1" width="100%" align="center">
		<tr bgcolor="#E0E0E0">
			<th>姓名</th><th>性别</th><th>年龄</th><th>电话</th><th>E-mail</th>
			<th>简介</th><th>添加日期</th>
		</tr>
		<% 
		Do While Not rs.Eof									'只要不是结尾就执行循环
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
			rs.MoveNext										'将记录指针移动到下一条记录
		Loop
		%>
	</table>
</body> 
</html> 
