<html>
<body>
	<h2 align="center">排序显示数据示例</h2>
	<% 
	'首先获取传递过来的排序字段名称
	Dim strField
	If Request.QueryString("varField")="" Then
		strField="strName"						'如果为空，就默认按姓名字段排序
	Else
		strField=Request.QueryString("varField")		'如果不为空，就按相应字段排序
	End If
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立记录集，建立一个RecordSet对象实例rs
	Dim rs,strSql 
	strSql="Select * From tbAddress Order By " & strField			'按指定字段排序
	Set rs=conn.Execute(strSql)
	'以下利用表格显示记录集中的记录
	%>
	<table border="1" width="100%" >
		<tr bgcolor="#E0E0E0">
			<th><a href="9-1.asp?varField=strName">姓名</a></th>
			<th><a href="9-1.asp?varField=strSex">性别</a></th>
			<th><a href="9-1.asp?varField=intAge">年龄</a></th>
			<th><a href="9-1.asp?varField=strTel">电话</a></th>
			<th><a href="9-1.asp?varField=strEmail">E-mail</a></th>
			<th><a href="9-1.asp?varField=strIntro">简介</a></th>
			<th><a href="9-1.asp?varField=dtmSubmit">添加日期</a></th>
		</tr>
		<% 
		Do While Not rs.Eof							'只要不是结尾就执行循环
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
			rs.MoveNext							'将记录指针移动到下一条记录
		Loop
		%>
	</table>
</body> 
</html>
