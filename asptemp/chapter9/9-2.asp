<html>
<body>
	<h2 align="center">查找记录示例</h2>
	<form name="frmSearch" method="POST" action="">
		请输入要查找的姓名：<input type="text" name="txtName">
		<input type="submit" name="btnSubmit" value=" 确 定 ">
	</form>
	<%
	If Request.Form("txtName")<>"" Then
		'以下连接数据库，建立一个Connection对象实例conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn 
		'以下建立一个RecordSet对象实例rs。注意Select语句中要用到提交的姓名
		Dim rs,strSql 
		strSql="Select * From tbAddress Where strName Like '%" & Request.Form("txtName") & "%'"
		Set rs=conn.Execute(strSql)
		'以下利用表格显示查找到的记录
		%>
		<table border="1" width="100%" >
			<tr bgcolor="#E0E0E0">
				<th>姓名</th><th>性别</th><th>年龄</th><th>电话</th><th>E-mail</th>
				<th>简介</th><th>添加日期</th>
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
	<%End If%>
</body> 
</html>
