<html>
<body>
	<h2 align="center">利用Command对象进行数据库的基本操作</h2>
	<%
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立Command对象实例cmd
	Dim cmd
	Set cmd= Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn						'指定Connection对象
	'以下查询记录
	Dim rs,strSql
	strSql="Select * From tbAddress"
	cmd.CommandType=1								'表示查询指令为SQL语句
	cmd.CommandText=strSql							'指定查询指令
	Set rs=cmd.Execute								'执行查询，返回一个Recordset对象
	'以下利用循环简单列出人员姓名和电话
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	'以下添加记录
	strSql ="Insert Into tbAddress(strName,strTel,intAge) Values('李玫','88888888',21)"
	cmd.CommandText=strSql							'指定查询指令
	cmd.Execute										'执行操作，添加记录
	'以下更新记录
	strSql ="Update tbAddress Set strTel='66666666' Where ID=2"
	cmd.CommandText=strSql							'指定查询指令
	cmd.Execute										'执行操作，更新记录
	'以下删除记录
	strSql="Delete From tbAddress Where ID=2"
	cmd.CommandText=strSql							'指定查询指令
	cmd.Execute										'执行操作，删除记录
	%>
</body>
</html>
