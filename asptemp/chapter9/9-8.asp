<html>
<body>
	<h2 align="center">非参数查询示例</h2>
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
	'执行查询qryList
	Dim rs
	cmd.CommandType=4							'表示查询指令是查询名
	cmd.CommandText="qryList"						'指定查询名称
	Set rs=cmd.Execute								'执行查询，返回Recordset对象
	'以下利用循环简单列出人员姓名和电话
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	%>
</body>
</html>

