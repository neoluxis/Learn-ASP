<html>
<body>
	<h2 align="center">对多个表进行组合查询示例</h2>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("userinfo.mdb")
	conn.Open strConn 
	'以下建立记录集对象实例rs
	Dim strSql,rs                                       
	strSql="Select tbUsers.strUserId,tbUsers.strName,tbLog.strIP,tbLog.dtmLog From tbUsers,tbLog Where tbUsers.strUserId=tbLog.strUserId Order By tbLog.dtmLog DESC"   
	Set rs=conn.Execute(strSql)
	'以下利用循环显示记录集，并显示在表格中
	Response.Write "<table border='1' width='100%'><tr bgcolor='#E0E0E0'>"
	Response.Write "<th>用户名</th><th>姓名</th><th>登录IP</th><th>登录时间</th></tr>"
	Do While Not rs.Eof                            
		Response.Write "<tr>"
		Response.Write "<td>" & rs("strUserId") & "</td>"
		Response.Write "<td>" & rs("strName") & "</td>"
		Response.Write "<td>" & rs("strIP") & "</td>"
		Response.Write "<td>" & rs("dtmLog") & "</td>"
		Response.Write "</tr>"
		rs.MoveNext                        
	Loop
	Response.Write "</table>"
	%>
</body> 
</html>

