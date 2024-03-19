<html>
<body>
	<h2 align="center">最新消息</h2>
	<center><a href="insert_form.asp">发布新消息</a></center>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim strConn,conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("news.mdb")
	conn.Open strConn
	'以下建立Recordset对象实例rs
	Dim strSql,rs 
	strSql="Select * From tbNews Order By dtmSubmit DESC"        
	Set rs=conn.Execute(strSql)
	'以下利用循环显示数据库记录
	Do while Not rs.Eof
	%>
		<table border="0" width="80%" align="center">
			<tr><td colspan=2><hr></td></tr>
			<tr><td width="20%">标题</td><td><%=rs("strTitle")%></td></tr>
			<tr><td>内容</td><td><%=rs("strBody")%></td></tr>
			<tr><td>附件</td><td>
			<a href="upfiles/<%=rs("strFilename")%>"><%=rs("strFilename")%></a></td></tr>
			<tr><td>时间</td><td><%=rs("dtmSubmit")%></td></tr>
		</table>
	<%
		rs.MoveNext 
	Loop 
	%>
	</center>
</body> 
</html>






