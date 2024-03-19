<html>
<body>
	<h2 align="center">Fields集合与Field对象示例</h2>
	<%
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立Recordset对象实例rs
	Dim rs,strSql
	Set rs=Server.CreateObject("ADODB.Recordset")
	strSql ="Select * From tbAddress"						
	rs.Open strSql,conn
	'用一个循环将所有字段属性写出
	Dim I,fld
	For I=0 to rs.Fields.Count-1
		Set fld=rs.Fields.Item(I)					'建立当前字段的Field对象实例fld
		Response.Write "字段名称：" & fld.Name & "<br>"
		Response.Write "字段值：" & fld.Value & "<br>"
		Response.Write "字段类型：" & fld.Type & "<br>"
		Response.Write "字段大小：" & fld.Definedsize & "<br>"
		Response.Write "字段最大位数：" & fld.Precision & "<br>"
	Next
	%>
</body>
</html>
