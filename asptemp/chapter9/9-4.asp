<html>
<body>
	<h2 align="center">详细页面</h2>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立一个RecordSet对象实例rs，注意要用到传递过来的ID值
	Dim rs,strSql 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID") 	        
	Set rs=conn.Execute(strSql)
	'以下显示查找的数据
	Response.Write "<br>姓名：" & rs("strName")
	Response.Write "<br>性别：" & rs("strSex")
	Response.Write "<br>年龄：" & rs("intAge")
	Response.Write "<br>电话：" & rs("strTel")
	Response.Write "<br>E-mail：" & rs("strEmail")
	Response.Write "<br>简介：" & rs("strIntro")
	Response.Write "<br>添加日期：" & rs("dtmSubmit")
	%>
	<p align="center"><a href="#" onclick="window.close()">关闭窗口</a>
</body> 
</html> 