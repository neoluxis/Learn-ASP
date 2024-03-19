<%
'===============================================================================================
'particular.asp说明
'1. 本页面和9.2.5节的详细页面示例基本上是一样的，只是简单的显示一个人员的详细信息。
'2. 为了美观，这里将信息显示在了表格中。
'3. 关闭有关对象的语句也可以省略掉，因为执行完本页面后也会自动关闭的。
'4. 本页面中的数据库连接语句也放在了connection.asp中
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>人员详细信息</title>
<head>
</body>
	<h2 align="center">人员详细信息</h2>

	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<% 
		'以下建立一个RecordSet对象实例rs，注意要用到传递过来的ID值
		Dim rs,strSql 
		strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID") 	        
		Set rs=conn.Execute(strSql)

		'以下显示查找的数据
		Response.Write "<tr><td width='30%' height='25'>姓名</td><td>" & rs("strName") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>性别</td><td>" & rs("strSex") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>年龄</td><td>" & rs("intAge") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>电话</td><td>" & rs("strTel") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>E-mail</td><td>" & rs("strEmail") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>简介</td><td>" & rs("strIntro") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>添加日期</td><td>" & rs("dtmSubmit") & "</td></tr>"

		'现在关闭有关对象
		rs.Close
		Set rs=Nothing
		conn.Close
		Set conn=Nothing
		%>
	</table>

	<p align="center"><a href="#" onclick="window.close()">关闭窗口</a>
</body> 
</html> 