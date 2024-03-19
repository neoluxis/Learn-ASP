<%
'===============================================================================================
'search.asp说明
'1. 本页面实际上是在9.2.4节查找记录示例的基础上修改而成的。
'2. 这里只是将表格等HTML代码都用Response.Write 方法输出了， 另外，添加了”更新“、”删除“和链接到详细页面的超链接。 这些部分实际上和index.asp差不多。
'3. 需要注意的是本程序还不是十分完美，比如在本页面中单击”删除“或”更新“超链接， 执行完毕后不会返回本页面， 而是会直接返回首页index.asp。 
'4. 关闭有关对象的语句也可以省略掉，因为执行完本页面后也会自动关闭的。
'5. 本页面中的数据库连接语句也放在了connection.asp中
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>查找人员</title>
</head>
<body>
	<h2 align="center">查找人员</h2>
	<form name="frmSearch" method="POST" action="">
		请输入要查找的姓名：<input type="text" name="txtName">
		<input type="submit" name="btnSubmit" value=" 确 定 ">
	</form>
	<%
	'如果输入了姓名，就执行下面的查找过程
	If Request.Form("txtName")<>"" Then

		'以下建立一个RecordSet对象实例rs。注意Select语句中要用到提交的姓名-----------------------------------
		Dim rs,strSql 
		strSql="Select ID,strName,strTel,strEmail,dtmSubmit From tbAddress Where strName Like '%" & Request.Form("txtName") & "%'"
		Set rs=conn.Execute(strSql)

		'以下利用表格显示查找到的记录------------------------------------------------------------------------
		'下面首先输出表格的标题栏
		Response.Write "<table width='100%' border='1' bordercolorlight='#B0B0B0' bordercolordark='#FFFFFF' cellspacing='0' cellpadding='0' align='center'>"
		Response.Write "<tr bgcolor='#E0E0E0' height='25'>"
		Response.Write "<th width='10%'>编号</th>"
		Response.Write "<th width='10%'>姓名</th>"
		Response.Write "<th width='20%'>电话</th>"
		Response.Write "<th width='20%'>E-mail</th>"
		Response.Write "<th width='20%'>添加日期</th>"
		Response.Write "<th width='10%'>更新</th>"
		Response.Write "<th width='10%'>删除</th>"
		Response.Write "</tr>"

		'下面利用循环显示所有记录
		Do While Not rs.Eof							'只要不是结尾就执行循环
			Response.Write "<tr align='center' height='25'>"								
			Response.Write "<td>" & rs("ID") & "</td>"
			Response.Write "<td><a href='particular.asp?ID=" & rs("ID") & "' target='_blank'>" & rs("strName") & "</td>"
			Response.Write "<td>" & rs("strTel") & "</td>"
			Response.Write "<td><a href='mailto:" & rs("strEmail") & "'>" & rs("strEmail") & "</td>"
			Response.Write "<td>" & rs("dtmSubmit") & "</td>"
			Response.Write "<td><a href='update.asp?ID=" & rs("ID") & "'>更新</td>"
			Response.Write "<td><a href='delete.asp?ID=" & rs("ID") & "'>删除</td>"
			Response.Write "</tr>"
			rs.MoveNext								'将记录指针移动到下一条记录
		Loop
		
		'下面输出表格的结束标记，表格到此结束
		Response.Write "</table>"
	
	End If

	'现在关闭Connection对象-----------------------------------------------------------------------------------
	conn.Close
	Set conn=Nothing
	%>

	<p align='center'><a href="index.asp">返回首页</a>
</body> 
</html>
