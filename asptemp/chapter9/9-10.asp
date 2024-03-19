<html>
<body>
	<h2 align="center">利用Recordset对象存取数据库</h2>
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
	rs.Open strSql,conn,1,2				'注意后两个参数，表示键盘指针和可以更新
	'以下查询记录，利用循环简单列出人员姓名和电话
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
	'以下添加记录
	rs.AddNew									'添加一个新记录，指针会指向该记录
	rs("strName")="李玫"						'给新记录的字段赋值
	rs("strTel")="88888888"						'同上
	rs("strSex")="女"							'同上
	rs("intAge")=22								'同上
	rs("strTel")="limei@371.net"				'同上
	rs("strIntro")="一个大学生"					'同上
	rs("dtmSubmit")=Date()						'同上，令字段值为当天日期
	rs.Update									'更新数据库
	'以下更新记录（更新ID=5的人员的电话）
	rs.MoveFirst								'将记录指针移动到第1条记录
	rs.Find "ID=5"								'找到这条记录，指针会指向该记录
	If Not rs.Eof Then							'判断是否找到了记录
		rs("strTel")="66666666"					'更新当前记录的字段值
		rs.Update								'更新数据库 
	End If
	'以下删除记录（删除ID=6的人员）
	rs.MoveFirst								'将记录指针移动到第1条记录
	rs.Find "ID=6"								'找到这条记录，指针会指向该记录
	If Not rs.Eof Then							'判断是否找到了记录
		rs.Delete								'删除当前记录	
		rs.Update								'更新数据库
	End If
	%>
</body>
</html>

