<html>
<body>
<h2 align="center">分页显示数据示例</h2>
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
rs.Open strSql,conn,1									'注意参数设置为键盘指针
'如果记录集不是空的，就执行分页显示
If Not rs.Bof And Not rs.Eof Then
	'-------------------------------------------------------------------------------
	'下面一段判断当前显示第几页，如果是第一次打开，为1，否则由传回参数决定
	Dim intPage                            
	If Request.QueryString("varPage")="" Then   
		intPage=1 
	Else
		intPage=CInt(Request.QueryString("varPage"))	'用CInt转换为整数
	End If
	'-------------------------------------------------------------------------------
	'下面一段开始分页显示，指向要显示的页，然后逐条显示当前页的所有记录。
	rs.PageSize=5										'设置每页5条记录
	rs.AbsolutePage=intPage								'设置当前显示第几页
	'下面利用循环显示当前页的所有记录，并显示在表格中
	Response.Write "<table border='1' width='100%'><tr bgcolor='#E0E0E0'>"
	Response.Write "<th>姓名</th><th>电话</th><th>E-mail</th> <th>添加日期</th></tr>"
	Dim I												'定义一个变量I
	For I=1 To rs.PageSize
		If rs.Eof Then Exit For							'如果到了记录集结尾，就跳出循环
		Response.Write "<tr>"								
		Response.Write "<td>" & rs("strName") & "</td>"
		Response.Write "<td>" & rs("strTel") & "</td>"
		Response.Write "<td>" & rs("strEmail") & "</td>"
		Response.Write "<td>" & rs("dtmSubmit") & "</td>"
		Response.Write "</tr>"
		rs.MoveNext
	Next
	Response.Write "</table>"
	'-------------------------------------------------------------------------------
	'下面一段依次输出第1页、上一页、下一页和最后页的超链接
	Response.Write "<p><a href='9-12.asp?varPage=1'>第1页</a>&nbsp;"
	If intPage>1 Then
		Response.Write "<a href='9-12.asp?varPage=" & (intPage-1) & "'>上一页</a>&nbsp;"
	Else
		Response.Write "上一页&nbsp;"
	End If
	If intPage<rs.PageCount Then
		Response.Write "<a href='9-12.asp?varPage=" & (intPage+1) & "'>下一页</a>&nbsp;"
	Else
		Response.Write "下一页&nbsp;"
	End If
	Response.Write "<a href='9-12.asp?varPage=" & rs.PageCount & "'>最后页</a>&nbsp;"
	'-------------------------------------------------------------------------------
Else
	Response.Write "该记录集为空"
End If
%>
</body>
</html>
