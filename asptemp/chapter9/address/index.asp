<%
'===============================================================================================
'index.asp说明
'1. 本页面是将分页显示数据、排序显示数据、链接到详细页面三个示例整合到了一起。 另外，增加了“添加记录、 查找记录、 更新记录、 删除记录”的超链接。
'2. 本页面的一个主要难点就是要同时考虑排序和分页显示数据， 也就是要同时考虑 strField 和 intPage参数。 实际上这两者是有关系的。 比如，当选择按”姓名“排序后， 点击”下一页“超链接再次打开本页面时， 不仅要记住当前要显示页码， 还要记住当前的排序字断。 所以在后面几个超链接中不仅要传回来页码， 还要传回来排序字段。
'3. 当然，如果再次在标题栏中选择了排序字段， 因为改变了记录集中的记录的顺序， 所以此时没有必要记住原来所在的页码， 直接显示第1页即可。
'4. 本程序在设置每页显示多少条记录时使用了 config.asp中的定义的常数。
'5. 本程序还在表格的美观方面做了一些修改，请大家结合第2章内容来理解。
'===============================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="config.asp"-->
<html>
<head>
	<title>我的通信录</title>
</head>
<body>
	<h1 align="center">通信录</h1>

	<!--下面首先给添加记录和查找记录的超链接 -->
	<a href="insert.asp">添加记录</a>&nbsp;&nbsp;<a href="search.asp">查找记录</a><p>

	<%
	'下面首先获取传递过来的排序字段名称-------------------------------------------------------------
	Dim strField
	If Request.QueryString("varField")="" Then
		strField="ID"										'如果为空，就默认按编号字段排序
	Else
		strField=Request.QueryString("varField")			'如果不为空，就按相应字段排序
	End If
	'下面一段判断当前显示第几页，如是第一次打开，为1，否则由传回参数决定
	Dim intPage                            
	If Request.QueryString("varPage")="" Then   
		intPage=1 
	Else
		intPage=CInt(Request.QueryString("varPage"))		'用CInt转换为整数
	End If
	
	'以下建立Recordset对象实例rs，注意SQL语句中要用到上面获取的排序字段-----------------------------
	Dim rs,strSql
	Set rs=Server.CreateObject("ADODB.Recordset")
	strSql ="Select ID,strName,strTel,strEmail,dtmSubmit From tbAddress Order By " & strField & " DESC"
	rs.Open strSql,conn,1									'注意参数设置为键盘指针

	'如果记录集不是空的，就执行分页显示
	If Not rs.Bof And Not rs.Eof Then
		'-------------------------------------------------------------------------------------------
		'下面一段开始分页显示，指向要显示的页，然后逐条显示当前页的所有记录
		rs.PageSize=conPageSize								'这里调用config.asp中的常数设置每页显示多少条记录
		rs.AbsolutePage=intPage								'设置当前显示第几页

		'下面首先输出表格的标题栏
		Response.Write "<table width='100%' border='1' bordercolorlight='#B0B0B0' bordercolordark='#FFFFFF' cellspacing='0' cellpadding='0' align='center'>"
		Response.Write "<tr bgcolor='#E0E0E0' height='25'>"
		Response.Write "<th width='10%'><a href='index.asp?varField=ID'>编号</a></th>"
		Response.Write "<th width='10%'><a href='index.asp?varField=strName'>姓名</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=strTel'>电话</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=strEmail'>E-mail</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=dtmSubmit'>添加日期</a></th>"
		Response.Write "<th width='10%'>更新</th>"
		Response.Write "<th width='10%'>删除</th>"
		Response.Write "</tr>"

		'下面利用循环显示当前页的所有记录
		Dim I
		For I=1 To rs.PageSize
			If rs.Eof Then Exit For							'如果到了记录集结尾，就跳出循环
			Response.Write "<tr align='center' height='25'>"								
			Response.Write "<td>" & rs("ID") & "</td>"
			Response.Write "<td><a href='particular.asp?ID=" & rs("ID") & "' target='_blank'>" & rs("strName") & "</td>"
			Response.Write "<td>" & rs("strTel") & "</td>"
			Response.Write "<td><a href='mailto:" & rs("strEmail") & "'>" & rs("strEmail") & "</td>"
			Response.Write "<td>" & rs("dtmSubmit") & "</td>"
			Response.Write "<td><a href='update.asp?ID=" & rs("ID") & "'>更新</td>"
			Response.Write "<td><a href='delete.asp?ID=" & rs("ID") & "'>删除</td>"
			Response.Write "</tr>"
			rs.MoveNext
		Next
		'下面输出表格的结束标记，表格到此结束
		Response.Write "</table>"

		'-------------------------------------------------------------------------------
		'下面首先输出当前页和总页数
		Response.Write "<p align='right'>当前显示第" & intPage & "页/共" & rs.PageCount & "页"

		'下面一段依次输出第1页、上一页、下一页和最后页的超链接
		Response.Write "&nbsp;&nbsp;<a href='index.asp?varPage=1&varField=" & strField & "'>第1页</a>&nbsp;"
		If intPage>1 Then
			Response.Write "<a href='index.asp?varPage=" & (intPage-1) & "&varField=" & strField & "'>上一页</a>&nbsp;"
		Else
			Response.Write "上一页&nbsp;"
		End If
		If intPage<rs.PageCount Then
			Response.Write "<a href='index.asp?varPage=" & (intPage+1) & "&varField=" & strField & "'>下一页</a>&nbsp;"
		Else
			Response.Write "下一页&nbsp;"
		End If
		Response.Write "<a href='index.asp?varPage=" & rs.PageCount  & "&varField=" & strField & "'>最后页</a>&nbsp;"
		'-------------------------------------------------------------------------------
	End If
	'现在关闭有关对象
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>
</body>
</html>

