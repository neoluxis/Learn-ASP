<%
'=========================================================================================
'BBS列表页
'1. 本页用来分页显示某个栏目对应的文章，显示时按发表日期倒序排列。
'2. 注意本页只显示第一层文章，不显示回复文章。
'3. 要注意两个特殊的变量intForumId和intPage，前者用来确定显示哪个栏目？后者用来确定显示哪一页。 这两个变量是传递过来的， 而且从本页链接到别的页面时， 也会将这两个变量传递过去， 之后再传递回来。 总之，在不同的页面之间要始终记得传递这两个变量。
'4. 有两种情况下，intPage要设为1；第一就是从首页打开本栏目时，第二是发表新文章后， 这样可以确保马上看到新发表的文章。
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>bbs论坛</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css">
</head>

<body>
	<h1 align="center"><font face="黑体"><%=conBBSTitle%></font></h1>
	<%
	'下面一段分别获取本页面需要的两个变量，首先获取论坛栏目编号变量---------------------------
	Dim intForumId 
	intForumId=CInt(Request.QueryString("intForumId"))			'获取传递过来的论坛编号 

	'获取数据页码变量
	Dim intPage                              
	If Request.QueryString("intPage")<>"" Then
		intPage=CInt(Request.QueryString("intPage"))			'获取传递过来的页码，转换为数字
	Else
		intPage=1												'其他情况下设为1，当从首页过来或发表文章后就设为1
	End If
	'下面输出发表文章和返回首页的超链接---------------------------------------------------
	Response.Write "<a href='bbs_insert.asp?intForumId=" & intForumId & "'>【发表文章】</a>"
	Response.Write "<a href='index.asp'>【返回论坛】</a>"
	Response.Write "<p>"										'为了美观，输出一个空行
	'下面首先输出表格的标题栏-----------------------------------------------------------------
	Response.Write "<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>"
	Response.Write "<tr bgcolor='#C5EDE7' height='30' align='center'>"
	Response.Write "<th width='5%'>编号</th>"
	Response.Write "<th width='55%'>主题</th>"
	Response.Write "<th width='5%'>点击</th>"
	Response.Write "<th width='5%'>回复</th>"
	Response.Write "<th width='20%'>发表时间</th>"
	Response.Write "<th width='10%'>作者</th>"
	Response.Write "</tr>"
	'下面建立记录集实例rs，请注意Open方法的参数-----------------------------------------------
	Dim rs,strSql
	strSql="Select * From tbBBS Where intForumId=" & intForumId & " And intLayer=1 Order By dtmSubmit Desc"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open strSql,conn,1									'因为要分页显示，所以用键盘指针
	'下面如果非空就显示记录
	If Not rs.Bof And Not rs.Eof Then   
	'以下主要为了分页显示
		rs.PageSize=conPageSize								'设置每页显示多少条记录，从配置文件中读取
		rs.AbsolutePage=intPage								'设置当前显示第几页
		'下面利用循环显示当前页的所有记录
		Dim I
		For I=1 To rs.PageSize								'循环输出当前页的所有记录
			If rs.Eof Then Exit For							'如果到了记录集结尾，就跳出循环
			Response.Write "<tr bgcolor='#E1F3F4' height='30' valign='middle'>"  
			Response.Write "<td align='center'>" & rs("ID") & "</td>"
			Response.Write "<td align='left'><a href='bbs_particular.asp?ID=" & rs("ID") & "&intForumId=" & intForumId & "&intPage=" & intPage & "'>" & myHTMLEncode(rs("strTitle")) & "</a></td>"
			Response.Write "<td align='center'>" & rs("intHits") & "</td>"
			Response.Write "<td align='center'>" & rs("intChild") & "</td>"  
			Response.Write "<td align='center'>" & rs("dtmSubmit") & "</td>" 
			Response.Write "<td align='center'><a href='mailto:" & rs("strEmail") & "'>" & rs("strUserId") & "</a></td>"
			Response.Write "</tr>"
			rS.MoveNext                
		Next
		'下面输出表格的结束标记，表格到此结束
		Response.Write "</table>"
		'下面输出当前页和总页数--------------------------------------------------------------------------
		Response.Write "<p align='right'>当前显示第" & intPage & "页/共" & rs.PageCount & "页"
		'下面一段依次输出第1页、上一页、下一页和最后页的超链接
		Response.Write "&nbsp;&nbsp;<a href='bbs_list.asp?intPage=1&intForumId=" & intForumId & "'>【第1页】</a>"
		If intPage>1 Then
			Response.Write "<a href='bbs_list.asp?intPage=" & (intPage-1) & "&intForumId=" & intForumId & "'>【上一页】</a>"
		Else
			Response.Write "【上一页】"
		End If
		If intPage<rs.PageCount Then
			Response.Write "<a href='bbs_list.asp?intPage=" & (intPage+1) & "&intForumId=" & intForumId & "'>【下一页】</a>"
		Else
			Response.Write "【下一页】"
		End If
		Response.Write "<a href='bbs_list.asp?intPage=" & rs.PageCount  & "&intForumId=" & intForumId & "'>【最后页】</a>"
	End If
	'关闭对象
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>   
</body>
</html>