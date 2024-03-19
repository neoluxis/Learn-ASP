<%
'=========================================================================================
'BBS详细页
'1. 本页面中心思想是：首先根据传递过来的文章编号，将文章及回复文章全部列出来；然后在下面再给出一个回复表单。 如果用户提交了回复表单，就提交到bbs_reinsert.asp中去执行回复操作。
'2. 在注意，在显示文章之前，会首先将本文章的点击数增加1.
'3. 要注意本页也会获取intForumId和intPage两个变量，然后返回BBS列表时都会将这两个变量再传递过去。
'4. 在显示作者个人信息时，会调用config.asp中的函数，获取个人信息数组。
'5. 注意，在回复表单中，会将几个重要的变量用隐藏文本框传递了过去。
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>bbs论坛</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JAVASCRIPT">
		function check_Null(){
			if (document.frmRe.txtTitle.value==""){
				alert("主题不能为空!");
				return false;
			}
			if (document.frmRe.txtUserId.value==""){
				alert("用户名不能为空!");
				return false;
			}
			return true;
		}
	</script>
</head>
<body>
	<h1 align="center">详细页面</h1>
	<%
	'下面首先获取传递过来的论坛编码intForumId，页码信息intPage和记录编号ID----------------------
	Dim ID,intForumId,intPage
	ID=Request.QueryString("ID")
	intForumId=Request.QueryString("intForumId")
	intPage=Request.QueryString("intPage")
	'下面将该记录的点击数加1--------------------------------------------------------------------
	Dim strSql,rs
	strSql="Update tbBBS Set intHits=intHits+1 Where id=" & id
	conn.Execute(strSql)
	'下面开始输出内容，首先输出表格-------------------------------------------------------------
	Response.Write "<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>"
	'下面建立记录集，并利用循环显示所有记录
	Dim strTitle
	'请注意这条SQL语句，它会获取当前文章及其所有回复文章的记录，并按日期升序排列
	strSql="Select * From tbBBS Where ID=" & ID & " Or intFatherId=" & ID & " Order By dtmSubmit"
	Set rs=conn.Execute(strSql)
	strTitle=rs("strTitle")
	Dim I														'变量I用来判断是第几楼
	Do While not rs.Eof 
		I=I+1
		'下面每条记录都用两行两列的单元格来显示
		'下面是第1行
		Response.Write "<tr bgcolor='#C5EDE7' height='30'>"
		'下面是第1行第1列
		Response.Write "<td width='20%' align='center'>"
		'下面判断是否楼主，分别显示不同的信息
		If I=1 Then
			Response.Write "楼主"						
		Else
			Response.Write "第" & I & "楼"				
		End If
		Response.Write "</td>"
		'下面是第1行第2列
		Response.Write "<td>"
		'下面判断是否楼主，如果是楼主，显示发布时间、点击数、回复数，否则只显示发布时间
		If I=1 Then
			Response.Write "发布时间：" & rs("dtmSubmit") & "&nbsp;点击：" & rs("intHits") & "&nbsp;回复：" & rs("intChild")
		Else
			Response.Write "发布时间：" & rs("dtmSubmit") 
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		'下面是第2行
		Response.Write "<tr valign='top'>"
		'下面是第2行第1列
		Response.Write "<td>"
		Response.Write "作者：<a href='mailto:" & rs("strEmail") & "'>" & rs("strUserId") & "</a>"
		'如果不是过客，就调用function.asp中的函数，返回该作者的一些信息
		If rs("strUserId")<>"过客" Then
			Dim strTemp
			strTemp=PersonalInfo(rs("strUserId"))
			Response.Write "<br>发表：" & strTemp(4) & "篇"
			Response.Write "<br>回复：" & strTemp(5) & "篇"
			Response.Write "<br>QQ：" & strTemp(2) 
		End If
		Response.Write "</td>"
		'下面是第2行第2列，其中会调用文本处理函数
		Response.Write "<td>"
		Response.Write "<b><font color='#b40000'>" & myHTMLEncode(rs("strTitle")) & "</font></b>"
		Response.Write "<p>" & myHTMLEncode(rs("strBody"))
		Response.Write "</td>"
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	'输出表格结束标记，输出内容到此结束
	Response.Write "</table>"
	%>

	<!--注释：下面是返回BBS列表页面的超链接，注意其中将两个变量传递了回去-->
	<p align="center"><a href="bbs_list.asp?intForumId=<%=intForumId%>&intPage=<%=intPage%>">【返回论坛】</a>

	<!--注释：下面是回复文章用的表格-->
	<form name="frmRe" action="bbs_reinsert.asp" method="POST" onsubmit="javascript: return check_Null();">
	<table border="0" width="80%" cellspacing="0" cellpadding="0" align="center">
		<tr>                                        
			<td width="20%">文章主题：</td>
			<!--注释：下面将回复标题显示在文本框中-->
			<td><input type="text" name="txtTitle" size="50" value="re：<%=strTitle%>"></td>
		</tr>
		<tr>
			<td>文章内容：</td>
			<td><textarea rows="6" name="txtBody" cols="60"></textarea></td>
		</tr>
		<tr>                                        
			<td>用户名：</td>
			<td>
				<!--注释：下面根据是否已经登录显示不同的用户名，如没有登录，则为“过客”-->
				<%If Session("strUserId")<>"" Then%>
					<input type="text" name="txtUserId" size="20" value="<%=Session("strUserId")%>" disabled>
				<%Else%>
					<input type="text" name="txtUserId" size="20" value="过客" disabled>
				<%End If%>
			</td>               
		</tr>
		<tr>                                        
			<td>E-mail：</td>
			<!--注释：如果登录了，就将用户的E-mail显示在文本框中-->
			<td><input type="text" name="txtEmail" size="30" value="<%=Session("strEmail")%>"></td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<br>
				<!--注释：这里用隐藏文本框将几个重要的变量传递了过去-->
				<input type="hidden" name="txtFatherId" value="<%=ID%>">
				<input type="hidden" name="txtForumId" value="<%=intForumId%>">
				<input type="hidden" name="txtPage" value="<%=intPage%>">
				<input type="submit" value="确定提交">&nbsp; 
				<!--注释：下面是返回BBS列表页面的按钮，注意其中将两个变量传递了回去-->
				<input type="button" value="返回论坛" name="返回论坛" onclick="window.open('bbs_list.asp?intForumId=<%=intForumId%>&intPage=<%=intPage%>','_self')" >
			</td>
		</tr>
	</table>
	</form>    
</body>
</html>
