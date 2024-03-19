<%
'=========================================================================================
'BBS首页
'1. 这是BBS首页，其中只是显示论坛栏目列表，并且显示用户登录框
'2. 在用户登录框中，会根据用户是否已经登录显示不同的内容。如果没有登录，就显示登录框；如果已经登录了， 就显示退出登录、修改密码和修改个人信息的超链接。
'3. 在论坛栏目列表中，注意单击图片或栏目名称都会打开相应的栏目，另外，会将栏目的ID编号传递到bbs_list.asp中。
'4. 在显示图片时用了一个小技巧，图片文件名字依次为1.gif、2.gif、3.gif... 这里用循环输出每一个图片文件的路径。
'=========================================================================================
%>

<% Option Explicit						'强制声明变量 %>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>bbs论坛</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css" >
	<script language="JAVASCRIPT">
		function check_Null(){
			if (document.frmReg.txtUserId.value==""){
				alert("用户名不能为空!");
				return false;
			}
			if (document.frmReg.txtPwd.value==""){
				alert("密码不能为空!");
				return false;
			}
			return true;
		}
	</script>
</head>

<body>
	<h1 align="center"><font face="黑体"><%=conBBSTitle%></font></h1>

	<form name="frmReg" method="POST" action="log_in.asp" onsubmit="Javascript: return check_Null();">
	<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>
		<tr bgcolor="#E1F3F4" height="40">    
			<td valign="middle">
				<%If Session("strUserId")<>"" Then%>
					已登录用户<input type="text" name="txtUserId" size="15" value="<%=Session("strUserId")%>" disabled>
					<a href="log_out.asp">【注销】</a>
					<a href="log_updatePWD.asp">【修改密码】</a>
					<a href="log_update.asp">【修改个人信息】</a>
				<%Else%>
					用户名<input type="text" name="txtUserId" size="15">
					密码<input type="PassWord" name="txtPwd" size="15">
					<input type="submit" value="登 录"> 
					<input type="button" value="注 册" onclick="window.open('log_register1.asp','_self')">
				<%End If%>
			</td>  
		</tr>  
	</table>
	</form>

	<br>
	<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>
		<tr bgcolor="#C5EDE7" height="30" align="center">    
			<td width="20%">&nbsp;</td>  
			<td width="40%">论坛栏目</td>
			<td width="40%">文章数</td>
		</tr>  
		<%
		'以下建立记录集对象实例rs
		Dim rs,strSql
		strSql="Select * From tbForum"
		Set rs=conn.Execute(strSql)
		'下面利用循环输出所有栏目，其中I变量用来顺序显示图片
		Dim I
		Do While Not rs.Eof
			I=I+1
			Response.Write "<tr height='60' valign='middle' align='center'>"
			Response.Write "<td><a href='bbs_list.asp?intForumId=" & rs("ID") & "'><img src='images/" & I & ".gif' border='0'></a></td>"
			Response.Write "<td><a href='bbs_list.asp?intForumId=" & rs("ID") & "'>" & rs("strForumName") & "</a></td>"
			Response.Write "<td>" & rs("lngForumCount") & "</td>"
			Response.Write "</tr>"  
			rS.MoveNext                
		Loop
		'关闭对象
		rs.Close
		Set rs=Nothing
		conn.Close
		Set conn=Nothing
		%>   
	</table> 
</body>
</html>
