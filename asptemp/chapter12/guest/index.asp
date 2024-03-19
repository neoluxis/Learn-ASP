<%
'================================================================================================
'首页文件
'1. 本页主要分为两部分：上面是一个添加留言的表单，表单会被提交到insert.asp；下面是显示所有留言的部分， 就是利用循环显示所有记录而已。
'2. 要注意这里在表单中使用了客户端的JavaScript验证，通过验证后才会继续提交表单，否则就提示用户重新填写。
'3. 本页会调用样式文件guest.css设置有关文字、超链接等的样式。
'4. 本页会读取config.asp中的配置，显示留言板的名称
'5. 在下面显示留言时，会调用function.asp中的函数，对留言主题和内容进行编码，以便显示HTML代码和实现换行效果。
'================================================================================================
%>

<% Option Explicit						'强制声明变量%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>欢迎访问我的留言板</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="guest.css">
	<script language="JavaScript">
		//该函数用来进行客户端验证
		function check_Null(){
			if (document.frmGuest.txtTitle.value==""){
				alert("主题不能为空!");
				return false;
			}
			if (document.frmGuest.txtName.value==""){
				alert("姓名不能为空!");
				return false;
			}
			if (document.frmGuest.txtTitle.value.length>50){
				alert("主题不能超过50个字符");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<!--注释：下面要从配置文件中读取常量conGuestTitle-->
	<h1 align="center"><font face="黑体"><%=conGuestTitle%></font></h1>

	<!--注释：下面是提交留言表单，提交后，首先会调用上面的客户端验证函数验证，验证通过后，再传送到insert.asp-->
	<form name="frmGuest" method="POST" action="insert.asp" onsubmit="JavaScript:return check_Null();">
		<table border="0" width="80%" bgcolor="#203F80" align="center">
			<tr>
				<td><font color="white">主题：</font></td>
				<td><input type="text" name="txtTitle" size="60">*</td>
			</tr><tr>
				<td><font color="white">内容：</font></td>
				<td><Textarea Name="txtBody" Rows="4" Cols="60"></TextArea></td>
			</tr><tr>
				<td><font color="white">姓名：</font></td>
				<td><input type="text" name="txtName" size="13">*</td>
			</tr><tr>
				<td><font color="white">E-mail：</font></td>
				<td><input type="text" name="txtEmail" size="40"></td>
			</tr><tr>
				<td></td>
				<td><input type="submit" value="　提   交　" Size="20" ></td>
			</tr>
		</table>
	</form>
	<%
	'以下开始显示原有留言，请注意每条留言会显示在一个表格中
	Dim rs,strSql
	strSql ="Select * From tbGuest Order By dtmSubmit Desc"	
	Set rs=conn.Execute(strSql)
	Do While Not rs.Eof
		%>
		<table border="0" width="80%" align="center">
			<tr><td colspan="2"><hr></td></tr>
			<tr><td width="20%">主题</td><td><%=myHTMLEncode(rs("strTitle"))%></td></tr>
			<tr><td>内容</td><td><%=myHTMLEncode(rs("strBody"))%></td></tr>
			<tr><td>留言人</td><td><a href="mailto:<%=rs("strEmail")%>"><%=myHTMLEncode(rs("strName"))%></a></td></tr>
			<tr><td>时间</td><td><%=rs("dtmSubmit")%></td></tr>
			<tr><td></td><td><a href="delete.asp?ID=<%=rs("ID")%>">删除</a></td></tr>
		</table>
		<%
		rs.MoveNext
	Loop 
	'关闭对象
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>
</body>
</html>