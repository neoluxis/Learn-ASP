<%
'=========================================================================================
'用户修改个人信息页
'1. 这其实就是一个普通的更新记录页面。首先将原有内容显示在表单中，提交表单后再更新记录。
'2. 在更新记录时有些信息可以省略，所以SQL语句较为复杂。以QQ号码为例，如果用户原来提交了QQ号码，在这里删除了QQ号码。 那么此时就需要将该字段值清空， 这里使用NULL关键字，这样该字段值就被清空了。 事实上此时也可以用空字符串""将其清空。
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>修改个人信息</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmInfo.txtName.value==""){
				alert("真实姓名不能为空!");
				return false;
			}
			if (document.frmInfo.txtEmail.value==""){
				alert("E-mail不能为空!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">修改个人信息</h2>
	<%
	'下面读取该用户的信息，然后将其显示在后面的表格中
	Dim strSql,rs
	strSql="Select * From tbUsers Where strUserId='" & Session("strUserId") & "'"
	Set rs=conn.Execute(strSql)
	%>
	<form name="frmInfo" onsubmit="javascript: return check_Null();" action="" method="post">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%">用户名</td>
			<td width="80%"><%=Session("strUserId")%></td>
		</tr>
		<tr height="25"> 
			<td>真实姓名</td>
			<td><input type="text" name="txtName" size=15 value="<%=rs("strName")%>">*</td>
		</tr>
		<tr height="25"> 
			<td>性别</td>
			<td>
				<input type="radio" name="rdoSex" value="男" <%If rs("strSex")="男" Then Response.Write "checked"%>>男
				<input type="radio" name="rdoSex" value="女" <%If rs("strSex")="女" Then Response.Write "checked"%>>女
			</td>
		</tr>
		<tr height="25"> 
			<td>电话</td>
			<td><input type="text" name="txtTel" size="25" value="<%=rs("strTel")%>"></td>
		</tr>
		<tr height="25"> 
			<td>E-mail</td>
			<td><input type="text" name="txtEmail" size="40" value="<%=rs("strEmail")%>">*</td>
		</tr>
		<tr height="25"> 
			<td>QQ号码 </td>
			<td><input type="text" name="txtQQ" size="15" value="<%=rs("strQQ")%>"></td>
		</tr>
		<tr height="25"> 
			<td>个人简介</td>
			<td><textarea name="txtIntro" rows="4" cols="50" ><%=rs("strIntro")%></textarea></td>
		</tr>
		<tr height="25"> 
			<td>发表文章</td>
			<td><%=rs("intArticle")%>篇</td>
		</tr>
		<tr height="25"> 
			<td>回复文章</td>
			<td><%=rs("intReArticle")%>篇</td>
		</tr>
	</table>
	<p align="center"><input type="submit" value="  确　定  ">
	</form>

	<%
	If Request.Form("txtName")<>"" And Request.Form("txtEmail")<>"" Then
		'下面建立SQL语句，因为某些字段允许为空，所以需要判断一下
		strSql="Update tbUsers Set strName='" & Request.Form("txtName") & "'"
		strSql=strSql & ",strEmail='" & Request.Form("txtEmail") & "'"
		strSql=strSql & ",strSex='" & Request.Form("rdoSex") & "'"
		strSql=strSql & ",dtmSubmit=#" & Date() & "#"
		'注意：如果用户没有提交QQ号码，那么不管原来有没有QQ号码，都将该字段值用NULL清空了。
		If Request.Form("txtQQ") <> "" Then
			strSql = strSql & ",strQQ='" & Request.Form("txtQQ") & "'"
		Else
			strSql = strSql & ",strQQ=NULL"
		End If
		'注意：关于电话的解释同上面的QQ
		If Request.Form("txtTel") <> "" Then
			strSql = strSql & ",strTel='" & Request.Form("txtTel") & "'"
		Else
			strSql = strSql & ",strTel=NULL"
		End If
		'注意：关于备注的解释同上面的QQ
		If Request.Form("txtIntro") <> "" Then
			strSql = strSql & ",strIntro='" & Request.Form("txtIntro") & "'"
		Else
			strSql = strSql & ",strIntro=NULL"
		End If	
		strSql=strSql & " Where strUserId='" & Session("strUserId") & "'"
		conn.Execute(strSql)
		'关闭对象
		conn.close
		Set conn=Nothing
		Response.Redirect "index.asp"
	End If
	%>
</body>
</html>