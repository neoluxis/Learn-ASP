<%
'=========================================================================================
'用户注册页---第二步
'1. 这一步要用户填写详细信息。其实就是更新记录页面。
'2. 因为有的信息可以省略，所以这里的SQL字符串比较复杂。另外，这里用的是Update语句，最后也会组成一条标准的SQL语句。
'3. 如果用户名已经被使用，就提醒用户修改用户名。
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>注册：第二步</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmReg.txtName.value==""){
				alert("真实姓名不能为空!");
				return false;
			}
			if (document.frmReg.txtEmail.value==""){
				alert("E-mail不能为空!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">用户注册</h2>
	<p align="center">第二步 填写详细信息 （注意：所有带*号的项目必须填写） 

	<form name="frmReg" onsubmit="javascript: return check_Null();" action="" method="POST">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%" >用户名</td>
			<td width="80%"><%=Session("strUserId")%></td>
		</tr>
		<tr height="25"> 
			<td>真实姓名</td>
			<td><input type="text" name="txtName" size="15">*</td>
		</tr>
		<tr height="25"> 
			<td>性别</td>
			<td><input type="radio" name="rdoSex" value="男" checked>男
				<input type="radio" name="rdoSex" value="女">女</td>
		</tr>
		<tr height="25"> 
			<td>电话 </td>
			<td><input type="text" name="txtTel" size="25"></td>
		</tr>
		<tr height="25"> 
			<td>E-mail </td>
			<td><input type="text" name="txtEmail" size="40">*</td>
		</tr>
		<tr height="25"> 
			<td>QQ号码 </td>
			<td><input type="text" name="txtQQ" size="15"></td>
		</tr>
		<tr height="25"> 
			<td>个人简介</td>
			<td><textarea name="txtIntro" rows="4" cols="50"></textarea></td>
		</tr>
	</table>
	<p align="center"><input type="submit" value="  确　定  ">
	</form>

	<%
	If Request.Form("txtName")<>"" And Request.Form("txtEmail")<>"" Then
		Dim strUserId,strSql
		strUserId=Session("strUserId")						'从Session中获取用户名
		'因为有的字段是允许忽略的，所以下面要根据情况来建立SQL语句
		'首先更新肯定有的字段值，下面4行也可以写在一行中，这样写是为了看得更清楚
		strSql="Update tbUsers Set strName='" & Request.Form("txtName") & "'"
		strSql=strSql & ",strEmail='" & Request.Form("txtEmail") & "'"
		strSql=strSql & ",strSex='" & Request.Form("rdoSex") & "'"
		strSql=strSql & ",dtmSubmit=#" & Date() & "#"
		'下面根据是否提交了信息，从而更新相应字段值
		If Request.Form("txtQQ") <> "" Then
			strSql=strSql & ",strQQ='" & Request.Form("txtQQ") & "'"
		End If
		If Request.Form("txtTel") <> "" Then
			strSql=strSql & ",strTel='" & Request.Form("txtTel") & "'"
		End If
		If Request.Form("txtIntro") <> "" Then
			strSql=strSql & ",strIntro='" & Request.Form("txtIntro") & "'"
		End If
		strSql=strSql & " Where strUserId='" & strUserId & "'"
		conn.Execute(strSql)
		'关闭对象
		conn.close
		Set conn=Nothing
		Response.Redirect "log_register3.asp"				'重定向到下一页
	End If
	%>
</body>
</html>