<%
'=========================================================================================
'用户修改密码页
'1. 首先判断用户输入的旧密码是否正确？如果正确，就更新为新密码，如果不正确，就提醒重新填写。
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>修改个人信息</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmPwd.txtOldPwd.value==""){
				alert("旧密码不能为空!");
				return false;
			}
			if (document.frmPwd.txtNewPwd.value==""){
				alert("新密码不能为空!");
				return false;
			}
			if (document.frmPwd.txtNewPwd.value!=document.frmPwd.txtNewPwd2.value){
				alert("新密码和确认密码必须一致!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">修改密码</h2>
	<form name="frmPwd" onsubmit="javascript: return check_Null();" action="" method="post">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr bgcolor="#FFFFFF"> 
			<td height="25">旧 密 码</td>
			<td><input type="password" name="txtOldPwd" size="15">*</td>
		</tr>
		<tr bgcolor="#FFFFFF"> 
			<td height="25">新 密 码</td>
			<td><input type="password" name="txtNewPwd" size="15">*</td>
		</tr>
		<tr bgcolor="#FFFFFF"> 
			<td height="25">确认密码</td>
			<td><input type="password" name="txtNewPwd2" size="15">*</td>
		</tr>
	</table>
	<p align="center">
	<input type="submit" value="  确　定  " name="submit" class="inputbutton">
	</form>
	<%
	'如果提交了表单，就执行下面更新操作
	If Request.Form("txtOldPwd")<>"" And Request.Form("txtNewPwd")<>"" Then
		'下面先判断旧密码是否正确
		Dim strSql,rs
		strSql="Select strPwd From tbUsers Where strUserId='" & Session("strUserId") & "'"
		Set rs=conn.Execute(strSql)
		If rs("strPwd")<>Request.Form("txtOldPwd") Then
			Response.Write "<p align='center'>对不起，旧密码不正确，请重新输入！"
		Else
			'下面更新密码
			strSql="Update tbUsers Set strPwd='" & Request.Form("txtNewPwd") & "' Where strUserId='" & Session("strUserId") & "'"
			conn.Execute(strSql)
			'关闭对象
			conn.close
			Set conn=Nothing
			'重定向到首页
			Response.Redirect "index.asp"
		End If
	End If
	%>
</body>
</html>