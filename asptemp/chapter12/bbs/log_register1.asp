<%
'=========================================================================================
'用户注册页---第一步
'1. 这一步要让用户输入用户名和密码。
'2. 如果用户名可以使用，就添加到数据库中，继续下一步。要注意此时会将用户名保存到Session中，以备下一个页面使用。
'3. 如果用户名已经被使用，就提醒用户修改用户名。
'=========================================================================================
%>

<% Option Explicit %>
<!--#INCLUDE FILE="odbc_connection.asp"-->
<html>
<head>
	<title>注册：第一步</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="javascript">
		<!--
		function check_Null(){
			if (document.frmReg.txtUserId.value==""){
				alert("用户名不能为空!");
				return false;
			}
			if (document.frmReg.txtUserId.value.length<4 || document.frmReg.txtUserId.value.length>20){
				alert("用户名不能少于4个字符，也不能多于20个字符");
				return false;
			}
			if (document.frmReg.txtPwd.value==""){
				alert("密码不能为空!");
				return false;
			}
			if (document.frmReg.txtPwd.value!=document.frmReg.txtPwd2.value){
				alert("密码和密码确认必须一致!");
				return false;
			}
			return true;
			}
		// -->
		</script>
</head>
<body>
	<h2 align="center">用户注册</h2>
	<p align="center">第一步 申请用户名（注意：所有带*号的项目必须填写） 

	<form name="frmReg" action="" method="POST" onsubmit="javascript: return check_Null();">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%">用户名</td>
			<td width="80%"><input size="15" name="txtUserId">*（大于或等于4位，小于20位）</td>
		</tr>
		<tr height="25"> 
			<td>密 码</td>
			<td><input type="password" size="15" name="txtPwd">*</td>
		</tr>
		<tr height="25"> 
			<td>确认密码</td>
			<td><input type="password" size="15" name="txtPwd2">*</td>
		</tr>
	</table>
	<br>
	<input type=submit value=" 确 定 " name="submit">
	</form>
	<br>
	<%
	'各项验证正确无误，则可继续注册，否则返回
	If Request("txtUserId")<>"" Then
		'下面首先获取提交的用户名和密码
		Dim strUserId,strPwd
		strUserId=Request.Form("txtUserId")
		strPwd=Request.Form("txtPwd")
		'以下检查该用户是否已经存在，如存在，则需要更换用户名
		Dim strSql,rs
		strSql="Select * From tbUsers Where strUserId='" & Request.Form("txtUserId") & "'"
		Set rs=conn.execute(strSql)
		If Not rs.Eof And Not rs.Bof Then
			Response.Write "<p align='center'>提示：已有人使用该用户名,请重新填写</p>"
		Else
			strSql="Insert Into tbUsers(strUserId,strPwd) Values('" & strUserId & "','" & strPwd & "')" 
			conn.execute(strSql)
			Session("strUserId")=strUserId				'记住用户名，以备后面使用。
			Response.Redirect "log_register2.asp"		'重定向到下一个页面
		End If
	End If
	%>
</body>
</html>