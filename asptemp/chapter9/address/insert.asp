<%
'===============================================================================================
'insert.asp说明
'1. 本页面实际上是在9.4.4节添加不完整的记录示例的基础上简单修改而成的。
'2. 实际上这里也可以使用8.3.2节讲述的方法添加记录。
'3. 本页面中的数据库连接语句也放在了connection.asp中
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>添加人员</title>
</head>
<html>
<body>
	<h2 align="center">添加人员</h2>

	<!--  下面是提交信息用的表单   -->
	<form name="frmInsert" method="POST" action="">
		<p align="center">其中带*号的必须填写
		<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td>姓名</td><td><input type="text" name="txtName" size="20">*</td>
			</tr><tr>
				<td>性别</td><td><input type="radio" name="rdoSex" value="男">男
				<input type="radio" name="rdoSex" value="女">女</td>
			</tr><tr>
				<td>年龄</td><td><input type="text" name="txtAge" size="4"></td>
			</tr><tr>
				<td>电话</td><td><input type="text" name="txtTel" size="20">*</td>
			</tr><tr>
				<td>E-mail</td><td><input type="text" name="txtEmail" size="50"></td>
			</tr><tr>
				<td>个人简介</td><td><textarea name="txtIntro" rows="4" cols="50"></textarea></td>
			</tr><tr>
				<td>&nbsp;</td><td><input type="submit" name="btnSubmit" value=" 确 定 "></td>
			</tr>
		</table>
		<p align='center'><a href="index.asp">返回首页</a>
	</form>

	<%
	'只要添加了姓名和电话，就添加记录
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then

		'以下建立Recordset对象实例rs-------------------------------------------------------------------------
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		strSql ="Select * From tbAddress"						
		rs.Open strSql,conn,1,2							'注意后两个参数，表示键盘指针和可以更新

		'以下添加记录----------------------------------------------------------------------------------------
		rs.AddNew										
		'以下几个字段肯定有值，所以一定要添加
		rs("strName")=Request.Form("txtName")			
		rs("strTel")=Request.Form("txtTel")				
		rs("dtmSubmit")=Date()												
		'以下几个字段将根据是否提交而决定是否添加
		If Request.Form("rdoSex")<>"" Then rs("strSex")=Request.Form("rdoSex")
		If Request.Form("txtAge")<>"" Then rs("intAge")=Request.Form("txtAge")
		If Request.Form("txtEmail")<>"" Then rs("strEmail")=Request.Form("txtEmail")
		If Request.Form("txtIntro")<>"" Then rs("strIntro")=Request.Form("txtIntro")
		rs.Update										'更新数据库

		'添加成功后，重定向回首页----------------------------------------------------------------------------
		Response.Redirect "index.asp"				

	End If
	%>
</body> 
</html> 
