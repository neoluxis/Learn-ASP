<html>
<body>
	<h2 align="center">添加新成员</h2>
	<form name="frmInsert" method="POST" action="">
	<p align="center">其中带*号的必须填写
	<table border="1" width="80%" align="center">
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
			<td></td><td><input type="submit" name="btnSubmit" value=" 确 定 "></td>
		</tr>
	</table>
	</form>
	<% 
	'只要添加了姓名和电话，就添加记录
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then
		 '以下首先获取提交的数据
		Dim strName,strSex,intAge,strTel,strEmail,strIntro			'声明几个变量
		strName=Request.Form("txtName")								'获取姓名
		strSex=Request.Form("rdoSex")								'获取性别
		If Request.Form("txtAge")<>"" Then							'获取年龄
			intAge=Request.Form("txtAge")
		Else
			intAge=0
		End If
		strTel=Request.Form("txtTel")								'获取电话
		strEmail=Request.Form("txtEmail")							'获取E-mail
		strIntro=Request.Form("txtIntro")							'获取个人简介
		'以下连接数据库，建立一个Connection对象实例conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection") 
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn
		'下面利用Execute方法添加记录
		Dim strSql
		strSql="Insert Into tbAddress(strName,strSex,intAge,strTel,strEmail,strIntro,dtmSubmit) Values('" & strName & "','" & strSex & "'," & intAge & ",'" & strTel & "','" & strEmail & "','" & strIntro & "',#" & Date() & "#)"
		conn.Execute(strSql)  
		'添加成功后，则返回首页
		Response.Redirect "index.asp"                     
	End If
	%>
</body> 
</html> 
