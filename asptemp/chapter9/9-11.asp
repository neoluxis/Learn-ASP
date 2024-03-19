<html>
<body>
	<h2 align="center">添加不完整记录示例</h2>
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
		'以下连接数据库，建立一个Connection对象实例conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn 
		'以下建立Recordset对象实例rs
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		strSql ="Select * From tbAddress"						
		rs.Open strSql,conn,1,2							'注意后两个参数，表示键盘指针和可以更新
		'以下添加记录--------------------------------------------------------------------------
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
		rs.Update										'更新数据库----------------------------
		Response.Write "<p align='center'>已经成功添加记录"
	End If
	%>
</body> 
</html> 
