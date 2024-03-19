<%
'===============================================================================================
'update.asp说明
'本页面实际上是在8.3.4节更新记录示例的基础上修改的， 只不过将两个文件合并成了一个文件。 不过合并成一个文件后， 理解难度有所增加。 请大家注意以下几点：

'1. 当从首页index.asp（或search.asp）单击"更新"超链接时， 就会从头到尾执行本页面一次。 此时首先利用Request.QueryString("ID")获取传递过来的ID值， 然后从数据库中查找相应的记录， 并显示在表单中。 要注意这里也将ID值保存在了隐藏文本框中。
'因为此时还没有提交表单，所以后面的更新部分的条件不成立，因此不会执行。

'2.  当用户提交表单后， 会再次从头到尾执行一次本页面， 此时更新部分的条件成立了， 所以会执行后面的更新记录部分。 注意， 此时会从隐藏文本框中获取ID值。

'3. 细心的同学可能会想到， 当提交表单后， 前面的表单部分是否也会再执行一次呢？ 确实是， 不过因为后面用Response.Redirect重定向到首页了， 所以表单部分执行也没有意义了。 事实上这里还隐藏着ASP的一个"秘密"， 当利用action=""向自身提交表单时， 查询字符串中的信息还自动保留在URL后面， 也会再次提交给自身。 也就是说，第二次执行本文件时仍然可以用Request.QueryString方法获取ID值， 这样其实就不需要用隐藏文本框传递ID值。 不过这种方法容易混淆， 建议大家用示例中的方法。

'4. 本页面中的数据库连接语句也放在了connection.asp中
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>更新成员信息</title>
</head>
<body>
	<h2 align="center">更新成员信息</h2>
	<%
	'以下建立一个记录集对象实例rs，注意会用到从首页index.asp或search.asp传过来的ID
	Dim strSql,rs 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID")  
	Set rs=conn.Execute(strSql)
	%>
	<!-- 下面将记录中已有的数据显示在表单中，要特别注意action属性  -->
	<form name="frmUpdate" method="POST" action="">
		<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td>姓名</td>
				<td><input type="text" name="txtName" size="20" value="<%=rs("strName")%>">*</td>
			</tr><tr>
				<td>性别</td><td>
				<input type="radio" name="rdoSex" value="男" <%If rs("strSex")="男" Then Response.Write "checked"%>>男
				<input type="radio" name="rdoSex" value="女" <%If rs("strSex")="女" Then Response.Write "checked"%>>女</td>
			</tr><tr>
				<td>年龄</td>
				<td><input type="text" name="txtAge" size="4" value="<%=rs("intAge")%>"></td>
			</tr><tr>
				<td>电话</td>
				<td><input type="text" name="txtTel" size="20" value="<%=rs("strTel")%>">*</td>
			</tr><tr>
				<td>E-mail</td>
				<td><input type="text" name="txtEmail" size="50" value="<%=rs("strEmail")%>"></td>
			</tr><tr>
				<td>个人简介</td>
				<td><textarea name="txtIntro" rows="4" cols="50"><%=rs("strIntro")%></textarea></td>
			</tr><tr>
				<td><input type="hidden" name="txtID" value="<%=rs("ID")%>">&nbsp;</td>
				<td><input type="submit" name="btnSubmit" value=" 确 定 "></td>
			</tr>
		</table>
		<p align='center'><a href="index.asp">返回首页</a>
	</form>

	<%
	'当用户提交表单后，下面部分就会更新数据--------------------------------------------------------

	'只要有姓名和电话，就更新记录，否则给出错误信息
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then

		'以下获取提交的表单数据
		Dim ID,strName,strSex,intAge,strTel,strEmail,strIntro		'声明几个变量
		ID=Request.Form("txtID")									'获取隐藏文本框传递过来的ID值
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

		'下面利用Execute方法更新记录，注意更新条件为隐藏文本框传过来的ID值
		strSql="Update tbAddress Set strName='" & strName & "',strSex='" & strSex & "',intAge=" & intAge & ",strTel='" & strTel & "',strEmail='" & strEmail & "',strIntro='" & strIntro & "' Where ID=" & ID
		conn.Execute(strSql)
		
		'更新完毕，重定向回首页----------------------------------------------------------------------
		Response.Redirect "index.asp"    
	End If
	%>
</body> 
</html>
