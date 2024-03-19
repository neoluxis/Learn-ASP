<html>
<body>
	<h2 align="center">更新成员信息</h2>
	<%
	'以下连接数据库，建立一个Connection对象实例rs
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn
	'以下建立一个记录集对象实例rs，注意会用到传过来的ID
	Dim strSql,rs 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID")  
	Set rs=conn.Execute(strSql)
	'下面将记录的数据显示在表单中
	%>
	<form name="frmUpdate" method="POST" action="update.asp" >
	<table border="1" width="80%" align="center">
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
			<td><input type="hidden" name="txtID" value="<%=rs("ID")%>"></td>
			<td><input type="submit" name="btnSubmit" value=" 确 定 "></td>
		</tr>
	</table>
	</form>
</body> 
</html>
