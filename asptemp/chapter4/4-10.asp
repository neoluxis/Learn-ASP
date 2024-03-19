<% Response.Buffer=True								'注意，最好有该语句%>
<html>
<body>
	<form name="frmTest" method="POST" action="">
		请选择用户类型：
		<input type="radio" name="rdoUserType" value="teacher">教师
		<input type="radio" name="rdoUserType" value="student">学生
		<input type="submit" name="btnSubmit" value="确定">
	</form>
	<% 
	If Request.Form("rdoUserType")="teacher" Then
		Response.Redirect "teacher.asp"     		'将教师用户引导至教师网页
	Elseif Request.Form("rdoUserType")="student" Then
		Response.Redirect "student.asp"     		'将学生用户引导至学生网页
	End If
	%>
</body>
</html>


