<% Response.Buffer=True								'ע�⣬����и����%>
<html>
<body>
	<form name="frmTest" method="POST" action="">
		��ѡ���û����ͣ�
		<input type="radio" name="rdoUserType" value="teacher">��ʦ
		<input type="radio" name="rdoUserType" value="student">ѧ��
		<input type="submit" name="btnSubmit" value="ȷ��">
	</form>
	<% 
	If Request.Form("rdoUserType")="teacher" Then
		Response.Redirect "teacher.asp"     		'����ʦ�û���������ʦ��ҳ
	Elseif Request.Form("rdoUserType")="student" Then
		Response.Redirect "student.asp"     		'��ѧ���û�������ѧ����ҳ
	End If
	%>
</body>
</html>


