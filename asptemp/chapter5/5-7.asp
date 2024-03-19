<html>
<body>
	<form name="form1" method="post" action="">
		昵称：<input type="text" name="txtName" size="10">
		发言：<input type="text" name="txtSay" size="30">
		<input type="submit" value=" 发送 ">
	</form>
	<%
	'如果提交了表单，就将发言内容添加到Application对象中
	If Trim(Request.Form("txtName"))<>"" And Trim(Request.Form("txtSay"))<>"" Then
		'下面先获取本次发言字符串，包括发言人和发言内容
		Dim strSay
		strSay=Request.Form("txtName") & "说：" & Request.Form("txtSay") & "<br>"
		'下面将本次发言添加到聊天内容中
		Application.Lock						'先锁定
		Application("strChat")=strSay & Application("strChat")
		Application.Unlock						'解除锁定
	End If
	%>
</body> 
</html>
