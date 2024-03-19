<html>
<body>
	<h2 align="center">在线发送Email示例</h2>
	<form name="frmEmail" method="POST" action="">
	<table border="0" align="center">
		<tr><td>发件人</td><td>aspbook@163.com</td></tr>
		<tr><td>收件人</td><td><input type="text" name="txtRecipient" size="30"></td></tr>
		<tr><td>主题</td><td><input type="text" name="txtSubject" size="40"></td></tr>
		<tr><td>内容</td>
		<td><textarea name="txtBody" rows="4" cols="40"></textarea></td></tr>
		<tr><td></td><td><input type="submit" value=" 发送 "></td></tr>
	</table>
	</form>
	<%
	'判断，如果收件人和主题都不为空，则继续发信
	If Request("txtRecipient")<>"" And Request("txtSubject")<>"" Then
		Dim jmail										'声明Message对象实例
		Set jmail=Server.CreateObject("Jmail.Message")		    
		jmail.From="aspbook@163.com"					'发件人 
		jmail.AddRecipient Request("txtRecipient")		'收件人 
		jmail.Subject=Request("txtSubject")				'主题 
		jmail.Body=Request("txtBody")					'内容                         
		jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"		'执行发送
		jmail.Close										'关闭对象
		Response.Write "<p align='center'>成功发送"
	End If
	%>
</body> 
</html>


