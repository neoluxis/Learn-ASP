<html>
<body>
	<h2 align="center">在线发送附件示例</h2>
	<form name="frmUp" action="11-10.asp" method="POST" enctype="multipart/form-data">
	<table border="0" align="center">
		<tr><td>发件人</td><td>aspbook@163.com</td></tr>
		<tr><td>收件人</td><td><input type="text" name="txtRecipient" size="40"></td></tr>
		<tr><td>主题</td><td><input type="text" name="txtSubject" size="40"></td></tr>
		<tr><td>内容</td>
		<td><textarea name="txtBody" rows="4" cols="40"></textarea></td></tr>
		<tr><td>附件</td><td><input type="file" name="fleUpload" size="40"></td></tr>
		<tr><td></td><td><input type="submit" value=" 发送 "></td></tr>
	</table>
	</form>
</body> 
</html>
