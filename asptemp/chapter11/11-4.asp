<html>
<body>
	<h2 align="center">文件已安全上传</h2>
	<%
	Dim upload											'建立upload对象实例
	Set upload=Server.CreateObject("Persits.Upload")   
	upload.Save Server.MapPath("upfiles")				'上传到指定文件夹
	Dim objItem											'声明一个对象实例变量
	For Each objItem In upload.Files					'用循环输出所有文件的信息
		Response.Write objItem.Name & "=" & objItem.Path & "<br>"
	Next
	For Each objItem In upload.Form						'用循环输出所有表单元素信息
		Response.Write objItem.Name & "=" & objItem.Value & "<br>"
	Next
	%>
</body> 
</html>
