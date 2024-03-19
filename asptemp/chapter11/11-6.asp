<html>
<body>
	<h2 align="center">文件上传结果</h2>
	<%
	Dim upload											'声明一个upload对象实例
	Set upload=Server.CreateObject("Persits.Upload")
	upload.OverwriteFiles=False							'False表示文件不可以覆盖
	upload.Save											'上传文件到内存中
	'先建立一个上传文件对象，后面可以直接引用该对象
	Dim fle
	Set fle=upload.Files("fleUpload")
	'下面判断文件是否已经存在
	If upload.FileExists(Server.MapPath("upfiles" & "/" & fle.FileName)) Then
		Response.Write "文件已经存在，请返回<a href='11-5.asp'>重新上传</a>"
	Else
		'从内存中保存到文件夹下
		fle.SaveAs Server.MapPath("upfiles" & "/" & fle.FileName)	
		Response.Write "文件路径：" & fle.Path
	End If
	%>
</body> 
</html>


