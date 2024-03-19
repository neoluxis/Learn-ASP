<html>
<body>
	<h2 align="center">文件已安全上传</h2>
	<%
	Dim upload														'声明一个upload对象实例
	Set upload=Server.CreateObject("Persits.Upload")
	upload.SetMaxSize 10*1024*1024,True								'限制文件不超过10MB，否则报错
	upload.OverWriteFiles=False										'False表示文件不可以被覆盖
	upload.Save "C:\Inetpub\wwwroot\asptemp\chapter11\upfiles"		'上传到指定文件夹
	'下面判断一下，如果确实上传了文件，则输出文件信息
	If upload.Files.Count>0 Then
		Response.Write "<br>上传文件：" & upload.Files("fleUpload").Path 
		Response.Write "<br>文件名：" & upload.Files("fleUpload").FileName 
		Response.Write "<br>扩展名：" & upload.Files("fleUpload").Ext
		Response.Write "<br>文件大小：" & upload.Files("fleUpload").Size & "字节"
		Response.Write "<br>最后修改时间：" & upload.Files("fleUpload").LastWriteTime
	End If
	'下面输出表单元素信息
	Response.Write "<br>文件说明：" & upload.Form("txtIntro").Value 
	Response.Write "<br>作者姓名：" & upload.Form("txtAuthor").Value 
	%>
</body> 
</html>


