<html>
<body>
	<h2 align="center">在线发送附件</h2>
	<%
	'下面首先保存上传的文件
	Dim upload													'声明upload对象实例
	Set upload=Server.CreateObject("Persits.Upload")     
	upload.Save "C:\Inetpub\wwwroot\asptemp\chapter11\upfiles"	'上传到指定文件夹
	'下面开始发送邮件
	Dim jmail													'声明Message对象实例
	Set jmail=Server.CreateObject("Jmail.Message")    
	jmail.From="aspbook@163.com"								'发件人
	jmail.AddRecipient upload.Form("txtRecipient")				'收件人 
	jmail.Subject=upload.Form("txtSubject")						'主题
	jmail.Body=upload.Form("txtBody")							'内容
	'下面判断一下，如果确实上传了文件，就发送附件
	If upload.Files.Count>0 Then
		jmail.AddAttachment upload.Files("fleUpload").Path		'添加附件
	End If
	jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"					'执行发送
	jmail.Close                                                  
	Response.Write "已经成功发送"
	%>
</body> 
</html>



