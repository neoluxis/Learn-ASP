<%
Dim jmail
Set jmail=Server.CreateObject("Jmail.Message")		'建立Message对象实例
jmail.From="aspbook@163.com"						'发件人地址
jmail.FromName="李静"								'发件人姓名
jmail.AddRecipient "jjshang@263.net"				'第一个收件人
jmail.AddRecipient "jiananlearning@qq.com"				'第二个收件人
jmail.Subject="你好"								'信件主题
jmail.Body="测试一下发信组件"						'信件内容
jmail.AddAttachment "C:\Inetpub\wwwroot\asptemp\chapter11\music.mid"		'添加附件一
jmail.AddAttachment "C:\Inetpub\wwwroot\asptemp\chapter11\beauty.jpg"		'添加附件二
jmail.Charset="gb2312"								'信件编码格式
jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"			'执行发送，参数是发信服务器
jmail.Close											'发送完毕，关闭该对象
Response.Write "已经成功发送"
%>
