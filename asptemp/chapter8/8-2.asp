<html>
<body>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"									
	'以下利用Connection对象的Execute方法添加新记录
	conn.Execute("Insert Into tbAddress(strName,strSex,intAge,strTel,strEmail,strIntro,dtmSubmit) Values('萌萌','女',21,'6112211','mm@372.net','金融系同学',#2008-8-8#)")
	Response.Write "已经成功添加，你可以自己打开address.mdb查看结果。"
	%>
</body> 
</html>

