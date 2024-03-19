<html>
<body>
	<% 
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	conn.Open "address"                                
	'以下利用Connection对象的Execute方法更新记录
	conn.Execute("Update tbAddress Set intAge=22,strTel='61223211' Where ID=2")
	Response.Write "已经成功更新，你可以自己打开address.mdb查看结果。"
	%>
</body> 
</html> 
