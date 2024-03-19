<html>
<body>
	<h2 align="center">事务处理用法示例</h2>
	<%
	On Error Resume Next								'如果发生错误，跳过执行下一句
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'下面开始事务处理
	conn.BeginTrans										
	'删除记录
	strSql="Delete From tbAddress Where strName='李玫'"
	conn.Execute(strSql)
	'添加记录
	strSql="Insert Into tbAddress(strName,strTel) Values('李玫','88888888')"
	conn.Execute(strSql)
	'下面根据是否发生错误来决定是否提交处理结果
	If  conn.Errors.Count=0 Then							
		conn.CommitTrans								'如果无错误，就提交事务处理结果
		Response.Write "成功执行"
	Else
		conn.RollbackTrans								'如果有错误，就取消事务处理结果
		Response.Write "有错误发生，取消处理结果"
	End If
	%>
</body>
</html>

