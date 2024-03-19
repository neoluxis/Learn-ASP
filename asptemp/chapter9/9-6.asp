<html>
<body>
	<h2 align="center"> Error对象和Errors集合示例</h2>
	<%
	On Error Resume Next						'如果发生错误，跳过执行下一句
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open "Dsn=temp9"						'故意给一个错误的数据源
	'下面利用循环输出所有的错误对象的属性
	Dim I,objErr								'I是循环变量，objErr是错误对象变量
	For I =0 To conn.Errors.Count-1                                 
		Set objErr=conn.Errors.Item(I)				'建立错误对象实例objErr
		Response.Write "<br>错误编号：" & objErr.Number 
		Response.Write "<br>错误描述：" & objErr.Description 
		Response.Write "<br>错误原因：" & objErr.Source 
		Response.Write "<br>提示文字：" & objErr.HelpContext 
		Response.Write "<br>帮助文件：" & objErr.HelpFile 
		Response.Write "<br>原始错误：" & objErr.NativeError 
	Next
	%>
</body>
</html>

