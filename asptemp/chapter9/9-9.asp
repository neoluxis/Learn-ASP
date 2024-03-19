<html>
<body>
<h2 align="center">参数查询示例</h2>
<form name="frmSearch" method="POST" action="">
	请输入要查找的姓名：<input type="text" name="txtName">
	<input type="submit" name="btnSubmit" value=" 确 定 ">
</form>
<%
If Request.Form("txtName")<>"" Then
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'以下建立Command对象实例cmd
	Dim cmd
	Set cmd= Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn						'指定Connection对象
	'下面建立一个参数对象prm
	Dim prmName,prmType,prmDirection,prmSize,prmValue
	prmName="varName"							'参数名称，在qryList2中的变量
	prmType=200									'参数类型，200表示是变长字符串
	prmDirection=1								'参数方向，1表示输入
	prmSize=10									'参数大小，最大字节数为10
	prmValue=Request.Form("txtName")			'要传入的参数值
	Dim prm										'声明一个参数对象变量
	Set prm=cmd.CreateParameter(prmName,prmType,prmDirection,prmSize,prmValue)
	'下面将该参数对象加入到参数集合中
	cmd.Parameters.Append(prm)					
	'下面执行查询qryList2
	Dim rs
	cmd.CommandType=4							'表示查询指令是查询名
	cmd.CommandText="qryList2"					'指定查询名称
	Set rs=cmd.Execute							'执行查询
	'以下利用循环简单列出人员姓名和电话
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
End If
%>
</body>
</html>
