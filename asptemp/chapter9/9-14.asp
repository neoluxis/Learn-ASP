<html>
<body>
	<h2 align="center">�Զ���������ϲ�ѯʾ��</h2>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("userinfo.mdb")
	conn.Open strConn 
	'���½�����¼������ʵ��rs
	Dim strSql,rs                                       
	strSql="Select tbUsers.strUserId,tbUsers.strName,tbLog.strIP,tbLog.dtmLog From tbUsers,tbLog Where tbUsers.strUserId=tbLog.strUserId Order By tbLog.dtmLog DESC"   
	Set rs=conn.Execute(strSql)
	'��������ѭ����ʾ��¼��������ʾ�ڱ����
	Response.Write "<table border='1' width='100%'><tr bgcolor='#E0E0E0'>"
	Response.Write "<th>�û���</th><th>����</th><th>��¼IP</th><th>��¼ʱ��</th></tr>"
	Do While Not rs.Eof                            
		Response.Write "<tr>"
		Response.Write "<td>" & rs("strUserId") & "</td>"
		Response.Write "<td>" & rs("strName") & "</td>"
		Response.Write "<td>" & rs("strIP") & "</td>"
		Response.Write "<td>" & rs("dtmLog") & "</td>"
		Response.Write "</tr>"
		rs.MoveNext                        
	Loop
	Response.Write "</table>"
	%>
</body> 
</html>

