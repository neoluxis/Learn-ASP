<html>
<body>
	<h2 align="center">��ϸҳ��</h2>
	<% 
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½���һ��RecordSet����ʵ��rs��ע��Ҫ�õ����ݹ�����IDֵ
	Dim rs,strSql 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID") 	        
	Set rs=conn.Execute(strSql)
	'������ʾ���ҵ�����
	Response.Write "<br>������" & rs("strName")
	Response.Write "<br>�Ա�" & rs("strSex")
	Response.Write "<br>���䣺" & rs("intAge")
	Response.Write "<br>�绰��" & rs("strTel")
	Response.Write "<br>E-mail��" & rs("strEmail")
	Response.Write "<br>��飺" & rs("strIntro")
	Response.Write "<br>������ڣ�" & rs("dtmSubmit")
	%>
	<p align="center"><a href="#" onclick="window.close()">�رմ���</a>
</body> 
</html> 