<html>
<body>
	<h2 align="center">Fields������Field����ʾ��</h2>
	<%
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½���Recordset����ʵ��rs
	Dim rs,strSql
	Set rs=Server.CreateObject("ADODB.Recordset")
	strSql ="Select * From tbAddress"						
	rs.Open strSql,conn
	'��һ��ѭ���������ֶ�����д��
	Dim I,fld
	For I=0 to rs.Fields.Count-1
		Set fld=rs.Fields.Item(I)					'������ǰ�ֶε�Field����ʵ��fld
		Response.Write "�ֶ����ƣ�" & fld.Name & "<br>"
		Response.Write "�ֶ�ֵ��" & fld.Value & "<br>"
		Response.Write "�ֶ����ͣ�" & fld.Type & "<br>"
		Response.Write "�ֶδ�С��" & fld.Definedsize & "<br>"
		Response.Write "�ֶ����λ����" & fld.Precision & "<br>"
	Next
	%>
</body>
</html>
