<html>
<body>
<h2 align="center">��ҳ��ʾ����ʾ��</h2>
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
rs.Open strSql,conn,1									'ע���������Ϊ����ָ��
'�����¼�����ǿյģ���ִ�з�ҳ��ʾ
If Not rs.Bof And Not rs.Eof Then
	'-------------------------------------------------------------------------------
	'����һ���жϵ�ǰ��ʾ�ڼ�ҳ������ǵ�һ�δ򿪣�Ϊ1�������ɴ��ز�������
	Dim intPage                            
	If Request.QueryString("varPage")="" Then   
		intPage=1 
	Else
		intPage=CInt(Request.QueryString("varPage"))	'��CIntת��Ϊ����
	End If
	'-------------------------------------------------------------------------------
	'����һ�ο�ʼ��ҳ��ʾ��ָ��Ҫ��ʾ��ҳ��Ȼ��������ʾ��ǰҳ�����м�¼��
	rs.PageSize=5										'����ÿҳ5����¼
	rs.AbsolutePage=intPage								'���õ�ǰ��ʾ�ڼ�ҳ
	'��������ѭ����ʾ��ǰҳ�����м�¼������ʾ�ڱ����
	Response.Write "<table border='1' width='100%'><tr bgcolor='#E0E0E0'>"
	Response.Write "<th>����</th><th>�绰</th><th>E-mail</th> <th>�������</th></tr>"
	Dim I												'����һ������I
	For I=1 To rs.PageSize
		If rs.Eof Then Exit For							'������˼�¼����β��������ѭ��
		Response.Write "<tr>"								
		Response.Write "<td>" & rs("strName") & "</td>"
		Response.Write "<td>" & rs("strTel") & "</td>"
		Response.Write "<td>" & rs("strEmail") & "</td>"
		Response.Write "<td>" & rs("dtmSubmit") & "</td>"
		Response.Write "</tr>"
		rs.MoveNext
	Next
	Response.Write "</table>"
	'-------------------------------------------------------------------------------
	'����һ�����������1ҳ����һҳ����һҳ�����ҳ�ĳ�����
	Response.Write "<p><a href='9-12.asp?varPage=1'>��1ҳ</a>&nbsp;"
	If intPage>1 Then
		Response.Write "<a href='9-12.asp?varPage=" & (intPage-1) & "'>��һҳ</a>&nbsp;"
	Else
		Response.Write "��һҳ&nbsp;"
	End If
	If intPage<rs.PageCount Then
		Response.Write "<a href='9-12.asp?varPage=" & (intPage+1) & "'>��һҳ</a>&nbsp;"
	Else
		Response.Write "��һҳ&nbsp;"
	End If
	Response.Write "<a href='9-12.asp?varPage=" & rs.PageCount & "'>���ҳ</a>&nbsp;"
	'-------------------------------------------------------------------------------
Else
	Response.Write "�ü�¼��Ϊ��"
End If
%>
</body>
</html>
