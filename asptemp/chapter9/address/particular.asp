<%
'===============================================================================================
'particular.asp˵��
'1. ��ҳ���9.2.5�ڵ���ϸҳ��ʾ����������һ���ģ�ֻ�Ǽ򵥵���ʾһ����Ա����ϸ��Ϣ��
'2. Ϊ�����ۣ����ｫ��Ϣ��ʾ���˱���С�
'3. �ر��йض�������Ҳ����ʡ�Ե�����Ϊִ���걾ҳ���Ҳ���Զ��رյġ�
'4. ��ҳ���е����ݿ��������Ҳ������connection.asp��
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>��Ա��ϸ��Ϣ</title>
<head>
</body>
	<h2 align="center">��Ա��ϸ��Ϣ</h2>

	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<% 
		'���½���һ��RecordSet����ʵ��rs��ע��Ҫ�õ����ݹ�����IDֵ
		Dim rs,strSql 
		strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID") 	        
		Set rs=conn.Execute(strSql)

		'������ʾ���ҵ�����
		Response.Write "<tr><td width='30%' height='25'>����</td><td>" & rs("strName") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>�Ա�</td><td>" & rs("strSex") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>����</td><td>" & rs("intAge") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>�绰</td><td>" & rs("strTel") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>E-mail</td><td>" & rs("strEmail") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>���</td><td>" & rs("strIntro") & "</td></tr>"
		Response.Write "<tr><td width='30%' height='25'>�������</td><td>" & rs("dtmSubmit") & "</td></tr>"

		'���ڹر��йض���
		rs.Close
		Set rs=Nothing
		conn.Close
		Set conn=Nothing
		%>
	</table>

	<p align="center"><a href="#" onclick="window.close()">�رմ���</a>
</body> 
</html> 