<html>
<body>
	<h2 align="center">���³�Ա��Ϣ</h2>
	<%
	'�����������ݿ⣬����һ��Connection����ʵ��rs
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn
	'���½���һ����¼������ʵ��rs��ע����õ���������ID
	Dim strSql,rs 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID")  
	Set rs=conn.Execute(strSql)
	'���潫��¼��������ʾ�ڱ���
	%>
	<form name="frmUpdate" method="POST" action="update.asp" >
	<table border="1" width="80%" align="center">
		<tr>
			<td>����</td>
			<td><input type="text" name="txtName" size="20" value="<%=rs("strName")%>">*</td>
		</tr><tr>
			<td>�Ա�</td><td>
			<input type="radio" name="rdoSex" value="��" <%If rs("strSex")="��" Then Response.Write "checked"%>>��
			<input type="radio" name="rdoSex" value="Ů" <%If rs("strSex")="Ů" Then Response.Write "checked"%>>Ů</td>
		</tr><tr>
			<td>����</td>
			<td><input type="text" name="txtAge" size="4" value="<%=rs("intAge")%>"></td>
		</tr><tr>
			<td>�绰</td>
			<td><input type="text" name="txtTel" size="20" value="<%=rs("strTel")%>">*</td>
		</tr><tr>
			<td>E-mail</td>
			<td><input type="text" name="txtEmail" size="50" value="<%=rs("strEmail")%>"></td>
		</tr><tr>
			<td>���˼��</td>
			<td><textarea name="txtIntro" rows="4" cols="50"><%=rs("strIntro")%></textarea></td>
		</tr><tr>
			<td><input type="hidden" name="txtID" value="<%=rs("ID")%>"></td>
			<td><input type="submit" name="btnSubmit" value=" ȷ �� "></td>
		</tr>
	</table>
	</form>
</body> 
</html>
