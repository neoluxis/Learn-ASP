<html>
<body>
	<h2 align="center">��Ӳ�������¼ʾ��</h2>
	<form name="frmInsert" method="POST" action="">
	<p align="center">���д�*�ŵı�����д
	<table border="1" width="80%" align="center">
		<tr>
			<td>����</td><td><input type="text" name="txtName" size="20">*</td>
		</tr><tr>
			<td>�Ա�</td><td><input type="radio" name="rdoSex" value="��">��
			<input type="radio" name="rdoSex" value="Ů">Ů</td>
		</tr><tr>
			<td>����</td><td><input type="text" name="txtAge" size="4"></td>
		</tr><tr>
			<td>�绰</td><td><input type="text" name="txtTel" size="20">*</td>
		</tr><tr>
			<td>E-mail</td><td><input type="text" name="txtEmail" size="50"></td>
		</tr><tr>
			<td>���˼��</td><td><textarea name="txtIntro" rows="4" cols="50"></textarea></td>
		</tr><tr>
			<td></td><td><input type="submit" name="btnSubmit" value=" ȷ �� "></td>
		</tr>
	</table>
	</form>
	<% 
	'ֻҪ����������͵绰������Ӽ�¼
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then
		'�����������ݿ⣬����һ��Connection����ʵ��conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn 
		'���½���Recordset����ʵ��rs
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		strSql ="Select * From tbAddress"						
		rs.Open strSql,conn,1,2							'ע���������������ʾ����ָ��Ϳ��Ը���
		'������Ӽ�¼--------------------------------------------------------------------------
		rs.AddNew										
		'���¼����ֶο϶���ֵ������һ��Ҫ���
		rs("strName")=Request.Form("txtName")			
		rs("strTel")=Request.Form("txtTel")				
		rs("dtmSubmit")=Date()												
		'���¼����ֶν������Ƿ��ύ�������Ƿ����
		If Request.Form("rdoSex")<>"" Then rs("strSex")=Request.Form("rdoSex")
		If Request.Form("txtAge")<>"" Then rs("intAge")=Request.Form("txtAge")
		If Request.Form("txtEmail")<>"" Then rs("strEmail")=Request.Form("txtEmail")
		If Request.Form("txtIntro")<>"" Then rs("strIntro")=Request.Form("txtIntro")
		rs.Update										'�������ݿ�----------------------------
		Response.Write "<p align='center'>�Ѿ��ɹ���Ӽ�¼"
	End If
	%>
</body> 
</html> 
