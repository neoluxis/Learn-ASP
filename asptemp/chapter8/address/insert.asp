<html>
<body>
	<h2 align="center">����³�Ա</h2>
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
		 '�������Ȼ�ȡ�ύ������
		Dim strName,strSex,intAge,strTel,strEmail,strIntro			'������������
		strName=Request.Form("txtName")								'��ȡ����
		strSex=Request.Form("rdoSex")								'��ȡ�Ա�
		If Request.Form("txtAge")<>"" Then							'��ȡ����
			intAge=Request.Form("txtAge")
		Else
			intAge=0
		End If
		strTel=Request.Form("txtTel")								'��ȡ�绰
		strEmail=Request.Form("txtEmail")							'��ȡE-mail
		strIntro=Request.Form("txtIntro")							'��ȡ���˼��
		'�����������ݿ⣬����һ��Connection����ʵ��conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection") 
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
		conn.Open strConn
		'��������Execute������Ӽ�¼
		Dim strSql
		strSql="Insert Into tbAddress(strName,strSex,intAge,strTel,strEmail,strIntro,dtmSubmit) Values('" & strName & "','" & strSex & "'," & intAge & ",'" & strTel & "','" & strEmail & "','" & strIntro & "',#" & Date() & "#)"
		conn.Execute(strSql)  
		'��ӳɹ����򷵻���ҳ
		Response.Redirect "index.asp"                     
	End If
	%>
</body> 
</html> 
