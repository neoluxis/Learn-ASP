<%
'===============================================================================================
'insert.asp˵��
'1. ��ҳ��ʵ��������9.4.4����Ӳ������ļ�¼ʾ���Ļ����ϼ��޸Ķ��ɵġ�
'2. ʵ��������Ҳ����ʹ��8.3.2�ڽ����ķ�����Ӽ�¼��
'3. ��ҳ���е����ݿ��������Ҳ������connection.asp��
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>�����Ա</title>
</head>
<html>
<body>
	<h2 align="center">�����Ա</h2>

	<!--  �������ύ��Ϣ�õı�   -->
	<form name="frmInsert" method="POST" action="">
		<p align="center">���д�*�ŵı�����д
		<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
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
				<td>&nbsp;</td><td><input type="submit" name="btnSubmit" value=" ȷ �� "></td>
			</tr>
		</table>
		<p align='center'><a href="index.asp">������ҳ</a>
	</form>

	<%
	'ֻҪ����������͵绰������Ӽ�¼
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then

		'���½���Recordset����ʵ��rs-------------------------------------------------------------------------
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		strSql ="Select * From tbAddress"						
		rs.Open strSql,conn,1,2							'ע���������������ʾ����ָ��Ϳ��Ը���

		'������Ӽ�¼----------------------------------------------------------------------------------------
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
		rs.Update										'�������ݿ�

		'��ӳɹ����ض������ҳ----------------------------------------------------------------------------
		Response.Redirect "index.asp"				

	End If
	%>
</body> 
</html> 
