<%
'===============================================================================================
'update.asp˵��
'��ҳ��ʵ��������8.3.4�ڸ��¼�¼ʾ���Ļ������޸ĵģ� ֻ�����������ļ��ϲ�����һ���ļ��� �����ϲ���һ���ļ��� ����Ѷ��������ӡ� ����ע�����¼��㣺

'1. ������ҳindex.asp����search.asp������"����"������ʱ�� �ͻ��ͷ��βִ�б�ҳ��һ�Ρ� ��ʱ��������Request.QueryString("ID")��ȡ���ݹ�����IDֵ�� Ȼ������ݿ��в�����Ӧ�ļ�¼�� ����ʾ�ڱ��С� Ҫע������Ҳ��IDֵ�������������ı����С�
'��Ϊ��ʱ��û���ύ�������Ժ���ĸ��²��ֵ���������������˲���ִ�С�

'2.  ���û��ύ���� ���ٴδ�ͷ��βִ��һ�α�ҳ�棬 ��ʱ���²��ֵ����������ˣ� ���Ի�ִ�к���ĸ��¼�¼���֡� ע�⣬ ��ʱ��������ı����л�ȡIDֵ��

'3. ϸ�ĵ�ͬѧ���ܻ��뵽�� ���ύ���� ǰ��ı������Ƿ�Ҳ����ִ��һ���أ� ȷʵ�ǣ� ������Ϊ������Response.Redirect�ض�����ҳ�ˣ� ���Ա�����ִ��Ҳû�������ˡ� ��ʵ�����ﻹ������ASP��һ��"����"�� ������action=""�������ύ��ʱ�� ��ѯ�ַ����е���Ϣ���Զ�������URL���棬 Ҳ���ٴ��ύ������ Ҳ����˵���ڶ���ִ�б��ļ�ʱ��Ȼ������Request.QueryString������ȡIDֵ�� ������ʵ�Ͳ���Ҫ�������ı��򴫵�IDֵ�� �������ַ������׻����� ��������ʾ���еķ�����

'4. ��ҳ���е����ݿ��������Ҳ������connection.asp��
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>���³�Ա��Ϣ</title>
</head>
<body>
	<h2 align="center">���³�Ա��Ϣ</h2>
	<%
	'���½���һ����¼������ʵ��rs��ע����õ�����ҳindex.asp��search.asp��������ID
	Dim strSql,rs 
	strSql="Select * From tbAddress Where ID=" & Request.QueryString("ID")  
	Set rs=conn.Execute(strSql)
	%>
	<!-- ���潫��¼�����е�������ʾ�ڱ��У�Ҫ�ر�ע��action����  -->
	<form name="frmUpdate" method="POST" action="">
		<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
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
				<td><input type="hidden" name="txtID" value="<%=rs("ID")%>">&nbsp;</td>
				<td><input type="submit" name="btnSubmit" value=" ȷ �� "></td>
			</tr>
		</table>
		<p align='center'><a href="index.asp">������ҳ</a>
	</form>

	<%
	'���û��ύ�������沿�־ͻ��������--------------------------------------------------------

	'ֻҪ�������͵绰���͸��¼�¼���������������Ϣ
	If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then

		'���»�ȡ�ύ�ı�����
		Dim ID,strName,strSex,intAge,strTel,strEmail,strIntro		'������������
		ID=Request.Form("txtID")									'��ȡ�����ı��򴫵ݹ�����IDֵ
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

		'��������Execute�������¼�¼��ע���������Ϊ�����ı��򴫹�����IDֵ
		strSql="Update tbAddress Set strName='" & strName & "',strSex='" & strSex & "',intAge=" & intAge & ",strTel='" & strTel & "',strEmail='" & strEmail & "',strIntro='" & strIntro & "' Where ID=" & ID
		conn.Execute(strSql)
		
		'������ϣ��ض������ҳ----------------------------------------------------------------------
		Response.Redirect "index.asp"    
	End If
	%>
</body> 
</html>
