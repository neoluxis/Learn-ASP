<%
'================================================================================================
'ɾ�������ļ�
'1. ����û�������ȷ��ɾ�����룬�Ϳ��Խ�����ɾ����
'2. ע�⣬�������Ƚ���ȡ�����ļ�¼ID��Ŵ����һ�������ı����У�Ȼ���ύ����Ϳ��Ի�ȡ��ID�� �Ӷ�ɾ���ü�¼��
'================================================================================================
%>

<% Option Explicit						'ǿ����������%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>ɾ������</title>
	<link rel="stylesheet" href="guest.css">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body >
	<h2 align="center">ɾ������</h2>
	<!-- ע�����н����ݹ�����ID��ŵ������ı�������  -->
	<form name="frmDelete" method="POST" action="">
		<center>������ɾ�����룺
		<input type="password" name="txtPwd" size="10" value="">
		<input type="hidden" name="txtID" value="<%=Request.QueryString("ID")%>">
		<input type="submit" value="���ύ��" size="20">
		</center>
	</form>
	<%
	'�����ж�һ�£��������������ļ��е�������ȣ���ɾ��������
	If Request.Form("txtPwd")=conPwd Then
		Dim strSql
		strSql="Delete From tbGuest where ID=" & Request.Form("txtID") 
		conn.Execute(strSql)
		Response.Redirect("index.asp")
	End If
	%>
</body>
</html>