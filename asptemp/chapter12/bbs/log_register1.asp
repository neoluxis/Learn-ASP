<%
'=========================================================================================
'�û�ע��ҳ---��һ��
'1. ��һ��Ҫ���û������û��������롣
'2. ����û�������ʹ�ã�����ӵ����ݿ��У�������һ����Ҫע���ʱ�Ὣ�û������浽Session�У��Ա���һ��ҳ��ʹ�á�
'3. ����û����Ѿ���ʹ�ã��������û��޸��û�����
'=========================================================================================
%>

<% Option Explicit %>
<!--#INCLUDE FILE="odbc_connection.asp"-->
<html>
<head>
	<title>ע�᣺��һ��</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="javascript">
		<!--
		function check_Null(){
			if (document.frmReg.txtUserId.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			if (document.frmReg.txtUserId.value.length<4 || document.frmReg.txtUserId.value.length>20){
				alert("�û�����������4���ַ���Ҳ���ܶ���20���ַ�");
				return false;
			}
			if (document.frmReg.txtPwd.value==""){
				alert("���벻��Ϊ��!");
				return false;
			}
			if (document.frmReg.txtPwd.value!=document.frmReg.txtPwd2.value){
				alert("���������ȷ�ϱ���һ��!");
				return false;
			}
			return true;
			}
		// -->
		</script>
</head>
<body>
	<h2 align="center">�û�ע��</h2>
	<p align="center">��һ�� �����û�����ע�⣺���д�*�ŵ���Ŀ������д�� 

	<form name="frmReg" action="" method="POST" onsubmit="javascript: return check_Null();">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%">�û���</td>
			<td width="80%"><input size="15" name="txtUserId">*�����ڻ����4λ��С��20λ��</td>
		</tr>
		<tr height="25"> 
			<td>�� ��</td>
			<td><input type="password" size="15" name="txtPwd">*</td>
		</tr>
		<tr height="25"> 
			<td>ȷ������</td>
			<td><input type="password" size="15" name="txtPwd2">*</td>
		</tr>
	</table>
	<br>
	<input type=submit value=" ȷ �� " name="submit">
	</form>
	<br>
	<%
	'������֤��ȷ������ɼ���ע�ᣬ���򷵻�
	If Request("txtUserId")<>"" Then
		'�������Ȼ�ȡ�ύ���û���������
		Dim strUserId,strPwd
		strUserId=Request.Form("txtUserId")
		strPwd=Request.Form("txtPwd")
		'���¼����û��Ƿ��Ѿ����ڣ�����ڣ�����Ҫ�����û���
		Dim strSql,rs
		strSql="Select * From tbUsers Where strUserId='" & Request.Form("txtUserId") & "'"
		Set rs=conn.execute(strSql)
		If Not rs.Eof And Not rs.Bof Then
			Response.Write "<p align='center'>��ʾ��������ʹ�ø��û���,��������д</p>"
		Else
			strSql="Insert Into tbUsers(strUserId,strPwd) Values('" & strUserId & "','" & strPwd & "')" 
			conn.execute(strSql)
			Session("strUserId")=strUserId				'��ס�û������Ա�����ʹ�á�
			Response.Redirect "log_register2.asp"		'�ض�����һ��ҳ��
		End If
	End If
	%>
</body>
</html>