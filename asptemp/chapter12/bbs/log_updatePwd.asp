<%
'=========================================================================================
'�û��޸�����ҳ
'1. �����ж��û�����ľ������Ƿ���ȷ�������ȷ���͸���Ϊ�����룬�������ȷ��������������д��
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>�޸ĸ�����Ϣ</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmPwd.txtOldPwd.value==""){
				alert("�����벻��Ϊ��!");
				return false;
			}
			if (document.frmPwd.txtNewPwd.value==""){
				alert("�����벻��Ϊ��!");
				return false;
			}
			if (document.frmPwd.txtNewPwd.value!=document.frmPwd.txtNewPwd2.value){
				alert("�������ȷ���������һ��!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">�޸�����</h2>
	<form name="frmPwd" onsubmit="javascript: return check_Null();" action="" method="post">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr bgcolor="#FFFFFF"> 
			<td height="25">�� �� ��</td>
			<td><input type="password" name="txtOldPwd" size="15">*</td>
		</tr>
		<tr bgcolor="#FFFFFF"> 
			<td height="25">�� �� ��</td>
			<td><input type="password" name="txtNewPwd" size="15">*</td>
		</tr>
		<tr bgcolor="#FFFFFF"> 
			<td height="25">ȷ������</td>
			<td><input type="password" name="txtNewPwd2" size="15">*</td>
		</tr>
	</table>
	<p align="center">
	<input type="submit" value="  ȷ����  " name="submit" class="inputbutton">
	</form>
	<%
	'����ύ�˱�����ִ��������²���
	If Request.Form("txtOldPwd")<>"" And Request.Form("txtNewPwd")<>"" Then
		'�������жϾ������Ƿ���ȷ
		Dim strSql,rs
		strSql="Select strPwd From tbUsers Where strUserId='" & Session("strUserId") & "'"
		Set rs=conn.Execute(strSql)
		If rs("strPwd")<>Request.Form("txtOldPwd") Then
			Response.Write "<p align='center'>�Բ��𣬾����벻��ȷ�����������룡"
		Else
			'�����������
			strSql="Update tbUsers Set strPwd='" & Request.Form("txtNewPwd") & "' Where strUserId='" & Session("strUserId") & "'"
			conn.Execute(strSql)
			'�رն���
			conn.close
			Set conn=Nothing
			'�ض�����ҳ
			Response.Redirect "index.asp"
		End If
	End If
	%>
</body>
</html>