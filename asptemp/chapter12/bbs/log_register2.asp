<%
'=========================================================================================
'�û�ע��ҳ---�ڶ���
'1. ��һ��Ҫ�û���д��ϸ��Ϣ����ʵ���Ǹ��¼�¼ҳ�档
'2. ��Ϊ�е���Ϣ����ʡ�ԣ����������SQL�ַ����Ƚϸ��ӡ����⣬�����õ���Update��䣬���Ҳ�����һ����׼��SQL��䡣
'3. ����û����Ѿ���ʹ�ã��������û��޸��û�����
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>ע�᣺�ڶ���</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmReg.txtName.value==""){
				alert("��ʵ��������Ϊ��!");
				return false;
			}
			if (document.frmReg.txtEmail.value==""){
				alert("E-mail����Ϊ��!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">�û�ע��</h2>
	<p align="center">�ڶ��� ��д��ϸ��Ϣ ��ע�⣺���д�*�ŵ���Ŀ������д�� 

	<form name="frmReg" onsubmit="javascript: return check_Null();" action="" method="POST">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%" >�û���</td>
			<td width="80%"><%=Session("strUserId")%></td>
		</tr>
		<tr height="25"> 
			<td>��ʵ����</td>
			<td><input type="text" name="txtName" size="15">*</td>
		</tr>
		<tr height="25"> 
			<td>�Ա�</td>
			<td><input type="radio" name="rdoSex" value="��" checked>��
				<input type="radio" name="rdoSex" value="Ů">Ů</td>
		</tr>
		<tr height="25"> 
			<td>�绰 </td>
			<td><input type="text" name="txtTel" size="25"></td>
		</tr>
		<tr height="25"> 
			<td>E-mail </td>
			<td><input type="text" name="txtEmail" size="40">*</td>
		</tr>
		<tr height="25"> 
			<td>QQ���� </td>
			<td><input type="text" name="txtQQ" size="15"></td>
		</tr>
		<tr height="25"> 
			<td>���˼��</td>
			<td><textarea name="txtIntro" rows="4" cols="50"></textarea></td>
		</tr>
	</table>
	<p align="center"><input type="submit" value="  ȷ����  ">
	</form>

	<%
	If Request.Form("txtName")<>"" And Request.Form("txtEmail")<>"" Then
		Dim strUserId,strSql
		strUserId=Session("strUserId")						'��Session�л�ȡ�û���
		'��Ϊ�е��ֶ���������Եģ���������Ҫ�������������SQL���
		'���ȸ��¿϶��е��ֶ�ֵ������4��Ҳ����д��һ���У�����д��Ϊ�˿��ø����
		strSql="Update tbUsers Set strName='" & Request.Form("txtName") & "'"
		strSql=strSql & ",strEmail='" & Request.Form("txtEmail") & "'"
		strSql=strSql & ",strSex='" & Request.Form("rdoSex") & "'"
		strSql=strSql & ",dtmSubmit=#" & Date() & "#"
		'��������Ƿ��ύ����Ϣ���Ӷ�������Ӧ�ֶ�ֵ
		If Request.Form("txtQQ") <> "" Then
			strSql=strSql & ",strQQ='" & Request.Form("txtQQ") & "'"
		End If
		If Request.Form("txtTel") <> "" Then
			strSql=strSql & ",strTel='" & Request.Form("txtTel") & "'"
		End If
		If Request.Form("txtIntro") <> "" Then
			strSql=strSql & ",strIntro='" & Request.Form("txtIntro") & "'"
		End If
		strSql=strSql & " Where strUserId='" & strUserId & "'"
		conn.Execute(strSql)
		'�رն���
		conn.close
		Set conn=Nothing
		Response.Redirect "log_register3.asp"				'�ض�����һҳ
	End If
	%>
</body>
</html>