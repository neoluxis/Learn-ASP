<%
'=========================================================================================
'�û��޸ĸ�����Ϣҳ
'1. ����ʵ����һ����ͨ�ĸ��¼�¼ҳ�档���Ƚ�ԭ��������ʾ�ڱ��У��ύ�����ٸ��¼�¼��
'2. �ڸ��¼�¼ʱ��Щ��Ϣ����ʡ�ԣ�����SQL����Ϊ���ӡ���QQ����Ϊ��������û�ԭ���ύ��QQ���룬������ɾ����QQ���롣 ��ô��ʱ����Ҫ�����ֶ�ֵ��գ� ����ʹ��NULL�ؼ��֣��������ֶ�ֵ�ͱ�����ˡ� ��ʵ�ϴ�ʱҲ�����ÿ��ַ���""������ա�
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<html>
<head>
	<title>�޸ĸ�����Ϣ</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmInfo.txtName.value==""){
				alert("��ʵ��������Ϊ��!");
				return false;
			}
			if (document.frmInfo.txtEmail.value==""){
				alert("E-mail����Ϊ��!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>
<body>
	<h2 align="center">�޸ĸ�����Ϣ</h2>
	<%
	'�����ȡ���û�����Ϣ��Ȼ������ʾ�ں���ı����
	Dim strSql,rs
	strSql="Select * From tbUsers Where strUserId='" & Session("strUserId") & "'"
	Set rs=conn.Execute(strSql)
	%>
	<form name="frmInfo" onsubmit="javascript: return check_Null();" action="" method="post">
	<table width="80%" border="1" bordercolorlight="#B0B0B0" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
		<tr height="25"> 
			<td width="20%">�û���</td>
			<td width="80%"><%=Session("strUserId")%></td>
		</tr>
		<tr height="25"> 
			<td>��ʵ����</td>
			<td><input type="text" name="txtName" size=15 value="<%=rs("strName")%>">*</td>
		</tr>
		<tr height="25"> 
			<td>�Ա�</td>
			<td>
				<input type="radio" name="rdoSex" value="��" <%If rs("strSex")="��" Then Response.Write "checked"%>>��
				<input type="radio" name="rdoSex" value="Ů" <%If rs("strSex")="Ů" Then Response.Write "checked"%>>Ů
			</td>
		</tr>
		<tr height="25"> 
			<td>�绰</td>
			<td><input type="text" name="txtTel" size="25" value="<%=rs("strTel")%>"></td>
		</tr>
		<tr height="25"> 
			<td>E-mail</td>
			<td><input type="text" name="txtEmail" size="40" value="<%=rs("strEmail")%>">*</td>
		</tr>
		<tr height="25"> 
			<td>QQ���� </td>
			<td><input type="text" name="txtQQ" size="15" value="<%=rs("strQQ")%>"></td>
		</tr>
		<tr height="25"> 
			<td>���˼��</td>
			<td><textarea name="txtIntro" rows="4" cols="50" ><%=rs("strIntro")%></textarea></td>
		</tr>
		<tr height="25"> 
			<td>��������</td>
			<td><%=rs("intArticle")%>ƪ</td>
		</tr>
		<tr height="25"> 
			<td>�ظ�����</td>
			<td><%=rs("intReArticle")%>ƪ</td>
		</tr>
	</table>
	<p align="center"><input type="submit" value="  ȷ����  ">
	</form>

	<%
	If Request.Form("txtName")<>"" And Request.Form("txtEmail")<>"" Then
		'���潨��SQL��䣬��ΪĳЩ�ֶ�����Ϊ�գ�������Ҫ�ж�һ��
		strSql="Update tbUsers Set strName='" & Request.Form("txtName") & "'"
		strSql=strSql & ",strEmail='" & Request.Form("txtEmail") & "'"
		strSql=strSql & ",strSex='" & Request.Form("rdoSex") & "'"
		strSql=strSql & ",dtmSubmit=#" & Date() & "#"
		'ע�⣺����û�û���ύQQ���룬��ô����ԭ����û��QQ���룬�������ֶ�ֵ��NULL����ˡ�
		If Request.Form("txtQQ") <> "" Then
			strSql = strSql & ",strQQ='" & Request.Form("txtQQ") & "'"
		Else
			strSql = strSql & ",strQQ=NULL"
		End If
		'ע�⣺���ڵ绰�Ľ���ͬ�����QQ
		If Request.Form("txtTel") <> "" Then
			strSql = strSql & ",strTel='" & Request.Form("txtTel") & "'"
		Else
			strSql = strSql & ",strTel=NULL"
		End If
		'ע�⣺���ڱ�ע�Ľ���ͬ�����QQ
		If Request.Form("txtIntro") <> "" Then
			strSql = strSql & ",strIntro='" & Request.Form("txtIntro") & "'"
		Else
			strSql = strSql & ",strIntro=NULL"
		End If	
		strSql=strSql & " Where strUserId='" & Session("strUserId") & "'"
		conn.Execute(strSql)
		'�رն���
		conn.close
		Set conn=Nothing
		Response.Redirect "index.asp"
	End If
	%>
</body>
</html>