<%
'================================================================================================
'��������ҳ
'1. �ڸ�ҳ���������������û���������û���û�б�ռ�ã��Ϳ��Խ��������ҡ�
'2. �ڽ���������ǰ������ú�����������Ϣ����ӻ�ӭ�����ߵ���Ϣ�����⣬�Ὣ��ǰ�û���ӵ�������Ա�����С�
'3. ������Ա����ʵ�����Ǳ�����Application�е�һ�����顣
'================================================================================================
%>

<%option explicit%>
<!--#Include File="function.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>��ӭ�����ҵ�������</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="chat.css">
	<script language="JavaScript">
		//�ú����������пͻ�����֤
		function check_Null(){
			if (document.frmLogIn.txtUserName.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<!--ע�ͣ�����Ҫ�������ļ��ж�ȡ����conChatTitle-->
	<h1 align="center"><font face="����"><%=conChatTitle%></font></h1>

	<!--ע�ͣ�������ú�����ʾ��������-->
	<p align="center">��ǰ����<%=AllUserName()%>������

	<form name="frmLogIn" method="POST" action="index.asp" onsubmit="JavaScript:return check_Null();">
		<center>
		�����������û���<input type="text" name="txtUserName" size="30">
		<input type="submit" value="����������">
		</center>
	</form>

	<%
	'����û��������û�������ִ������ĳ���
	If Request.Form("txtUserName")<>"" Then
		'���·����û������û�IP��ַ�������浽Session��
		Dim strUserName,IP
		strUserName=Request.Form("txtUserName")				'��ȡ�û�����
		IP=Request.ServerVariables("REMOTE_ADDR")			'��ȡIP��ַ
		Session("strUserName")=strUserName					'���浽Session�д���                   
		Session("IP")=IP									'ͬ��

		'������ú�������Ƿ��Ѿ�����ʹ�ø��û�����������ʹ�ã�����ӵ��û�������
		If GetUserName(strUserName)=True Then
			Response.Write "<p align='center'>�Ѿ�����ʹ�ø����ƣ�����������"
		Else
			'������ú��������û������浽������Ա������
			Call addUserName(strUserName)
			'������ú��������û���������Ϣ��ӵ�������Ϣ��
			Call AddUserCome(strUserName,IP)
			'�ض���������ҳ��
			Response.Redirect "whole.asp"
		End If
	End If
	%>
</body>
</html>