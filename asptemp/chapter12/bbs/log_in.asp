<%
'=========================================================================================
'�û���¼ҳ
'1. ���о����ж��û�������û����������Ƿ���ȷ��
'2. �����ȷ���ͽ��û�����E-mail���浽Session�У�Ȼ���ض������ҳ��
'3. �������ȷ�������������ʾ��Ϣ
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<%
'�������Ȳ鿴�û����������Ƿ���ȷ
Dim strSql,rs
strSql="Select strUserId,strEmail From tbUsers Where strUserId='" & Request.Form("txtUserId") & "' And strPwd='" & Request.Form("txtPwd") & "'"
Set rs=conn.Execute(strSql)
If Not rs.Eof And Not rs.Bof Then
	'����м�¼����ʾ�и��û������û�����Email���浽Session��
	Session("strUserId")=rs("strUserId")
	Session("strEmail")=rs("strEmail")
	Response.Redirect "index.asp"					'�ض�����ҳ
Else
	'���û�м�¼����ʾ�û�����������ܲ���ȷ���������ʾ��Ϣ
	Response.Write "�Բ����û���������������<a href='index.asp'>������ҳ</a>���µ�¼"
End If
%>
