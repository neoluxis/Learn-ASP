<%
'================================================================================================
'������Աҳ��
'1. ��ҳֻ�Ǵ�Application�ж�ȡ������Ա����������һ�����飬Ȼ����ѭ���г�������Ա��
'2. ���е��������ļ��еĳ������Ӷ���������Զ�ˢ��һ�Σ�����ʾ����������Ա��Ĭ��60��ˢ��һ�Ρ�
'3. ����������Ϣ�Ǵ������¹����ģ�����JavaScript���Ϳ��Խ����ڹ��������±ߣ��Ӷ���ʾ���µ�������Ϣ��
'================================================================================================
%>

<%option explicit%>
<!--#Include file="Config.asp"-->
<!--#INCLUDE FILE="function.asp"-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<meta http-equiv="refresh" content="<%=conRefreshOnline%>">
	<link rel="stylesheet" href="chat.css">
</head>
<body background="images/bg.gif">
	<p align="center">������Ա
	<br><hr>
	<%
	'�������л�ȡ������Ա
	Dim strUsers
	strUsers=Application("strUsers")
	'����ѭ���г�������������Ա
	Dim I
	For I=0 To Ubound(strUsers)
		Response.Write strUsers(I) & "<br>" 
	Next
	%>
</body>
</html>