<%
'================================================================================================
'��������ҳ��
'1. ��ҳֻ�Ǵ�Application�ж�ȡ������Ϣ������ʾ��ҳ���С�
'2. ���е��������ļ��еĳ������Ӷ���������Զ�ˢ��һ�Σ�����ʾ����������Ϣ��Ĭ��5��ˢ��һ�Ρ�
'3. ����������Ϣ�Ǵ������¹����ģ�����JavaScript���Ϳ��Խ����ڹ��������±ߣ��Ӷ���ʾ���µ�������Ϣ��
'================================================================================================
%>

<%option explicit%>
<!--#Include File="config.asp"-->
<html>
<head>
	<meta http-equiv="refresh" content="<%=conRefresh%>">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="chat.css">
</head>
<body bgcolor="#EEEEEE" bottomMargin="50">
	<%
	'�����Application�л�ȡ������Ϣ���������ҳ����
	Response.Write Application("strChat")
	%>
	<!--ע�ͣ���Ϊ���ϵ�����ʾ������Ҫ����JavaScript���������±�-->
	<script language="JavaScript" for="window" event="onload">window.scroll(0,60000);</script>
</body>
</html>