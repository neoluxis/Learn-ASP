<%
'=========================================================================================
'BBS�б�ҳ
'1. ��ҳ������ҳ��ʾĳ����Ŀ��Ӧ�����£���ʾʱ���������ڵ������С�
'2. ע�Ȿҳֻ��ʾ��һ�����£�����ʾ�ظ����¡�
'3. Ҫע����������ı���intForumId��intPage��ǰ������ȷ����ʾ�ĸ���Ŀ����������ȷ����ʾ��һҳ�� �����������Ǵ��ݹ����ģ� ���Ҵӱ�ҳ���ӵ����ҳ��ʱ�� Ҳ�Ὣ�������������ݹ�ȥ�� ֮���ٴ��ݻ����� ��֮���ڲ�ͬ��ҳ��֮��Ҫʼ�ռǵô���������������
'4. ����������£�intPageҪ��Ϊ1����һ���Ǵ���ҳ�򿪱���Ŀʱ���ڶ��Ƿ��������º� ��������ȷ�����Ͽ����·�������¡�
'=========================================================================================
%>

<%Option Explicit%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>bbs��̳</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css">
</head>

<body>
	<h1 align="center"><font face="����"><%=conBBSTitle%></font></h1>
	<%
	'����һ�ηֱ��ȡ��ҳ����Ҫ���������������Ȼ�ȡ��̳��Ŀ��ű���---------------------------
	Dim intForumId 
	intForumId=CInt(Request.QueryString("intForumId"))			'��ȡ���ݹ�������̳��� 

	'��ȡ����ҳ�����
	Dim intPage                              
	If Request.QueryString("intPage")<>"" Then
		intPage=CInt(Request.QueryString("intPage"))			'��ȡ���ݹ�����ҳ�룬ת��Ϊ����
	Else
		intPage=1												'�����������Ϊ1��������ҳ�����򷢱����º����Ϊ1
	End If
	'��������������ºͷ�����ҳ�ĳ�����---------------------------------------------------
	Response.Write "<a href='bbs_insert.asp?intForumId=" & intForumId & "'>���������¡�</a>"
	Response.Write "<a href='index.asp'>��������̳��</a>"
	Response.Write "<p>"										'Ϊ�����ۣ����һ������
	'��������������ı�����-----------------------------------------------------------------
	Response.Write "<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>"
	Response.Write "<tr bgcolor='#C5EDE7' height='30' align='center'>"
	Response.Write "<th width='5%'>���</th>"
	Response.Write "<th width='55%'>����</th>"
	Response.Write "<th width='5%'>���</th>"
	Response.Write "<th width='5%'>�ظ�</th>"
	Response.Write "<th width='20%'>����ʱ��</th>"
	Response.Write "<th width='10%'>����</th>"
	Response.Write "</tr>"
	'���潨����¼��ʵ��rs����ע��Open�����Ĳ���-----------------------------------------------
	Dim rs,strSql
	strSql="Select * From tbBBS Where intForumId=" & intForumId & " And intLayer=1 Order By dtmSubmit Desc"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open strSql,conn,1									'��ΪҪ��ҳ��ʾ�������ü���ָ��
	'��������ǿվ���ʾ��¼
	If Not rs.Bof And Not rs.Eof Then   
	'������ҪΪ�˷�ҳ��ʾ
		rs.PageSize=conPageSize								'����ÿҳ��ʾ��������¼���������ļ��ж�ȡ
		rs.AbsolutePage=intPage								'���õ�ǰ��ʾ�ڼ�ҳ
		'��������ѭ����ʾ��ǰҳ�����м�¼
		Dim I
		For I=1 To rs.PageSize								'ѭ�������ǰҳ�����м�¼
			If rs.Eof Then Exit For							'������˼�¼����β��������ѭ��
			Response.Write "<tr bgcolor='#E1F3F4' height='30' valign='middle'>"  
			Response.Write "<td align='center'>" & rs("ID") & "</td>"
			Response.Write "<td align='left'><a href='bbs_particular.asp?ID=" & rs("ID") & "&intForumId=" & intForumId & "&intPage=" & intPage & "'>" & myHTMLEncode(rs("strTitle")) & "</a></td>"
			Response.Write "<td align='center'>" & rs("intHits") & "</td>"
			Response.Write "<td align='center'>" & rs("intChild") & "</td>"  
			Response.Write "<td align='center'>" & rs("dtmSubmit") & "</td>" 
			Response.Write "<td align='center'><a href='mailto:" & rs("strEmail") & "'>" & rs("strUserId") & "</a></td>"
			Response.Write "</tr>"
			rS.MoveNext                
		Next
		'����������Ľ�����ǣ���񵽴˽���
		Response.Write "</table>"
		'���������ǰҳ����ҳ��--------------------------------------------------------------------------
		Response.Write "<p align='right'>��ǰ��ʾ��" & intPage & "ҳ/��" & rs.PageCount & "ҳ"
		'����һ�����������1ҳ����һҳ����һҳ�����ҳ�ĳ�����
		Response.Write "&nbsp;&nbsp;<a href='bbs_list.asp?intPage=1&intForumId=" & intForumId & "'>����1ҳ��</a>"
		If intPage>1 Then
			Response.Write "<a href='bbs_list.asp?intPage=" & (intPage-1) & "&intForumId=" & intForumId & "'>����һҳ��</a>"
		Else
			Response.Write "����һҳ��"
		End If
		If intPage<rs.PageCount Then
			Response.Write "<a href='bbs_list.asp?intPage=" & (intPage+1) & "&intForumId=" & intForumId & "'>����һҳ��</a>"
		Else
			Response.Write "����һҳ��"
		End If
		Response.Write "<a href='bbs_list.asp?intPage=" & rs.PageCount  & "&intForumId=" & intForumId & "'>�����ҳ��</a>"
	End If
	'�رն���
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>   
</body>
</html>