<%
'===============================================================================================
'index.asp˵��
'1. ��ҳ���ǽ���ҳ��ʾ���ݡ�������ʾ���ݡ����ӵ���ϸҳ������ʾ�����ϵ���һ�� ���⣬�����ˡ���Ӽ�¼�� ���Ҽ�¼�� ���¼�¼�� ɾ����¼���ĳ����ӡ�
'2. ��ҳ���һ����Ҫ�ѵ����Ҫͬʱ��������ͷ�ҳ��ʾ���ݣ� Ҳ����Ҫͬʱ���� strField �� intPage������ ʵ�������������й�ϵ�ġ� ���磬��ѡ�񰴡������������ �������һҳ���������ٴδ򿪱�ҳ��ʱ�� ����Ҫ��ס��ǰҪ��ʾҳ�룬 ��Ҫ��ס��ǰ�������ֶϡ� �����ں��漸���������в���Ҫ������ҳ�룬 ��Ҫ�����������ֶΡ�
'3. ��Ȼ������ٴ��ڱ�������ѡ���������ֶΣ� ��Ϊ�ı��˼�¼���еļ�¼��˳�� ���Դ�ʱû�б�Ҫ��סԭ�����ڵ�ҳ�룬 ֱ����ʾ��1ҳ���ɡ�
'4. ������������ÿҳ��ʾ��������¼ʱʹ���� config.asp�еĶ���ĳ�����
'5. �������ڱ������۷�������һЩ�޸ģ����ҽ�ϵ�2����������⡣
'===============================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="config.asp"-->
<html>
<head>
	<title>�ҵ�ͨ��¼</title>
</head>
<body>
	<h1 align="center">ͨ��¼</h1>

	<!--�������ȸ���Ӽ�¼�Ͳ��Ҽ�¼�ĳ����� -->
	<a href="insert.asp">��Ӽ�¼</a>&nbsp;&nbsp;<a href="search.asp">���Ҽ�¼</a><p>

	<%
	'�������Ȼ�ȡ���ݹ����������ֶ�����-------------------------------------------------------------
	Dim strField
	If Request.QueryString("varField")="" Then
		strField="ID"										'���Ϊ�գ���Ĭ�ϰ�����ֶ�����
	Else
		strField=Request.QueryString("varField")			'�����Ϊ�գ��Ͱ���Ӧ�ֶ�����
	End If
	'����һ���жϵ�ǰ��ʾ�ڼ�ҳ�����ǵ�һ�δ򿪣�Ϊ1�������ɴ��ز�������
	Dim intPage                            
	If Request.QueryString("varPage")="" Then   
		intPage=1 
	Else
		intPage=CInt(Request.QueryString("varPage"))		'��CIntת��Ϊ����
	End If
	
	'���½���Recordset����ʵ��rs��ע��SQL�����Ҫ�õ������ȡ�������ֶ�-----------------------------
	Dim rs,strSql
	Set rs=Server.CreateObject("ADODB.Recordset")
	strSql ="Select ID,strName,strTel,strEmail,dtmSubmit From tbAddress Order By " & strField & " DESC"
	rs.Open strSql,conn,1									'ע���������Ϊ����ָ��

	'�����¼�����ǿյģ���ִ�з�ҳ��ʾ
	If Not rs.Bof And Not rs.Eof Then
		'-------------------------------------------------------------------------------------------
		'����һ�ο�ʼ��ҳ��ʾ��ָ��Ҫ��ʾ��ҳ��Ȼ��������ʾ��ǰҳ�����м�¼
		rs.PageSize=conPageSize								'�������config.asp�еĳ�������ÿҳ��ʾ��������¼
		rs.AbsolutePage=intPage								'���õ�ǰ��ʾ�ڼ�ҳ

		'��������������ı�����
		Response.Write "<table width='100%' border='1' bordercolorlight='#B0B0B0' bordercolordark='#FFFFFF' cellspacing='0' cellpadding='0' align='center'>"
		Response.Write "<tr bgcolor='#E0E0E0' height='25'>"
		Response.Write "<th width='10%'><a href='index.asp?varField=ID'>���</a></th>"
		Response.Write "<th width='10%'><a href='index.asp?varField=strName'>����</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=strTel'>�绰</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=strEmail'>E-mail</a></th>"
		Response.Write "<th width='20%'><a href='index.asp?varField=dtmSubmit'>�������</a></th>"
		Response.Write "<th width='10%'>����</th>"
		Response.Write "<th width='10%'>ɾ��</th>"
		Response.Write "</tr>"

		'��������ѭ����ʾ��ǰҳ�����м�¼
		Dim I
		For I=1 To rs.PageSize
			If rs.Eof Then Exit For							'������˼�¼����β��������ѭ��
			Response.Write "<tr align='center' height='25'>"								
			Response.Write "<td>" & rs("ID") & "</td>"
			Response.Write "<td><a href='particular.asp?ID=" & rs("ID") & "' target='_blank'>" & rs("strName") & "</td>"
			Response.Write "<td>" & rs("strTel") & "</td>"
			Response.Write "<td><a href='mailto:" & rs("strEmail") & "'>" & rs("strEmail") & "</td>"
			Response.Write "<td>" & rs("dtmSubmit") & "</td>"
			Response.Write "<td><a href='update.asp?ID=" & rs("ID") & "'>����</td>"
			Response.Write "<td><a href='delete.asp?ID=" & rs("ID") & "'>ɾ��</td>"
			Response.Write "</tr>"
			rs.MoveNext
		Next
		'����������Ľ�����ǣ���񵽴˽���
		Response.Write "</table>"

		'-------------------------------------------------------------------------------
		'�������������ǰҳ����ҳ��
		Response.Write "<p align='right'>��ǰ��ʾ��" & intPage & "ҳ/��" & rs.PageCount & "ҳ"

		'����һ�����������1ҳ����һҳ����һҳ�����ҳ�ĳ�����
		Response.Write "&nbsp;&nbsp;<a href='index.asp?varPage=1&varField=" & strField & "'>��1ҳ</a>&nbsp;"
		If intPage>1 Then
			Response.Write "<a href='index.asp?varPage=" & (intPage-1) & "&varField=" & strField & "'>��һҳ</a>&nbsp;"
		Else
			Response.Write "��һҳ&nbsp;"
		End If
		If intPage<rs.PageCount Then
			Response.Write "<a href='index.asp?varPage=" & (intPage+1) & "&varField=" & strField & "'>��һҳ</a>&nbsp;"
		Else
			Response.Write "��һҳ&nbsp;"
		End If
		Response.Write "<a href='index.asp?varPage=" & rs.PageCount  & "&varField=" & strField & "'>���ҳ</a>&nbsp;"
		'-------------------------------------------------------------------------------
	End If
	'���ڹر��йض���
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>
</body>
</html>

