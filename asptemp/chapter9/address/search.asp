<%
'===============================================================================================
'search.asp˵��
'1. ��ҳ��ʵ��������9.2.4�ڲ��Ҽ�¼ʾ���Ļ������޸Ķ��ɵġ�
'2. ����ֻ�ǽ�����HTML���붼��Response.Write ��������ˣ� ���⣬����ˡ����¡�����ɾ���������ӵ���ϸҳ��ĳ����ӡ� ��Щ����ʵ���Ϻ�index.asp��ࡣ
'3. ��Ҫע����Ǳ����򻹲���ʮ�������������ڱ�ҳ���е�����ɾ�����򡱸��¡������ӣ� ִ����Ϻ󲻻᷵�ر�ҳ�棬 ���ǻ�ֱ�ӷ�����ҳindex.asp�� 
'4. �ر��йض�������Ҳ����ʡ�Ե�����Ϊִ���걾ҳ���Ҳ���Զ��رյġ�
'5. ��ҳ���е����ݿ��������Ҳ������connection.asp��
'===============================================================================================
%>


<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<html>
<head>
	<title>������Ա</title>
</head>
<body>
	<h2 align="center">������Ա</h2>
	<form name="frmSearch" method="POST" action="">
		������Ҫ���ҵ�������<input type="text" name="txtName">
		<input type="submit" name="btnSubmit" value=" ȷ �� ">
	</form>
	<%
	'�����������������ִ������Ĳ��ҹ���
	If Request.Form("txtName")<>"" Then

		'���½���һ��RecordSet����ʵ��rs��ע��Select�����Ҫ�õ��ύ������-----------------------------------
		Dim rs,strSql 
		strSql="Select ID,strName,strTel,strEmail,dtmSubmit From tbAddress Where strName Like '%" & Request.Form("txtName") & "%'"
		Set rs=conn.Execute(strSql)

		'�������ñ����ʾ���ҵ��ļ�¼------------------------------------------------------------------------
		'��������������ı�����
		Response.Write "<table width='100%' border='1' bordercolorlight='#B0B0B0' bordercolordark='#FFFFFF' cellspacing='0' cellpadding='0' align='center'>"
		Response.Write "<tr bgcolor='#E0E0E0' height='25'>"
		Response.Write "<th width='10%'>���</th>"
		Response.Write "<th width='10%'>����</th>"
		Response.Write "<th width='20%'>�绰</th>"
		Response.Write "<th width='20%'>E-mail</th>"
		Response.Write "<th width='20%'>�������</th>"
		Response.Write "<th width='10%'>����</th>"
		Response.Write "<th width='10%'>ɾ��</th>"
		Response.Write "</tr>"

		'��������ѭ����ʾ���м�¼
		Do While Not rs.Eof							'ֻҪ���ǽ�β��ִ��ѭ��
			Response.Write "<tr align='center' height='25'>"								
			Response.Write "<td>" & rs("ID") & "</td>"
			Response.Write "<td><a href='particular.asp?ID=" & rs("ID") & "' target='_blank'>" & rs("strName") & "</td>"
			Response.Write "<td>" & rs("strTel") & "</td>"
			Response.Write "<td><a href='mailto:" & rs("strEmail") & "'>" & rs("strEmail") & "</td>"
			Response.Write "<td>" & rs("dtmSubmit") & "</td>"
			Response.Write "<td><a href='update.asp?ID=" & rs("ID") & "'>����</td>"
			Response.Write "<td><a href='delete.asp?ID=" & rs("ID") & "'>ɾ��</td>"
			Response.Write "</tr>"
			rs.MoveNext								'����¼ָ���ƶ�����һ����¼
		Loop
		
		'����������Ľ�����ǣ���񵽴˽���
		Response.Write "</table>"
	
	End If

	'���ڹر�Connection����-----------------------------------------------------------------------------------
	conn.Close
	Set conn=Nothing
	%>

	<p align='center'><a href="index.asp">������ҳ</a>
</body> 
</html>
