<%
'================================================================================================
'��ҳ�ļ�
'1. ��ҳ��Ҫ��Ϊ�����֣�������һ��������Եı������ᱻ�ύ��insert.asp����������ʾ�������ԵĲ��֣� ��������ѭ����ʾ���м�¼���ѡ�
'2. Ҫע�������ڱ���ʹ���˿ͻ��˵�JavaScript��֤��ͨ����֤��Ż�����ύ�����������ʾ�û�������д��
'3. ��ҳ�������ʽ�ļ�guest.css�����й����֡������ӵȵ���ʽ��
'4. ��ҳ���ȡconfig.asp�е����ã���ʾ���԰������
'5. ��������ʾ����ʱ�������function.asp�еĺ�������������������ݽ��б��룬�Ա���ʾHTML�����ʵ�ֻ���Ч����
'================================================================================================
%>

<% Option Explicit						'ǿ����������%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>��ӭ�����ҵ����԰�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="guest.css">
	<script language="JavaScript">
		//�ú����������пͻ�����֤
		function check_Null(){
			if (document.frmGuest.txtTitle.value==""){
				alert("���ⲻ��Ϊ��!");
				return false;
			}
			if (document.frmGuest.txtName.value==""){
				alert("��������Ϊ��!");
				return false;
			}
			if (document.frmGuest.txtTitle.value.length>50){
				alert("���ⲻ�ܳ���50���ַ�");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<!--ע�ͣ�����Ҫ�������ļ��ж�ȡ����conGuestTitle-->
	<h1 align="center"><font face="����"><%=conGuestTitle%></font></h1>

	<!--ע�ͣ��������ύ���Ա����ύ�����Ȼ��������Ŀͻ�����֤������֤����֤ͨ�����ٴ��͵�insert.asp-->
	<form name="frmGuest" method="POST" action="insert.asp" onsubmit="JavaScript:return check_Null();">
		<table border="0" width="80%" bgcolor="#203F80" align="center">
			<tr>
				<td><font color="white">���⣺</font></td>
				<td><input type="text" name="txtTitle" size="60">*</td>
			</tr><tr>
				<td><font color="white">���ݣ�</font></td>
				<td><Textarea Name="txtBody" Rows="4" Cols="60"></TextArea></td>
			</tr><tr>
				<td><font color="white">������</font></td>
				<td><input type="text" name="txtName" size="13">*</td>
			</tr><tr>
				<td><font color="white">E-mail��</font></td>
				<td><input type="text" name="txtEmail" size="40"></td>
			</tr><tr>
				<td></td>
				<td><input type="submit" value="����   ����" Size="20" ></td>
			</tr>
		</table>
	</form>
	<%
	'���¿�ʼ��ʾԭ�����ԣ���ע��ÿ�����Ի���ʾ��һ�������
	Dim rs,strSql
	strSql ="Select * From tbGuest Order By dtmSubmit Desc"	
	Set rs=conn.Execute(strSql)
	Do While Not rs.Eof
		%>
		<table border="0" width="80%" align="center">
			<tr><td colspan="2"><hr></td></tr>
			<tr><td width="20%">����</td><td><%=myHTMLEncode(rs("strTitle"))%></td></tr>
			<tr><td>����</td><td><%=myHTMLEncode(rs("strBody"))%></td></tr>
			<tr><td>������</td><td><a href="mailto:<%=rs("strEmail")%>"><%=myHTMLEncode(rs("strName"))%></a></td></tr>
			<tr><td>ʱ��</td><td><%=rs("dtmSubmit")%></td></tr>
			<tr><td></td><td><a href="delete.asp?ID=<%=rs("ID")%>">ɾ��</a></td></tr>
		</table>
		<%
		rs.MoveNext
	Loop 
	'�رն���
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>
</body>
</html>