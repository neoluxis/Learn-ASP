<%
'=========================================================================================
'BBS��ҳ
'1. ����BBS��ҳ������ֻ����ʾ��̳��Ŀ�б�������ʾ�û���¼��
'2. ���û���¼���У�������û��Ƿ��Ѿ���¼��ʾ��ͬ�����ݡ����û�е�¼������ʾ��¼������Ѿ���¼�ˣ� ����ʾ�˳���¼���޸�������޸ĸ�����Ϣ�ĳ����ӡ�
'3. ����̳��Ŀ�б��У�ע�ⵥ��ͼƬ����Ŀ���ƶ������Ӧ����Ŀ�����⣬�Ὣ��Ŀ��ID��Ŵ��ݵ�bbs_list.asp�С�
'4. ����ʾͼƬʱ����һ��С���ɣ�ͼƬ�ļ���������Ϊ1.gif��2.gif��3.gif... ������ѭ�����ÿһ��ͼƬ�ļ���·����
'=========================================================================================
%>

<% Option Explicit						'ǿ���������� %>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>bbs��̳</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" type="text/css" href="bbs.css" >
	<script language="JAVASCRIPT">
		function check_Null(){
			if (document.frmReg.txtUserId.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			if (document.frmReg.txtPwd.value==""){
				alert("���벻��Ϊ��!");
				return false;
			}
			return true;
		}
	</script>
</head>

<body>
	<h1 align="center"><font face="����"><%=conBBSTitle%></font></h1>

	<form name="frmReg" method="POST" action="log_in.asp" onsubmit="Javascript: return check_Null();">
	<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>
		<tr bgcolor="#E1F3F4" height="40">    
			<td valign="middle">
				<%If Session("strUserId")<>"" Then%>
					�ѵ�¼�û�<input type="text" name="txtUserId" size="15" value="<%=Session("strUserId")%>" disabled>
					<a href="log_out.asp">��ע����</a>
					<a href="log_updatePWD.asp">���޸����롿</a>
					<a href="log_update.asp">���޸ĸ�����Ϣ��</a>
				<%Else%>
					�û���<input type="text" name="txtUserId" size="15">
					����<input type="PassWord" name="txtPwd" size="15">
					<input type="submit" value="�� ¼"> 
					<input type="button" value="ע ��" onclick="window.open('log_register1.asp','_self')">
				<%End If%>
			</td>  
		</tr>  
	</table>
	</form>

	<br>
	<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>
		<tr bgcolor="#C5EDE7" height="30" align="center">    
			<td width="20%">&nbsp;</td>  
			<td width="40%">��̳��Ŀ</td>
			<td width="40%">������</td>
		</tr>  
		<%
		'���½�����¼������ʵ��rs
		Dim rs,strSql
		strSql="Select * From tbForum"
		Set rs=conn.Execute(strSql)
		'��������ѭ�����������Ŀ������I��������˳����ʾͼƬ
		Dim I
		Do While Not rs.Eof
			I=I+1
			Response.Write "<tr height='60' valign='middle' align='center'>"
			Response.Write "<td><a href='bbs_list.asp?intForumId=" & rs("ID") & "'><img src='images/" & I & ".gif' border='0'></a></td>"
			Response.Write "<td><a href='bbs_list.asp?intForumId=" & rs("ID") & "'>" & rs("strForumName") & "</a></td>"
			Response.Write "<td>" & rs("lngForumCount") & "</td>"
			Response.Write "</tr>"  
			rS.MoveNext                
		Loop
		'�رն���
		rs.Close
		Set rs=Nothing
		conn.Close
		Set conn=Nothing
		%>   
	</table> 
</body>
</html>
