<%
'=========================================================================================
'BBS��ϸҳ
'1. ��ҳ������˼���ǣ����ȸ��ݴ��ݹ��������±�ţ������¼��ظ�����ȫ���г�����Ȼ���������ٸ���һ���ظ����� ����û��ύ�˻ظ��������ύ��bbs_reinsert.asp��ȥִ�лظ�������
'2. ��ע�⣬����ʾ����֮ǰ�������Ƚ������µĵ��������1.
'3. Ҫע�ⱾҳҲ���ȡintForumId��intPage����������Ȼ�󷵻�BBS�б�ʱ���Ὣ�����������ٴ��ݹ�ȥ��
'4. ����ʾ���߸�����Ϣʱ�������config.asp�еĺ�������ȡ������Ϣ���顣
'5. ע�⣬�ڻظ����У��Ὣ������Ҫ�ı����������ı��򴫵��˹�ȥ��
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<title>bbs��̳</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JAVASCRIPT">
		function check_Null(){
			if (document.frmRe.txtTitle.value==""){
				alert("���ⲻ��Ϊ��!");
				return false;
			}
			if (document.frmRe.txtUserId.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			return true;
		}
	</script>
</head>
<body>
	<h1 align="center">��ϸҳ��</h1>
	<%
	'�������Ȼ�ȡ���ݹ�������̳����intForumId��ҳ����ϢintPage�ͼ�¼���ID----------------------
	Dim ID,intForumId,intPage
	ID=Request.QueryString("ID")
	intForumId=Request.QueryString("intForumId")
	intPage=Request.QueryString("intPage")
	'���潫�ü�¼�ĵ������1--------------------------------------------------------------------
	Dim strSql,rs
	strSql="Update tbBBS Set intHits=intHits+1 Where id=" & id
	conn.Execute(strSql)
	'���濪ʼ������ݣ�����������-------------------------------------------------------------
	Response.Write "<table width='100%' border='1' bordercolorlight='#80BFFF' bgcolor='#E1F3F4' bordercolordark='#E1F3F4' cellspacing='0' cellpadding='0' align='center'>"
	'���潨����¼����������ѭ����ʾ���м�¼
	Dim strTitle
	'��ע������SQL��䣬�����ȡ��ǰ���¼������лظ����µļ�¼������������������
	strSql="Select * From tbBBS Where ID=" & ID & " Or intFatherId=" & ID & " Order By dtmSubmit"
	Set rs=conn.Execute(strSql)
	strTitle=rs("strTitle")
	Dim I														'����I�����ж��ǵڼ�¥
	Do While not rs.Eof 
		I=I+1
		'����ÿ����¼�����������еĵ�Ԫ������ʾ
		'�����ǵ�1��
		Response.Write "<tr bgcolor='#C5EDE7' height='30'>"
		'�����ǵ�1�е�1��
		Response.Write "<td width='20%' align='center'>"
		'�����ж��Ƿ�¥�����ֱ���ʾ��ͬ����Ϣ
		If I=1 Then
			Response.Write "¥��"						
		Else
			Response.Write "��" & I & "¥"				
		End If
		Response.Write "</td>"
		'�����ǵ�1�е�2��
		Response.Write "<td>"
		'�����ж��Ƿ�¥���������¥������ʾ����ʱ�䡢��������ظ���������ֻ��ʾ����ʱ��
		If I=1 Then
			Response.Write "����ʱ�䣺" & rs("dtmSubmit") & "&nbsp;�����" & rs("intHits") & "&nbsp;�ظ���" & rs("intChild")
		Else
			Response.Write "����ʱ�䣺" & rs("dtmSubmit") 
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		'�����ǵ�2��
		Response.Write "<tr valign='top'>"
		'�����ǵ�2�е�1��
		Response.Write "<td>"
		Response.Write "���ߣ�<a href='mailto:" & rs("strEmail") & "'>" & rs("strUserId") & "</a>"
		'������ǹ��ͣ��͵���function.asp�еĺ��������ظ����ߵ�һЩ��Ϣ
		If rs("strUserId")<>"����" Then
			Dim strTemp
			strTemp=PersonalInfo(rs("strUserId"))
			Response.Write "<br>����" & strTemp(4) & "ƪ"
			Response.Write "<br>�ظ���" & strTemp(5) & "ƪ"
			Response.Write "<br>QQ��" & strTemp(2) 
		End If
		Response.Write "</td>"
		'�����ǵ�2�е�2�У����л�����ı�������
		Response.Write "<td>"
		Response.Write "<b><font color='#b40000'>" & myHTMLEncode(rs("strTitle")) & "</font></b>"
		Response.Write "<p>" & myHTMLEncode(rs("strBody"))
		Response.Write "</td>"
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	'�����������ǣ�������ݵ��˽���
	Response.Write "</table>"
	%>

	<!--ע�ͣ������Ƿ���BBS�б�ҳ��ĳ����ӣ�ע�����н��������������˻�ȥ-->
	<p align="center"><a href="bbs_list.asp?intForumId=<%=intForumId%>&intPage=<%=intPage%>">��������̳��</a>

	<!--ע�ͣ������ǻظ������õı��-->
	<form name="frmRe" action="bbs_reinsert.asp" method="POST" onsubmit="javascript: return check_Null();">
	<table border="0" width="80%" cellspacing="0" cellpadding="0" align="center">
		<tr>                                        
			<td width="20%">�������⣺</td>
			<!--ע�ͣ����潫�ظ�������ʾ���ı�����-->
			<td><input type="text" name="txtTitle" size="50" value="re��<%=strTitle%>"></td>
		</tr>
		<tr>
			<td>�������ݣ�</td>
			<td><textarea rows="6" name="txtBody" cols="60"></textarea></td>
		</tr>
		<tr>                                        
			<td>�û�����</td>
			<td>
				<!--ע�ͣ���������Ƿ��Ѿ���¼��ʾ��ͬ���û�������û�е�¼����Ϊ�����͡�-->
				<%If Session("strUserId")<>"" Then%>
					<input type="text" name="txtUserId" size="20" value="<%=Session("strUserId")%>" disabled>
				<%Else%>
					<input type="text" name="txtUserId" size="20" value="����" disabled>
				<%End If%>
			</td>               
		</tr>
		<tr>                                        
			<td>E-mail��</td>
			<!--ע�ͣ������¼�ˣ��ͽ��û���E-mail��ʾ���ı�����-->
			<td><input type="text" name="txtEmail" size="30" value="<%=Session("strEmail")%>"></td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<br>
				<!--ע�ͣ������������ı��򽫼�����Ҫ�ı��������˹�ȥ-->
				<input type="hidden" name="txtFatherId" value="<%=ID%>">
				<input type="hidden" name="txtForumId" value="<%=intForumId%>">
				<input type="hidden" name="txtPage" value="<%=intPage%>">
				<input type="submit" value="ȷ���ύ">&nbsp; 
				<!--ע�ͣ������Ƿ���BBS�б�ҳ��İ�ť��ע�����н��������������˻�ȥ-->
				<input type="button" value="������̳" name="������̳" onclick="window.open('bbs_list.asp?intForumId=<%=intForumId%>&intPage=<%=intPage%>','_self')" >
			</td>
		</tr>
	</table>
	</form>    
</body>
</html>
