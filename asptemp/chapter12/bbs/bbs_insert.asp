<%
'=========================================================================================
'����������ҳ��
'1. ����ʵ����һ����ͨ����Ӽ�¼��ҳ�档��Ӽ�¼�󣬻���Ҫ�������������� �Ա���¸���Ŀ������Ŀ�͸��û�����������Ŀ��
'2. ����Ӽ�¼ʱ����Ϊ���ݺ�E-mail����Ϊ�գ������������SQL�����һЩ���ӣ��������SQL����ǰ�벿�֣� Ȼ�����SQL���ĺ�벿�֡� ������һ��������SQL��䡣
'3. ����Ҫע�⣬����Ὣ���ݹ�����intForumId�ȴ���������ı����У�Ȼ���ύ���󣬿��Լ�����ȡ�ñ����� Ȼ�����ض���BBS�б�ҳʱ�������ñ������ݻ�ȥ��
'=========================================================================================
%>

<%option explicit%>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="function.asp"-->
<html>
<head>
	<title>bbs��̳</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmInsert.txtTitle.value==""){
				alert("���ⲻ��Ϊ��!");
				return false;
			}
			if (document.frmInsert.txtUserId.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>

<body>
	<h1 align="center">��������</h1>
	<form name="frmInsert" action="" method="POST" onsubmit="javascript: return check_Null();">
	<table border="0" width="80%" cellspacing="0" cellpadding="0" class="s2" align="center">
		<tr>                                        
			<td width="20%">�������⣺</td>
			<td><input type="text" name="txtTitle" size="50" ></td>
		</tr>
		<tr>
			<td>�������ݣ�</td>
			<td><textarea rows="10" name="txtBody" cols="60" ></textarea></td>
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
			<td><input type="text" name="txtEmail" size="30" value="<%=Session("strEmail")%>"></td>
		</tr>
		<tr>
			<td colspan="2"  align="center">
				<br>
				<!--ע�ͣ������������ı���intForumId���������˹�ȥ-->
				<input type="hidden" name="txtForumId" value="<%=Request.QueryString("intForumId")%>">
				<input type="submit" value="ȷ���ύ">&nbsp; 
				<input type="button" value="������̳" onclick="javascript:history.back();">
			</td>
		</tr>
	</table>
	</form>    
	
	<%
	'����ύ�˱�����ִ�������������
	If Trim(Request.Form("txtTitle"))<>"" Then
		Dim strTitle,strBody,intLayer,intFatherId,intChild,intHits,strIP,strUserId,strEmail,intForumId
		strTitle=myDangerEncode(Trim(Request.Form("txtTitle")))		'�������±���
		strBody=myDangerEncode(Request.Form("txtBody"))				'������������
		If Session("strUserId")<>"" Then
			strUserId=Session("strUserId")							'���������û���
		Else
			strUserId="����"										'�����δ��¼�û�����ͳһ����Ϊ ����
		End If
		strEmail=myDangerEncode(Request.Form("txtEmail"))			'��������E-mail
		intForumId=Request.Form("txtForumId")						'��ȡ�����ı��򴫵ݻ�������Ŀ���
		intLayer=1													'���ǵ�һ��
		intFatherId=0												'��Ϊ�ǵ�һ�㣬�������Ϊ0
		intChild=0													'�ظ�������ĿΪ0
		intHits=0													'�����Ϊ0
		strIP=Request.ServerVariables("REMOTE_ADDR")				'����IP��ַ

		'���½��������±��浽���ݿ�
		Dim sqlA,sqlB,strSql
		sqlA="Insert Into tbBbs(strTitle,intLayer,intFatherId,intChild,intHits,strIP,strUserId,dtmSubmit,intForumId"
		sqlB = "Values('" & strTitle & "'," & intLayer & "," & intFatherId & "," & intChild & "," & intHits & ",'" & strIP & "','" & strUserId & "',#" & Now() & "#," & intForumId 
		If strBody<>"" Then											'��������ݣ������
			sqlA = sqlA & ",strBody"
			sqlB = sqlB & ",'" & strBody & "'"
		End If
		If strEmail<>"" Then										'�����E-mail�������
			sqlA = sqlA & ",strEmail"
			sqlB = sqlB & ",'" & strEmail & "'"
		End If
		strSql = sqlA & ") " & sqlB & ")"
		conn.Execute(strSql)

		'���潫����Ŀ����������1
		strSql="Update tbForum Set lngForumCount=lngForumCount+1 Where ID = " & intForumId
		conn.Execute(strSql)

		'���潫���û��ķ�����������1
		strSql="Update tbUsers Set intArticle=intArticle+1 Where strUserId='" & Session("strUserId") & "'"
		conn.Execute(strSql)

		'�رն���
		conn.Close
		Set conn=Nothing

		'�ض����BBS�б�ҳ�����ס�ٽ���ǰ��Ŀ���봫�ݻ�ȥ��
		Response.Redirect "bbs_list.asp?intForumId=" & Request.Form("txtForumId")
	End If
	%>
</body>
</html>
