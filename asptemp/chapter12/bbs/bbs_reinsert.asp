<%
'=========================================================================================
'�ظ�����ҳ��
'1. ����bbs_particular.asp���ύ�ظ����º󣬾ͻ�򿪱�ҳ�档���о������һ����¼�����Ҹ�����صı�
'2. ����Ӽ�¼ʱ����Ϊ���ݺ�E-mail����Ϊ�գ������������SQL�����һЩ���ӣ��������SQL����ǰ�벿�֣� Ȼ�����SQL���ĺ�벿�֡� ������һ��������SQL��䡣
'3. Ҫע�⣬�ظ����µĸ�����ID���������ı��򴫵ݹ����ġ�
'4. ���⣬���ȡ���������ı��򴫵ݹ�����ID��Ȼ�����ض���bbs_list.aspʱ���������������������ݻ�ȥ��
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="function.asp"-->
<%
'�������Ȼ�ȡ�й�ֵ
Dim strTitle,strBody,intLayer,intFatherId,intChild,intHits,strIP,strUserId,strEmail,intForumId,intPage
strTitle=myDangerEncode(Request.Form("txtTitle"))						'�������±���
strBody=myDangerEncode(Request.Form("txtBody"))							'������������
If Session("strUserId")<>"" Then
	strUserId=Session("strUserId")										'���������û���
Else
	strUserId="����"													'�����δ��¼�û�����ͳһ����Ϊ ����
End If
strEmail=myDangerEncode(Request.Form("txtEmail"))						'��������E-mail
intFatherId=Request.Form("txtFatherId")									'��ȡ�������ı��򴫻����ĸ�����Id
intLayer=2																'���ǵ�2��
intChild=0																'�ظ�������ĿΪ0����2�㲻����ظ�
intHits=0																'�����Ϊ0
strIP=Request.ServerVariables("remote_addr")							'����IP��ַ
intForumId=Request.Form("txtForumId")									'�������ı����ȡ��ǰ��Ŀ���
intPage=Request.Form("txtPage")											'�������ı����ȡҳ�룬�����ض�����õ�

'���½����±��浽���ݿ�
Dim sqlA,sqlB,strSql
sqlA="Insert Into tbBBS(strTitle,intLayer,intFatherId,intChild,intHits,strIP,strUserId,dtmSubmit,intForumId"
sqlB="Values('" & strTitle & "'," & intLayer & "," & intFatherId & "," & intChild & "," & intHits & ",'" & strIP & "','" & strUserId & "',#" & Now() & "#," & intForumId 
If strBody<>"" Then													'��������ݣ������
	sqlA=sqlA & ",strBody"
	sqlB=sqlB & ",'" & strBody & "'"
End If
If strEmail<>"" Then												'�����E-mail�������
	sqlA=sqlA & ",strEmail"
	sqlB=sqlB & ",'" & strEmail & "'"
End If
strSql=sqlA & ") " & sqlB & ")"
conn.Execute(strSql)

'���潫�����µĻظ�����1
strSql="Update tbBBS Set intChild=intChild+1 where ID=" & intFatherId
conn.Execute(strSql)

'�������佫��̳����������1
strSql="Update tbForum Set lngForumCount=lngForumCount+1 where ID=" & intForumId
conn.Execute(strSql)

'���潫���û��Ļظ���������1
strSql="Update tbUsers Set intReArticle=intReArticle+1 Where strUserId='" & Session("strUserId") & "'"
conn.Execute(strSql)

'�رն���
conn.Close
Set conn=Nothing

'�ض����BBS�б�ҳ�����ס�ٽ���ǰ��Ŀ�����ҳ�봫��ȥ
Response.Redirect "bbs_list.asp?intForumId=" & intForumId & "&intPage=" & intPage
%>
