<%
'===============================================================================================
'delete.asp˵��
'1. ��ҳ��ֻ�Ǽ򵥵�ɾ���˼�¼�����е����ݿ�������������connection.asp�ļ��С�
'2. �����йر�connection��������ʵ����Ҳ����ʡ�ԡ���Ϊ���ļ�ִ����Ϻ�Ҳ���Զ��رյġ�
'===============================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<% 
'����ɾ����¼��ע������Ҫ�õ���index.asp��������Ҫɾ���ļ�¼��ID
Dim strSql
strSql="Delete From tbAddress Where ID=" & Request.QueryString("ID") 
conn.Execute(strSql)

'���ڹر�Connection����
conn.Close
Set conn=Nothing

'ɾ����Ϻ󣬷�����ҳ
Response.Redirect "index.asp"								
%>
