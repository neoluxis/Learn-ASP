<%
'================================================================================================
'��������ļ�
'1. ������ȡ��ҳ�ύ�����ı���Ȼ��������ӵ����ݿ��С�
'2. ��Ϊ���ݺ�E-mail����ʡ�ԣ������������SQL���Ĵ���Ƚϸ��ӡ�ʵ�����Ƿֳ�ǰ�����벿�֣�Ȼ��ֱ���֯�� ������γ�һ����׼��Insert��䡣
'3. ����û��ύ�ı��а�����Ӣ�ĵĵ����ţ��ͻ��SQL����еĵ����ŷ�����ͻ���Ӷ����´��� �����������myDangerEncode�����еĵ������滻Ϊ�����������ĵ����š��������ᷢ������ ����д�����ݿ��к���Ȼ�����1�������š�
'================================================================================================
%>

<% Option Explicit						'ǿ����������%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="function.asp"-->
<%
'�������Ȼ�ȡ�ύ���������ݣ�ע�����л���ú�������Σ���ַ�
Dim strTitle,strBody,strName,strEmail
strTitle=myDangerEncode(Request.Form("txtTitle"))				
strBody=myDangerEncode(Request.Form("txtBody"))
strName=myDangerEncode(Request.Form("txtName"))
strEmail=myDangerEncode(Request.Form("txtEmail"))
'���濪ʼ��Ӽ�¼����Ϊ���ݺ�E-mail����ʡ�ԣ����Խ�SQL���ֳ�ǰ�����ηֱ���
Dim sqlA,sqlB,strSql
sqlA="Insert Into tbGuest(strName,strTitle,dtmSubmit"
sqlB="values('" & strName & "','" & strTitle & "',#" & Now() & "#"
If strBody<>"" Then										'����ύ�����ݣ������
	sqlA=sqlA & ",strBody"
	sqlB=sqlB & ",'" & strBody & "'"
End If
If strEmail<>"" Then									'����ύ��E-mail�������
	sqlA=sqlA & ",strEmail"
	sqlB=sqlB & ",'" & strEmail & "'"
End If
'���潫ǰ���������������SQL��䣬��ִ�����
strSql=sqlA & ") " & sqlB & ")"
conn.Execute(strSql)											
'�رն���
conn.Close
Set conn=Nothing
Response.Redirect "index.asp"									
%>
