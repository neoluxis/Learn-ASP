<% 
'ֻҪ�������͵绰���͸��¼�¼���������������Ϣ
If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then
	'���»�ȡ�ύ�ı�����
	Dim ID,strName,strSex,intAge,strTel,strEmail,strIntro		'������������
	ID=Request.Form("txtID")									'��ȡ�����ı��򴫵ݹ�����IDֵ
	strName=Request.Form("txtName")								'��ȡ����
	strSex=Request.Form("rdoSex")								'��ȡ�Ա�
	If Request.Form("txtAge")<>"" Then							'��ȡ����
		intAge=Request.Form("txtAge")
	Else
		intAge=0
	End If
	strTel=Request.Form("txtTel")								'��ȡ�绰
	strEmail=Request.Form("txtEmail")							'��ȡE-mail
	strIntro=Request.Form("txtIntro")							'��ȡ���˼��
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn
	'��������Execute�������¼�¼��ע���������Ϊ�����ı��򴫹�����IDֵ
	Dim strSql
	strSql="Update tbAddress Set strName='" & strName & "',strSex='" & strSex & "',intAge=" & intAge & ",strTel='" & strTel & "',strEmail='" & strEmail & "',strIntro='" & strIntro & "' Where ID=" & ID
	conn.Execute(strSql) 
	'������ϣ��ض������ҳ
	Response.Redirect "index.asp"    
Else
	Response.Write "�����͵绰������д"
	Response.Write "<a href='index.asp'>������д</a>"
End If
%>
