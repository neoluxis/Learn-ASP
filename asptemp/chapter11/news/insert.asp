<%
'�������ȱ����ϴ����ļ�
Dim upload												'����upload����ʵ��
Set upload=Server.CreateObject("Persits.Upload") 
upload.OverWriteFiles=False								'False��ʾ�ļ������Ա�����
upload.Save Server.MapPath("upfiles")					'�ϴ���ָ���ļ���
'���»�ȡ�ύ�����ı��⡢���ݺ��ļ�����
Dim strTitle,strBody,strFileName						'����3����������
strTitle=upload.Form("txtTitle")
strBody=upload.Form("txtBody")
'�����ж��Ƿ��ϴ����ļ������û���ϴ�����ֵ���ַ���
If upload.Files.Count>0 Then
	strFileName=upload.Files("fleUpload").FileName		'��ȡ�ϴ��ļ�������
Else
	strFileName=""										'��ֵ���ַ���
End If
'���½����⡢���ݺ��ļ�����ӵ����ݿ���
Dim strConn,conn 
Set conn=Server.CreateObject("ADODB.Connection") 
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("news.mdb")
conn.Open strConn
Dim strSql
strSql="Insert Into tbNews(strTitle,strBody,strFileName,dtmSubmit) Values('" & strTitle & "','" & strBody & "','" & strFileName & "',#" & date() & "#)"
conn.execute(strSql)
Response.Redirect "index.asp"							'�ɹ���ӣ�������ҳ
%>
