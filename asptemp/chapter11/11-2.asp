<html>
<body>
	<h2 align="center">�ļ��Ѱ�ȫ�ϴ�</h2>
	<%
	Dim upload														'����һ��upload����ʵ��
	Set upload=Server.CreateObject("Persits.Upload")
	upload.SetMaxSize 10*1024*1024,True								'�����ļ�������10MB�����򱨴�
	upload.OverWriteFiles=False										'False��ʾ�ļ������Ա�����
	upload.Save "C:\Inetpub\wwwroot\asptemp\chapter11\upfiles"		'�ϴ���ָ���ļ���
	'�����ж�һ�£����ȷʵ�ϴ����ļ���������ļ���Ϣ
	If upload.Files.Count>0 Then
		Response.Write "<br>�ϴ��ļ���" & upload.Files("fleUpload").Path 
		Response.Write "<br>�ļ�����" & upload.Files("fleUpload").FileName 
		Response.Write "<br>��չ����" & upload.Files("fleUpload").Ext
		Response.Write "<br>�ļ���С��" & upload.Files("fleUpload").Size & "�ֽ�"
		Response.Write "<br>����޸�ʱ�䣺" & upload.Files("fleUpload").LastWriteTime
	End If
	'���������Ԫ����Ϣ
	Response.Write "<br>�ļ�˵����" & upload.Form("txtIntro").Value 
	Response.Write "<br>����������" & upload.Form("txtAuthor").Value 
	%>
</body> 
</html>


