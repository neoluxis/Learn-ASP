<html>
<body>
	<h2 align="center">���߷��͸���</h2>
	<%
	'�������ȱ����ϴ����ļ�
	Dim upload													'����upload����ʵ��
	Set upload=Server.CreateObject("Persits.Upload")     
	upload.Save "C:\Inetpub\wwwroot\asptemp\chapter11\upfiles"	'�ϴ���ָ���ļ���
	'���濪ʼ�����ʼ�
	Dim jmail													'����Message����ʵ��
	Set jmail=Server.CreateObject("Jmail.Message")    
	jmail.From="aspbook@163.com"								'������
	jmail.AddRecipient upload.Form("txtRecipient")				'�ռ��� 
	jmail.Subject=upload.Form("txtSubject")						'����
	jmail.Body=upload.Form("txtBody")							'����
	'�����ж�һ�£����ȷʵ�ϴ����ļ����ͷ��͸���
	If upload.Files.Count>0 Then
		jmail.AddAttachment upload.Files("fleUpload").Path		'��Ӹ���
	End If
	jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"					'ִ�з���
	jmail.Close                                                  
	Response.Write "�Ѿ��ɹ�����"
	%>
</body> 
</html>



