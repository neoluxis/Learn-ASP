<html>
<body>
	<h2 align="center">�ļ��Ѱ�ȫ�ϴ�</h2>
	<%
	Dim upload											'����upload����ʵ��
	Set upload=Server.CreateObject("Persits.Upload")   
	upload.Save Server.MapPath("upfiles")				'�ϴ���ָ���ļ���
	Dim objItem											'����һ������ʵ������
	For Each objItem In upload.Files					'��ѭ����������ļ�����Ϣ
		Response.Write objItem.Name & "=" & objItem.Path & "<br>"
	Next
	For Each objItem In upload.Form						'��ѭ��������б�Ԫ����Ϣ
		Response.Write objItem.Name & "=" & objItem.Value & "<br>"
	Next
	%>
</body> 
</html>
