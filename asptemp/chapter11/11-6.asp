<html>
<body>
	<h2 align="center">�ļ��ϴ����</h2>
	<%
	Dim upload											'����һ��upload����ʵ��
	Set upload=Server.CreateObject("Persits.Upload")
	upload.OverwriteFiles=False							'False��ʾ�ļ������Ը���
	upload.Save											'�ϴ��ļ����ڴ���
	'�Ƚ���һ���ϴ��ļ����󣬺������ֱ�����øö���
	Dim fle
	Set fle=upload.Files("fleUpload")
	'�����ж��ļ��Ƿ��Ѿ�����
	If upload.FileExists(Server.MapPath("upfiles" & "/" & fle.FileName)) Then
		Response.Write "�ļ��Ѿ����ڣ��뷵��<a href='11-5.asp'>�����ϴ�</a>"
	Else
		'���ڴ��б��浽�ļ�����
		fle.SaveAs Server.MapPath("upfiles" & "/" & fle.FileName)	
		Response.Write "�ļ�·����" & fle.Path
	End If
	%>
</body> 
</html>


