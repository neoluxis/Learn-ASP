<html>
<body>
	<h2 align="center">��ʾ���ͼƬʾ��</h2>
	<%
	Dim ad											'����һ��AdRotator����ʵ��
	Set ad=Server.CreateObject("MSWC.AdRotator")
	ad.Border=2										'����ͼƬ�߿�Ϊ2����
	ad.Clickable=True                          		'����ͼƬ�ṩ�����ӹ���
	ad.TargetFrame="_blank"							'�������´����д���ַ
	Response.Write Ad.GetAdvertisement("adver.txt")	'��ȡ�����ϢHTML�ַ���
	%>
</body> 
</html> 
