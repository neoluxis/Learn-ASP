<html>
<body>
	<h2 align="center">���������Ӧ��ʾ��</h2>
	<%
	Dim count											'����һ�����ʵ������
	Set count=Server.CreateObject("MSWC.PageCounter")
	count.PageHit										'����ǰ��ҳ���ʴ�����1
	Dim intVisit
	intVisit=count.Hits()								'��ȡ��ǰ��ҳ���ʴ���
	Response.Write "���ǵ�" & intVisit & "λ�ÿ�"
	%> 
</body> 
</html>
 


