<% Option Explicit                      		'���ڳ������У�ǿ����������%>
<html>
<body>
	<%
	Dim intMyRnd
	Randomize Timer()			'��������ڳ�ʼ��������ӣ�����ÿ�β�����ֵ��һ��
	intMyRnd = Int((10 * Rnd()) + 1)			'�����������
	Response.Write intMyRnd
	%>
</body> 
</html>
