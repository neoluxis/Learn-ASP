<% Option Explicit                      			'���ڳ������У�ǿ����������%>
<html>
<body>
	<%
	Dim strPwd										'����һ���ַ�������
	strPwd="   abcd*abcd   "						'���ַ���������ֵ
	strPwd=Trim(strPwd)								'ȥ���ַ������ߵĿո�
	Response.Write "<p>��ʱ�ַ����ǣ�" & strPwd 
	strPwd=Replace(strPwd,"*","")					'��*�滻Ϊ���ַ�����Ҳ����ȥ��*
	Response.Write "<p>��ʱ�ַ����ǣ�" & strPwd 
	strPwd=Left(strPwd,6)							'���ַ�����߿�ʼȡ6���ַ�
	Response.Write "<p>����ַ����ǣ�" & strPwd 
	%>
</body> 
</html>

