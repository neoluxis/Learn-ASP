<html>
<body>
	<%
	Dim intA,intB,intC
	intA=Request.Form("txtA")				'��ȡ��Ԫ��txtA��ֵ
	intB=Request.Form("txtB")				'��ȡ��Ԫ��txtB��ֵ
	intC=CInt(intA)+CInt(intB)				'��Ϊ��ȡ�����ַ��������Ա���ת������
	Response.Write "�������ĺ�=" & intC	
	%>
</body>
</html>

