<html>
<body>
	<% 
	Dim strName,intAge                                   
	strName=Request.QueryString("strName")				'��ȡ����
	intAge=Request.QueryString("intAge")				'��ȡ����
	Response.Write "���������ǣ�" & strName & "�����������ǣ�" & intAge
	%>
</body>
</html>

