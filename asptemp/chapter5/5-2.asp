<html>
<body>
	<%
	Dim strName,intAge
	strName=Session("strName")						'��ȡSession����
	intAge=Session("intAge")							'��ȡSession����
	Response.Write strName & "���ã���ӭ��<br>"
	Response.Write "����������" & intAge
	%>
</body> 
</html>

