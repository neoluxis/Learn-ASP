<html>
<body>
	<%
	Dim strNames										'ע�⣺��������һ����ͨ����
	strNames=Session("strNames")						'��ȡSession����
	Response.Write strNames(0) & "���ã���ӭ��<br>"	
	Response.Write strNames(1) & "���ã���ӭ��<br>"  
	%>
</body> 
</html>

