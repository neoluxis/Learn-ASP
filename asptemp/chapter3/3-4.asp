<% Option Explicit								'���ڳ������У�ǿ�Ʊ�������  %>
<html>
<body>
	<%
	Dim strA
	strA = Split("VBScript*is*good!", "*")		'��*Ϊ�ֽ�����
	Response.Write strA(0) & "<P>"
	Response.Write strA(1) & "<P>"
	Response.Write strA(2)
	%>
</body> 
</html> 
