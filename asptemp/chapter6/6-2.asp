<html>
<body>
	<% 
	Response.Write "http://localhost/index.asp?name=�� ��&age=22"
	Response.Write "<p>"										
	Response.Write Server.URLEncode("http://localhost/index.asp?name=�� ��&age=22")
	%>
</body> 
</html> 
