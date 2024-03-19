<html>
<body>
	<% 
	Response.Write "http://localhost/index.asp?name=ÕÅ ÔÆ&age=22"
	Response.Write "<p>"										
	Response.Write Server.URLEncode("http://localhost/index.asp?name=ÕÅ ÔÆ&age=22")
	%>
</body> 
</html> 
