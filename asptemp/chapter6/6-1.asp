<html>
<body>
	<% 
	Response.Write "<a href='http://www.sohu.com'>หับü</a>"
	Response.Write "<p>"										
	Response.Write Server.HTMLEncode("<a href='http://www.sohu.com'>หับü</a>")
	%>
</body> 
</html> 
