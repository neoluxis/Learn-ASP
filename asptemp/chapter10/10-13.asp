<html>
<body>
	<h2 align="center">客户端浏览器特性</h2>
	<%
	Dim bc										'声明一个BrowserType对象实例
	Set bc=Server.CreateObject("MSWC.BrowserType")
	Response.Write "<br>浏览器类型：" & bc.Browser
	Response.Write "<br>浏览器版本：" & bc.Version
	Response.Write "<br>支持Cookies否：" & bc.Cookies
	Response.Write "<br>支持Java小程序否：" & bc.JavaApplets 
	%>
</body> 
</html>
 

