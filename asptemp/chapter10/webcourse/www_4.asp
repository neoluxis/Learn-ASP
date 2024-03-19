<html>
<body>
	<h2 align='center'>第四讲 利用ASP开发动态网页</h2>
	<table width="80%" border="1" align="center">
		<tr><td>本节主要讲授利用ASP开发动态网页......<br><br><br><br><br></td></tr>
	</table>
	<p align="center">
	<%
	Dim link										'声明一个NextLink对象实例
	Set link=Server.CreateObject("MSWC.NextLink")
	'下面是上一篇的超链接
	Response.Write "【<a href='" & link.GetPreviousURL("link.txt") & "'>" & link.GetPreviousDescription("link.txt") & "</a>】"
	'下面是返回目录的超链接
	Response.Write "【<a href='10-15.asp'>目录</a>】"
	'下面是下一篇的超链接
	Response.Write "【<a href='" & link.GetNextURL("link.txt") & "'>" & link.GetNextDescription("link.txt") & "</a>】"
	%>
</body>