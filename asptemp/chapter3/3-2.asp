<% Option Explicit                      		'放在程序首行，强制声明变量%>
<html>
<body>
	<%
	Dim intMyRnd
	Randomize Timer()			'该语句用于初始化随机种子，否则每次产生的值都一样
	intMyRnd = Int((10 * Rnd()) + 1)			'产生随机整数
	Response.Write intMyRnd
	%>
</body> 
</html>
