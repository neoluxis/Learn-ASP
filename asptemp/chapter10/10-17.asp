<html>
<body>
	<h2 align="center">计数器组件应用示例</h2>
	<%
	Dim count											'声明一个组件实例变量
	Set count=Server.CreateObject("MSWC.PageCounter")
	count.PageHit										'将当前网页访问次数加1
	Dim intVisit
	intVisit=count.Hits()								'获取当前网页访问次数
	Response.Write "您是第" & intVisit & "位访客"
	%> 
</body> 
</html>
 


