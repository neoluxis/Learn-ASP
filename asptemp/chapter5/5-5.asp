<html>
<body>
	<h2 align="center">我的主页</h2>
	<%
	Application.Lock										'先锁定
	Application("intAll")=Application("intAll")+1			'给Application变量赋值
	Application.Unlock										'解除锁定
	Dim intVisit
	intVisit=Application("intAll")							'获取Application变量
	Response.Write "<p>您是第" & intVisit & "位访客。"	
	%>
</body>
</html>



