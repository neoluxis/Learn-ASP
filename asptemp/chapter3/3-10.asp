<% Option Explicit									'强制声明变量%>
<html>
<body>
	<%
	Dim strGrade
	strGrade="B"			'这里为了简单，直接赋值了，一般来说应该是传过来的参数，
							'比如从数据库中读出，或由程序计算得出。
	Select Case strGrade
	Case "A"       
		Response.Write "你的成绩很优秀，恭喜你。"
	Case "B"
		Response.Write "不错啊，继续努力。"
	Case "C"
		Response.Write "有点差，还需努力。"
	Case Else
		Response.Write "你需要加油了。"
	End Select
	Response.Write "<p>程序运行结束。"
	%>
</body> 
</html>



