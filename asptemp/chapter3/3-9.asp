<% Option Explicit									'强制声明变量%>
<html>
<body>
	<%
	Dim intGrade
	intGrade=86			'这里为了简单，直接赋值了，一般来说应该是传过来的参数，
						'比如从数据库中读出，或由程序计算得出。
	If intGrade>=85 Then       
		Response.Write "你的成绩很优秀，恭喜你。"
	Elseif intGrade>=70 And intGrade<85 Then
		Response.Write "不错啊，继续努力。"
	Elseif intGrade>=60 And intGrade<70 Then
		Response.Write "有点差，还需努力。"
	Else
		Response.Write "你需要加油了。"
	End If
	Response.Write "<p>程序运行结束。"
	%>
</body> 
</html>


