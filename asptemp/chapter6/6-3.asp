<html>
<body>
	<%
	'下面转换两个特殊的路径
	Response.Write Server.MapPath("/") & "<br>"				'转换应用程序根目录
	Response.Write Server.MapPath(".") & "<br>"				'转换当前目录
	'下面将绝对路径转换为物理路径
	Response.Write Server.MapPath("/asptemp/chapter6/6-3.asp") & "<br>"	
	Response.Write Server.MapPath("/asptemp/chapter6/6-1.asp") & "<br>"		
	Response.Write Server.MapPath("/asptemp/chapter6/temp6/01.jpg") & "<br>"	
	Response.Write Server.MapPath("/asptemp/chapter5/5-1.asp") & "<br>"	
	'下面将相对路径转换为物理路径
	Response.Write Server.MapPath("6-3.asp") & "<br>"		'转换自身
	Response.Write Server.MapPath("6-1.asp") & "<br>"		'转换同文件夹下的文件
	Response.Write Server.MapPath("temp6/01.jpg") & "<br>"	'转换子文件夹下的文件
	Response.Write Server.MapPath("../chapter5/5-1.asp")	'转换其他文件夹下的文件
	%>
</body>
</html>
