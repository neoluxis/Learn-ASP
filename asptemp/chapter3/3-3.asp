<% Option Explicit                      			'放在程序首行，强制声明变量%>
<html>
<body>
	<%
	Dim strPwd										'声明一个字符串变量
	strPwd="   abcd*abcd   "						'给字符串变量赋值
	strPwd=Trim(strPwd)								'去掉字符串两边的空格
	Response.Write "<p>此时字符串是：" & strPwd 
	strPwd=Replace(strPwd,"*","")					'将*替换为空字符串，也就是去掉*
	Response.Write "<p>此时字符串是：" & strPwd 
	strPwd=Left(strPwd,6)							'从字符串左边开始取6个字符
	Response.Write "<p>最后字符串是：" & strPwd 
	%>
</body> 
</html>

