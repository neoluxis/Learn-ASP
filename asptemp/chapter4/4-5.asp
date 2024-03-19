<html>
<body>
	<h2 align="center">下面是您的个人信息</h2>
	<%
	Dim strName,strPwd,strSex,strLove,strCareer,strIntro    '为了引用方便，声明变量
	strName=Request.Form("txtName")     
	strPwd=Request.Form("txtPwd")
	strSex=Request.Form("rdoSex")
	strLove=Request.Form("chkLove")
	strCareer=Request.Form("sltCareer")
	strIntro=Request.Form("txtIntro")
	Response.Write "姓名：" & strName
	Response.Write "<br>密码：" & strPwd
	Response.Write "<br>性别：" & strSex
	Response.Write "<br>爱好：" & strLove
	Response.Write "<br>职业：" & strCareer
	Response.Write "<br>简介：" & strIntro
	%>
</body> 
</html>

