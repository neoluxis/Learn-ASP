<html>
<body>
	<h2 align="center">���������ĸ�����Ϣ</h2>
	<%
	Dim strName,strPwd,strSex,strLove,strCareer,strIntro    'Ϊ�����÷��㣬��������
	strName=Request.Form("txtName")     
	strPwd=Request.Form("txtPwd")
	strSex=Request.Form("rdoSex")
	strLove=Request.Form("chkLove")
	strCareer=Request.Form("sltCareer")
	strIntro=Request.Form("txtIntro")
	Response.Write "������" & strName
	Response.Write "<br>���룺" & strPwd
	Response.Write "<br>�Ա�" & strSex
	Response.Write "<br>���ã�" & strLove
	Response.Write "<br>ְҵ��" & strCareer
	Response.Write "<br>��飺" & strIntro
	%>
</body> 
</html>

