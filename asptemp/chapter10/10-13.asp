<html>
<body>
	<h2 align="center">�ͻ������������</h2>
	<%
	Dim bc										'����һ��BrowserType����ʵ��
	Set bc=Server.CreateObject("MSWC.BrowserType")
	Response.Write "<br>��������ͣ�" & bc.Browser
	Response.Write "<br>������汾��" & bc.Version
	Response.Write "<br>֧��Cookies��" & bc.Cookies
	Response.Write "<br>֧��JavaС�����" & bc.JavaApplets 
	%>
</body> 
</html>
 

