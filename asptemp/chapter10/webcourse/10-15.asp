<html>
<body>
	<h2 align="center">�����ļ��������������Ŀ¼ҳ</h2>
	<%
	Dim link													'����һ��NextLink����ʵ��
	Set link=Server.CreateObject("MSWC.NextLink")
	Dim I
	For I=1 To link.GetListCount("link.txt")					'��ѭ������д�����е��ļ�����
		Response.Write "<a href='" & link.GetNthURL("link.txt",I) & "'>" & link.GetNthDescription("link.txt",I) & "</a><p>"
	Next
	%>
</body> 
</html>



