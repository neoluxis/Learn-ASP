<html>
<body>
	<h2 align='center'>���Ľ� ����ASP������̬��ҳ</h2>
	<table width="80%" border="1" align="center">
		<tr><td>������Ҫ��������ASP������̬��ҳ......<br><br><br><br><br></td></tr>
	</table>
	<p align="center">
	<%
	Dim link										'����һ��NextLink����ʵ��
	Set link=Server.CreateObject("MSWC.NextLink")
	'��������һƪ�ĳ�����
	Response.Write "��<a href='" & link.GetPreviousURL("link.txt") & "'>" & link.GetPreviousDescription("link.txt") & "</a>��"
	'�����Ƿ���Ŀ¼�ĳ�����
	Response.Write "��<a href='10-15.asp'>Ŀ¼</a>��"
	'��������һƪ�ĳ�����
	Response.Write "��<a href='" & link.GetNextURL("link.txt") & "'>" & link.GetNextDescription("link.txt") & "</a>��"
	%>
</body>