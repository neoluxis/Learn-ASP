<html>
<body>
	<form name="form1" method="post" action="">
		�ǳƣ�<input type="text" name="txtName" size="10">
		���ԣ�<input type="text" name="txtSay" size="30">
		<input type="submit" value=" ���� ">
	</form>
	<%
	'����ύ�˱����ͽ�����������ӵ�Application������
	If Trim(Request.Form("txtName"))<>"" And Trim(Request.Form("txtSay"))<>"" Then
		'�����Ȼ�ȡ���η����ַ��������������˺ͷ�������
		Dim strSay
		strSay=Request.Form("txtName") & "˵��" & Request.Form("txtSay") & "<br>"
		'���潫���η�����ӵ�����������
		Application.Lock						'������
		Application("strChat")=strSay & Application("strChat")
		Application.Unlock						'�������
	End If
	%>
</body> 
</html>
