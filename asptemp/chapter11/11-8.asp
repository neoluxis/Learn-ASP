<html>
<body>
	<h2 align="center">���߷���Emailʾ��</h2>
	<form name="frmEmail" method="POST" action="">
	<table border="0" align="center">
		<tr><td>������</td><td>aspbook@163.com</td></tr>
		<tr><td>�ռ���</td><td><input type="text" name="txtRecipient" size="30"></td></tr>
		<tr><td>����</td><td><input type="text" name="txtSubject" size="40"></td></tr>
		<tr><td>����</td>
		<td><textarea name="txtBody" rows="4" cols="40"></textarea></td></tr>
		<tr><td></td><td><input type="submit" value=" ���� "></td></tr>
	</table>
	</form>
	<%
	'�жϣ�����ռ��˺����ⶼ��Ϊ�գ����������
	If Request("txtRecipient")<>"" And Request("txtSubject")<>"" Then
		Dim jmail										'����Message����ʵ��
		Set jmail=Server.CreateObject("Jmail.Message")		    
		jmail.From="aspbook@163.com"					'������ 
		jmail.AddRecipient Request("txtRecipient")		'�ռ��� 
		jmail.Subject=Request("txtSubject")				'���� 
		jmail.Body=Request("txtBody")					'����                         
		jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"		'ִ�з���
		jmail.Close										'�رն���
		Response.Write "<p align='center'>�ɹ�����"
	End If
	%>
</body> 
</html>


