<html>
<body>
	<h2 align="center"> Error�����Errors����ʾ��</h2>
	<%
	On Error Resume Next						'���������������ִ����һ��
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open "Dsn=temp9"						'�����һ�����������Դ
	'��������ѭ��������еĴ�����������
	Dim I,objErr								'I��ѭ��������objErr�Ǵ���������
	For I =0 To conn.Errors.Count-1                                 
		Set objErr=conn.Errors.Item(I)				'�����������ʵ��objErr
		Response.Write "<br>�����ţ�" & objErr.Number 
		Response.Write "<br>����������" & objErr.Description 
		Response.Write "<br>����ԭ��" & objErr.Source 
		Response.Write "<br>��ʾ���֣�" & objErr.HelpContext 
		Response.Write "<br>�����ļ���" & objErr.HelpFile 
		Response.Write "<br>ԭʼ����" & objErr.NativeError 
	Next
	%>
</body>
</html>

