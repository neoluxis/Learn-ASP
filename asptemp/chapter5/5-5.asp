<html>
<body>
	<h2 align="center">�ҵ���ҳ</h2>
	<%
	Application.Lock										'������
	Application("intAll")=Application("intAll")+1			'��Application������ֵ
	Application.Unlock										'�������
	Dim intVisit
	intVisit=Application("intAll")							'��ȡApplication����
	Response.Write "<p>���ǵ�" & intVisit & "λ�ÿ͡�"	
	%>
</body>
</html>



