<% Option Explicit									'ǿ����������%>
<html>
<body>
	<%
	Dim strGrade
	strGrade="B"			'����Ϊ�˼򵥣�ֱ�Ӹ�ֵ�ˣ�һ����˵Ӧ���Ǵ������Ĳ�����
							'��������ݿ��ж��������ɳ������ó���
	Select Case strGrade
	Case "A"       
		Response.Write "��ĳɼ������㣬��ϲ�㡣"
	Case "B"
		Response.Write "����������Ŭ����"
	Case "C"
		Response.Write "�е�����Ŭ����"
	Case Else
		Response.Write "����Ҫ�����ˡ�"
	End Select
	Response.Write "<p>�������н�����"
	%>
</body> 
</html>



