<% Option Explicit									'ǿ����������%>
<html>
<body>
	<%
	Dim intGrade
	intGrade=86			'����Ϊ�˼򵥣�ֱ�Ӹ�ֵ�ˣ�һ����˵Ӧ���Ǵ������Ĳ�����
						'��������ݿ��ж��������ɳ������ó���
	If intGrade>=85 Then       
		Response.Write "��ĳɼ������㣬��ϲ�㡣"
	Elseif intGrade>=70 And intGrade<85 Then
		Response.Write "����������Ŭ����"
	Elseif intGrade>=60 And intGrade<70 Then
		Response.Write "�е�����Ŭ����"
	Else
		Response.Write "����Ҫ�����ˡ�"
	End If
	Response.Write "<p>�������н�����"
	%>
</body> 
</html>


