<% Option Explicit					'ǿ����������%>                    
<html>
<body>
	<%
	Dim intA(9,9)					'����һ��10��10�еĶ�ά����
	Dim I,J							'I�����ѭ��������������J���ڲ�ѭ������������
	For I=0 To 9					'���ѭ��
		For J=0 To 9				'�ڲ�ѭ��
			intA(I,J)=10			'��ÿһ��Ԫ�ظ���ֵ10
		Next
	Next
	Response.Write "�������н�����" 
	%>
</body> 
</html>
