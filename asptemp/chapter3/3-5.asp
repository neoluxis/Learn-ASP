<% Option Explicit								'ǿ�Ʊ������� %>
<html>
<body>
	<%
	Dim intM,intN								'intM��intNΪʵ�ʲ���
	intM=3
	intN=4
	Call mySquare(intM,intN)					'�����ӳ�����ʾ���
	'�������ӳ�������������������ƽ����
	Sub mySquare(intA,intB)						'intA��intB����ʽ����
		Dim lngSum
		lngSum=intA^2+intB^2  
		Response.Write "3��4��ƽ�����ǣ�" & lngSum    
	End Sub
	%>   
</body>
</html>

