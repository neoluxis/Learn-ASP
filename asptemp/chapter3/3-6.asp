<% Option Explicit										'ǿ�Ʊ������� %>
<html>
<body>
	<%
	Dim intM,intN,lngResult								'intM��intNΪʵ�ʲ���
	intM=3
	intN=4
	lngResult=mySquare(intM,intN)						'���ú���������ƽ����
	Response.Write "3��4��ƽ�����ǣ�" & lngResult
	'�����Ǻ�����������ʾ��������ƽ����
	Function mySquare(intA,intB)						'intA��intB����ʽ����
		Dim lngSum
		lngSum=intA^2+intB^2  
		mySquare=lngSum									'��ֵ��������������Ҫ
	End Function
	%>   
</body>
</html>
