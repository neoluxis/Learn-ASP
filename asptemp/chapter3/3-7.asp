<% Option Explicit										'ǿ�Ʊ������� %>
<!--#Include file="3-8.asp"-->
<html>
<body>
	<%
	Dim intM,intN,lngResult								'intM��intNΪʵ�ʲ���
	intM=3
	intN=4
	lngResult=mySquare(intM,intN)						'���ú���������ƽ����
	Response.Write "3��4��ƽ�����ǣ�" & lngResult
	%>   
</body>
</html>
