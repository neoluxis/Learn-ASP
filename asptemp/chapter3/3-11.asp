<% Option Explicit						'ǿ����������%> 
<html>
<body>
	<%
	Dim lngSum,I						'lngSum������Ž����I��ѭ������������
	lngSum=0							'��lngSum����ֵ0 
	For I=1 To 100						'����������I��1ѭ����100
		lngSum=lngSum+I^2									
	Next
	Response.Write "1��100��ƽ����=" & lngSum           
	%>
</body> 
</html>
