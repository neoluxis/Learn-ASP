<% Option Explicit							'ǿ����������%>                    
<html>
<body>
	<%
	Dim lngSum,I							'lngSum������������I��������ѭ��
	lngSum=0								'��lngSum����ֵ 
	I=1												
	Do While I<=100							'��IС�ڵ���100ʱִ��ѭ��
		lngSum=lngSum+I^2					
		I=I+1								'I��ֵ����1
	Loop
	Response.Write "1��100��ƽ����=" & lngSum           
	%>
</body> 
</html>
