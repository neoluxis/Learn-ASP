<% Option Explicit						'ǿ����������%>                    
<html>
<body>
	<%
	Dim strA(2),strSum,Item				'strSum������������Item������������Ԫ��
	strA(0)="a"							'������Ԫ�ظ�ֵ
	strA(1)="b"
	strA(2)="c"
	For Each Item in strA				'ִ��ѭ����ȡ��ÿ��Ԫ��
		strSum=strSum & Item			
	Next
	Response.Write "ȫ������Ԫ���γɵ��ַ����ǣ�" & strSum
	%>
</body> 
</html>

