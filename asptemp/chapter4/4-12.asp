<% Response.Buffer=True							'ע�⣺����и����       %>
<html>
<body>
	<% 
	Dim varNumber                        		'����һ�����ʴ�������
	varNumber=Request.Cookies("intVisit") 		'��ȡCookieֵ
	If varNumber="" then						'���varNumber=""����ʾ��û�ж����Cookie
		varNumber=1                    			'����ǵ�һ�Σ�������ʴ���Ϊ1
	Else
		varNumber=CInt(varNumber)+1            	'������ǵ�һ�Σ�������ʴ�����1
	End If
	Response.Write "���ǵ�" & varNumber & "�η��ʱ�վ"   
	Response.Cookies("intVisit")=varNumber   						'���µķ��ʴ������浽Cookie��
	Response.Cookies("intVisit").Expires=DateAdd("d",30,Date())   	'������Ч��Ϊ30��
	%>
</body>
</html>

