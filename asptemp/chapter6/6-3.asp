<html>
<body>
	<%
	'����ת�����������·��
	Response.Write Server.MapPath("/") & "<br>"				'ת��Ӧ�ó����Ŀ¼
	Response.Write Server.MapPath(".") & "<br>"				'ת����ǰĿ¼
	'���潫����·��ת��Ϊ����·��
	Response.Write Server.MapPath("/asptemp/chapter6/6-3.asp") & "<br>"	
	Response.Write Server.MapPath("/asptemp/chapter6/6-1.asp") & "<br>"		
	Response.Write Server.MapPath("/asptemp/chapter6/temp6/01.jpg") & "<br>"	
	Response.Write Server.MapPath("/asptemp/chapter5/5-1.asp") & "<br>"	
	'���潫���·��ת��Ϊ����·��
	Response.Write Server.MapPath("6-3.asp") & "<br>"		'ת������
	Response.Write Server.MapPath("6-1.asp") & "<br>"		'ת��ͬ�ļ����µ��ļ�
	Response.Write Server.MapPath("temp6/01.jpg") & "<br>"	'ת�����ļ����µ��ļ�
	Response.Write Server.MapPath("../chapter5/5-1.asp")	'ת�������ļ����µ��ļ�
	%>
</body>
</html>
