<html>
<body>
   <h2 align="center">�Զ�����HTML�ļ�</h2>
	<%
	Dim fso										'����һ��FileSystemObject����ʵ��
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim tsm										'����һ��TextStream����ʵ��
	Set tsm=fso.CreateTextFile("C:\Inetpub\wwwroot\asptemp\chapter10\temp.htm",True)
	tsm.WriteLine "<html>"						'���ļ���д�����ݣ�����ͬ
	tsm.WriteLine "<head>"						
	tsm.WriteLine "<title>�ҵ���ҳ</title>"	
	tsm.WriteLine "</head>"					
	tsm.WriteLine "<body>"
	tsm.WriteLine "<h2 align='center'>�ҵ���ҳ</h2>"
	tsm.WriteLine "<p align='center'>��ӭ��ҷ���"
	tsm.WriteLine "</body>"						
	tsm.WriteLine "</html>"						
	tsm.Close									'�ر�TextStream����
	Response.Write "�Ѿ��ɹ������ļ������Լ��򿪲鿴��"
	%>
</body> 
</html>
