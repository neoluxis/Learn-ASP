<% '�����Ǽ�������SQL���ݿ�ķ�ʽ����ҿ��Էֱ���ʹ�� %>

<%
Dim conn 
Set conn=Server.CreateObject("ADODB.Connection") 
conn.Open "Dsn=test;Uid=jjshang;Pwd=123456"          
%>

<%
Dim conn 
Set conn=Server.CreateObject("ADODB.Connection") 
conn.Open "test","jjshang", "123456"  
%>

<%
Dim conn 
Set conn=Server.CreateObject("ADODB.Connection") 
conn.Open "Driver={SQL Server};Server=localhost;Database=sqltest;Uid=jjshang;Pwd=123456"
%>

<%
Dim conn 
Set conn=Server.CreateObject("ADODB.Connection") 
conn.Open "Provider=SQLOLEDB;Data Source=localhost;initial Catalog=sqltest;Uid=jjshang;Pwd=123456"
%>
