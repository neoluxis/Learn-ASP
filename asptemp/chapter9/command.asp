<% '���������ֽ���Command����ķ�ʽ����ҿ��Էֱ���ʹ�� %>

<%
Dim conn 
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Dsn=address2"
Dim cmd
Set cmd=Server.CreateObject("ADODB.Command")
cmd.ActiveConnection=conn
%>

<%
Dim cmd
Set cmd= Server.CreateObject("ADODB.Command")
cmd.ActiveConnection="Dsn=address2"                     
%>
