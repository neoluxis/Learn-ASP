<% '下面是两种建立Command对象的方式，大家可以分别复制使用 %>

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
