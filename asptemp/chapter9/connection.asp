<% '下面是六种数据库连接方式，大家可以分别复制使用 %>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Dsn=address2"
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "address2"
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\Inetpub\wwwroot\asptemp\chapter9\address.mdb"
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.MapPath("address.mdb")
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\asptemp\chapter9\address.mdb"
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
%>
