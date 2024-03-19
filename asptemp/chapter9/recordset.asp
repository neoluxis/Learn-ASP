<% '下面是五种建立Recordset对象的方式，大家可以分别复制使用 %>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open("Dsn=address2")
Dim rs
Set rs=conn.Execute("Select * From tbAddress")
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open("Dsn=address2")
Dim cmd
Set cmd= Server.CreateObject("ADODB.Command")
cmd.ActiveConnection=conn
cmd.CommandType=1
cmd.CommandText="Select * From tbAddress"
Dim rs
Set rs=cmd.Execute
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open("Dsn=address2")
Dim rs
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From tbAddress",conn,1,2,1 
%>

<%
Dim conn
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open("Dsn=address2")
Dim cmd
Set cmd= Server.CreateObject("ADODB.Command")
cmd.ActiveConnection=conn
cmd.CommandType=1
cmd.CommandText="Select * From tbAddress"
Dim rs
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open cmd,,1,3 
%>

<%
Dim rs
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From tbAddress","Dsn=address2",1,2,1 
%>
