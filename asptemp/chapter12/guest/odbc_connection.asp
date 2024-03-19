<%
'========================================================================================================
'这是数据库连接文件，专门用来连接数据库。在其他页面中可以包含本页面，就相当于将如下语句写到别的页面中一样。
'=========================================================================================================

'以下连接数据库，建立一个Connection对象实例conn
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("guest.mdb")
conn.Open strConn 
%>

