<% 
'以下连接数据库，建立一个Connection对象实例conn
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection") 
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
conn.Open strConn
'以下删除记录，注意这里要用到由index.asp传过来的要删除的记录的ID
Dim strSql
strSql="Delete From tbAddress Where ID=" & Request.QueryString("ID") 
conn.Execute(strSql)
'删除完毕后，返回首页
Response.Redirect "index.asp"								
%>
