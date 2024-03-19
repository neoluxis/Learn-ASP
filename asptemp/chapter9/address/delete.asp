<%
'===============================================================================================
'delete.asp说明
'1. 本页面只是简单地删除了记录。其中的数据库连接语句放在了connection.asp文件中。
'2. 本节中关闭connection对象的语句实际上也可以省略。因为该文件执行完毕后也会自动关闭的。
'===============================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<% 
'以下删除记录，注意这里要用到由index.asp传过来的要删除的记录的ID
Dim strSql
strSql="Delete From tbAddress Where ID=" & Request.QueryString("ID") 
conn.Execute(strSql)

'现在关闭Connection对象
conn.Close
Set conn=Nothing

'删除完毕后，返回首页
Response.Redirect "index.asp"								
%>
