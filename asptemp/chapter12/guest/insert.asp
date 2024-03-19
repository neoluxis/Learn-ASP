<%
'================================================================================================
'添加留言文件
'1. 这里会获取首页提交过来的表单，然后将留言添加到数据库中。
'2. 因为内容和E-mail可以省略，所以这里组成SQL语句的代码比较复杂。实际上是分成前后两半部分，然后分别组织， 最后再形成一个标准的Insert语句。
'3. 如果用户提交的表单中包括了英文的单引号，就会和SQL语句中的单引号发生冲突，从而导致错误。 所以这里调用myDangerEncode将其中的单引号替换为了两个连续的单引号。这样不会发生错误， 而且写到数据库中后自然变成了1个单引号。
'================================================================================================
%>

<% Option Explicit						'强制声明变量%>
<!--#Include File="odbc_connection.asp"-->
<!--#Include File="function.asp"-->
<%
'下面首先获取提交过来的数据，注意其中会调用函数处理危险字符
Dim strTitle,strBody,strName,strEmail
strTitle=myDangerEncode(Request.Form("txtTitle"))				
strBody=myDangerEncode(Request.Form("txtBody"))
strName=myDangerEncode(Request.Form("txtName"))
strEmail=myDangerEncode(Request.Form("txtEmail"))
'下面开始添加记录，因为内容和E-mail可以省略，所以将SQL语句分成前后两段分别建立
Dim sqlA,sqlB,strSql
sqlA="Insert Into tbGuest(strName,strTitle,dtmSubmit"
sqlB="values('" & strName & "','" & strTitle & "',#" & Now() & "#"
If strBody<>"" Then										'如果提交了内容，才添加
	sqlA=sqlA & ",strBody"
	sqlB=sqlB & ",'" & strBody & "'"
End If
If strEmail<>"" Then									'如果提交了E-mail，才添加
	sqlA=sqlA & ",strEmail"
	sqlB=sqlB & ",'" & strEmail & "'"
End If
'下面将前后两段组成完整的SQL语句，并执行添加
strSql=sqlA & ") " & sqlB & ")"
conn.Execute(strSql)											
'关闭对象
conn.Close
Set conn=Nothing
Response.Redirect "index.asp"									
%>
