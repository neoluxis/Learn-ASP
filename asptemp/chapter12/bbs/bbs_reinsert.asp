<%
'=========================================================================================
'回复文章页面
'1. 当在bbs_particular.asp中提交回复文章后，就会打开本页面。其中就是添加一条记录，并且更新相关的表。
'2. 在添加记录时，因为内容和E-mail允许为空，所以这里组成SQL语句有一些复杂，首先组成SQL语句的前半部分， 然后组成SQL语句的后半部分。 最后组成一条完整的SQL语句。
'3. 要注意，回复文章的父文章ID是用隐藏文本框传递过来的。
'4. 另外，会获取利用隐藏文本框传递过来的ID，然后在重定向bbs_list.asp时，继续将这两个变量传递回去。
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="function.asp"-->
<%
'下面首先获取有关值
Dim strTitle,strBody,intLayer,intFatherId,intChild,intHits,strIP,strUserId,strEmail,intForumId,intPage
strTitle=myDangerEncode(Request.Form("txtTitle"))						'返回文章标题
strBody=myDangerEncode(Request.Form("txtBody"))							'返回文章内容
If Session("strUserId")<>"" Then
	strUserId=Session("strUserId")										'返回作者用户名
Else
	strUserId="过客"													'如果是未登录用户，则统一命名为 过客
End If
strEmail=myDangerEncode(Request.Form("txtEmail"))						'返回作者E-mail
intFatherId=Request.Form("txtFatherId")									'获取从隐藏文本框传回来的父文章Id
intLayer=2																'这是第2层
intChild=0																'回复文章数目为0，第2层不允许回复
intHits=0																'点击数为0
strIP=Request.ServerVariables("remote_addr")							'作者IP地址
intForumId=Request.Form("txtForumId")									'从隐藏文本框获取当前栏目编号
intPage=Request.Form("txtPage")											'从隐藏文本框获取页码，后面重定向会用到

'以下将文章保存到数据库
Dim sqlA,sqlB,strSql
sqlA="Insert Into tbBBS(strTitle,intLayer,intFatherId,intChild,intHits,strIP,strUserId,dtmSubmit,intForumId"
sqlB="Values('" & strTitle & "'," & intLayer & "," & intFatherId & "," & intChild & "," & intHits & ",'" & strIP & "','" & strUserId & "',#" & Now() & "#," & intForumId 
If strBody<>"" Then													'如果有内容，则添加
	sqlA=sqlA & ",strBody"
	sqlB=sqlB & ",'" & strBody & "'"
End If
If strEmail<>"" Then												'如果有E-mail，则添加
	sqlA=sqlA & ",strEmail"
	sqlB=sqlB & ",'" & strEmail & "'"
End If
strSql=sqlA & ") " & sqlB & ")"
conn.Execute(strSql)

'下面将父文章的回复数加1
strSql="Update tbBBS Set intChild=intChild+1 where ID=" & intFatherId
conn.Execute(strSql)

'下面两句将论坛的文章数加1
strSql="Update tbForum Set lngForumCount=lngForumCount+1 where ID=" & intForumId
conn.Execute(strSql)

'下面将该用户的回复文章数加1
strSql="Update tbUsers Set intReArticle=intReArticle+1 Where strUserId='" & Session("strUserId") & "'"
conn.Execute(strSql)

'关闭对象
conn.Close
Set conn=Nothing

'重定向回BBS列表页，请记住再将当前栏目号码和页码传回去
Response.Redirect "bbs_list.asp?intForumId=" & intForumId & "&intPage=" & intPage
%>
