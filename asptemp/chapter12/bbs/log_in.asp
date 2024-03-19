<%
'=========================================================================================
'用户登录页
'1. 其中就是判断用户输入的用户名和密码是否正确？
'2. 如果正确，就将用户名、E-mail保存到Session中，然后重定向回首页。
'3. 如果不正确，就输出错误提示信息
'=========================================================================================
%>

<% Option Explicit %>
<!--#Include File="odbc_connection.asp"-->
<%
'下面首先查看用户名和密码是否正确
Dim strSql,rs
strSql="Select strUserId,strEmail From tbUsers Where strUserId='" & Request.Form("txtUserId") & "' And strPwd='" & Request.Form("txtPwd") & "'"
Set rs=conn.Execute(strSql)
If Not rs.Eof And Not rs.Bof Then
	'如果有记录，表示有该用户，则将用户名和Email保存到Session中
	Session("strUserId")=rs("strUserId")
	Session("strEmail")=rs("strEmail")
	Response.Redirect "index.asp"					'重定向到首页
Else
	'如果没有记录，表示用户名或密码可能不正确，请给出提示信息
	Response.Write "对不起，用户名或密码有误，请<a href='index.asp'>返回首页</a>重新登录"
End If
%>
