<%
'=========================================================================================
'退出登录页
'1. 这里只是将Session清空，然后重定向回首页。
'=========================================================================================
%>

<%
'注销登录用户
Session.Abandon					'清空Session，这样所有Session信息就没有了
Response.Redirect "index.asp"	'重定向回首页
%>