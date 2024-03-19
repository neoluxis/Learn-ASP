<%
'================================================================================================
'聊天室首页
'1. 在该页面让来访者输入用户名，如果用户名没有被占用，就可以进入聊天室。
'2. 在进入聊天室前，会调用函数在聊天信息中添加欢迎来访者的信息，另外，会将当前用户添加到在线人员名单中。
'3. 在线人员名单实际上是保存在Application中的一个数组。
'================================================================================================
%>

<%option explicit%>
<!--#Include File="function.asp"-->
<!--#Include File="config.asp"-->
<html>
<head>
	<title>欢迎进行我的聊天室</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="chat.css">
	<script language="JavaScript">
		//该函数用来进行客户端验证
		function check_Null(){
			if (document.frmLogIn.txtUserName.value==""){
				alert("用户名不能为空!");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<!--注释：下面要从配置文件中读取常量conChatTitle-->
	<h1 align="center"><font face="黑体"><%=conChatTitle%></font></h1>

	<!--注释：下面调用函数显示在线人数-->
	<p align="center">当前共有<%=AllUserName()%>人在线

	<form name="frmLogIn" method="POST" action="index.asp" onsubmit="JavaScript:return check_Null();">
		<center>
		请输入您的用户名<input type="text" name="txtUserName" size="30">
		<input type="submit" value="进入聊天室">
		</center>
	</form>

	<%
	'如果用户输入了用户名，就执行下面的程序
	If Request.Form("txtUserName")<>"" Then
		'以下返回用户名和用户IP地址，并保存到Session中
		Dim strUserName,IP
		strUserName=Request.Form("txtUserName")				'获取用户姓名
		IP=Request.ServerVariables("REMOTE_ADDR")			'获取IP地址
		Session("strUserName")=strUserName					'保存到Session中待用                   
		Session("IP")=IP									'同上

		'下面调用函数检查是否已经有人使用该用户名，如无人使用，则添加到用户名单中
		If GetUserName(strUserName)=True Then
			Response.Write "<p align='center'>已经有人使用该名称，请重新输入"
		Else
			'下面调用函数将该用户名保存到在线人员名单中
			Call addUserName(strUserName)
			'下面调用函数将该用户来到的信息添加到聊天信息中
			Call AddUserCome(strUserName,IP)
			'重定向到聊天主页面
			Response.Redirect "whole.asp"
		End If
	End If
	%>
</body>
</html>