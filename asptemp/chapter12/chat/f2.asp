<%
'================================================================================================
'发言页面
'1. 该页只是将用户提交的信息，通过调用函数添加到聊天信息中。
'2. 注意说话颜色框，其中要保留上一次使用的颜色，所以这里会默认选中上一次使用的颜色。
'3. 最后的JavaScript语句用于定位光标，另外，发言后立即刷新f2.asp，以显示最新发言。
'================================================================================================
%>

<%option explicit											'强制声明变量%>
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<meta http-equiv='content-type' content='text/html; charset=gb2312'>
	<title>发言区</title>
	<link rel="stylesheet" href="chat.css">
</head>
<body  background="images/bg.gif">
	<form name="frmInput" method="POST" action="">
	<br><%=Session("strUserName")%>说:
	<input type="text" name="txtSays" size="50" maxlength="100" >       
	<input type="submit" value=" 发 言 " > 
	<p>说话颜色:        
	<select name="txtSaysColor">        
		<option value="#000000" 
		<%If Request.Form("txtSaysColor")="#000000" Then Response.Write "selected" %>
		>默认</option>    
		<option value="#0000FF" 
		<%If Request.Form("txtSaysColor")="#0000FF" Then Response.Write "selected" %>
		>蓝色</option>    
		<option value="#FF0000" 
		<%If Request.Form("txtSaysColor")="#FF0000" Then Response.Write "selected" %>
		>红色</option>    
		<option value="#FFFF00" 
		<%If Request.Form("txtSaysColor")="#FFFF00" Then Response.Write "selected" %>
		>黄色</option>    
	</select>    
	表情:       
	<select name="txtEmote">       
		<option value="" selected>无</option>    
		<option value="高兴地">高兴</option>    
		<option value="伤心地">伤心</option> 
		<option value="气愤地">气愤</option>    
		<option value="紧张兮兮地">紧张</option>    
	</select> 
	<a href="#" onclick="JavaScript:window.open('about:blank','_top').close();">离开聊天室</a>
	</form>
	<%
	'如果提交了发言信息，就执行下面的语句
	If Request.Form("txtSays")<>"" Then
		'下面先获取有关数据，这里调用了函数
		Dim strSaysColor,strEmote,strSays
		strSaysColor=Request.Form("txtSaysColor")				'获取发言者颜色
		strEmote=Request.Form("txtEmote")						'获取发言者表情
		strSays=Trim(Request.Form("txtSays"))					'去掉前后的空格
		'下面调用函数将本次发言添加到聊天信息中
		Call AddUserSay(Session("strUserName"),strSaysColor,strEmote,strSays)
	End If
	%>
	<!--注释：下面利用JavaScript将光标定位到发言文本框，并刷新聊天内容框f1.asp-->
	<Script language="JavaScript">
		document.frmInput.txtSays.focus();						//定位光标
		top.main.location.href="f1.asp";						//自动刷新f1.asp
	</Script>
</body>  
</html>  
