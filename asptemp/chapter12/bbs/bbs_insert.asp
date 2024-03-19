<%
'=========================================================================================
'发表新文章页面
'1. 这其实就是一个普通的添加记录的页面。添加记录后，还需要更新另外两个表， 以便更新该栏目文章数目和该用户发表文章数目。
'2. 在添加记录时，因为内容和E-mail允许为空，所以这里组成SQL语句有一些复杂，首先组成SQL语句的前半部分， 然后组成SQL语句的后半部分。 最后组成一条完整的SQL语句。
'3. 另外要注意，这里会将传递过来的intForumId先存放在隐藏文本框中，然后提交表单后，可以继续获取该变量， 然后在重定向BBS列表页时继续将该变量传递回去。
'=========================================================================================
%>

<%option explicit%>
<!--#Include file="odbc_connection.asp"-->
<!--#Include file="function.asp"-->
<html>
<head>
	<title>bbs论坛</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="bbs.css" rel="stylesheet" type="text/css">
	<script language="JavaScript">
		<!--
		function check_Null(){
			if (document.frmInsert.txtTitle.value==""){
				alert("主题不能为空!");
				return false;
			}
			if (document.frmInsert.txtUserId.value==""){
				alert("用户名不能为空!");
				return false;
			}
			return true;
		}
		// -->
	</script>
</head>

<body>
	<h1 align="center">发表文章</h1>
	<form name="frmInsert" action="" method="POST" onsubmit="javascript: return check_Null();">
	<table border="0" width="80%" cellspacing="0" cellpadding="0" class="s2" align="center">
		<tr>                                        
			<td width="20%">文章主题：</td>
			<td><input type="text" name="txtTitle" size="50" ></td>
		</tr>
		<tr>
			<td>文章内容：</td>
			<td><textarea rows="10" name="txtBody" cols="60" ></textarea></td>
		</tr>
		<tr>                                        
			<td>用户名：</td>
			<td>
				<!--注释：下面根据是否已经登录显示不同的用户名，如没有登录，则为“过客”-->
				<%If Session("strUserId")<>"" Then%>
					<input type="text" name="txtUserId" size="20" value="<%=Session("strUserId")%>" disabled>
				<%Else%>
					<input type="text" name="txtUserId" size="20" value="过客" disabled>
				<%End If%>
			</td>               
		</tr>
		<tr>                                        
			<td>E-mail：</td>
			<td><input type="text" name="txtEmail" size="30" value="<%=Session("strEmail")%>"></td>
		</tr>
		<tr>
			<td colspan="2"  align="center">
				<br>
				<!--注释：这里用隐藏文本框将intForumId变量传递了过去-->
				<input type="hidden" name="txtForumId" value="<%=Request.QueryString("intForumId")%>">
				<input type="submit" value="确定提交">&nbsp; 
				<input type="button" value="返回论坛" onclick="javascript:history.back();">
			</td>
		</tr>
	</table>
	</form>    
	
	<%
	'如果提交了表单，就执行下面的添加语句
	If Trim(Request.Form("txtTitle"))<>"" Then
		Dim strTitle,strBody,intLayer,intFatherId,intChild,intHits,strIP,strUserId,strEmail,intForumId
		strTitle=myDangerEncode(Trim(Request.Form("txtTitle")))		'返回文章标题
		strBody=myDangerEncode(Request.Form("txtBody"))				'返回文章内容
		If Session("strUserId")<>"" Then
			strUserId=Session("strUserId")							'返回作者用户名
		Else
			strUserId="过客"										'如果是未登录用户，则统一命名为 过客
		End If
		strEmail=myDangerEncode(Request.Form("txtEmail"))			'返回作者E-mail
		intForumId=Request.Form("txtForumId")						'获取隐藏文本框传递回来的栏目编号
		intLayer=1													'这是第一层
		intFatherId=0												'因为是第一层，父编号设为0
		intChild=0													'回复文章数目为0
		intHits=0													'点击数为0
		strIP=Request.ServerVariables("REMOTE_ADDR")				'作者IP地址

		'以下将发表文章保存到数据库
		Dim sqlA,sqlB,strSql
		sqlA="Insert Into tbBbs(strTitle,intLayer,intFatherId,intChild,intHits,strIP,strUserId,dtmSubmit,intForumId"
		sqlB = "Values('" & strTitle & "'," & intLayer & "," & intFatherId & "," & intChild & "," & intHits & ",'" & strIP & "','" & strUserId & "',#" & Now() & "#," & intForumId 
		If strBody<>"" Then											'如果有内容，则添加
			sqlA = sqlA & ",strBody"
			sqlB = sqlB & ",'" & strBody & "'"
		End If
		If strEmail<>"" Then										'如果有E-mail，则添加
			sqlA = sqlA & ",strEmail"
			sqlB = sqlB & ",'" & strEmail & "'"
		End If
		strSql = sqlA & ") " & sqlB & ")"
		conn.Execute(strSql)

		'下面将该栏目的文章数加1
		strSql="Update tbForum Set lngForumCount=lngForumCount+1 Where ID = " & intForumId
		conn.Execute(strSql)

		'下面将该用户的发表文章数加1
		strSql="Update tbUsers Set intArticle=intArticle+1 Where strUserId='" & Session("strUserId") & "'"
		conn.Execute(strSql)

		'关闭对象
		conn.Close
		Set conn=Nothing

		'重定向回BBS列表页，请记住再将当前栏目号码传递回去。
		Response.Redirect "bbs_list.asp?intForumId=" & Request.Form("txtForumId")
	End If
	%>
</body>
</html>
