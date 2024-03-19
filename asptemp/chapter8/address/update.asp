<% 
'只要有姓名和电话，就更新记录，否则给出错误信息
If Request.Form("txtName")<>"" And Request.Form("txtTel")<>"" Then
	'以下获取提交的表单数据
	Dim ID,strName,strSex,intAge,strTel,strEmail,strIntro		'声明几个变量
	ID=Request.Form("txtID")									'获取隐藏文本框传递过来的ID值
	strName=Request.Form("txtName")								'获取姓名
	strSex=Request.Form("rdoSex")								'获取性别
	If Request.Form("txtAge")<>"" Then							'获取年龄
		intAge=Request.Form("txtAge")
	Else
		intAge=0
	End If
	strTel=Request.Form("txtTel")								'获取电话
	strEmail=Request.Form("txtEmail")							'获取E-mail
	strIntro=Request.Form("txtIntro")							'获取个人简介
	'以下连接数据库，建立一个Connection对象实例conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection") 
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn
	'下面利用Execute方法更新记录，注意更新条件为隐藏文本框传过来的ID值
	Dim strSql
	strSql="Update tbAddress Set strName='" & strName & "',strSex='" & strSex & "',intAge=" & intAge & ",strTel='" & strTel & "',strEmail='" & strEmail & "',strIntro='" & strIntro & "' Where ID=" & ID
	conn.Execute(strSql) 
	'更新完毕，重定向回首页
	Response.Redirect "index.asp"    
Else
	Response.Write "姓名和电话必须填写"
	Response.Write "<a href='index.asp'>重新填写</a>"
End If
%>
