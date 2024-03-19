<%
'===========================================================================================
'函数文件
'===========================================================================================

'该函数用来返回用户的信息,这里返回的信息比较多，不过并不是所有的信息都会用到
'注意，这里会返回一个数组
Function PersonalInfo(strUserId)
	Dim strSql,rs
	strSql="Select * From tbUsers Where strUserId='" & strUserId & "'"
	Set rs=conn.Execute(strSql)
	Dim strTemp(6)
	If Not rs.Bof And Not rs.Eof Then
		'如果该用户存在，就将有关数据赋值到一个数组中
		strTemp(0)=strUserId
		strTemp(1)=rs("strEmail")
		strTemp(2)=rs("strQQ")
		strTemp(3)=rs("strTel")
		strTemp(4)=rs("intArticle")
		strTemp(5)=rs("intReArticle")
		strTemp(6)=rs("dtmSubmit")
	Else
		'如果该用户不存在，就表示已经被删除
		strTemp(0)=strUserId
		strTemp(1)=""
		strTemp(2)=""
		strTemp(3)=""
		strTemp(4)=""
		strTemp(5)=""
		strTemp(6)=""
	End If
	PersonalInfo=strTemp               '返回数组
End Function

'该函数用来对字符串进行HTML编码，而且，要替换其中的空格和换行符号，以实现更佳的排版效果
Function myHTMLEncode(myString)
	If IsNull(myString) Then
		myHTMLEncode=""									'如果myString为空，则赋值空字符串
	Else
		myString=Replace(myString,"&","&amp;")			'替换&为字符实体&amp;
		myString=Replace(myString,"<","&lt;")			'替换<为字符实体&lt;
		myString=Replace(myString,">","&gt;")			'替换>为字符实体&gt;
		myString=Replace(myString,Chr(32),"&nbsp;")		'替换空格符为字符实体&nbsp;
		myString=Replace(myString,Chr(13),"<br>")		'替换回车符为换行标记<br>  
		myHTMLEncode=myString							'返回函数值
	End If
End Function

'该函数用来对字符串中的危险字符进行处理。
Function myDangerEncode(myString)
	If IsNull(myString) Then
		myDangerEncode=""								'如果myString为空，则赋值空字符串
	Else
		myString=Trim(myString)							'去掉前后的空格
		myString=Replace(myString,"'","''")				'将单引号替换为连续两个单引号
		myDangerEncode=myString							'返回函数值
	End If
End Function

%>