<%
'=========================================================================================
'这里是函数文件，用来保存在各页面中用到的函数
'=========================================================================================

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
%>