<% Response.Buffer=True							'注意：最好有该语句       %>
<html>
<body>
	<% 
	Dim varNumber                        		'定义一个访问次数变量
	varNumber=Request.Cookies("intVisit") 		'读取Cookie值
	If varNumber="" then						'如果varNumber=""，表示还没有定义该Cookie
		varNumber=1                    			'如果是第一次，则令访问次数为1
	Else
		varNumber=CInt(varNumber)+1            	'如果不是第一次，则令访问次数加1
	End If
	Response.Write "您是第" & varNumber & "次访问本站"   
	Response.Cookies("intVisit")=varNumber   						'将新的访问次数保存到Cookie中
	Response.Cookies("intVisit").Expires=DateAdd("d",30,Date())   	'设置有效期为30天
	%>
</body>
</html>

