<Script language="VBScript" runat="server">
'==================================================================================================
'这是函数文件
'1. 其中的函数供其他页面来调用
'2. 注意：因为Global.asa会调用该文件，相当于将本文件内容全部复制到Global.asa文件中， 而Global.asa中不允许使用<%和%>的形式，所以这里用了现在的形式。
'==================================================================================================

'该函数返回该聊天室共有多少人
Function AllUserName()
	Dim strUsers,I
	strUsers=Application("strUsers")				'获取在线人员数组
	If IsArray(strUsers)=True Then
		AllUserName=Ubound(strUsers)+1				'条件成立，表示确实是数组
	Else
		AllUserName=0								'表示还没有人在线
	End If
End Function

'该函数用来判断该用户名是否已被使用,存在返回True，不存在返回False
Function GetUserName(strUserName)
	Dim strUsers,I
	strUsers=Application("strUsers")				'获取在线人员数组
	'下面判断strUsers是否是数组，如果是才查找；否则表示无人在线，不必查找
	If IsArray(strUsers)=True Then
		For I=0 To Ubound(strUsers)					'在数组中循环查找该用户名
			If strUserName=strUsers(I) Then   
				GetUserName=True					'找到该用户名，返回True
				Exit Function						'跳出函数
			End IF
		Next
	Else
		GetUserName=False							'无人在线，直接返回False
	End If
End Function

'该子程序用来将该用户加入在线用户列表
Sub AddUserName(strUserName)
	Dim strUsers,intTemp
	strUsers=Application("strUsers")
	'下面判断strUsers是否是数组，如果是就添加；否则表示无人在线，就创建一个数组
	If IsArray(strUsers)=True Then
		'条件不成立，表示已经有人在线，只要将其添加在后面即可
		'重定义原来数组的大小，令其比原数组多1项，将新用户加在后面
		intTemp=Ubound(strUsers)					'获取原来数组下标
		Redim Preserve strUsers(intTemp+1)			'重定义数组，使数组长度加1
		strUsers(intTemp+1)=strUserName				'将新用户名添加到后面
		'下面保存到Application中
		Application.Lock
		Application("strUsers")=strUsers			'将数组保存到Application中
		Application.Unlock
	Else
		'如果条件成立，表示此时根本还不是数组，说明是第一个人访问，所以定义一个新的数组，并将该用户添加进去即可。
		Dim strNew(0)								'定义一个长度为1的数组
		strNew(0)=strUserName						'给数组赋值
		Application.Lock
		Application("strUsers")=strNew				'将数组保存到Application中
		Application.Unlock
	End If
End Sub

'该子程序将用户从在线用户中删除
Sub DelUserName(strUserName)
	Dim strUsers,intTemp,I,J
	strUsers=Application("strUsers")				'获取在线人员数组
	intTemp=Ubound(strUsers)					'返回数组最大下标
	'下面分两种情况删除
	If intTemp>0 Then
		'intTemp>0表示有多个人，首先查找该用户，用变量J记住该用户所在位置
		For I=0 To intTemp
			If strUserName=strUsers(I) Then		'该条件表示找到了用户名
				J=I							'用J记住所在位置
			End IF
		Next
		'将该位置之后的所有人向前移动一个位置
		For I=J To intTemp-1
			strUsers(I)=strUsers(I+1)
		Next
		'重定义该数组的大小，令其比原数组少1
		Redim Preserve strUsers(intTemp-1)			'重定义数组，保留原来值
		'下面将新的数组保存到Application中
		Application.Lock
		Application("strUsers")=strUsers			'将在线人员保存到数组中
		Application.Unlock
	Else
		'intTemp>0不成立表示就只有他一个人，只要将Application中的值清空即可
		Application.Lock
		Application("strUsers")=""
		Application.Unlock
	End If
End Sub

'该子程序将某人来到的信息添加到聊天信息中
Sub AddUserCome(strUserName,IP)
	'下面先组成该用户来到的信息字符串
	Dim strSays
	strSays=Time() & "<font color='red'>来自" & IP & "的" & strUserName & "大驾光临</font>"
	'下面将该用户来到的信息保存到聊天信息中
	Application.Lock													'先锁定
	Application("strChat")= Application("strChat") & "<br>" & strSays	'添加到Application中    
	Application.Unlock													'解除锁定
End Sub

'该子程序将某人离去的信息添加聊天信息中
Sub AddUserGo(strUserName,IP)
	'下面先组成该用户离去的信息字符串
	Dim strSays
	strSays=Time() & "<font color='blue'>来自" & IP & "的" & strUserName & "悄悄走了</font>"
	'下面将该用户离去的信息保存到聊天信息中
	Application.Lock													'先锁定
	Application("strChat")= Application("strChat") & "<br>" & strSays	'添加到Application中    
	Application.Unlock													'解除锁定
End Sub

'该子程序将某人的发言添加到聊天信息中
'其中strUserName表示发言用户，strSaysColor表示发言颜色，strEmote表示发言表情
'strSays表示本次发言内容
Sub AddUserSay(strUserName,strSaysColor,strEmote,strSays)
	'下面将组成本次发言的字符串
	strSays=Time() & strUserName & strEmote & "说：" & "<font color='" & strSaysColor & "'>" & strSays & "</font>" 
	'下面几句将本次发言信息保存到Application中
	Application.Lock											'先锁定
	Application("strChat")=Application("strChat") & "<br>" & strSays		'添加本次发言
	'如果聊天信息总长度超过10000个字符，则截短
	If Len(Application("strChat"))>10000 Then    
		Application("strChat")=Right(Application("strChat"),10000)
	End If
	Application.Unlock											'解除锁定
End Sub

</Script>