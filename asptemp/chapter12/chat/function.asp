<Script language="VBScript" runat="server">
'==================================================================================================
'���Ǻ����ļ�
'1. ���еĺ���������ҳ��������
'2. ע�⣺��ΪGlobal.asa����ø��ļ����൱�ڽ����ļ�����ȫ�����Ƶ�Global.asa�ļ��У� ��Global.asa�в�����ʹ��<%��%>����ʽ�����������������ڵ���ʽ��
'==================================================================================================

'�ú������ظ������ҹ��ж�����
Function AllUserName()
	Dim strUsers,I
	strUsers=Application("strUsers")				'��ȡ������Ա����
	If IsArray(strUsers)=True Then
		AllUserName=Ubound(strUsers)+1				'������������ʾȷʵ������
	Else
		AllUserName=0								'��ʾ��û��������
	End If
End Function

'�ú��������жϸ��û����Ƿ��ѱ�ʹ��,���ڷ���True�������ڷ���False
Function GetUserName(strUserName)
	Dim strUsers,I
	strUsers=Application("strUsers")				'��ȡ������Ա����
	'�����ж�strUsers�Ƿ������飬����ǲŲ��ң������ʾ�������ߣ����ز���
	If IsArray(strUsers)=True Then
		For I=0 To Ubound(strUsers)					'��������ѭ�����Ҹ��û���
			If strUserName=strUsers(I) Then   
				GetUserName=True					'�ҵ����û���������True
				Exit Function						'��������
			End IF
		Next
	Else
		GetUserName=False							'�������ߣ�ֱ�ӷ���False
	End If
End Function

'���ӳ������������û����������û��б�
Sub AddUserName(strUserName)
	Dim strUsers,intTemp
	strUsers=Application("strUsers")
	'�����ж�strUsers�Ƿ������飬����Ǿ���ӣ������ʾ�������ߣ��ʹ���һ������
	If IsArray(strUsers)=True Then
		'��������������ʾ�Ѿ��������ߣ�ֻҪ��������ں��漴��
		'�ض���ԭ������Ĵ�С�������ԭ�����1������û����ں���
		intTemp=Ubound(strUsers)					'��ȡԭ�������±�
		Redim Preserve strUsers(intTemp+1)			'�ض������飬ʹ���鳤�ȼ�1
		strUsers(intTemp+1)=strUserName				'�����û�����ӵ�����
		'���汣�浽Application��
		Application.Lock
		Application("strUsers")=strUsers			'�����鱣�浽Application��
		Application.Unlock
	Else
		'���������������ʾ��ʱ�������������飬˵���ǵ�һ���˷��ʣ����Զ���һ���µ����飬�������û���ӽ�ȥ���ɡ�
		Dim strNew(0)								'����һ������Ϊ1������
		strNew(0)=strUserName						'�����鸳ֵ
		Application.Lock
		Application("strUsers")=strNew				'�����鱣�浽Application��
		Application.Unlock
	End If
End Sub

'���ӳ����û��������û���ɾ��
Sub DelUserName(strUserName)
	Dim strUsers,intTemp,I,J
	strUsers=Application("strUsers")				'��ȡ������Ա����
	intTemp=Ubound(strUsers)					'������������±�
	'������������ɾ��
	If intTemp>0 Then
		'intTemp>0��ʾ�ж���ˣ����Ȳ��Ҹ��û����ñ���J��ס���û�����λ��
		For I=0 To intTemp
			If strUserName=strUsers(I) Then		'��������ʾ�ҵ����û���
				J=I							'��J��ס����λ��
			End IF
		Next
		'����λ��֮�����������ǰ�ƶ�һ��λ��
		For I=J To intTemp-1
			strUsers(I)=strUsers(I+1)
		Next
		'�ض��������Ĵ�С�������ԭ������1
		Redim Preserve strUsers(intTemp-1)			'�ض������飬����ԭ��ֵ
		'���潫�µ����鱣�浽Application��
		Application.Lock
		Application("strUsers")=strUsers			'��������Ա���浽������
		Application.Unlock
	Else
		'intTemp>0��������ʾ��ֻ����һ���ˣ�ֻҪ��Application�е�ֵ��ռ���
		Application.Lock
		Application("strUsers")=""
		Application.Unlock
	End If
End Sub

'���ӳ���ĳ����������Ϣ��ӵ�������Ϣ��
Sub AddUserCome(strUserName,IP)
	'��������ɸ��û���������Ϣ�ַ���
	Dim strSays
	strSays=Time() & "<font color='red'>����" & IP & "��" & strUserName & "��ݹ���</font>"
	'���潫���û���������Ϣ���浽������Ϣ��
	Application.Lock													'������
	Application("strChat")= Application("strChat") & "<br>" & strSays	'��ӵ�Application��    
	Application.Unlock													'�������
End Sub

'���ӳ���ĳ����ȥ����Ϣ���������Ϣ��
Sub AddUserGo(strUserName,IP)
	'��������ɸ��û���ȥ����Ϣ�ַ���
	Dim strSays
	strSays=Time() & "<font color='blue'>����" & IP & "��" & strUserName & "��������</font>"
	'���潫���û���ȥ����Ϣ���浽������Ϣ��
	Application.Lock													'������
	Application("strChat")= Application("strChat") & "<br>" & strSays	'��ӵ�Application��    
	Application.Unlock													'�������
End Sub

'���ӳ���ĳ�˵ķ�����ӵ�������Ϣ��
'����strUserName��ʾ�����û���strSaysColor��ʾ������ɫ��strEmote��ʾ���Ա���
'strSays��ʾ���η�������
Sub AddUserSay(strUserName,strSaysColor,strEmote,strSays)
	'���潫��ɱ��η��Ե��ַ���
	strSays=Time() & strUserName & strEmote & "˵��" & "<font color='" & strSaysColor & "'>" & strSays & "</font>" 
	'���漸�佫���η�����Ϣ���浽Application��
	Application.Lock											'������
	Application("strChat")=Application("strChat") & "<br>" & strSays		'��ӱ��η���
	'���������Ϣ�ܳ��ȳ���10000���ַ�����ض�
	If Len(Application("strChat"))>10000 Then    
		Application("strChat")=Right(Application("strChat"),10000)
	End If
	Application.Unlock											'�������
End Sub

</Script>