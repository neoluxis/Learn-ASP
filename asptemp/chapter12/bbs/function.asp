<%
'===========================================================================================
'�����ļ�
'===========================================================================================

'�ú������������û�����Ϣ,���ﷵ�ص���Ϣ�Ƚ϶࣬�������������е���Ϣ�����õ�
'ע�⣬����᷵��һ������
Function PersonalInfo(strUserId)
	Dim strSql,rs
	strSql="Select * From tbUsers Where strUserId='" & strUserId & "'"
	Set rs=conn.Execute(strSql)
	Dim strTemp(6)
	If Not rs.Bof And Not rs.Eof Then
		'������û����ڣ��ͽ��й����ݸ�ֵ��һ��������
		strTemp(0)=strUserId
		strTemp(1)=rs("strEmail")
		strTemp(2)=rs("strQQ")
		strTemp(3)=rs("strTel")
		strTemp(4)=rs("intArticle")
		strTemp(5)=rs("intReArticle")
		strTemp(6)=rs("dtmSubmit")
	Else
		'������û������ڣ��ͱ�ʾ�Ѿ���ɾ��
		strTemp(0)=strUserId
		strTemp(1)=""
		strTemp(2)=""
		strTemp(3)=""
		strTemp(4)=""
		strTemp(5)=""
		strTemp(6)=""
	End If
	PersonalInfo=strTemp               '��������
End Function

'�ú����������ַ�������HTML���룬���ң�Ҫ�滻���еĿո�ͻ��з��ţ���ʵ�ָ��ѵ��Ű�Ч��
Function myHTMLEncode(myString)
	If IsNull(myString) Then
		myHTMLEncode=""									'���myStringΪ�գ���ֵ���ַ���
	Else
		myString=Replace(myString,"&","&amp;")			'�滻&Ϊ�ַ�ʵ��&amp;
		myString=Replace(myString,"<","&lt;")			'�滻<Ϊ�ַ�ʵ��&lt;
		myString=Replace(myString,">","&gt;")			'�滻>Ϊ�ַ�ʵ��&gt;
		myString=Replace(myString,Chr(32),"&nbsp;")		'�滻�ո��Ϊ�ַ�ʵ��&nbsp;
		myString=Replace(myString,Chr(13),"<br>")		'�滻�س���Ϊ���б��<br>  
		myHTMLEncode=myString							'���غ���ֵ
	End If
End Function

'�ú����������ַ����е�Σ���ַ����д���
Function myDangerEncode(myString)
	If IsNull(myString) Then
		myDangerEncode=""								'���myStringΪ�գ���ֵ���ַ���
	Else
		myString=Trim(myString)							'ȥ��ǰ��Ŀո�
		myString=Replace(myString,"'","''")				'���������滻Ϊ��������������
		myDangerEncode=myString							'���غ���ֵ
	End If
End Function

%>