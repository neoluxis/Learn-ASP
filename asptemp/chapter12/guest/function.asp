<%
'=========================================================================================
'�����Ǻ����ļ������������ڸ�ҳ�����õ��ĺ���
'=========================================================================================

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
%>