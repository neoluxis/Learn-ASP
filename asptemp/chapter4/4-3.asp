<html>
<body>
	<form name="frmTest" method="POST" action="">
		��1����<input type="text" name="txtA">
		+
		��2����<input type="text" name="txtB">
		<p><input type="submit" name="btnSubmit" value="ȷ��">
	</form>
	<%
	'�������������ʾֻ���ύ�˱��Ž��м���
	If Request.Form("txtA")<>"" And Request.Form("txtB")<>"" Then      
		Dim intA,intB,intC
		intA=Request.Form("txtA")				'��ȡ��Ԫ��txtA��ֵ
		intB=Request.Form("txtB")				'��ȡ��Ԫ��txtB��ֵ
		intC=CInt(intA)+CInt(intB)				'��Ϊ���͵����ַ��������Ա���ת������
		Response.Write "�������ĺ�=" & intC	
	Else
		Response.Write "��������������ȷ����ť"
	End If
	%>
</body>
</html>
