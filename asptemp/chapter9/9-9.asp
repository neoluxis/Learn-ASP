<html>
<body>
<h2 align="center">������ѯʾ��</h2>
<form name="frmSearch" method="POST" action="">
	������Ҫ���ҵ�������<input type="text" name="txtName">
	<input type="submit" name="btnSubmit" value=" ȷ �� ">
</form>
<%
If Request.Form("txtName")<>"" Then
	'�����������ݿ⣬����һ��Connection����ʵ��conn
	Dim conn,strConn 
	Set conn=Server.CreateObject("ADODB.Connection")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
	conn.Open strConn 
	'���½���Command����ʵ��cmd
	Dim cmd
	Set cmd= Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn						'ָ��Connection����
	'���潨��һ����������prm
	Dim prmName,prmType,prmDirection,prmSize,prmValue
	prmName="varName"							'�������ƣ���qryList2�еı���
	prmType=200									'�������ͣ�200��ʾ�Ǳ䳤�ַ���
	prmDirection=1								'��������1��ʾ����
	prmSize=10									'������С������ֽ���Ϊ10
	prmValue=Request.Form("txtName")			'Ҫ����Ĳ���ֵ
	Dim prm										'����һ�������������
	Set prm=cmd.CreateParameter(prmName,prmType,prmDirection,prmSize,prmValue)
	'���潫�ò���������뵽����������
	cmd.Parameters.Append(prm)					
	'����ִ�в�ѯqryList2
	Dim rs
	cmd.CommandType=4							'��ʾ��ѯָ���ǲ�ѯ��
	cmd.CommandText="qryList2"					'ָ����ѯ����
	Set rs=cmd.Execute							'ִ�в�ѯ
	'��������ѭ�����г���Ա�����͵绰
	Do While Not rs.Eof 
		Response.Write rs("strName") & "," & rs("strTel") & "<br>"
		rs.MoveNext
	Loop
End If
%>
</body>
</html>
