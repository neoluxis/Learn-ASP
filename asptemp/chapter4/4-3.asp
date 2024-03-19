<html>
<body>
	<form name="frmTest" method="POST" action="">
		第1个数<input type="text" name="txtA">
		+
		第2个数<input type="text" name="txtB">
		<p><input type="submit" name="btnSubmit" value="确定">
	</form>
	<%
	'下面的条件语句表示只有提交了表单才进行计算
	If Request.Form("txtA")<>"" And Request.Form("txtB")<>"" Then      
		Dim intA,intB,intC
		intA=Request.Form("txtA")				'获取表单元素txtA的值
		intB=Request.Form("txtB")				'获取表单元素txtB的值
		intC=CInt(intA)+CInt(intB)				'因为传送的是字符串，所以必须转换类型
		Response.Write "两个数的和=" & intC	
	Else
		Response.Write "请输入两个数后按确定按钮"
	End If
	%>
</body>
</html>
