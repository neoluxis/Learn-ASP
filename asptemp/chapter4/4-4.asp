<html>
<body>
	<h2 align="center">请填写个人信息</h2>
	<form name="frmInfor" method="POST" action="4-5.asp">
		姓名：<input type="text" name="txtName"><br>
		密码：<input type="password" name="txtPwd"><br>
		性别：<input type="radio" name="rdoSex" value="男">男
			  <input type="radio" name="rdoSex" value="女">女<br>
		爱好：<input type="checkbox" name="chkLove" value="计算机">计算机
			  <input type="checkbox" name="chkLove" value="音乐">音乐
			  <input type="checkbox" name="chkLove" value="旅游">旅游<br>
		职业：<select name="sltCareer">
				  <option value="教育业">教育业</option>
				  <option value="金融业">金融业</option>
				  <option value="其他">其他</option>
			  </select><br>
		简述：<textarea name="txtIntro" rows="3" cols="40"></textarea><br>
		<input type="submit" name="btnSubmit" value="  确 定　">
		<input type="reset" name="btnReset" value="重新填写">
	</form>
</body> 
</html>

