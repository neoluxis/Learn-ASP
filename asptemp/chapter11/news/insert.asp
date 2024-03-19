<%
'下面首先保存上传的文件
Dim upload												'声明upload对象实例
Set upload=Server.CreateObject("Persits.Upload") 
upload.OverWriteFiles=False								'False表示文件不可以被覆盖
upload.Save Server.MapPath("upfiles")					'上传到指定文件夹
'以下获取提交过来的标题、内容和文件名称
Dim strTitle,strBody,strFileName						'声明3个变量待用
strTitle=upload.Form("txtTitle")
strBody=upload.Form("txtBody")
'下面判断是否上传了文件？如果没有上传，则赋值空字符串
If upload.Files.Count>0 Then
	strFileName=upload.Files("fleUpload").FileName		'获取上传文件的名称
Else
	strFileName=""										'赋值空字符串
End If
'以下将标题、内容和文件名添加到数据库中
Dim strConn,conn 
Set conn=Server.CreateObject("ADODB.Connection") 
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("news.mdb")
conn.Open strConn
Dim strSql
strSql="Insert Into tbNews(strTitle,strBody,strFileName,dtmSubmit) Values('" & strTitle & "','" & strBody & "','" & strFileName & "',#" & date() & "#)"
conn.execute(strSql)
Response.Redirect "index.asp"							'成功添加，返回首页
%>
