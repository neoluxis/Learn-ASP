<html>
<body>
	<h2 align="center">显示广告图片示例</h2>
	<%
	Dim ad											'声明一个AdRotator对象实例
	Set ad=Server.CreateObject("MSWC.AdRotator")
	ad.Border=2										'设置图片边框为2像素
	ad.Clickable=True                          		'设置图片提供超链接功能
	ad.TargetFrame="_blank"							'设置在新窗口中打开网址
	Response.Write Ad.GetAdvertisement("adver.txt")	'获取广告信息HTML字符串
	%>
</body> 
</html> 
