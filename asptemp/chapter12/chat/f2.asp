<%
'================================================================================================
'����ҳ��
'1. ��ҳֻ�ǽ��û��ύ����Ϣ��ͨ�����ú�����ӵ�������Ϣ�С�
'2. ע��˵����ɫ������Ҫ������һ��ʹ�õ���ɫ�����������Ĭ��ѡ����һ��ʹ�õ���ɫ��
'3. ����JavaScript������ڶ�λ��꣬���⣬���Ժ�����ˢ��f2.asp������ʾ���·��ԡ�
'================================================================================================
%>

<%option explicit											'ǿ����������%>
<!--#Include File="config.asp"-->
<!--#Include File="function.asp"-->
<html>
<head>
	<meta http-equiv='content-type' content='text/html; charset=gb2312'>
	<title>������</title>
	<link rel="stylesheet" href="chat.css">
</head>
<body  background="images/bg.gif">
	<form name="frmInput" method="POST" action="">
	<br><%=Session("strUserName")%>˵:
	<input type="text" name="txtSays" size="50" maxlength="100" >       
	<input type="submit" value=" �� �� " > 
	<p>˵����ɫ:        
	<select name="txtSaysColor">        
		<option value="#000000" 
		<%If Request.Form("txtSaysColor")="#000000" Then Response.Write "selected" %>
		>Ĭ��</option>    
		<option value="#0000FF" 
		<%If Request.Form("txtSaysColor")="#0000FF" Then Response.Write "selected" %>
		>��ɫ</option>    
		<option value="#FF0000" 
		<%If Request.Form("txtSaysColor")="#FF0000" Then Response.Write "selected" %>
		>��ɫ</option>    
		<option value="#FFFF00" 
		<%If Request.Form("txtSaysColor")="#FFFF00" Then Response.Write "selected" %>
		>��ɫ</option>    
	</select>    
	����:       
	<select name="txtEmote">       
		<option value="" selected>��</option>    
		<option value="���˵�">����</option>    
		<option value="���ĵ�">����</option> 
		<option value="���ߵ�">����</option>    
		<option value="���������">����</option>    
	</select> 
	<a href="#" onclick="JavaScript:window.open('about:blank','_top').close();">�뿪������</a>
	</form>
	<%
	'����ύ�˷�����Ϣ����ִ����������
	If Request.Form("txtSays")<>"" Then
		'�����Ȼ�ȡ�й����ݣ���������˺���
		Dim strSaysColor,strEmote,strSays
		strSaysColor=Request.Form("txtSaysColor")				'��ȡ��������ɫ
		strEmote=Request.Form("txtEmote")						'��ȡ�����߱���
		strSays=Trim(Request.Form("txtSays"))					'ȥ��ǰ��Ŀո�
		'������ú��������η�����ӵ�������Ϣ��
		Call AddUserSay(Session("strUserName"),strSaysColor,strEmote,strSays)
	End If
	%>
	<!--ע�ͣ���������JavaScript����궨λ�������ı��򣬲�ˢ���������ݿ�f1.asp-->
	<Script language="JavaScript">
		document.frmInput.txtSays.focus();						//��λ���
		top.main.location.href="f1.asp";						//�Զ�ˢ��f1.asp
	</Script>
</body>  
</html>  
