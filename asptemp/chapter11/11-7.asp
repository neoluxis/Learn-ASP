<%
Dim jmail
Set jmail=Server.CreateObject("Jmail.Message")		'����Message����ʵ��
jmail.From="aspbook@163.com"						'�����˵�ַ
jmail.FromName="�"								'����������
jmail.AddRecipient "jjshang@263.net"				'��һ���ռ���
jmail.AddRecipient "jiananlearning@qq.com"				'�ڶ����ռ���
jmail.Subject="���"								'�ż�����
jmail.Body="����һ�·������"						'�ż�����
jmail.AddAttachment "C:\Inetpub\wwwroot\asptemp\chapter11\music.mid"		'��Ӹ���һ
jmail.AddAttachment "C:\Inetpub\wwwroot\asptemp\chapter11\beauty.jpg"		'��Ӹ�����
jmail.Charset="gb2312"								'�ż������ʽ
jmail.Send "aspbook:MVBYNUILCPUISVQM@smtp.163.com"			'ִ�з��ͣ������Ƿ��ŷ�����
jmail.Close											'������ϣ��رոö���
Response.Write "�Ѿ��ɹ�����"
%>
