<% 
'�����������ݿ⣬����һ��Connection����ʵ��conn
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection") 
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("address.mdb")
conn.Open strConn
'����ɾ����¼��ע������Ҫ�õ���index.asp��������Ҫɾ���ļ�¼��ID
Dim strSql
strSql="Delete From tbAddress Where ID=" & Request.QueryString("ID") 
conn.Execute(strSql)
'ɾ����Ϻ󣬷�����ҳ
Response.Redirect "index.asp"								
%>
