<%
set c1=server.createobject("adodb.connection")
p1="provider=microsoft.jet.oledb.4.0;"
path1="data source="&server.mappath("user.mdb")
c1.open p1&path1
set rs=server.createobject("adodb.recordset")
rs.open "select * from ������" ,c1,1,3
rs.addnew
rs("������")=trim(request.querystring("liuyan"))
rs("������")=trim(request.querystring("jieshou"))
rs("��������")=trim(request.querystring("content"))
rs("����ʱ��")=now()
rs.update
rs.close
set rs=nothing
c1.close
set c1=nothing%>