<%Response.Charset = "GB2312"
set c1=server.createobject("adodb.connection")
p1="provider=microsoft.jet.oledb.4.0;"
path1="data source="&server.mappath("user.mdb")
c1.open p1&path1

set rs=server.createobject("adodb.recordset")
rs.open "select * from ������ order by ����ʱ�� asc" ,c1,1,1
response.write "<table border=0>"
for i=1 to rs.recordcount%>

	<tr>
		<td>��<%=trim(rs("������"))%>���ԡ�<%=trim(rs("������"))%>��˵:<%=trim(rs("��������"))%>&nbsp;&nbsp;&nbsp;<font color=#cccccc><%=rs("����ʱ��")%></font></td>
	</tr>

<%rs.movenext
next
response.write "</table>"
rs.close
set rs=nothing
c1.close
set c1=nothing%>