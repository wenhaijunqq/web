<% 
Response.ContentType="text/xml"

DbPath = SERVER.MapPath("192.mdb")     '�����ݿ�ACCESS
	Set c1 = Server.CreateObject("ADODB.Connection")
	c1.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DbPath
set rs=server.createobject("adodb.recordset")
rs.open "select * from �û��� where �û���='"&trim(request.querystring("username"))&"'" ,c1,1,1
if rs.recordcount=0 then
response.write("<SUCCESS/>")
else
response.write("<ERROR/>")
end if
rs.close
set rs=nothing
c1.close
set c1=nothing%>
  