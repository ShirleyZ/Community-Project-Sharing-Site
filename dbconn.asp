<%
  dim conn
  set conn = server.createobject ("ADODB.Connection")
  conn.open "Provider=Microsoft.jet.oledb.4.0;data source="&Server.MapPath("assn4.mdb")
%>