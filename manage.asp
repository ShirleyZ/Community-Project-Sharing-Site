<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Manage</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <div id="mngUpperBar">
          <a href="addProject.asp">Add</a>
            <a onclick="show(3)">Search</a>
          <div id="status">
          </div>
        </div>
        <div id="info">
          <div id="mngProjectThumbTray">
            <!-- These will probably be in ASP code -->
            <%
              dim sql, info
              dim userNum, userProj
              
              userNum = Session("sz_uID")
              userProj = Session("sz_uID")
            %>
            <!-- #include file="codeChunkListProj.asp" -->
          </div>
        </div>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>