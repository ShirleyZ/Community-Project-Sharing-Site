<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Site Settings</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <% dim sql, info %>
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <p class="pageTitle">
          Site Settings
        </p>
        <%
        '-- Checking user has permission
        if Session("sz_uAccess") <> 1 then
          Session("sz_errorMsg") = "You don't have permission to view that"
          response.redirect "error.asp"
        end if
        
        %>
        <ul>
          <li><a href="userList.asp">Users List</a> - [Edit, Delete, Toggle Permissions]</li>
          <li>Manage Competitions</li>
        </ul>
        
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>