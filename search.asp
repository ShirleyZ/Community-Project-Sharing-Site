<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Template</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <div id="searchBar">
          <form action='search.asp' method='post'>
            Search term: <input type='text' name='searchTerm'><br>
            <input type='submit' value='Search'>
          </form>
          <!-- #include file="codeChunkMsgDisp.asp" -->
        </div>
        <%
          dim userProj
        %>
        <!-- #include file="codeChunkSearch.asp" -->
      
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>