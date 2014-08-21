
<!-- Navigation bar that changes based on your -->
<!--  access level -->
    <script language="javascript" type="text/javascript" src="script.js">
    </script>
<nav>
  <div id="navMenu" onmouseover='openMenu()' onmouseout='closeMenu()'>
    <a>V</a>
    <div id="navMenuBody">
      <ul>
        <%
          '-- If this user is an admin
            if Session("sz_uAccess") = 1 then
        %>
            <a href="siteSettings.asp"><li>Site Settings</li></a>
        <%
            end if
        %>
      </ul>    
    </div>
  </div>
  <ul class="navRow" id="navUpper">
    <a href="default.asp"><li>Home</li></a>
    <a href="addProject.asp"><li class="navAdd">+</li></a>
    <a href="search.asp"><li>Projects</li></a>
    <a href="search.asp"><li>Designs</li></a>
    <a href="forums.asp"><li>News</li></a>
    <a href="forums.asp"><li>Discussion</li></a>
    <a href="forums.asp"><li>Q & A</li></a>
<%
    '-- If this user is an admin
    if Session("sz_uAccess") = 1 then
%>
    <a href="siteSettings.asp"><li>Site Settings</li></a>
<%
    end if
%>
  </ul>
  <ul class="navRow" id="navLower">
  </ul>
  <!-- #include file="userTab.asp" -->
  <br clear="both">
</nav>