<!-- Code to check if the user is logged in or not -->
<!-- If the user tries to access a page that requires
    login access, then it will redirect to login page -->
    
<%
  dim uAccess
  uAccess = Session("sz_uAccess")
  
  if (uAccess = -1) or (uAccess = "") then
    Session("sz_errorMsg") = "Please login to access that area"
    response.redirect "login.asp"
  end if
%>