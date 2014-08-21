<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Logout</title>
  </head>
  
  <body>
  	<%
      Session("sz_uAccess") = -1
			Session("sz_uName") = "-1"
			Session("sz_uID") = -1
			
			Session("sz_infoMsg") = "You have been logged out"
			response.redirect "login.asp"
		%>
  
  </body>
</html>