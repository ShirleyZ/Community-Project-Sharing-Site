<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Register</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <%
        '-- If the user is logged in
        if Session("sz_uAccess") >= 1 then
      %>
        <!-- #include file="nav.asp" -->      
        <div id="content">
        <p>You're already logged in!</p>

      <%
        else
          dim stage
          stage = request.form("stage")
          
          '-- If this is the first stage then show register form
          if stage = "" then
      %>
        <!-- #include file="nav.asp" -->      
        <div id="content">
        <!-- #include file="codeChunkMsgDisp.asp" -->

        <form action="register.asp" method="post">
          <input type="hidden" name="stage" value="2">
          Desired Username: 
          <input type="text" name="dUserName">*
          <br>
          Desired Password:
          <input type="password" name="dUserPass">*
          <br>
          Confirm Password:
          <input type="password" name="dUserPassConf">*
          <br>
          Email: 
          <input type="text" name="userEmail">*
          <br>
          <input type="submit" value="Register!">
        </form>
      <%
          '-- If this is the second stage
          elseif stage = 2 then

            dim sql, info
            dim dUserName, dUserPass, userEmail, dUserPassConf
            
            dUserName = request.form("dUserName")
            dUserPass = request.form("dUserPass")
            dUserPassConf = request.form("dUserPassConf")
            userEmail = request.form("userEmail")
            
            '--------------------------------------------------------------'
            '-- ERROR CHECKING: Redirecting the user if there is a problem '
            '--------------------------------------------------------------'
            
            '-- Checking to see if this username is taken
            sql = "SELECT COUNT(username) " &_
                  "FROM UsersTable " &_
                  "WHERE username='" & dUserName & "'"
            set info = conn.execute(sql)
            
            '-- If we've found the username in the table, error and redirect
            if info(0) <> 0 then
              Session("sz_errorMsg") = "Sorry! This user already exists"
              response.redirect "register.asp"
            end if
            
            '-- Checking they've filled everything out
            if dUserName = "" or dUserPass = "" or dUserPassConf = "" or userEmail = "" then
              Session("sz_errorMsg") = "Oops! You didn't fill in all the fields"
              response.redirect "register.asp"
            end if
            
            '-- Checking to see if the passwords match
            if dUserPass <> dUserPassConf then
              Session("sz_errorMsg") = "Oops! Your passwords don't match"
              response.redirect "register.asp"
            end if
            
            '--------------------'
            '-- IF ALL IS WELL --'
            '--------------------'
            
            sql = "INSERT INTO UsersTable(username, userPass, userEmail, accessLvl) " &_
                  "VALUES ('" & dUserName & "', '" & dUserPass & "', '" & userEmail &"', 2)"
            response.write sql
            conn.execute(sql)
      %>
        <!-- #include file="nav.asp" -->
        
      <%
            response.write "<p>Success! Hello " & dUserName & ". You can now login.</p>"
          end if
          
        end if
        conn.close
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>