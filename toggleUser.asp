<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Toggle User</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        dim userID, stage
        userID = request.querystring("userID")
        stage = request.form("stage")
        
        dim sql, info
        
        if stage = "" then
          '---------------------------'
          '-- STAGE 1: Confirmation --'
          '---------------------------'
          if userID <> "" then
            '-- Checking that they have permission to do this
            if Session("sz_uAccess") <> 1 then
              response.write "You're not an admin! Get outta here, shoo!"
            else
              '-- Getting user info
              '                0        1          2         3
              sql = "SELECT userID, username, accessLvl, userEmail " &_
                    "FROM UsersTable " &_
                    "WHERE userID=" & userID
              set info = conn.execute(sql)
              
              '-- Checking user exists
              if info.eof then
                response.write "User #" & userID & " Does not exist"
              else
              '-- Printing out user
                response.write "<b>" &_
                               info(1) &_
                               "</b> #" &_
                               info(0) &_
                               "<br>Access: " &_
                               info(2) &_
                               "<br>Email: " &_
                               info(3) &_
                               "<br>"
                
                response.write "<p>What would you like to toggle user to?</p>" 
                
                response.write "<form action='toggleUser.asp?userID=" &_
                                userID &_
                                "' method='post'><input type='hidden' name='stage' value='2'>" &_
                                "<select name='uAccess'>"
                '-- Getting options
                '                 0         1
                sql = "SELECT accessID, accessName " &_
                      "FROM AccessLvlTable"
                set info = conn.execute(sql)
                
                do until info.eof
                  response.write "<option value='" &_
                                 info(0) &_
                                 "'>" &_
                                 info(1) &_
                                 "</option>"
                  info.movenext
                loop
                
                response.write "</select><br><input type='submit' value='Toggle'></form>" 
              end if
              '-- End of checking users
            end if
            '-- End of checking access
          else
            response.write "I think that you are lost (no userID)"
          end if      
          '-- End of checking userID
        elseif stage = 2 then
          '-----------------------'
          '-- STAGE 2: Toggling --'
          '-----------------------'
          dim uAccess
          uAccess = request.form("uAccess")
          
          sql = "UPDATE UsersTable " &_
                "SET accessLvl=" & uAccess & " " &_
                "WHERE userID=" & userID
          conn.execute(sql)
          
          '-- Getting user info
          '                0        1          2         3
          sql = "SELECT userID, username, accessLvl, userEmail " &_
                "FROM UsersTable " &_
                "WHERE userID=" & userID
          set info = conn.execute(sql)
          
          '-- Printing out user
            response.write "<b>" &_
                           info(1) &_
                           "</b> #" &_
                           info(0) &_
                           "<br>Access: " &_
                           info(2) &_
                           "<br>Email: " &_
                           info(3) &_
                           "<br>"
            response.write "<br>Done!"
        else
          response.write "I think that you are lost (no/wrong stage)"
        end if  
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>