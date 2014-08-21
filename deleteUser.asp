<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Delete User Item</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        dim userID, userCmtID, stage
        userID = request.querystring("userID")
        userCmtID = request.querystring("userCmtID")
        stage = request.form("stage")
        
        dim sql, info
        
        if stage = "" then
          '-------------------------------------------'
          '-- STAGE 1: Confirmation of User Profile --'
          '-------------------------------------------'
          if userID <> "" then
            '-- Checking that they have permission to do this
            if Session("sz_uAccess") <> 1 then
              response.write "You're not an admin! Get outta here, shoo!"
            else
              response.write "<p>Are you <b>certain</b> you want to delete this user?</p>" 
              
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
                
                '-- Making sure you don't delete an admin
                if info(2) = 1 then
                  response.write "This user is an admin! Please <a href='editUser.asp?userID=" &_
                                 info(0) &_
                                 "'>change them</a> to a regular user first."
                else
                  response.write "<form action='deleteUser.asp?userID=" &_
                                 info(0) &_
                                 "' method='post'>" &_
                                 "<input type='hidden' name='stage' value='2'>" &_
                                 "<input type=""submit"" value=""Yes I'm sure!"">" &_
                                 "</form>"
                end if
              end if
              '-- End of checking users
            end if
            '-- End of checking access
            
          '---------------------------------------'
          '-- STAGE 1 changing profile commente --'
          '---------------------------------------'
          elseif userCmtID <> "" then 
            response.write "<p>Deleting Profile Comment</p>"
            '-- Getting info about the comment
            '                 0         1         2
            sql = "SELECT userNum, profileNum, comment " &_
                  "FROM UserPgCommentsTable " &_
                  "WHERE userCmtID=" & userCmtID
            set info = conn.execute(sql)
            
            '-- Checking that they have permission to do this
            if (info(0) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
              response.write "<p>Are you <b>certain</b> you want to delete this profile comment?</p>" 
              response.write "<form action='deleteUser.asp?userCmtID=" &_
                             userCmtID &_
                             "' method='post'>" &_
                             "<input type='hidden' name='stage' value='2'>" &_
                             "Comment: <br>" &_
                             info(2) &_
                             "<br>" &_
                             "<input type=""submit"" value=""Yes I'm sure!"">" &_
                             "</form><a class='jsButton' href='viewProfile.asp?userID=" &_
                             info(1) &_
                             "'>No, take me back!</a>"
              
            else
              response.write "You don't have permission to do that! Get outta here, shoo!"
            end if     
          else
            response.write "I think that you are lost (no userID)"
          end if      
          '-- End of checking userID
        elseif stage = 2 then
          '-----------------------'
          '-- STAGE 2: Deletion --'
          '-----------------------'
          
          if userID <> "" then
            '-- Deleting image comments
            sql = "DELETE FROM ImgCommentsTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting images
            sql = "DELETE FROM ImgTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting projects
            sql = "DELETE FROM ProjTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting profile comments
            sql = "DELETE FROM UserPgCommentsTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting forum threads
            sql = "DELETE FROM ForumThreadsTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting forum posts
            sql = "DELETE FROM ForumPostsTable " &_
                  "WHERE userNum=" &_
                  userID
            conn.execute(sql)
            
            '-- Deleting User
            sql = "DELETE FROM UsersTable " &_
                  "WHERE userID=" & userID
            conn.execute(sql)
            response.write "done"          
          elseif userCmtID <> "" then   
            '-- Grabbing where to redirect to
            sql = "SELECT profileNum " &_
                  "FROM UserPgCommentsTable " &_
                  "WHERE userCmtID=" & userCmtID
            set info = conn.execute(sql)
            
            '-- Deleting the comment
            sql = "DELETE FROM UserPgCommentsTable " &_
                  "WHERE userCmtID=" & userCmtID
            conn.execute(sql)
            response.write "Done"
            response.redirect "viewProfile.asp?userID=" & info(0)
          else
            response.write "No userID/userCmtID (stage 2)"
          end if
        else
          response.write "I think that you are lost (no/wrong stage)"
        end if  
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>