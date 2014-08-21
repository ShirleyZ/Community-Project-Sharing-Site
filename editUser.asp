<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Edit User Info</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      
      <%
      dim userID, userCmtID, userIDtmp
      userID = request.querystring("userID")
      userCmtID = request.querystring("userCmtID")
      dim stage
      stage = request.form("stage")
      dim sql, info
      '----------------------'
      '-- STAGE 1: Editing --'
      '----------------------'
      if stage = "" then
        '------------------------------------------'
        '-- STAGE 1 OF CHANGING PROFILE CONTENTS --'
        '------------------------------------------'
        if userID <> "" then
          userIDtmp = request.querystring("userID")
          userID = CLng(userIDtmp)
          '-- Check that you're the owner or admin
          if (Session("sz_uID") = userID) or (Session("sz_uAccess") = 1) then
            '-- Getting this user's info
            '                 0         1          2         3      4
            sql = "SELECT username, userEmail, userPass, userImg, about " &_
                  "FROM UsersTable " &_
                  "WHERE userID=" & userID
            set info = conn.execute(sql)
            
            response.write "<p>Edit User Profile</p>"
            
            '-- This is where the error message goes
            %>
            <!-- #include file="codeChunkMsgDisp.asp" -->
            <%
            '-- checking if User exists
            if info.eof then
              response.write "User #" &_
                             userID &_
                             " doesn't exist."
            else
              '-- Printing out the things
              response.write "<form action='editUser.asp?userID=" &_
                             userID &_
                             "' method='post'>" &_
                             "<input type='hidden' name='stage' value='2'>" &_
                             "<input type='hidden' name='userID' value='" &_
                             userID &_
                             "'><b>" &_
                             info(0) &_
                             "</b> #" &_
                             userID &_
                             "<br>Password: <input type='password' name='uPass' value='" &_
                             info(2) &_
                             "'>*<br>Email: <input type='text' name='uEmail' value='" &_
                             info(1) &_ 
                             "'>*<br>Image URL: <input type='text' name='imgURL' value='" &_
                             info(3) &_
                             "'><br>About: <input type='text' name='uAbout' value='" &_
                             info(4) &_
                             "'><br><input type='submit' value='Change'></form>"
            end if 
            
          else
            response.write "You're not allowed to change this user's profile! Get outta here, shoo!"
          end if
          
        '-----------------------------------------'
        '-- STAGE 1 OF CHANGING PROFILE COMMENT --'
        '-----------------------------------------'        
        elseif userCmtID <> "" then
          '-- Getting info for comment
          '                 0         1         2         3
          sql = "SELECT userNum, profileNum, comment, username " &_
                "FROM UserPgCommentsTable, UsersTable " &_
                "WHERE userCmtID=" & userCmtID & " AND userNum=userID"
          set info = conn.execute(sql)
          
          '-- Checking that you are the comment owner or admin
          if (Session("sz_uID") = info(0)) or (Session("sz_uAccess") = 1) then
            
            response.write "<p>Edit Profile Comment</p>"
            
            '-- This is where the error message goes
            %>
            <!-- #include file="codeChunkMsgDisp.asp" -->
            <%
            '-- Checking that the comment exists
            if info.eof then
              response.write "Comment #" & userCmtID & " doesn't exist"
            else
            '-- Print out the form
              response.write "<form action='editUser.asp?userCmtID=" &_
                              userCmtID &_
                              "' method='post'>" &_
                              "<input type='hidden' name='stage' value='2'>" &_
                              "Comment #" &_
                              userCmtID &_
                              " by " &_
                              info(3) &_
                              "<br>Comment: <br>"&_
                              info(2) &_
                              "<br>" &_
                              "<input type=""text"" name=""cmtBody"" value=""" &_
                              info(2) &_
                              """><br><input type='submit' value='Change'></form>"
            end if
          else
            response.write "Types: " & info(0) & "<br>Values: " & Session("sz_uID") & " and " & info(0)
            response.write "You're not allowed to change this user's comment! Get outta here, shoo!"
          end if
        else
          Session("sz_errorMsg") = "(Missing ID: editProfile.asp stage 1)"
          response.redirect "error.asp"
        end if
      
      '-----------------------'
      '-- STAGE 2: Changing --'
      '-----------------------'
      elseif stage = 2 then
      
        '--------------------------------------'
        '-- STAGE 2 OF CHANGING USER PROFILE --'
        '--------------------------------------'
        if userID <> "" then
          dim uPass, uEmail, uImg, uDesc
          
          uPass = request.form("uPass")
          uEmail = request.form("uEmail")
          uImg = request.form("imgURL")
          uDesc = request.form("uabout")
          
          '-- Error checking!
          '-- Cannot have empty password or email
          if (uPass = "") then
            Session("sz_errorMsg") = "Please don't leave your password blank!"
            response.redirect "editUser.asp?userID=" & userID
          elseif (uEmail = "") then
            Session("sz_errorMsg") = "Please don't leave your email blank!"
            response.redirect "editUser.asp?userID=" & userID
          end if
          
          '-- Changing the things
          sql = "UPDATE UsersTable " &_
                "SET userPass='" & uPass & "', " &_
                "userEmail='" & uEmail & "', " &_
                "userImg='" & uImg & "', " &_
                "about=""" & uDesc & """ " &_
                "WHERE userID=" & userID
          response.write sql
          conn.execute (sql)
          response.write "Done! Would you like to see <a href='viewProfile.asp?userID=" & userID &_
                         "'>your profile</a>?"
          
        '-----------------------------------------'
        '-- STAGE 2 OF CHANGING PROFILE COMMENT --'
        '-----------------------------------------'       
        elseif userCmtID <> "" then
          dim cmtBody
          cmtBody = request.form("cmtBody")
          
          '-- Error checking!
          '-- Cannot have empty comment
          if cmtBody = "" then
            Session("sz_errorMsg") = "You left the comment blank! Perhaps you want to " &_
                                  "<a href='deleteUser.asp?userCmtID='" & userCmtID & "'>"&_
                                  "delete</a> the comment instead?"
            response.redirect "editUser.asp?userCmtID=" & userCmtID
          end if
          
          '-- Changing the user profile comment
          sql = "UPDATE UserPgCommentsTable " &_
                "SET comment=""" & cmtBody & """ " &_
                "WHERE userCmtID=" & userCmtID
          response.write sql
          conn.execute(sql)
          
          '-- Redirecting back to profile
          sql = "SELECT profileNum " &_
                "FROM UserPgCommentsTable " &_
                "WHERE userCmtID=" & userCmtID
          set info = conn.execute(sql)
          response.redirect "viewProfile.asp?userID=" & info(0)
        else
          Session("sz_errorMsg") = "(Missing ID: editProfile.asp stage 2)"
          response.redirect "error.asp"
        end if
      else
        Session("sz_errorMsg") = "(Missing stage: editProfile.asp)"
        response.redirect "error.asp"
      end if
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>