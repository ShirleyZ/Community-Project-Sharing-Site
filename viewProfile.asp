<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: View Profile</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        dim userID
        userID = request.querystring("userID")
        
        dim sql, info
        '-- If there's no ID, show their own profile
        if userID = "" then
          userID = Session("sz_uID")
        end if
        %>
        <!-- #include file="codeChunkMsgDisp.asp" -->
        <%
        '-- Get info for profile
         '                0         1         2      3
        sql = "SELECT username, userEmail, userImg, about " &_
              "FROM UsersTable " &_
              "WHERE userID=" & userID
        set info = conn.execute(sql)
      %>
        <!-- Template for profile -->
        <div id="userInfoSheet">
          <%
          
          '-- Filling in stuff
          response.write "<img id='userInfoImg' src='" &_
                         info(2) &_
                         "' />"
                         
          response.write "<div id='userInfoCard'>" &_
                         "<span class='userInfoModuleTitle'>User Info</span>" 
          
          '-- When no ID is given in url, userID is set to Session("sz_uID") [both become 'long' type variable]
          '-- When an ID is given, comparing userID (str) to Session("sz_uID") (long) doesn't work, so we use seshValue (str)
          '-- Therefore, two conditions in the if clause to cover both situations
          dim seshValue
          seshValue = CStr(Session("sz_uID"))
          
          '-- If you're the owner, show owner options
          if (seshValue = userID) or (Session("sz_uID") = userID) then 
            response.write "  <span id='ownerOptions'><a href='editUser.asp?userID=" &_
                           userID & "'>Edit Profile</a></span>"
          end if
          response.write "<br>"
          response.write "<b>" &_
                         info(0) &_
                         "</b><br><br>Email: " &_
                         info(1) &_
                         "<br><br>About Me: <br>" &_
                         info(3) 
          
          %>
          </div>
        </div>
        <div id="userInfoProjects">
          <p class="userInfoModuleTitle">Projects by this user</p>
          <br>
          <%
            dim userNum
            
            userNum = userID
          %>
          <!-- #include file="codeChunkListProj.asp" -->
        </div>
        <div id="userInfoComments">
          <p class='userInfoModuleTitle'>Profile Comments</p>
          <br>
          <form action='makeComment.asp' method='post'>
            <input type='hidden' name='entryType' value='2'>
            <%
              response.write "<input type='hidden' name='profID' value='" &_
                              userID & "'>"
            %>
            <input type='text' name='cmtBody'>
            <input type='submit' value='Comment'>
          </form>
          <%
          '-- Getting profile comments
          dim uComments
          '                 0          1        2         3        4
          sql = "SELECT userCmtID, username, comment, datePost, userNum " &_
                "FROM UsersTable, UserPgCommentsTable " &_
                "WHERE profileNum=" & userID & " AND userNum=userID " &_
                "ORDER BY userCmtID"
          set uComments = conn.execute(sql)
          
          '-- Checking you got some comments
          if uComments.eof then
            response.write "<div class='dispCmt'>" &_
                              "There are no comments here!" &_
                            "</div>"
          else
          '-- Printing out list of comments!
            do until uComments.eof
              response.write "<div class='dispCmt'>" &_
                                "<p class='dispCmtBody'>" &_
                                uComments(2) &_
                                "</p>"&_
                                "<p class='dispCmtMeta'>" &_
                                "Comment #" &_
                                 uComments(0) &_
                                 " by: " &_
                                 uComments(1) &_
                                 "<br>Date Posted: " &_
                                 uComments(3)
              '-- Checking if admin or owner
              if (uComments(4) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
                response.write "<br>" &_
                               "<a href='editUser.asp?userCmtID=" &_
                               uComments(0) &_
                               "'>Edit</a>"
                response.write " &nbsp;&nbsp;" &_
                               "<a href='deleteUser.asp?userCmtID=" &_
                               uComments(0) &_
                               "'>Delete</a>"
              end if
              response.write   "</p>" &_
                              "</div>"
              uComments.movenext
            loop
          end if
          %>
        </div>
        <br clear="left">   
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>