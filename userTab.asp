<!-- Profile box that tells you who you're logged in as -->
<div id="userTab">
  <%
    '-- Setting variables if they've expired
    if Session("sz_uID") = "" then
      Session("sz_uID") = -1
    end if
    if Session("sz_uAccess") = "" then
      Session("sz_uAcess") = -1
    end if
    
    '-- If you are logged in
    if (Session("sz_uAccess") >= 0) then
    '-- Retrieve information
    sql = "SELECT userImg " &_
          "FROM UsersTable " &_
          "WHERE userID=" & Session("sz_uID")
    set info = conn.execute(sql)
    
    response.write "<img class='userTabImg' src='" &_
                    info(0) &_
                    "' />"
  %>
    <div class="userTabInfo">
    <%
      response.write "<a href='viewProfile.asp?userID=" & Session("sz_uID") &_
                     "'>" & Session("sz_uName") & "</a>" &_
                     " (#" & Session("sz_uID") & ") <br>" &_
                     "<ul class='userTabInfoOptions'><li><a href='editUser.asp?userID=" & Session("sz_uID") &_
                     "'>Edit Profile</a></li>" &_
                     "<li><a href='logout.asp'>Logout</a></li>" &_
                     "</ul>"
    %>
    </div>
  <%
    else
      response.write "<img class='userTabImg' src='' />" &_
                     "<p class='userTabInfo center'><br>You're not logged in!<br>" &_
                     "<a href='login.asp'>Login</a> or <a href='register.asp'>Register</a></p>"
    end if
  %>
</div clear="both">