<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->

<html>
  <head>
    <title>Assignment 4: Login</title>
    <link rel="stylesheet" href="style.css" type="text/css">
  </head>
  <body>
  	<div id="wrapper">
    
    <%
			dim stage, info
			
			stage = request.form("stage")
			if Session("sz_uAccess") >= 1 then
		%>
				<!-- #include file='nav.asp' -->
    <%
	      response.write "<div id='content'><p>Welcome back, " & Session("sz_uName") & "</p></div>"
			elseif stage = "" then
				Session("sz_uName") = "-1"
				Session("sz_uID") = -1
				Session("sz_uAccess") = -1
		%>
      <!-- #include file="nav.asp" -->
      <div id="content">
        <!-- #include file="codeChunkMsgDisp.asp" -->
        <p>
          <b>Admin login</b><br>
          Username: admin<br>
          Password: admin<br>
          <b>User login</b><br>
          Username: userA<br>
          Password: userA<br>
        </p>
        <p>
          Feel free to register your own. You can see what other users there are and how to login
          by going to Site Settings > Users List
        </p>
        <form action="login.asp" method="post">
          <input type="hidden" name="stage" value="2">
          Username: 
          <input type="text" name="uName">
          <br>
          Password: 
          <input type="password" name="uPass">
          <br>
          <input type="submit" value="Login">

        </form>
        <div id="signup">
          Not a user? <a href="register.asp">Register here</a>
        </div>
      </div>
    <%
			elseif stage = 2 then
				dim sql, uInfo
				dim uName, uPass
				
				uName = request.form("uName")
				uPass = request.form("uPass")
				
				Session("sz_errorMsg") = ""
				
				'-- Checking if user has missed a field 
				'   - if so then send message and redirect back to login
				if (uName = "") or (uPass = "") then
					Session("sz_errorMsg") = "Oops! You've left a field blank"
					response.redirect "login.asp"
				end if
				
				'-- Retrieving information to login user
				sql = "SELECT userID, accessLvl " &_
							"FROM UsersTable " &_
							"WHERE username='" & uName & "' AND userPass='" & uPass & "'"
				set uInfo = conn.execute(sql)
        
        '-- If matching username and pass wasn't found
        if uInfo.eof then
          sql = "SELECT userID " &_ 
                "FROM UsersTable " &_
                "WHERE username='" & uName & "'"
          set uInfo = conn.execute(sql)
          '-- Check if the user exists or not
          if uInfo.eof then
            Session("sz_errorMsg") = "Sorry, this user does not exist"
          else 
            Session("sz_errorMsg") = "Oops! Wrong password"
          end if
          response.redirect "login.asp"
        end if
        
        '-- Logging someone in if no problems arise
        Session("sz_uID") = uInfo(0)
        Session("sz_uName") = uName
        Session("sz_uAccess") = uInfo(1)
        Session("sz_errorMsg") = ""
%>

      <!-- #include file="nav.asp" -->
<%        
      response.write "<div id='content'><p>Welcome back, " + Session("sz_uName") + "</p>"
%>
    <p>What would you like to do?</p>
    <ul>
      <li><a href="addProject.asp">Add a new project</a></li>
      <li><a href="manage.asp">Manage my projects</a></li>
      <li><a href="forums.asp">See what's happening in the forums</a></li>
    </ul>
  </div>
<%
      '-------------'
      ' End Stage 2 '
      '-------------'
			end if
		%>
    <% conn.close %>

      <!-- #include file="footer.asp" -->
    </div>
  </body>
</html>
