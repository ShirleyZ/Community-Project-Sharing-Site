<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Make Comment</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
<%
  '-- Figuring out how to deal with this comment
  '## Entry Types are ##'
  '## o 1 - from viewProjects.asp
  '## o 2 - from viewProfile.asp

  dim sql, info
  dim entryType
  entryType = request.form("entryType")
  
  '-- Stuff used in both cases
  dim userID, datePost, cmtBody
  
  if entryType = 1 then '-- Commenting on a project
    '-- Yoinking info from form
    dim projID, imgID
    userID = Session("sz_uID")
    projID = request.form("projID")
    imgID = request.form("imgID")
    cmtBody = request.form("cmtBody")
    datePost = date()

    '-- Inserting into database
    sql = "INSERT INTO ImgCommentsTable(userNum, imgNum, comment, datePost) " &_
          "VALUES (""" & userID & """, " & imgID & ", """ & cmtBody & """, """ & datePost & """)"
          
    conn.execute(sql)
    
    response.write sql
    response.write "<br>" & datePost
    
    '-- Redirecting you back, now that it's done
    response.redirect "viewProject.asp?projID=" & projID & "&imgID=" & imgID
    
  elseif entryType = 2 then '-- Commenting on a profile
    '-- Yoinking info from form
    dim profID
    profID = request.form("profID")
    userID = Session("sz_uID")
    cmtBody = request.form("cmtBody")
    datePost = date()
    
    '-- Checking that they're not leaving an empty comment
    if cmtBody = "" then
      Session("sz_errorMsg") = "You can't leave an empty comment!"
      response.redirect "viewProfile.asp?userID" & profID
    end if
    
    '-- Inserting into Database
    sql = "INSERT INTO UserPgCommentsTable (userNum, profileNum, comment, datePost) " &_
          "VALUES (" & userID & ", " & profID & ", """ & cmtBody & """, """ & datePost & """)"
    response.write sql
    conn.execute(sql)
    
    '-- Redirect back to the profile
    response.redirect "viewProfile.asp?userID=" & profID
  else
    response.write "how did you get there!"
  end if
  
%>
      
      
      
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>