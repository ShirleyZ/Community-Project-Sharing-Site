<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Edit</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        '## THERE ARE # DIFFERENT TYPES OF PAGES COMING HERE ##'
        '## - Projects - edit.asp?projID=#
        '## - Images - edit.asp?imgID=#
        '## - Comments - edit.asp?imgCmtID=#
        
        '## NOT YET IMPLEMENTED ##'
        '## - Forum Posts
        '## - Forum threads
        
        '## TELL THESE TO GET THEIR OWN PAGE ##'
        '## - Profiles
        dim stage, sql, info
        dim projID, imgID, imgCmtID
        
        stage = request.form("stage")
        if stage = "" then
          projID = request.querystring("projID")
          imgID = request.querystring("imgID")
          imgCmtID = request.querystring("imgCmtID")
          
          if imgCmtID <> "" then
            '-- Getting info about the comment
            '                 0        1        2       3         4
            sql = "SELECT imgCmtID, username, imgNum, comment, userID " &_
                  "FROM imgCommentsTable, UsersTable " &_
                  "WHERE imgCmtID=" & imgCmtID & " AND userNum=userID"
            set info = conn.execute(sql)
            
            response.write "<p>Edit Comment</p>"
            '-- This is where the error message goes
            %>
            <!-- #include file="codeChunkMsgDisp.asp" -->
            <%
            '-- Checking that you ARE the owner, not some bozo changing other people's things
            '     (unless you're admin)
            if (Session("sz_uID") <> info(4)) AND (Session("sz_uAccess") <> 1) then
              response.write "This isn't your comment! Go away!"
            else
            '-- If this is yours, show form for changing
              response.write "<form action='edit.asp' method='post'>" &_
                              "<input type='hidden' name='stage' value='2'>" &_
                              "<input type='hidden' name='entryType' value='1'>" &_
                              "<input type='hidden' name='imgCmtID' value='" &_
                              info(0) & "'>#" &_
                              info(0) &_
                              " by " &_
                              info(1) &_
                              " on image #" &_
                              info(2) &_
                              "<br>Comment: <br>" &_
                              info(3) &_
                              "<br><input type='text' name='cmtBody' value='" &_
                              info(3) &_
                              "'><input type='submit' value='Change!'></form>"
            end if
          elseif projID <> "" then
            '-- Getting info about the project
            '                0         1        2        3         4        5        6      7   
            sql = "SELECT projID,  username, userID, projName, category, dateFin, status, ProjTable.about " &_
                  "FROM ProjTable, UsersTable " &_
                  "WHERE projID=" & projID & " AND userNum=userID"
            set info = conn.execute(sql)
            
            response.write "<p>Edit Project</p>"
            '-- This is where the error message goes
            %>
            <!-- #include file="codeChunkMsgDisp.asp" -->
            <%                    
            '-- Checking that you ARE the owner, not some bozo changing other people's things
            '     (unless you're admin)
            if (Session("sz_uID") <> info(2)) AND (Session("sz_uAccess") <> 1) then
              response.write "This isn't your project! Go away!"
            else
            '-- If this is yours, show form for changing
              response.write "<form action='edit.asp' method='post'>" &_
                             "<input type='hidden' name='entryType' value='3'>" &_
                              "<input type='hidden' name='stage' value='2'>" &_
                             "<input type='hidden' name='projID' value='" &_
                             info(0) &_
                             "'>Proj #" &_
                             info(0) &_
                             " by " &_
                             info(1) &_
                             "<br>Project Name: " &_
                             "<input type=""text"" name=""projName"" value="""&_
                             info(3) &_
                             """><br>" &_
                             "Category: <select name='projCateg'>"
              '-- Grabbing and printing categories
              dim tmpStuff
              '                0          1
              sql = "SELECT categID, categName " &_
                    "FROM ProjCategoriesTable "
              set tmpStuff = conn.execute(sql)
              do
                response.write "<option value='" &_
                               tmpStuff(0) &_
                               "'>" &_
                               tmpStuff(1) &_
                               "</option>"
                tmpStuff.movenext
              loop until tmpStuff.eof
              '------------------------------------
              
              response.write "</select><br>Status: <select name='projStatus'>"
              
              '-- Grabbing and printing statuses
              '                 0       1
              sql = "SELECT statusID, status " &_
                    "FROM ProjStatusTable " &_
                    "ORDER BY statusID DESC"
              set tmpStuff = conn.execute(sql)
              do
                response.write "<option value='" &_
                               tmpStuff(0) &_
                               "'>" &_
                               tmpStuff(1) &_
                               "</option>"
                tmpStuff.movenext
              loop until tmpStuff.eof
              '---------------------------------
              
              response.write "</select><br>Date Finished: " &_
                             "<input type='text' name='dateFin' value='" &_
                             info(5) &_                         
                             "'><br>" &_
                             "About: <input type='text' name='projDesc' value=""" &_
                             info(7) &_
                             """><br>" &_
                             "<input type='submit' value='Change'></form>"
              
            end if
          elseif imgID <> "" then
            '-- Getting info about the image
            '               0         1        2       3        4       5
            sql = "SELECT imgID,  username, userID, projNum, imgUrl, comment " &_
                  "FROM ImgTable, UsersTable " &_
                  "WHERE imgID=" & imgID & " AND userNum=userID"
            set info = conn.execute(sql)
            
            response.write "<p>Edit Image</p>"
            '-- This is where the error message goes
            %>
            <!-- #include file="codeChunkMsgDisp.asp" -->
            <%                                
            '-- Checking that you ARE the owner, not some bozo changing other people's things
            '     (unless you're admin)
            if (Session("sz_uID") <> info(2)) AND (Session("sz_uAccess") <> 1) then
              response.write "This isn't your image! Go away!"
            else
              response.write "Image #" &_
                             info(0) &_
                             " by " &_
                             info(1) &_
                             "<br>Currently: <br>url: " &_
                             info(4) &_
                             "<br>comment: " &_
                             info(5) &_
                             "<br><br>"
            '-- If this is yours, show form for changing
              response.write "<form action='edit.asp' method='post'>" &_
                             "<input type='hidden' name='entryType' value='2'>" &_
                             "<input type='hidden' name='stage' value='2'>" &_
                             "<input type='hidden' name='imgID' value='" &_
                             info(0) &_
                             "'>Image URL: <input type=""text"" name=""imgURL"" value=""" &_
                             info(4) &_
                             """><br>Comment: <input type=""text"" name=""imgCmt"" value=""" &_
                             info(5) &_
                             """><br><input type='submit' value='Change'></form>"
            end if
          else
            '-- May happen when you access edit.asp without an entrytype
            Session("sz_errorMsg") = "You didn't specify what you wanted to edit! (No ID: edit.asp)"
            response.redirect "error.asp"
          end if
        elseif stage = 2 then
          dim entryType
          entryType = request.form("entryType")
          
          if entryType = 1 then '-- Editing Img comment
            dim newCmt
            
            newCmt = request.form("cmtBody")
            imgCmtID = request.form("imgCmtID")
            
            '-- Saving where you came from
            dim fromProj, fromImg
            sql = "SELECT projNum, imgNum " &_
                  "FROM ImgCommentsTable, ImgTable " &_
                  "WHERE imgCmtID=" & imgCmtID & " AND imgNum=imgID"
            set info = conn.execute(sql)
            fromProj = info(0)
            fromImg = info(1)
            
            '-- Error Checking!
            '-- Don't leave an empty comment, just delete it!
            if newCmt = "" then
              Session("sz_errorMsg") = "You left the comment blank! Perhaps you want to " &_
                                    "<a href='delete.asp?imgCmtId='" & imgCmtID & "'>"&_
                                    "delete</a> the comment instead?"
              response.redirect "edit.asp?imgCmtID=" & imgCmtID
            end if
            
            sql = "UPDATE ImgCommentsTable " &_
                  "SET comment=""" & newCmt & """ " &_
                  "WHERE imgCmtID=" & imgCmtID
            conn.execute(sql)
            response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
                           fromProj &_
                           "&imgID=" &_
                           fromImg &_
                           "'>go back</a>?"
          elseif entryType = 2 then '-- Editing Image
            dim imgUrl, imgCmt
            
            imgID = request.form("imgID")
            imgUrl = request.form("imgUrl")
            imgCmt = request.form("imgCmt")
            
            '-- Saving where you came from
            sql = "SELECT projNum " &_
                  "FROM ImgTable " &_
                  "WHERE imgID=" & imgID 
            set info = conn.execute(sql)
            fromProj = info(0)
            
            '-- Error Checking!
            '-- Don't leave an empty url, just delete it!
            if imgUrl = "" then
              Session("sz_errorMsg") = "You left the url blank! Perhaps you want to " &_
                                    "<a href='delete.asp?imgId='" & imgID & "'>"&_
                                    "delete</a> the picture instead?"
              response.redirect "edit.asp?imgID=" & imgID
            end if
            
            sql = "UPDATE ImgTable " &_
                  "SET imgUrl=""" & imgUrl & """, comment=""" & imgCmt & """ " &_
                  "WHERE imgID=" & imgID
            conn.execute(sql)
            response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
                           fromProj &_
                           "&imgID=" &_
                           imgID &_
                           "'>go back</a>?"
          elseif entryType = 3 then '-- Editing Project
            dim projName, projCateg, projStatus, projDesc, dateFin
            
            projID = request.form("projID")
            projName = request.form("projName")
            projCateg = request.form("projCateg")
            projStatus = request.form("projStatus")
            projDesc = request.form("projDesc")
            dateFin = request.form("dateFin")
            
            '-- Error Checking!
            '-- Don't leave an empty comment, just delete it!
            if projName = "" then
              Session("sz_errorMsg") = "You left the project name blank! Perhaps you want to " &_
                                    "<a href='delete.asp?projID='" & projID & "'>"&_
                                    "delete</a> the project instead?"
              response.redirect "edit.asp?projID=" & projID
            end if
            
            sql = "UPDATE ProjTable " &_
                  "SET projName=""" & projName & """, " &_
                      "category=" & projCateg & ", " &_
                      "status=" & projStatus & ", " &_
                      "about=""" & projDesc & """, " &_
                      "dateFin=" & dateFin & " " &_
                  "WHERE projID=" & projID & ";"
            'response.write sql  
            conn.execute(sql)
            response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
                           projID &_
                           "'>go back</a>?"
          else
            '-- May happen when you access edit.asp
            Session("sz_errorMsg") = "Looks like something went wrong! (Invalid entryType: edit.asp stage 2"
            response.redirect "error.asp"
          end if
        else
          Session("sz_errorMsg") = "Looks like something went wrong! (Invalid stage: edit.asp"
          response.redirect "error.asp"
        end if
      
      
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>