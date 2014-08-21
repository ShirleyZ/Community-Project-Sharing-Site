<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Edit Forum Item</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <%
          '## editForum.asp ##'
          
          '## This page will be handling
          
          '-- User-side
          '## - Editing a Thread - editForum.asp?subCateg=#&editThread=#
          '## - Editing a Post - editForum.asp?subCateg=#&threadID=#&editPost=#
          
          '-- Admin-side
          '## - Editing a Subforum - editForum.asp?editCateg=#
          
          dim editPost, editThread, editCateg
          editPost = request.querystring("editPost")
          editThread = request.querystring("editThread")
          editCateg = request.querystring("editCateg")
          
          dim subCateg, threadID
          subCateg = request.querystring("subCateg")
          threadID = request.querystring("threadID")
          
          dim sql, info, stage, datePost
          stage = request.form("stage")
          datePost = date()
          
        '--------------------'
        '-- Editing a Post --'
        '--------------------'
          if editPost <> "" then
            '-- Error, if people are fucking around with the string
            if (subCateg = "") or (threadID = "") then
              Session("sz_errorMsg") = "Stop screwing around with the url!"
              response.redirect "error.asp"
            end if
            
            '-----------------------------'
            '-- STAGE 1 OF EDITING POST --'
            '-----------------------------'
            if stage = "" then
              '-- Getting info of post
              '                 0         1         2         3
              sql = "SELECT ownerNum, threadNum, comment, username " &_
                    "FROM ForumPostsTable, UsersTable " &_
                    "WHERE postID=" & editPost & " AND ownerNum=userID"
              set info = conn.execute(sql)
              
              '-- Error, this post doesn't exist
              if info.eof then
                Session("sz_errorMsg") = "This post doesn't exist!"
                response.redirect "error.asp"
              end if
              
              '-- Error, you don't have permission to do this
              if (Session("sz_uID") <> info(0)) AND (Session("sz_uAccess") <> 1) then
                Session("sz_errorMsg") = "You don't have permission to do that!"
                response.redirect "error.asp"
              end if
              
              '-- Printing out form
              %>
              <p class='pageTitle'>Editing Post</p>
              <!-- #include file='codeChunkMsgDisp.asp' -->
              <%
              response.write "<form action='editForum.asp?subCateg=" & subCateg &_
                             "&threadID=" & threadID & "&editPost=" & editPost & "' method='post'>" &_
                             "<input type='hidden' name='stage' value='2'>" &_
                             "Post #" & editPost & " by " &_
                             info(3) &_
                             "<br>Comment: " &_
                             info(2) &_
                             "<br><input type='text' name='cmtBody' value='" &_
                             info(2) &_
                             "'><br><input type='submit' value='Change'></form>"
              
            '-----------------------------'
            '-- STAGE 2 OF EDITING POST --'
            '-----------------------------'
            elseif stage = 2 then
              dim cmtBody
              cmtBody = request.form("cmtBody")
              
              '-- Checking that it isn't empty
              if cmtBody = "" then
                Session("sz_errorMsg") = "You left the post blank! Perhaps you want to " &_
                                  "<a href='deleteForum.asp?subCateg=" & subCateg &_
                                  "&threadID=" & threadID & "&deletePost=" & editPost & "'>" &_
                                  "delete</a> the post instead?"
                response.redirect "editForum.asp?subCateg=" & subCateg & "&threadID=" &_
                                  threadID & "&editPost=" & editPost
              end if
              
              '-- Changing the comment
              sql = "UPDATE ForumPostsTable " &_
                    "SET comment=""" & cmtBody & """, lastEdit=""" & datePost & """ " &_
                    "WHERE postID=" & editPost
              response.write sql
              conn.execute(sql)
              
              '-- Done, redirecting
              response.redirect "forums.asp?subCateg=" & subCateg & "&threadID=" & threadID
            else
              response.redirect "There is nothing here (wrong stage)"
            end if
          
        '----------------------'
        '-- Editing a Thread --'
        '----------------------'
          elseif editThread <> "" then
            '-- Error, if people are fucking around with the string
            if (subCateg = "") then
              Session("sz_errorMsg") = "Stop screwing around with the url!"
              response.redirect "error.asp"
            end if
            
            '-------------------------------'
            '-- STAGE 1 OF EDITING THREAD --'
            '-------------------------------'
            if stage = "" then
              '-- Getting info of thread
              '                 0         1         2         
              sql = "SELECT ownerNum, threadSubj, username " &_
                    "FROM ForumThreadsTable, UsersTable " &_
                    "WHERE threadID=" & editThread & " AND ownerNum=userID"
              set info = conn.execute(sql)
              
              '-- Error, this post doesn't exist
              if info.eof then
                Session("sz_errorMsg") = "This thread doesn't exist!"
                response.redirect "error.asp"
              end if
              
              '-- Error, you don't have permission to do this
              if (Session("sz_uID") <> info(0)) AND (Session("sz_uAccess") <> 1) then
                Session("sz_errorMsg") = "You don't have permission to do that!"
                response.redirect "error.asp"
              end if
              
              '-- Printing out form
              %>
              <p class='pageTitle'>Editing Thread</p>
              <!-- #include file='codeChunkMsgDisp.asp' -->
              <%
              response.write "<form action='editForum.asp?subCateg=" & subCateg &_
                             "&editThread=" & editThread & "' method='post'>" &_
                             "<input type='hidden' name='stage' value='2'>" &_
                             "Thread #" & editThread & " by " &_
                             info(2) &_
                             "<br>Subject: " &_
                             info(1) &_
                             "<br><input type='text' name='threadSubj' value='" &_
                             info(1) &_
                             "'><br>"
              '-- Moving the thread to another category
              if Session("sz_uAccess") = 1 then
                sql = "SELECT fCategID, fCategName " &_
                      "FROM ForumCategoriesTable " 
                dim listCateg
                set listCateg = conn.execute(sql)
                
                response.write "Move Category: <select name='moveToCateg'>"
                do until listCateg.eof
                  response.write "<option value=""" &_
                                  listCateg(0) &_
                                  """>" &_
                                  listCateg(1) &_
                                  "</option>"
                  listCateg.movenext
                loop
                response.write "</select><br>"
              end if
              
              response.write "<input type='submit' value='Change'></form>"
            '-------------------------------'
            '-- STAGE 2 OF EDITING THREAD --'
            '-------------------------------'
            elseif stage = 2 then
              dim threadSubj
              threadSubj = request.form("threadSubj")
              
              '-- Checking that it isn't empty
              if threadSubj = "" then
                Session("sz_errorMsg") = "You left the subject blank! Perhaps you want to " &_
                                  "<a href='deleteForum.asp?subCateg=" & subCateg &_
                                  "&deleteThread=" & editThread & "'>" &_
                                  "delete</a> the thread instead?"
                response.redirect "editForum.asp?subCateg=" & subCateg & "&editThread=" &_
                                  editThread
              end if
              
              '-- Changing the subject
              sql = "UPDATE ForumThreadsTable " &_
                    "SET threadSubj=""" & threadSubj & """ " &_
                    "WHERE threadID=" & editThread
              response.write sql
              conn.execute(sql)
              
              '-- Changing category
              dim moveToCateg
              moveToCateg = request.form("moveToCateg")
              
              if moveToCateg <> "" then
                sql = "UPDATE ForumThreadsTable " &_
                      "SET forumCat=" & moveToCateg & " " &_
                      "WHERE threadID=" & editThread
                conn.execute(sql)
              end if
              '-- Done, redirecting
              response.redirect "forums.asp?subCateg=" & subCateg & "&threadID=" & editThread
            
            else
              response.write "This is an error. You did this. (Thread, wrong stage)"
            end if
        '------------------------'
        '-- Editing a Category --'
        '------------------------'
          elseif editCateg <> "" then
            '-- Error, if you don't have permission to do this
            if Session("sz_uAccess") <> 1 then
              Session("sz_errorMsg") = "You don't have permission to do that! Go away, shoo!"
              response.redirect "error.asp"
            end if
            
            '---------------------------------'
            '-- STAGE 1 OF EDITING CATEGORY --'
            '---------------------------------'
            if stage = "" then
              '-- Getting info for Category
              '                  0            1
              sql = "SELECT fCategName, fCategAccessLvl " &_
                    "FROM ForumCategoriesTable " &_
                    "WHERE fCategID=" & editCateg
              set info = conn.execute(sql)
              %>
              <p class='pageTitle'>Editing Category</ap>
              <!-- #include file='codeChunkMsgDisp.asp' -->
              <%
              '-- Check category exists
              if info.eof then
                response.write "Category #" & editCateg & " doesn't exist!"
              else
              '-- Writing out the form
                response.write "<form action='editForum.asp?editCateg=" & editCateg & "' method='post'>"
              %>
                  <input type='hidden' name='stage' value='2'>
                  
                  Category Name: <input type='text' name='categName' value='<% response.write info(0) %>'><br>
                  Category permission: <input type='number' name='categAccess' value='<% response.write info(1) %>'><br>
                  <input type='submit' value='Change'>
                </form>
              <%
              end if
            '---------------------------------'
            '-- STAGE 2 OF EDITING CATEGORY --'
            '---------------------------------'
            elseif stage = 2 then
              dim categName, categAccess
              categName = request.form("categName")
              categAccess = request.form("categAccess")
              
              '-- Checking that it isn't empty
              if categName = "" then
                Session("sz_errorMsg") = "You left the name blank! Perhaps you want to " &_
                                  "<a href='deleteForum.asp?deleteCateg=" & editCateg &_
                                  "'>" &_
                                  "delete</a> the category instead?"
                response.redirect "editForum.asp?editCateg=" & editCateg 
              end if
              '-- Checking access isn't empty
              if categAccess = "" then
                Session("sz_errorMsg") = "Please don't leave the access level blank"
                response.redirect "editForum.asp?editCateg=" & editCateg 
              end if
              
              '-- Changing the category
              sql = "UPDATE ForumCategoriesTable " &_
                    "SET fCategName=""" & categName & """, " &_
                    "fCategAccessLvl=" & categAccess & " " &_
                    "WHERE fCategID=" & editCateg
              response.write sql
              conn.execute(sql)
              
              '-- Done, redirecting
              response.redirect "forums.asp"
            
            else
              response.write "This is an error. You did this. (categ, wrong stage)"
            end if
          else
            response.write "Another day another moron messing with the url"
          end if
        %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>