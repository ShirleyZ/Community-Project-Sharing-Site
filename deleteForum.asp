<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Delete Forum Item</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        '## deleteForum.asp ##'
        
        '## This page will be handling
        
        '-- User-side
        '## - Deleting a Thread - deleteForum.asp?subCateg=#&deleteThread=#
        '## - Deleting a Post - deleteForum.asp?subCateg=#&threadID=#&deletePost=#
        
        '-- Admin-side
        '## - Editing/Deleting a Subforum - editForum.asp?editCateg=#
        
        dim deletePost, deleteThread, deleteCateg
        deletePost = request.querystring("deletePost")
        deleteThread = request.querystring("deleteThread")
        deleteCateg = request.querystring("deleteCateg")
        
        dim subCateg, threadID
        subCateg = request.querystring("subCateg")
        threadID = request.querystring("threadID")
        
        dim sql, info, stage
        stage = request.form("stage")
        
        '---------------------'
        '-- Deleting a Post --'
        '---------------------'
        if deletePost <> "" then
          '------------------------------'
          '-- STAGE 1 OF DELETING POST --'
          '------------------------------'
          if stage = "" then
            '-- Get post
            '                 0        1
            sql = "SELECT ownerNum, comment " &_
                  "FROM ForumPostsTable " &_
                  "WHERE postID=" & deletePost
            set info = conn.execute(sql)
            
            response.write "<p class='pageTitle'>Deleting Post</p>"
            
            '-- Check post exists
            if info.eof then
              response.write "Post #" & deletePost & " doesn't exist"
              
            '-- Do you have permission
            elseif (Session("sz_uAccess") <> 1) AND (Session("sz_uID") <> info(0)) then
              response.write "You don't have permission to do this!"
            else
              '-- Print to confirm
              response.write "<p>Are you <b>certain</b> you want to delete this comment?</p>"
              response.write "<form action='deleteForum.asp?subCateg=" &_
                              subCateg &_
                              "&threadID=" & threadID &_
                              "&deletePost=" & deletePost &_
                             "' method='post'><input type='hidden' name='stage' value='2'>Comment: <br>" &_
                             info(1) &_
                             "<br><input type=""submit"" value=""Yes I'm sure!""></form>" &_
                             "<a href='forums.asp?subCateg=" & subCateg & "&threadID=" & threadID &_
                             "'>No, take me back!</a>"
            end if
          '------------------------------'
          '-- STAGE 2 OF DELETING POST --'
          '------------------------------'
          elseif stage = 2 then
            '-- Delete post 
            sql = "DELETE FROM ForumPostsTable " &_
                  "WHERE postID=" & deletePost
            response.write sql
            conn.execute(sql)
            response.redirect "forums.asp?subCateg=" & subCateg & "&threadID=" & threadID
          else
            Session("sz_errorMsg") = "Looks like something went wrong! (Invalid stage for Post: deleteForum.asp)"
            response.redirect "error.asp"
          end if
          
          
        '-----------------------'
        '-- Deleting a Thread --'
        '-----------------------'
        elseif deleteThread <> "" then
          '--------------------------------'
          '-- STAGE 1 OF DELETING THREAD --'
          '--------------------------------'
          if stage = "" then
            '-- Get thread
            '                 0        1
            sql = "SELECT ownerNum, threadSubj " &_
                  "FROM ForumThreadsTable " &_
                  "WHERE threadID=" & deleteThread
            set info = conn.execute(sql)
            
            response.write "<p class='pageTitle'>Deleting Thread</p>"
            
            '-- Check thread exists
            if info.eof then
              response.write "Thread #" & deleteThread & " doesn't exist"
              
            '-- Do you have permission
            elseif (Session("sz_uAccess") <> 1) AND (Session("sz_uID") <> info(0)) then
              response.write "You don't have permission to do this!"
            else
              '-- Print to confirm
              response.write "<p>Are you <b>certain</b> you want to delete this thread?</p>"
              response.write "<form action='deleteForum.asp?subCateg=" &_
                              subCateg &_
                              "&deleteThread=" & deleteThread &_
                             "' method='post'><input type='hidden' name='stage' value='2'>Subject: <br>" &_
                             info(1) &_
                             "<br><input type=""submit"" value=""Yes I'm sure!""></form>" &_
                             "<a href='forums.asp?subCateg=" & subCateg & "&threadID=" & threadID &_
                             "'>No, take me back!</a>"
            end if
            
          '--------------------------------'
          '-- STAGE 2 OF DELETING THREAD --'
          '--------------------------------'
          elseif stage = 2 then
            '-- Delete posts in thread
            sql = "DELETE FROM ForumPostsTable " &_
                  "WHERE threadNum=" & deleteThread
            response.write sql
            conn.execute(sql)
            
            '-- Delete thread 
            sql = "DELETE FROM ForumThreadsTable " &_
                  "WHERE threadID=" & deleteThread
            response.write sql
            conn.execute(sql)
            
            '-- Done, Redirecting
            response.redirect "forums.asp?subCateg=" & subCateg
          else
            Session("sz_errorMsg") = "Looks like something went wrong! (Invalid stage for Thread: deleteForum.asp)"
            response.redirect "error.asp"
          end if
        
        
        '-------------------------'
        '-- Deleting a Category --'
        '-------------------------'
        elseif deleteCateg <> "" then
          '----------------------------------'
          '-- STAGE 1 OF DELETING CATEGORY --'
          '----------------------------------'
          if stage = "" then
            '-- Get CATEGORY
            '                 0        1
            sql = "SELECT fCategName, fCategAccessLvl " &_
                  "FROM ForumCategoriesTable " &_
                  "WHERE fCategID=" & deleteCateg
            set info = conn.execute(sql)
            
            response.write "<p class='pageTitle'>Deleting Category</p>"
            
            '-- Check CATEGORY exists
            if info.eof then
              response.write "Category #" & deleteCateg & " doesn't exist"
              
            '-- Do you have permission
            elseif (Session("sz_uAccess") <> 1) then
              response.write "You don't have permission to do this!"
            else
              '-- Print to confirm
              response.write "<p>Are you <b>certain</b> you want to delete this category?</p>"
              response.write "<form action='deleteForum.asp?deleteCateg=" & deleteCateg &_
                             "' method='post'><input type='hidden' name='stage' value='2'>Name: <br>" &_
                             info(0) &_
                             "<br>Access Level: " &_
                             info(1) &_
                             "<br><input type=""submit"" value=""Yes I'm sure!""></form>" &_
                             "<a href='forums.asp?subCateg=" & deleteCateg &_
                             "'>No, take me back!</a>"
            end if
          '----------------------------------'
          '-- STAGE 2 OF DELETING CATEGORY --'
          '----------------------------------'
          elseif stage = 2 then
            '-- Delete posts in threads
            '-- get all threads in this category
            dim listThreads
            sql = "SELECT threadID " &_
                  "FROM ForumThreadsTable " &_
                  "WHERE forumCat=" & deleteCateg
            response.write sql & "<br>"
            set listThreads = conn.execute(sql)
            
            do until listThreads.eof
              sql = "DELETE FROM ForumPostsTable " &_
                    "WHERE threadNum=" & listThreads(0)
              response.write sql & "<br>"
              conn.execute(sql)
              listThreads.movenext
            loop
            
            '-- Delete threads in category
            sql = "DELETE FROM ForumThreadsTable " &_
                  "WHERE forumCat=" & deleteCateg
            response.write sql & "<br>"
            conn.execute(sql)
            
            '-- Delete category
            sql = "DELETE FROM ForumCategoriesTable " &_
                  "WHERE fCategID=" & deleteCateg
            response.write sql & "<br>"
            conn.execute(sql)
            
            '-- Done, Redirecting
            response.redirect "forums.asp"
          else
            Session("sz_errorMsg") = "Looks like something went wrong! (Invalid stage for Category: deleteForum.asp)"
            response.redirect "error.asp"
          end if
        else
          Session("sz_errorMsg") = "Looks like something went wrong! (Missing item to delete: deleteForum.asp)"
          response.redirect "error.asp"
        end if
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>