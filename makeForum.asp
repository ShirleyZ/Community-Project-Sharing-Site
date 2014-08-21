<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Create Forum Item</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <%
        '## makeForum.asp ##'
        
        '## This page will be handling
        
        '-- User-side
        '1 - Making a new Thread - makeForum.asp?subCateg=#&newThread=1
        '2- Making a new Post - makeForum.asp?subCateg=#&threadID=#&newPost=1
        
        '-- Admin-side
        '3 - Making a new Subforum - makeForum.asp?newCateg=1

        dim newThread, newPost, newCateg
        newThread = request.querystring("newThread")
        newPost = request.querystring("newPost")
        newCateg = request.querystring("newCateg")
        
        dim subCateg, threadID
        subCateg = request.querystring("subCateg")
        threadID = request.querystring("threadID")
        
        dim datePost
        datePost = date()
        
        dim stage
        stage = request.form("stage")
        
        dim sql, info
        '---------------------------'
        '-- Making a new Category --'
        '---------------------------'
        if newCateg <> "" then
          '-- Error, if people are fucking around with the string
          if (newThread <> "") or (newPost <> "") then
            Session("sz_errorMsg") = "Stop screwing around with the url!"
            response.redirect "error.asp"
          end if
          '-- Error, if you don't have permission to do this
          if Session("sz_uAccess") <> 1 then
            Session("sz_errorMsg") = "You don't have permission to do that! Go away, shoo!"
            response.redirect "error.asp"
          end if
          response.write "<p class='pageTitle'>Create Category</p>"
          
          '--------------------------------'
          '-- STAGE 1 OF CREATING CATEG --'
          '--------------------------------'
          if stage = "" then
            '-- Writing the form
            %>
            <!-- #include file='codeChunkMsgDisp.asp' -->
            <form action='makeForum.asp?newCateg=1' method='post'>
              <input type='hidden' name='stage' value='2'>
              Category name: <input type='text' name='categName'><br>
              Category permission: <input type='number' name='categAccess' value='2'><br>
              <input type='submit' value='Create'>
            </form>
            <%
          
          '--------------------------------'
          '-- STAGE 2 OF CREATING CATEG --'
          '--------------------------------'
          elseif stage = 2 then
            dim categName, categAccess
            categName = request.form("categName")
            categAccess = request.form("categAccess")
            
            '-- Checking if you left an empty subject
            if (categName = "") or (categAccess = "") then
              Session("sz_errorMsg") = "You need to fill in all fields!"
              response.redirect "makeForum.asp?newCateg=1"
            end if
            
            '-- Check it doesn't already exist
            sql = "SELECT fCategID, fCategName " &_
                  "FROM ForumCategoriesTable " &_
                  "WHERE fCategName=""" & categName & """"
            'response.write sql
            set info = conn.execute(sql)
            
            if NOT info.eof then
              Session("sz_errorMsg") = "This category already exists"
              response.redirect "makeForum.asp?newCateg=1"
            end if
            
            '-- Adding category
            sql = "INSERT INTO ForumCategoriesTable(fCategName, fCategAccessLvl) " &_
                  "VALUES(""" & categName & """, " & categAccess & ") "
            'response.write "<br>" & sql
            conn.execute(sql)
            
            '-- Done
            response.write "Done! Would you like to <a href='forums.asp'>view the forums</a>?"
          else
            response.write "something is wrong and you made it wrong (categ)"
          end if
          
        '-------------------------'
        '-- Making a new Thread --'
        '-------------------------'
        elseif newThread <> "" then
          '-- Error, if people are fucking around with the string
          if (newCateg <> "") or (newPost <> "") or (subCateg = "") then
            Session("sz_errorMsg") = "Stop screwing around with the url!"
            response.redirect "error.asp"
          end if
          response.write "<p class='pageTitle'>Create Thread</p>"
          
          '--------------------------------'
          '-- STAGE 1 OF CREATING THREAD --'
          '--------------------------------'
          if stage = "" then
            '-- Writing out form
          %>
          <!-- #include file='codeChunkMsgDisp.asp' -->
          <%
            response.write "<form action='makeForum.asp?subCateg=" &_
                            subCateg &_
                            "&newThread=1' method='post'>"
            %>
              <input type='hidden' name='stage' value='2'>
              Thread Title: <input type='text' name='threadName'><br>
              Thread Post: <input type='text' name='threadPost'><br>
              <input type='submit' value='Create'>
            </form>
            <%    
          '--------------------------------'
          '-- STAGE 2 OF CREATING THREAD --'
          '--------------------------------'
          elseif stage = 2 then
            dim threadName, threadPost
            threadName = request.form("threadName")
            threadPost = request.form("threadPost")
            
            '-- Checking if you left an empty subject or body
            if (threadName = "") or (threadPost = "") then
              Session("sz_errorMsg") = "Please don't leave the Subject or first post empty!"
              response.redirect "makeForum.asp?subCateg=" & subCateg & "&newThread=1"
            end if
            
            '-- Retrieving info abt the categ
            '                 0               1
            sql = "SELECT fCategName, fCategAccessLvl " &_
                  "FROM ForumCategoriesTable " &_
                  "WHERE fCategID=" & subCateg
            dim categInfo
            set categInfo = conn.execute(sql)
            
            '-- Checking if you have permission to post in that forum
            if (Session("sz_uAccess") > categInfo(1)) then
              Session("sz_errorMsg") = "You don't have permission to create threads in this sub forum!<br>"
              response.redirect "makeForum.asp?subCateg=" & subCateg & "&newThread=1"
            end if
            
            '-- Adding thread
            sql = "INSERT INTO ForumThreadsTable(ownerNum, forumCat, threadSubj, datePost, lastPost) " &_
                  "VALUES (" & Session("sz_uID") & ", " & subCateg & ", """ & threadName & """, """ &_
                  datePost & """, """ & datePost & """)"
            response.write sql
            conn.execute(sql)
            
            '-- Getting thread ID
            sql = "SELECT threadID " &_
                  "FROM ForumThreadsTable " &_
                  "WHERE forumCat=" & subCateg & " " &_
                  "ORDER BY threadID DESC"
            dim threadInfo
            set threadInfo = conn.execute(sql)
            
            '-- Adding first post
            sql = "INSERT INTO ForumPostsTable(ownerNum, threadNum, comment, datePost) " &_
                  "VALUES (" & Session("sz_uID") & ", " & threadInfo(0) & ", """ & threadPost & """, """ & datePost & """) "
            response.write sql
            response.write "<br>" & threadInfo(0)
            conn.execute(sql)
            
            '-- Redirecting to the new thread
            sql = "SELECT threadID, forumCat " &_
                  "FROM ForumThreadsTable " &_
                  "ORDER BY threadID DESC"
            set info = conn.execute(sql)
            response.redirect "forums.asp?subCateg=" & info(1) & "&threadID=" & info(0)
            
          else
            response.write "something is wrong and you made it wrong (thread stage incorrect)"
          end if
          
        '-----------------------'
        '-- Making a new Post --'
        '-----------------------'
        elseif newPost <> "" then
          '-- Error, if people are fucking around with the string
          if (newCateg <> "") or (newThread <> "") or (subCateg = "") or (threadID="") then
            Session("sz_errorMsg") = "Stop screwing around with the url!"
            response.redirect "error.asp"
          end if
          
          dim cmtBody
          cmtBody = request.form("cmtBody")
          
          '-- Checking it isn't empty
          if cmtBody = "" then
            Session("sz_errorMsg") = "You can't leave an empty reply!"
            response.redirect "error.asp"
          end if
          
          '-- Adding the comment
          sql = "INSERT INTO ForumPostsTable(ownerNum, threadNum, comment, datePost) " &_
                "VALUES (" & Session("sz_uID") & ", " & threadID & ", """ & cmtBody & """, """ & datePost & """) "
          response.write sql
          response.write "<br>" & datePost
          conn.execute(sql)
          
          '-- updating threads when the last post was
          sql = "UPDATE ForumThreadsTable " &_
                "SET datePost=""" & datePost & """ "
          response.write "<br>" & sql
          conn.execute(sql)
          
          '-- Done and redirect
          response.redirect "forums.asp?subCateg=" & subCateg & "&threadID=" & threadID
        else
          response.write "A thing went wrong (why are you here)"
        end if
        
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>