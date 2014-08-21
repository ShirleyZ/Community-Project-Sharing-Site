<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Forums</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="forumTableCSS.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">

  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <p class="pageTitle">Forums</p>
        <!-- #include file="codeChunkMsgDisp.asp" -->
        <div id="forumList">
      <%
        dim sql, info
        
        dim subCateg, threadID
        subCateg = request.querystring("subCateg")
        threadID = request.querystring("threadID")
        
        '----------------------'
        '-- Viewing a thread --'
        '----------------------'
        
        if threadID <> "" then
          dim threadInfo, threadPosts
          '-- Getting info about the thread
          '                  0         1          2         3
          sql = "SELECT threadSubj, ownerNum, forumCat, fCategName " &_
                "FROM ForumThreadsTable, ForumCategoriesTable " &_
                "WHERE threadID=" & threadID & " AND fCategID=forumCat"
          set threadInfo = conn.execute(sql)
          
          '-- Getting posts in this thread
          '                0        1         2         3        4         5        6
          sql = "SELECT postID, username, ownerNum, comment, datePost, lastEdit, userImg  " &_
                "From ForumPostsTable, UsersTable " &_
                "WHERE threadNum=" & threadID & " " &_
                "AND ownerNum=userID " &_
                "ORDER BY postID"
          set threadPosts = conn.execute(sql)
          
          '-- Navigation to go back to subforum
          %>
          <div class='forumPost'>
            <div class='forumPostRight'>
            <%
              '-- Navigation to previous pages
              response.write "<a href='forums.asp'>Forums</a> >> <a href='forums.asp?subCateg=" &_
                             threadInfo(2) &_
                             "'>" &_
                             threadInfo(3) &_
                             "</a> >> <a href='forums.asp?subCateg='" &_
                             threadInfo(2) &_
                             "&threadID=" &_
                             threadID &_
                             "'>" &_
                             threadInfo(0) &_
                             "</a>"
            %>
            </div>
            <br clear='left'>
          </div>
          <%
          '-- Options at Beginning for owner/admin
          if (Session("sz_uID") = threadInfo(1)) or (Session("sz_uAccess") = 1) then
            response.write "<div class='forumPost'><div class='forumPostRight'><a href='editForum.asp?subCateg=" &_
                            threadInfo(2) &_
                            "&editThread=" &_
                            threadID &_
                            "'>Edit Thread</a> <a href='deleteForum.asp?subCateg=" &_
                            threadInfo(2) &_
                            "&deleteThread=" &_
                            threadID &_
                            "'>Delete Thread</a></div><br clear='left'></div>"
          end if
          '-- Printing out thread name
          response.write "<div class='forumPost'><div class='forumPostTitle'>" &_
                          threadInfo(0) &_
                          "</div></div>"
          '-- Printing out posts
          do until threadPosts.eof
            %>
              <div class='forumPost'>
                <div class='forumPostLeft'>
                  <%
                    response.write "<img class='forumPostUserPic' src='" &_
                                    threadPosts(6) &_
                                    "' />"
                    response.write "<div class='forumPostUserInfo'>" &_
                                    threadPosts(1) &_
                                    "</div>"
                  %>
                </div>
                <div class='forumPostRight'>
                  <div class='forumPostCmtBody'>
                    <%
                      response.write threadPosts(3)
                    %>
                  </div>
                  <div class='forumPostMeta'>
                    <%
                      response.write "#" &_
                                      threadPosts(0) &_
                                      " Posted: " &_
                                      threadPosts(4)
                      if threadPosts(5) <> "" then
                        response.write " Last Edit: " &_
                                       threadPosts(5)
                      end if
                    %>
                    <%
                      '-- editing rites
                      if (Session("sz_uID") = threadPosts(2)) or (Session("sz_uAccess") = 1) then
                        response.write " <a href='editForum.asp?subCateg=" &_
                            threadInfo(2) &_
                            "&threadID=" &_
                            threadID &_
                            "&editPost=" &_
                            threadPosts(0) &_
                            "'>Edit Post</a> <a href='deleteForum.asp?subCateg=" &_
                            threadInfo(2) &_
                            "&threadID=" &_
                            threadID &_
                            "&deletePost=" &_
                            threadPosts(0) &_
                            "'>Delete Post</a>"
                      end if
                    %>
                  </div>
                </div>
                <br clear="left">
              </div>
            <%
            threadPosts.movenext
          loop
          '-- Box for replying to thing
          response.write "<form action='makeForum.asp?subCateg=" &_
                        threadInfo(2) &_
                        "&threadID=" &_
                        threadID &_
                        "&newPost=1' method='post'>"
          %>
            <input type='text' name='cmtBody'>
            <input type='submit' value='Reply'>
          </form>
          <%
          
          
        '----------------------------'
        '-- Viewing the sub forums --'
        '----------------------------'
        
        elseif subCateg <> "" then
          '-- Getting the different threads available
          '                 0        1          2          3          4
          sql = "SELECT threadID, username, forumCat, threadSubj, lastPost " &_
                "FROM ForumThreadsTable, UsersTable " &_
                "WHERE forumCat=" &_
                subCateg &_
                " AND userID=ownerNum " &_
                "ORDER BY lastPost DESC"
          set info = conn.execute(sql)
          
          '-- Navigation to previous pages
          response.write "<div class='forumModule'><a href='forums.asp'>Forums</a></div>"
                         
          '-- Options at beginning
          %>
        <div class='forumModule'>
          <%
          response.write "<a href='makeForum.asp?subCateg=" &_
                          subCateg &_
                          "&newThread=1'>New Thread</a>"
          %>
        </div>
        <br clear="both">
        <%
          
          '-- Checking that there are threads
          if info.eof then
            response.write "<div class='forumModule'>" &_
                           "Problem: No Threads found</div>"
          else
            '-- Printing out the threads
            do until info.eof
              '-- Getting the last post for this category
              '                 0       1
              sql = "SELECT postID, datePost " &_
                    "FROM ForumPostsTable " &_
                    "WHERE threadNum=" &_
                    info(0) &_
                    " ORDER BY datePost DESC"
              set fLastPost = conn.execute(sql)
              %>
          <div class='forumModule'>
            <p class='forumModuleName'>
          <%
              response.write "<a href='forums.asp?subCateg=" & subCateg &_
                             "&threadID=" &_
                              info(0) &_
                              "'>" &_
                              info(3) &_
                              "</a>"
          %>

            </p>
            <p class='forumModuleMeta'>
              <span class='forumModuleCreator'>
              <%
                response.write "created by " &_
                               info(1)
              %>
              </span>
              <span class='forumModuleLastPost'>
              <%
                '-- Checking that there are posts in here
                if fLastPost.eof then
                  response.write "No posts"
                else
                  response.write "Last post @ " &_
                                 fLastPost(1)
                end if
              %></span>
            </p>
          </div>
          <%
              info.movenext
            loop
          end if
          
          '-- Options at end
          %>
          <br clear="both">
        <div class='forumModule'>
          <%
          response.write "<a href='makeForum.asp?subCateg=" &_
                          subCateg &_
                          "&newThread=1'>New Thread</a>"
          %>
        </div>
        <%
        
        '------------------------'
        '-- Viewing the forums --'
        '------------------------'
        
        else
          '-- Getting the different subForums available
          '                 0          1
          sql = "SELECT fCategID, fCategName " &_
                "FROM ForumCategoriesTable " &_
                "ORDER BY fCategID"
          set info = conn.execute(sql)
          
          response.write "<br clear='all'>"
          '-- Options at beginning for admin
          if Session("sz_uAccess") = 1 then
          %>
        <div class='forumModule'>
          <a href='makeForum.asp?newCateg=1'>New Subforum</a> 
          <a href='editForum.asp?editCateg=1'>Edit/Delete a Subforum</a>
        </div>
        <br clear="both">
        <%
          end if
          
          '-- Checking there are infos
          if info.eof then 
            response.write "<div class='forumModule'>" &_
                           "Problem: No Categories found</div>"
          else
            dim fLastPost
            '-- Printing out forum categories
            do until info.eof
              '-- Getting the last post for this category
              '                 0          1          2
              sql = "SELECT threadID, threadSubj, datePost " &_
                    "FROM ForumThreadsTable " &_
                    "WHERE forumCat=" &_
                    info(0) &_
                    " ORDER BY datePost DESC"
              set fLastPost = conn.execute(sql)
          %>
          <div class='forumModule'>
            <p class='forumModuleName'>
          <%
              response.write "<a href='forums.asp?subCateg=" &_
                              info(0) &_
                              "'>" &_
                              info(1) &_
                              "</a>"
          %>

            </p>
            <p class='forumModuleMeta'>
              <span class='forumModuleLastPost'>
              <%
                '-- Checking that there are posts in here
                if fLastPost.eof then
                  response.write "No posts"
                else
                  response.write "Last post: <a href='forums.asp?subCateg=" &_
                                 info(0) &_
                                 "&threadID=" &_
                                 fLastPost(0) &_
                                 "'>" &_
                                 fLastPost(1)&_
                                 "</a> @ " &_
                                 fLastPost(2)
                end if
              %></span>
              
              <%
                if Session("sz_uAccess") = 1 then
                  response.write "<span><a href='editForum.asp?editCateg=" &_
                                  info(0) &_
                                  "'>Edit Category</a> <a href='deleteForum.asp?deleteCateg=" &_
                                  info(0) &_
                                  "'>Delete Category</a></span>"
                end if
              %>
            </p>
          </div>
          <%
            info.movenext
            loop
          end if
          
          response.write "<br clear='all'>"
          '-- Options at end for admin
          if Session("sz_uAccess") = 1 then
          %>
        <div class='forumModule'>
          <a href='makeForum.asp?newCateg=1'>New Subforum</a> 
        </div>
        <%
          end if
        end if
      %>
        <br clear="all">
        </div> <!-- End of #forumList -->
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div>
    <!-- End wrapper -->
  </body>
</html>