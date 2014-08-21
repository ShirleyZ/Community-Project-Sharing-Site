<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: Default</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <% dim sql, info %>
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <div id="defBanner" class="defRow">
          <p class="pageTitle">Welcome!</p>
          <div class="defText">
            This site is aimed towards people who would like to show off their projects involving furniture, or talk to people with the same interest. You can keep track of the projects that you've started and share them!
          </div>
        </div>
        <div id="defUploadProj" class="defRow">
          <div class="defText">
            <p>
              Upload projects for everyone to see! Create handy albums that lets you easily share your projects and progress photos with friends and family. Get feedback from the community on your projects or just create an archive for yourself.
            </p>
          </div>
          <div class="pullRight defExample">
            <p>Some recently created projects include:</p>
            <div class="defExampleEntry">
            <%
              '-- Getting recent projects
              sql = "SELECT projID, projName, userNum, username " &_
                    "FROM ProjTable, UsersTable " &_
                    "WHERE userNum = userID AND category=1 " &_
                    "ORDER BY projID DESC"
              set info = conn.execute(sql)
              
              dim i
              For i=0 to 4
                if NOT info.eof then
                response.write "<a href='viewProjest.asp?projID=" &_
                                info(0) &_
                                "'>" &_
                                info(1) &_
                                "</a> by <a href='viewProfile.asp?userID=" &_
                                info(2) &_
                                "'>" &_
                                info(3) &_
                                "</a><br>"
                info.movenext
                end if
              Next
            %>
            </div>
            <br clear="both">
          </div>
        </div>
        <div id="defShareIdeas" class="defRow">
          <div class="defText">
            <p>
              Share your ideas for DIY projects, or get some ideas! Upload designs and plans to keep track of how you made something, or just to let other's know how you created something.  
            </p>
          </div>
          <div class="pullRight defExample">
            <p>Some recently uploaded designs:</p>
            <%
              '-- Getting recent designs
              sql = "SELECT projID, projName, userNum, username " &_
                    "FROM ProjTable, UsersTable " &_
                    "WHERE userNum = userID AND category=2 " &_
                    "ORDER BY projID DESC"
              set info = conn.execute(sql)
              
              For i=0 to 4
                if NOT info.eof then
                response.write "<a href='viewProjest.asp?projID=" &_
                                info(0) &_
                                "'>" &_
                                info(1) &_
                                "</a> by <a href='viewProfile.asp?userID=" &_
                                info(2) &_
                                "'>" &_
                                info(3) &_
                                "</a><br>"
                info.movenext
                end if
              Next
            %>
          </div>
          <br clear="both">
        </div>
        <div id="defJoinComm" class="defRow">
          <div class="defText">
            <p>
              Join a community that shares your interests! Look in the forums for advice and tips on doing your own DIY project, or chat to friendly members of the community. 
            </p>
          </div>
          <div class="pullRight defExample">
            <p>Some recent forum threads:</p>
            <%
              '-- Getting recent forum threads
              '                 0         1           2         3          4
              sql = "SELECT threadID, threadSubj, ownerNum, username, forumCat " &_
                    "FROM ForumThreadsTable, UsersTable " &_
                    "WHERE ownerNum = userID " &_
                    "ORDER BY threadID DESC"
              set info = conn.execute(sql)
              
              For i=0 to 4
                if NOT info.eof then
                response.write "<a href='forums.asp?subCateg=" &_
                                info(4) &_
                                "&threadID=" &_
                                info(0) &_
                                "'>" &_
                                info(1) &_
                                "</a>" &_
                                " by <a href='viewProfile.asp?userID=" &_
                                info(2) &_
                                "'>" &_
                                info(3) &_
                                "</a><br>"
                info.movenext
                end if
              Next
            %>
          </div>
          <br clear="both">
        </div>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>