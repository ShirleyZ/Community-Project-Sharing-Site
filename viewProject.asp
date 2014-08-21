<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<html>
  <head>
    <title>Assn4: View Project</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        dim projID
        projID = request.querystring("projID")
        
        '-- If you didn't want to view any specific project
        if projID = "" then
          response.write "Apples"
        else
        '-- If you wanted to view a specific project
          dim sql, info, pictures
          '                0        1          2            3        4        5                           6          7
          sql = "SELECT username, projName, categName, datePost, dateFin, ProjStatusTable.status, ProjTable.about, userID " &_
                "FROM ProjTable, UsersTable, ProjCategoriesTable, ProjStatusTable " &_
                "WHERE projID=" & projID & " AND userNum = userID AND category = categID" &_
                  " AND ProjTable.status = statusID"
          
          set info = conn.execute(sql)
          '                0      1       2        3
          sql = "SELECT imgID, imgUrl, comment, userNum " &_
                "FROM ImgTable " &_
                "WHERE projNum=" & projID & " " &_
                "ORDER BY imgID"
        
          set pictures = conn.execute(sql)
      %>
      <!-- Row 1 of stuff -->
        <div id="viewProjPic" class='viewProjLeft viewProjRow1'>
          <%
            '-- Gets which image to display from the url
            dim test1, test2
            dim currPic0, currPicStr
            dim currPic
            currPic = request.querystring("imgID")
          
            '-- if there's nothing there, then show the first image
            if currPic = "" then
              '-- Check to see if there are images
              if pictures.eof then
                response.write "Sorry there's nothing here"
              else
                response.write "<div class='center'><img id='viewProjPicFrame' src='" &_
                             pictures(1) &_
                             "'></div>"
              end if
            else
              '-- else, show the image
              
              '-- check if there are images
              if pictures.eof then
                response.write "There aren't any pictures"
              else
                '-- for loop to find the image we want
                currPic0 = pictures(0)
                currPicStr = cStr(currPic0)
                do until (currPicStr = currPic) or pictures.eof
                  pictures.movenext
                  currPic0 = pictures(0)
                  currPicStr = cStr(currPic0)
                loop
              end if
              
              '-- Write it if we find it, handle error if we don't
              if pictures.eof then
                response.write "Sorry! We didn't find this image!"
              else
                response.write "<div class='center'><img id='viewProjPicFrame' src='" &_
                                pictures(1) &_
                                "'></div>"
              end if
            end if
            pictures.movefirst
          %>
        </div>
        <div id="viewProjThumbTray" class='viewProjRight viewProjRow1'>
        <%
          '-- Creating thumbnails of all the pictures
          do
            response.write "<a href='viewProject.asp?projID=" &_
                            projID &_
                            "&imgID=" &_
                            pictures(0) &_
                            "'>" &_
                           "<img class='viewProjThumbnails center' src='" &_
                            pictures(1) &_
                            "'/></a>"
            pictures.movenext
          loop until pictures.eof
          response.write "<br clear='left'>"
          
          '-- If this page want to view the first pic then move it to front
            pictures.movefirst
          if currPic <> "" then
            '-- If this page wants to view a specific pic, then move to that one
            currPic0 = pictures(0)
            currPicStr = cStr(currPic0)
            do until (currPicStr = currPic) or pictures.eof
              pictures.movenext
              currPic0 = pictures(0)
              currPicStr = cStr(currPic0)
            loop
          end if
        %>
      
          <img class="viewProjThumb" src='' 
              onclick='document.getElementById("viewProjPicFrame").src="";
                   '/>
        </div>
        <br clear='left'>
        <!-- Row 2 of stuff -->
        <div id="viewProjImgCmt" class='viewProjLeft viewProjRow2'>
          <p>
          <%
            '-- Meta info
            response.write "Date Posted: " &_
                           info(3)
                           
            if info(4) <> "" then
              response.write " Date Finished: " &_
                              info(4)
            end if
            '-- If you're the owner or admin, options to edit/delete image
            if (pictures(3) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
              response.write "&nbsp;&nbsp;&nbsp;<a href='edit.asp?imgID=" &_
                            pictures(0) &_
                            "' class='jsButton'>Edit Image</a>&nbsp;" &_
                            "<a href='delete.asp?imgID=" &_
                            pictures(0) &_
                            "' class='jsButton'>Delete Image</a>"
            end if
          %>
          </p>
          <p>
          <%
            '-- Printing out the comment
            if pictures(2) <> "" then
              response.write pictures(2)
            end if
          %>
          </p>
        </div>
        <div id="viewProjInfo" class='viewProjRight viewProjRow2'>
          <%
          '-- Printing out project information
          response.write "<p><span class='viewProjTitle'>" &_
                         info(1) &_
                         "</span><br>by <a href='viewProfile.asp?userID=" &_
                         info(7) &_
                         "'>" &_
                         info(0) &_
                         "</a><br>Category: " &_
                         info(2) &_
                         " Status: " &_
                         info(5) &_
                         "<br><br>" &_
                         info(6) 
          response.write "</p>"
          '-- Checking if the owner is viewing this project and provides options
          if (Session("sz_uID") = info(7)) or (Session("sz_uAccess") = 1) then
            response.write "<div id='ownerOptions'>" &_
                            "Owner Options <br><a href='edit.asp?projID=" &_
                            projID &_
                            "' class='jsButton'>Edit Project</a><br>" &_
                            "<a href='delete.asp?projID=" &_
                            projID &_
                            "' class='jsButton'>Delete Project</a></div>"
          end if
          %>
        </div>
        <br clear='left'>
        <!-- Row 3 of stuff -->
        <div id="viewProjCmts" class='viewProjRow3'>
          <!-- Comment box -->
          <form class="cmtForm" action="makeComment.asp" method="post">
            <%
              response.write "<input type='hidden' name='projID' value=" &_
                              projID &_
                              ">"
              response.write "<input type='hidden' name='imgID' value=" &_
                              pictures(0) &_
                              ">"
            %>
            <input type="hidden" name="entryType" value="1">
            <input type="text" name="cmtBody" class="cmtBody">
            <br>
            <input type="submit" value="Submit">
          </form>
          <%
            '-- Grabbing all the comments for this image
            dim imgCmts
            '                 0        1        2        3        4        5
            sql = "SELECT imgCmtID, imgNum, username, userNum, comment, datePost " &_
                  "FROM ImgCommentsTable, UsersTable " &_
                  "WHERE imgNum=" & pictures(0) & " AND userNum = userID " &_
                  "ORDER BY imgCmtId"
            
            set imgCmts = conn.execute(sql)
            
            '-- Printing out all the comments
            if imgCmts.eof then 
              response.write "<div class='dispCmt'>There are no comments here!</div>"
            else
              do
                response.write "<div class='dispCmt'><p class='viewProjImgCmtBody'>" &_
                               imgCmts(4) &_
                               "</p><p class='dispCmtMeta'>Comment #" &_
                               imgCmts(0) &_
                               " by: " &_
                               imgCmts(2) &_
                               " &nbsp; &nbsp; Date Posted: " &_
                               imgCmts(5)
                '-- Checking if admin or owner
                if (imgCmts(3) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
                  response.write " &nbsp;&nbsp;" &_
                                 "<a href='edit.asp?imgCmtID=" &_
                                 imgCmts(0) &_
                                 "'>Edit</a>"
                  response.write " &nbsp;&nbsp;" &_
                                 "<a href='delete.asp?imgCmtID=" &_
                                 imgCmts(0) &_
                                 "'>Delete</a>"
                end if
                response.write "</div>"
                imgCmts.movenext
              loop until imgCmts.eof
            end if
          %>
        </div>
      </div> <!-- End of #content -->
      <%
        end if
        conn.close
      %>
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>