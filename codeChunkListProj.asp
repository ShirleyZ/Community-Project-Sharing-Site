  <%
    '------------------------------------'
    '-- CODE CHUNK TO DISPLAY PROJECTS --'
    '------------------------------------'
    
    '-- Retrieving projects from this user
    dim pictures
    '                0        1         2       3       4        5
    sql = "SELECT projID, datePost, ProjTable.about, projName, userNum, username " &_
          "FROM ProjTable, UsersTable " &_
          "WHERE userNum=" & userNum & " AND userNum = userID " &_
          "ORDER BY projID"
          
    set info = conn.execute(sql)
    
    
    '-- Displaying all the projects found for this user
    if info.eof then
      response.write "<p>You don't have any projects. Would you like to <a onclick='show(1)'>add one</a>?</p>"
    else
    '-- Found projects
      do 
        '-- Getting images for this project
        '                0      1
        sql = "SELECT imgID, imgUrl " &_
              "FROM ImgTable " &_
              "WHERE projNum=" & info(0) & " " &_
              "ORDER BY imgID"
        
        set pictures = conn.execute(sql)
        
        '-- Writing out project module
        response.write "<div class='projectThumb'>"
        '-- Checking if pictures are found or not
        if pictures.eof then
          response.write "Images not found for Proj #" & info(0)
        else
          response.write "<img class='projectThumbPic' src='" &_
                         pictures(1) &_
                         "' />"
        end if
                
        response.write "<p><span class='projectThumbTitle' title='" &_
                       info(2) &_
                       "'><a href='viewProject.asp?projID=" &_
                       info(0) &_
                       "'>" &_
                       info(3) &_
                       "</a> (#" &_
                       info(0) &_
                       ")</span><br>" &_
                       "<span>Date: " &_
                       info(1) &_
                       " " &_ 
                       "by: " &_
                       info(5) &_
                       "</span><br>"
        if (info(4) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
          response.write "<span class='projectThumbEdit'>" &_
                         "<a href='edit.asp?projID=" &_
                         info(0) &_
                         "'>Edit</a> <a href='delete.asp?projID=" &_
                         info(0) &_
                         "'>Delete</a>" &_
                         "</span>" 
        end if
        response.write "</div>"
        info.movenext
      loop until info.eof
      response.write "<br clear='both'>"
    end if
  %>