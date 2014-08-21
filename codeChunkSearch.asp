<%
        '-- Getting all projects
        '-- Retrieving projects from this user
        dim sql, info, pictures
        dim searchTerm
        searchTerm = request.form("searchTerm")
        
        if searchTerm <> "" then
          '-- If there is a search, find projects like
          if userProj <> "" then
          '-- If there is a user ID, find user specific
          '                0        1         2       3        4   
          sql = "SELECT projID, datePost, about, projName, userNum " &_
                "FROM ProjTable " &_
                "WHERE projName LIKE '%" & searchTerm & "%' " &_
                "AND userNum=" & userProj
          else
          '                0        1         2       3        4   
          sql = "SELECT projID, datePost, about, projName, userNum " &_
                "FROM ProjTable " &_
                "WHERE projName LIKE '%" & searchTerm & "%'"
          end if
        else
          '-- If there's no search, then get all projects
          if userProj <> "" then
          '-- If user ID, find user specific
            '                0        1         2       3        4   
            sql = "SELECT projID, datePost, about, projName, userNum " &_
                  "FROM ProjTable " &_
                  "WHERE userNum=" & userProj
          else
            '                0        1         2       3        4   
            sql = "SELECT projID, datePost, about, projName, userNum " &_
                  "FROM ProjTable " 
          end if
        end if
        
        set info = conn.execute(sql)        
        
        '-- Displaying all the projects found
        if info.eof then
          if searchTerm = "" then
            response.write "<p>There aren't any projects )': Would you like to <a onclick='show(1)'>add one</a>?</p>"
          else
            response.write "<p>We didn't find any projects containing '" & searchTerm & "' in its name ):"
          end if
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
              
            '-- Writing the projects out
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
                           "# comments</span><br>"
            '-- Giving edit/delete options if you're the owner
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
        conn.close
        %>