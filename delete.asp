<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Delete</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
      <%
        '## THERE ARE # DIFFERENT TYPES OF PAGES COMING HERE ##'
        '## - Projects - delete.asp?projID=#
        '## - Images - delete.asp?imgID=#
        '## - Comments - delete.asp?imgCmtID=#
        
        '## TELL THESE TO GET THEIR OWN PAGE ##'
        '## - Profiles
        '## - Forum Posts
        '## - Forum threads
        
        dim projID, imgID, imgCmtID
        projID = request.querystring("projID")
        imgID = request.querystring("imgID")
        imgCmtID = request.querystring("imgCmtID")
        
        dim sql, info, stage
        stage = request.form("stage")
				if stage = "" then 
					'--------------------------------------'
					'-- STAGE 1: Asking for confirmation --'
					'--------------------------------------'
					if imgCmtID <> "" then
						response.write "<p>Deleting Comment</p>"
											
						'-- Getting the comment to delete
						'                 0        1        2                  3               4      5
						sql = "SELECT imgCmtID, username, imgNum, imgCommentsTable.comment, userID, projNum " &_
									"FROM imgCommentsTable, UsersTable, ImgTable " &_
									"WHERE imgCmtID=" & imgCmtID & " AND imgCommentsTable.userNum=userID AND imgNum = imgID"
						set info = conn.execute(sql)
						
						response.write "<p>Are you <b>certain</b> you want to delete this comment?</p>"
						
						'-- Checking I found the thing
						if info.eof then
							response.write "There's nothing here!"
						else 
						'-- Thing found
							'-- checking you're the owner or admin
							if (info(4) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
								response.write "#" &_
																info(0) &_
																" by " &_
																info(1) &_
																" on image #" &_
																info(2) &_
																"<br>Comment: <br>" &_
																info(3) & "<br><br>"
								response.write "<form action='delete.asp' method='post'>" &_
                								"<input type='hidden' name='stage' value='2'>" &_
																"<input type='hidden' name='imgCmtID' value='" &_
																imgCmtID &_
																"'><input type='hidden' name='entryType' value='1'>" &_
																"<input type='submit' value=""Yes I'm sure!""></form>"
								response.write "&nbsp;&nbsp;&nbsp;"
								response.write "<a class='jsButton' href='viewProject.asp?projID=" & info(5) & "'>No, take me back!</a>"
							else
								response.write "This isn't your comment! Go away!"
							end if
						end if
					elseif imgID <> "" then
						response.write "<p>Deleting Image</p>"
											
						'-- Getting the image to delete
						'              0        1        2        3       4        5
						sql = "SELECT imgID, username, userID, comment, imgUrl, projNum " &_
									"FROM UsersTable, ImgTable " &_
									"WHERE imgID=" & imgID & " AND userNum=userID "
						set info = conn.execute(sql)
						
						response.write "<p>Are you <b>certain</b> you want to delete this image?</p>"
						
						'-- Checking I found the thing
						if info.eof then
							response.write "There's nothing here!"
						else 
						'-- Thing found
							'-- checking you're the owner or admin
							if (info(2) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
								response.write "Image #" &_
																info(0) &_
																" by " &_
																info(1) &_
																"<br><img id='viewProjPicFrame' src='" &_
																info(4) &_
																"'><br>Comment: <br>" &_
																info(3) & "<br><br>"
								response.write "<form action='delete.asp' method='post'>" &_
                								"<input type='hidden' name='stage' value='2'>" &_
																"<input type='hidden' name='imgID' value='" &_
																imgID &_
																"'><input type='hidden' name='entryType' value='2'>" &_
																"<input type='submit' value=""Yes I'm sure!""></form>"
								response.write "&nbsp;&nbsp;&nbsp;"
								response.write "<a class='jsButton' href='viewProject.asp?projID=" & info(5) & "'>No, take me back!</a>"
							else
								response.write "This isn't your image! Go away!"
							end if
						end if
					elseif projID <> "" then
						response.write "<p>Deleting Project</p>"
											
						'-- Getting the project to delete
						'                0        1       2         3         4                5                 6       7
						sql = "SELECT projID, username, userID, projName, categName, ProjStatusTable.status, datePost, ProjTable.about " &_
									"FROM ProjTable, UsersTable, ProjCategoriesTable, ProjStatusTable " &_
									"WHERE projID=" & projID & " AND userNum=userID AND category=categID AND ProjTable.status=statusID "
						set info = conn.execute(sql)
						
						response.write "<p>Are you <b>certain</b> you want to delete this project?</p>"
						
						'-- Checking I found the thing
						if info.eof then
							response.write "There's nothing here!"
						else 
						'-- Thing found
							'-- checking you're the owner or admin
							if (info(2) = Session("sz_uID")) or (Session("sz_uAccess") = 1) then
								response.write "Project #" &_
																info(0) &_
																" by " &_
																info(1) &_
																"<br>Category: " &_
																info(4) &_
																"<br>Status: " &_
																info(5) &_
																"<br>Date posed: " &_
																info(6) &_
																"<br>Description:<br>" &_
																info(7) &_
																"<br><br>"
								response.write "<form action='delete.asp' method='post'>" &_
                								"<input type='hidden' name='stage' value='2'>" &_
																"<input type='hidden' name='projID' value='" &_
																projID &_
																"'><input type='hidden' name='entryType' value='3'>" &_
																"<input type='submit' value=""Yes I'm sure!""></form>"
								response.write "&nbsp;&nbsp;&nbsp;"
								response.write "<a class='jsButton' href='viewProject.asp?projID=" & info(0) & "'>No, take me back!</a>"
							else
								response.write "This isn't your project! Go away!"
							end if
						end if
					else
						Session("sz_errorMsg") = "Looks like something went wrong! (Nothing to delete: delete.asp stage 1)"
						response.redirect "error.asp"
					end if
        elseif stage = 2 then
						
					'----------------------------------------'
					'-- STAGE 2: Deleting the actual stuff --'
					'----------------------------------------'
					
					dim entryType
					
					entryType = request.form("entryType")
					
					if entryType = 1 then '-- Deleting Comments
						imgCmtID = request.form("imgCmtID")
						
						'-- Saving where you came from
						dim fromProj, fromImg
						sql = "SELECT projNum, imgNum " &_
									"FROM ImgCommentsTable, ImgTable " &_
									"WHERE imgCmtID=" & imgCmtID & " AND imgNum=imgID"
						set info = conn.execute(sql)
						fromProj = info(0)
						fromImg = info(1)
						
						'-- Deleting the thing
						sql= "DELETE FROM ImgCommentsTable " &_
								 "WHERE imgCmtID=" & imgCmtID
						conn.execute(sql)
						response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
													 fromProj &_
													 "&imgID=" &_
													 fromImg &_
													 "'>go back</a>?"
						
					elseif entryType = 2 then '-- Deleting Images
						imgID = request.form("imgID")
						
						'-- Saving where you came from
						sql = "SELECT projNum " &_
									"FROM ImgTable " &_
									"WHERE imgID=" & imgID
						set info = conn.execute(sql)
						fromProj = info(0)
						
						'-- Deleting the thing
						sql = "DELETE FROM ImgTable " &_
									"WHERE imgID=" & imgID
						conn.execute(sql)
						
						'-- Deleting all comments associated with this picture
						sql = "DELETE FROM ImgCommentsTable " &_
									"WHERE imgNum=" & imgID
						conn.execute(sql)
						
						'-- Done!
						response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
													 fromProj &_
													 "'>go back</a>?"
					elseif entryType = 3 then '-- Deleting Projects
						projID = request.form("projID")
						
						'-- Don't need to save where you from, it's gone!
						
						'-- Deleting comments of images in this thing
						dim projImg
						sql = "SELECT imgID " &_
									"FROM ImgTable " &_
									"WHERE projNum=" & projID
						set projImg = conn.execute(sql)
						
						do until projImg.eof
							sql = "DELETE FROM ImgCommentsTable " &_
										"WHERE imgNum=" & projImg(0)
							conn.execute(sql)
							projImg.movenext
						loop        
						
						'-- Deleting images in this thing
						sql = "DELETE FROM ImgTable " &_
									"WHERE projNum=" & projID
						conn.execute(sql)
						
						'-- Deleting this thing
						sql = "DELETE FROM ProjTable " &_
									"WHERE projID=" & projID
						conn.execute(sql)
						'-- Done!
						response.write "Done! Would you like to look at <a href='search.asp'>some other projects</a>?"
					else
						Session("sz_errorMsg") = "Looks like something went wrong! (Invalid entryType: delete.asp stage 2)"
						response.redirect "error.asp"
					end if
				else
					Session("sz_errorMsg") = "Looks like something went wrong! (Invalid stage: delete.asp)"
					response.redirect "error.asp"
				end if
      %>
      </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>