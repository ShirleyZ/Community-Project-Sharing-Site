<% option explicit%>
<!DOCTYPE HTML>
<!-- #include file="dbconn.asp" -->
<!-- #include file="redirect.asp" -->
<html>
  <head>
    <title>Assn4: Add Project</title>
    <link a href="cssTables.css" rel="stylesheet" type="text/css">
    <link a href="style.css" rel="stylesheet" type="text/css">
    <script language="javascript" type="text/javascript" src="addProjXML.js">
    </script>
  </head>
  
  <body>
    <div id="wrapper">
      <!-- #include file="nav.asp" -->      
      <div id="content">
        <div id="addProject">
          <!-- #include file="codeChunkMsgDisp.asp" -->
          <p class="pageTitle">Add a new project or design</p>
          <%
            dim stage, sql, info
            stage = request.form("stage")
          if stage = "" then
            
            '-------------------------------'
            '-- Stage 1: Gathering info ----'
            '-------------------------------'
            %>
            <form action="addProject.asp" method="post">
            <div class="left">
              <input type="hidden" name="stage" value="2">
              <input id="projNoImgs" type="hidden" name="projNoImgs" value="1">
              Project Name: 
              <input type="text" name="projName">
              <br>
              Description:
              <input type="text" name="projDesc">
              <br>
              Category: <select name="projCateg">
              <%
              '                0         1
              sql = "SELECT categID, categName " &_
                  "FROM ProjCategoriesTable " 
              
              set info = conn.execute(sql)
              
              do
                response.write "<option value='" & info(0) & "'>" &_
                       info(1) & "</option>"
                info.movenext
              loop until info.eof
              %>
              </select>
              <br>
              Status: <select name="projStatus">
              <!-- VBscript to drop down -->
              <%
              sql = "SELECT statusID, status " &_
                  "FROM ProjStatusTable " &_
                  "ORDER BY statusID DESC"
              set info = conn.execute(sql)
              
              do
                response.write "<option value='" & info(0) & "'>" &_
                       info(1) & "</option>"
                info.movenext
              loop until info.eof      
              %>
              </select>
              <br>
              <input type="submit" value="Add Project">
            </div>
            <div id="addProjImg" class="right">
              Images to add<br>
              <span class="jsButton" onclick="newProjImg()">Add Image</span><br>
              <span id="imgError"></span><br>
              1: <input type="text" name="projImg1"><br>
            </div>
            </form>
          </div>
          <%
          elseif stage = 2 then
            
            '-------------------------------'
            '-- Stage 2: Adding the thing --'
            '-------------------------------'
            
            '## RETRIEVING INFORMATION ##'
            
            dim projDatePost
            projDatePost = date()
            
            '-- Yoinking info from url string
            dim projName, projDesc, projCateg, projStatus, projNoImg
            
            projName = request.form("projName")
            projDesc = request.form("projDesc")
            projCateg = request.form("projCateg")
            projStatus = request.form("projStatus")
            projNoImg = request.form("projNoImgs")
            
            '-- Yoinking imgs from url string
            dim i 
            i = 1
            dim projImg(10)
            'response.write "ProjNoImg: " & projNoImg & "<br>"
            dim findThis
            
            for i = 1 to projNoImg
            findThis = "projImg" & i
            projImg(i) = request.form(findThis)
            'response.write i & ": Found " & projImg(i) & "<br>"
            next
            
            '-- Check that no fields are empty
            if (projName = "") or (projDesc = "") then
            Session("sz_errorMsg") = "Please fill out all fields"
            response.redirect "addProject.asp"
            end if
            
            '-- Checking that the first image is not blank
            if (projImg(1) = "") then
            Session("sz_errorMsg") = "You must fill out box 1 of images"
            response.redirect "addProject.asp"
            end if
            
            '## ADDING THE PROJECT ##'
            '-- Creating Project entry in DB
            sql = "INSERT INTO ProjTable (userNum, projName, category, datePost, status, about) " &_
              "VALUES (" & Session("sz_uID") & ", '" & projName & "', " &_
                projCateg & ", '" & projDatePost & "', " & projStatus & ", '" & projDesc & "')"
            'response.write "Inserting project: <br>" & sql & "<br>"
            conn.execute(sql)
            
            '-- Getting the project
            '                0         1        2         3        4          5        6       7
            sql = "SELECT projID,  userNum, projName, category, datePost, status, about " &_
              "FROM ProjTable " &_
              "WHERE userNum=" & Session("sz_uID") & " " &_
              "ORDER BY projID DESC"
            
            set info = conn.execute(sql)
            
            '-- Inserting first image into table
            sql = "INSERT INTO ImgTable (userNum, projNum, imgUrl) " &_
              "VALUES (" & Session("sz_uID") & ", " & info(0) & ", '" & projImg(1) & "')"
            'response.write "Inserting first image: <br>" & sql & "<br><br>"
            conn.execute(sql)
            
            '-- Inserting the rest of images into project
            i = 2
            
            for i = 2 to projNoImg
            if projImg(i) <> "" then
              sql = "INSERT INTO ImgTable (userNum, projNum, imgUrl) " &_
                "VALUES (" & Session("sz_uID") & ", " & info(0) & ", '" & projImg(i) & "')"
              'response.write i & ": Inserting " & projImg(i) & "<br>"
              'response.write sql & "<br><br>"
              conn.execute(sql)
            else
              'response.write "No value, skipping... <br><br>"
            end if
            next
            
            '-- Printing out project details
            'response.write "<br>Project Details<br>"
            'response.write "ProjID: " & info(0) & "<br>" &_
            '       "Name: " & info(2) & "<br>" &_
            '       "OwnerNum: " & info(1) & "<br>" &_
            '       "Category: " & info(3) & "<br>" &_
            '       "Status: " & info(5) & "<br>" &_
            '       "DatePost: " & info(4) & "<br>" &_
            '       "About: " & info(6) & "<br>"
                   
            '-- Giving a link to the project
            response.write "Done! Would you like to <a href='viewProject.asp?projID=" &_
                   info(0) & "'>view your project</a>?<br>"
          else 
            Session("sz_errorMsg") = "Something went wrong! (Invalid stage: addProject.asp)"
            response.redirect "error.asp"
          end if
          %>
       </div> <!-- End of #content -->
    <!-- #include file="footer.asp" -->
    </div> <!-- End of #wrapper -->
  </body>
</html>