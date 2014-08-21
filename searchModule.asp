<!-- #include file="dbconn.asp" -->
        <div id="searchBar">
          <form action='search.asp' method='post'>
            Search term: <input type='text' name='searchTerm'><br>
            <input type='submit' value='Search'>
          </form>
          <!-- #include file="codeChunkMsgDisp.asp" -->
        </div>
        <%
          dim userProj
          userProj = Session("sz_uID")
        %>
<!-- #include file="codeChunkSearch.asp" -->
