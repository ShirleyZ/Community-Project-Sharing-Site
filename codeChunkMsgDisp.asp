<%
        '-- This is where the error message goes
        if Session("sz_errorMsg") <> "" then
          response.write "<div id='errorMsgBox' class='msgBox'>" & Session("sz_errorMsg") & "</div>"
          Session("sz_errorMsg") = ""
        elseif Session("sz_infoMsg") <> "" then
          response.write "<div id='infoMsgBox' class='msgBox'>" & Session("sz_infoMsg") & "</div>"
          Session("sz_infoMsg") = ""
        end if
%>