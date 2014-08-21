// based on code from http://www.w3schools.com/ajax/ajax_source.asp
var xmlHttp;
//---------------------------------------------------------------
function GetXmlHttpObject() {
  var xmlHttp=null;
  try {
    //          V New object for XML code
    xmlHttp=new XMLHttpRequest();      // Firefox, Opera 8.0+, Safari
  } catch (e) {
    try {      // Internet Explorer
     xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
    } catch (e) {
     xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
    }
  }
  return xmlHttp;
}
//---------------------------------------------------------------
function stateChanged() {
  document.getElementById("status").innerHTML += xmlHttp.readyState;
  if (xmlHttp.readyState==4) {
    document.getElementById("info").innerHTML = xmlHttp.responseText;
  }
}
//---------------------------------------------------------------
function show(num) {
  // Creating object and checking if supported
  xmlHttp=GetXmlHttpObject()
  if (xmlHttp==null) {
    alert ("Your browser does not support AJAX!");
    return;
  }

  // Finding out which url you want to access
  var url = ""
  if (num == 1)
    url += "./addProject.asp"
  else if (num == 2)
    url += "./editProject.asp"
  else
    url += "./searchModule.asp"
  
  console.log(url);
  // Specifying the function to execute when the response is ready
  xmlHttp.onreadystatechange=stateChanged;
  try {
    xmlHttp.open("GET",url,true);
    xmlHttp.send(null);
  } catch(e) {
    document.getElementById("status").innerHTML += "exception error : "+e.name+" : "+
                                   e.message+"<br>";
  }
} 