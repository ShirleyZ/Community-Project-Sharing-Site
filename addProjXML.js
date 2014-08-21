// -------------------------------
// -- Code for XML Page Request --
// -------------------------------

// based on code from http://www.w3schools.com/ajax/ajax_source.asp
var xmlHttp;

//---------------------------------------------------------------
//-- Creating XML Object
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
//-- Specifying what to do when ready
function stateChanged() {
  document.getElementById("status").innerHTML += xmlHttp.readyState;
  if (xmlHttp.readyState==4) {
    document.getElementById("info").innerHTML = xmlHttp.responseText;
  }
}
//---------------------------------------------------------------
//-- The function that triggers everything
function addProj() {
  // Creating object and checking if supported
  xmlHttp=GetXmlHttpObject()
  if (xmlHttp==null) {
    alert ("Your browser does not support AJAX!");
    return;
  }

  // Finding out which url you want to access
  var projName = document.getElementsByName("projName");
  var projDesc = document.getElementsByName("projDesc");
  var projCateg = document.getElementsByName("projCateg");
  var projStatus = document.getElementsByName("projStatus");
  
  //- Getting image urls
  var imgUrls = new Array();
  var i = 0;
  var elementName;
  for (i = 1; i <= currProjImg; i++) {
    elementName = "projImg"+i;
    //console.log("Looking in "+elementName);
    imgUrls[i] = document.getElementsByName(elementName); 
    //console.log("Found: "+imgUrls[i][0].value);
  }
  
  
  //- Testing purposes
  console.log("No. Imgs: "+currProjImg);
  //console.log(imgUrls[1][0]);
  //console.log(imgUrls[2][0]);
  
  // Setting the url
  var url = "addProject2.asp?stage=2&projName="+projName[0].value+
              "&projDesc="+projDesc[0].value+
              "&projCateg="+projCateg[0].value+
              "&projStatus="+projStatus[0].value+
              "&projNoImgs="+currProjImg;
              
  // Setting imgs into url
  for (i = 1; i <= currProjImg; i++) {
    if (imgUrls[i][0].value != undefined) {
      url+= "&projImg"+i+"="+imgUrls[i][0].value;
    }
  }
  
  console.log(url);
  console.log(projName[0].value + " " + 
              projDesc[0].value + " " + 
              projCateg[0].value + " " + 
              projStatus[0].value)
  
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

// ----------------------------------------
// -- Adding new field for another image --
// ----------------------------------------

var currProjImg = 1;

function newProjImg() {
  console.log("Before thing: " + currProjImg);
  
  if (currProjImg <= 9) {
    currProjImg ++;
    document.getElementById('addProjImg').innerHTML += currProjImg + ": <input type='text' name='projImg"+currProjImg+"'><br>";
    document.getElementById('projNoImgs').value = currProjImg;
  } else {
    document.getElementById('imgError').innerHTML = "Can't add more than 10 images!";
  }
  
  console.log("After thing: " + currProjImg);
}