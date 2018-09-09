// JavaScript Document

sfHover = function() {
	var sfEls = document.getElementById("menu-container").getElementsByTagName("LI");
	for (var i=0; i<sfEls.length; i++) {
		sfEls[i].onmouseover=function() {
			this.className+=" sfhover";
		}
		sfEls[i].onmouseout=function() {
			this.className=this.className.replace(new RegExp(" sfhover\\b"), "");
		}
	}
}
if (window.attachEvent) window.attachEvent("onload", sfHover);


function openLayer(divID){
	var e;
	e = document.getElementById(divID);
	e.style.visibility = (e.style.visibility == 'visible' ? 'hidden' : 'visible');
	e.style.display = (e.style.display == 'block' ? 'none' : 'block');	
	}

function openLayer2(divID){
	var e;
	e = document.getElementById(divID);
	e.style.visibility = (e.style.visibility == 'hidden' ? 'hidden' : 'hidden');
	e.style.display = (e.style.display == 'none' ? 'none' : 'none');	
	}
	
function lightbox(){
	scroll(0,0); //This scrolls the window to top
	e = document.getElementById('light');
	f = document.getElementById('fade');
	e.style.display = (e.style.display == 'block' ? 'none' : 'block');
	f.style.display = (f.style.display == 'block' ? 'none' : 'block');
	}
	

function ChangeImage (ImageName,FileName,ImageURL,myAnchor) {
	document[ImageName].src = "/images/" + FileName;
	document.getElementById(myAnchor).href="/images/" + ImageURL;
	//document.getElementById('myAnchor').href="http://www.w3schools.com";
	//document.getElementById(ImageName).setAttribute('href',"/images/" + ImageURL);
	}











