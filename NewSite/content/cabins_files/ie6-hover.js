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
	//document.getElementById('myAnchor').href="https://web.archive.org/web/20171017221856/http://www.w3schools.com/";
	//document.getElementById(ImageName).setAttribute('href',"/images/" + ImageURL);
	}












/*
     FILE ARCHIVED ON 22:18:56 Oct 17, 2017 AND RETRIEVED FROM THE
     INTERNET ARCHIVE ON 00:25:53 Jul 09, 2018.
     JAVASCRIPT APPENDED BY WAYBACK MACHINE, COPYRIGHT INTERNET ARCHIVE.

     ALL OTHER CONTENT MAY ALSO BE PROTECTED BY COPYRIGHT (17 U.S.C.
     SECTION 108(a)(3)).
*/
/*
playback timings (ms):
  LoadShardBlock: 74.565 (3)
  esindex: 0.007
  captures_list: 92.321
  CDXLines.iter: 12.442 (3)
  PetaboxLoader3.datanode: 136.007 (5)
  exclusion.robots: 0.199
  exclusion.robots.policy: 0.179
  RedisCDXSource: 1.617
  PetaboxLoader3.resolve: 878.066 (2)
  load_resource: 981.024
*/