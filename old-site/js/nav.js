<!--//--><![CDATA[//><!--

sfHover = function() {
	var sfEls = document.getElementById("nav").getElementsByTagName("LI");
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

//--><!]]>


<!--
    if (document.images)
    {
		
        image2 = new Image(96, 34);
        image2.src = "images/menu02_roll.gif";
		
        image3 = new Image(736, 34);
        image3.src = "images/menu03_roll.gif";
		
        image4 = new Image(75, 34);
        image4.src = "images/menu04_roll.gif";
		
        image5 = new Image(74, 34);
        image5.src = "images/menu05_roll.gif";
		
        image6 = new Image(81, 34);
        image6.src = "images/menu06_roll.gif";
		
        image6 = new Image(102, 34);
        image6.src = "images/menu07_roll.gif";
		
        image6 = new Image(98, 34);
        image6.src = "images/menu08_roll.gif";
		
        image6 = new Image(104, 34);
        image6.src = "images/menu09_roll.gif";
    }
	
//-->
