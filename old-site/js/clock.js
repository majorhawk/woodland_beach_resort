
// Clock Script Generated By Maxx Blade's Clock v2.0d
// http://www.maxxblade.co.uk/clock
function tS(){ x=new Date(); x.setTime(x.getTime()); return x; } 
function lZ(x){ return (x>9)?x:'0'+x; } 
function tH(x){ if(x==0){ x=12; } return (x>12)?x-=12:x; } 
function dT(){ document.getElementById('tP').innerHTML=eval(oT); setTimeout('dT()',1000); } 
function aP(x){ return (x>11)?'pm':'am'; } 
function y4(x){ return (x<500)?x+1900:x; } 
var dN=new Array('Sun','Mon','Tue','Wed','Thu','Fri','Sat'),mN=new Array('January','February','March','April','May','June','July','August','September','October','November','December'),oT="mN[tS().getMonth()]+' '+tS().getDate()+','+' '+y4(tS().getYear())+' '+tH(tS().getHours())+':'+lZ(tS().getMinutes())+':'+lZ(tS().getSeconds())+' '+aP(tS().getHours())";
if(!document.all){ window.onload=dT; }else{ dT(); }
