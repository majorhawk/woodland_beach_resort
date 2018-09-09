
<ul id="nav">
	<li class="leftSide">&nbsp;</li>
	<li><a href="about.asp?PageID=1000" class="btnA">&nbsp;</a>
		<ul>
			<li><a href="about.asp?PageID=1000">Who We Are</a></li>
			<li><a href="about.asp?PageID=1005">Resort Map</a></li>
			<li><a href="about.asp?PageID=1010">Updates and Additions</a></li>
			<li><a href="about.asp?PageID=1020">Memories of Yesterday</a></li>
		</ul>
	</li>
	<li><a href="cabins.asp?PageID=2000" class="btnB">&nbsp;</a>
		<ul>
			<li><a href="cabins.asp?PageID=2000">Resort Cabins</a></li>
			<li><a href="cabins2.asp?PageID=2000">Private Lake Homes</a></li>
		</ul>
	</li>
	<li><a href="boats.asp" class="btnC">&nbsp;</a></li>
	<li><a href="rates.asp" class="btnD">&nbsp;</a>
		<ul>
			<li><a href="rates.asp">Cabin Rates</a></li>
			<li><a href="boats.asp">Boat/Pontoon Rates</a></li>
			<% If session("SP") = 1 then %>
			<li><a href="specials.asp">WBR Specials</a></li>
			<% End If %>
		</ul>
	</li>
	<li><a href="policies.asp" class="btnE">&nbsp;</a></li>
	<li><a href="contact.asp?PageID=6000" class="btnF">&nbsp;</a></li>
	<li><a href="store.asp" class="btnG">&nbsp;</a></li>
	<li><a href="http://www.upnorthdreams.com" class="btnH" target="_blank">&nbsp;</a></li>
	<li class="rightSide">&nbsp;</li>
</ul>
