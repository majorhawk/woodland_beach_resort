		  <div id="menu-container">
				<ul>
					<li style="width:307px; height:47px; display:block;"><img src="/images/spacer_nav.gif" width="307" height="47" border="0" alt=""></li>
					<li><a href="/cabins/" title="Cabins" class="cabins <%=iif( instr(path,"cabins") > 0,"active","")%>"><span class="hide">Cabins</span></a>
						<ul style="border:none;">
							<li style="border:none; padding-bottom:2px; width:100%">&nbsp;&nbsp;<img src="/images/pointer.gif" width="59" height="11" border="0" alt="" /></li>
							<li><a href="/cabins/">Resort Cabins</a></li>
							<li style="<%'=iif( session("WT") = 0,"border:none;","")%>"><a href="/cabins/?ID=map">Resort Map</a></li>
							<!-- <li><a href="/cabins/?ID=private">Private Lake Homes</a></li> -->		
							<% 'If session("WT") = 1 then %>
								<li><a href="/winter/">Winter Time Fun</a></li>
							<% 'End If %>
								<li><a href="/about/?ID=ac">Our Amenities</a></li>
								<li style="border:none;"><a href="http://ao4.availabilityonline.com/availtable.php?un=kevin412">Cabin Availability</a></li>
                            
						</ul>
					</li>
					<li><a href="/rates/" title="Rates" class="rates <%=iif( instr(path,"rates") > 0,"active","")%>"><span class="hide">Rates</span></a>
						<ul style="border:none;">
							<li style="border:none; padding-bottom:2px; width:100%">&nbsp;&nbsp;<img src="/images/pointer.gif" width="59" height="11" border="0" alt="" /></li>
							<li><a href="/rates/">Cabin Rates</a></li>
							<li style="<%=iif( session("SP") = 0,"border:none;","")%>"><a href="/boats/">Boat/Pontoon Rental Rates</a></li>			
							<% If session("SP") = 1 then %>
								<li style="border:none;"><a href="/specials/">WBR Specials</a></li>
							<% End If %>
						</ul>
					</li>
					<li><a href="/about/" title="About Us" class="about <%=iif( instr(path,"about") > 0,"active","")%>"><span class="hide">About Us</span></a>
						<ul style="border:none;">
							<li style="border:none; padding-bottom:2px; width:100%">&nbsp;&nbsp;<img src="/images/pointer.gif" width="59" height="11" border="0" alt="" /></li>
							<li><a href="/about/">Who We Are</a></li>
							<li><a href="/about/?ID=fishing">Fishing Guide Service</a></li>
							<li style="border:none;"><a href="/about/?ID=photo">Photo Album</a></li>
						</ul>
					</li>
					<li><a href="/boats/" title="Boats" class="boats <%=iif( instr(path,"boats") > 0,"active","")%>"><span class="hide">Boats</span></a></li>
					<li><a href="/policies/" title="Policies" class="policies <%=iif( instr(path,"policies") > 0,"active","")%>"><span class="hide">Policies</span></a></li>
					<li><a href="/contact/" title="Contact Us" class="contact <%=iif( instr(path,"contact") > 0,"active","")%>"><span class="hide">Contact Us</span></a>
						<ul style="border:none;">
							<li style="border:none; padding-bottom:2px; width:100%">&nbsp;&nbsp;<img src="/images/pointer.gif" width="59" height="11" border="0" alt="" /></li>
							<li><a href="/contact/">General Info</a></li>
							<li><a href="/thestore/">The Store</a></li>
							<li style="border:none;"><a href="http://upnorthdreams.com/" target="_blank">Real Estate for Sale</a></li>
						</ul>
					</li>
					<li><a href="/events/" title="Events" class="events <%=iif( instr(path,"events") > 0,"active","")%>"><span class="hide">Events</span></a></li>
					<li style="width:51px; height:47px; display:block;"><img src="/images/spacer_nav_2.gif" width="51" height="47" border="0" alt=""></li>
				</ul>
			</div>
			<br><br>