<div id="header">
			<div style="width:1000px; margin:0 auto;">
				<div id="header_left"><img src="<%=path%>images/logo.jpg" /></div>
				<div id="header_right">
					<div id="header_right_top"><!--<a href="<%'=path%>contact/message.html">在线留言</a> |--> <a href="<%=path%>registration/index.asp">报名指南</a> | <a href="<%=path%>contact/index.asp">联系我们</a></div>
					<div id="header_right_bottom">长征驾校欢迎您</div>
				</div>
				<div style="clear:both;"></div>
			</div>
			<div id="header_nav">

				<%response.Write(TopFirstNav(Config.SiteID,0,str,path))%>

			</div>
		</div>