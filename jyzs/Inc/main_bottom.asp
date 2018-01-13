<div style="clear:both;"></div>
			
			<div id="main_bottom">
				<div id="main_bottom_left">
					<ul id="menu1">
							<li class="sec2"><strong>产品图示</strong></li>
							<li class="sec1" style="width:400px;"></li>
							<div style="clear:both;"></div>
						</ul>
					
					<div style="clear:both;"></div>
					<div style="height:10px;"></div>
					
					<div id="bottom_left_box">
                    	<%Sub ProductList(InfoType,PageNum)%>
						<%
						sqlstr="select top "&PageNum&" * from [Products] where status=1 order by show_order desc,id desc"
											
						Dim rs_Products : Set rs_Products = db.getRecordBySQL(sqlstr)
					
						if not (rs_Products.eof or rs_Products.bof) then
							do until rs_Products.eof
						%>
                       
                        <div class="bottom_left_pic"><img src="../Pictures/<%=rs_Products("Pic_View")%>" width="125" height="75" /><span><%=rs_Products("Product_Name")%></span></div>
						<%
							rs_Products.movenext
							loop
						else
							
						%>
							没有数据
						<%
						end if											
						%>
                        
						
						<%
						db.C(rs_Products)
						%>
						<%end sub%>
                        

                        <%Call ProductList(0,4)%>
						
					</div>
				</div>
				
				<div id="main_bottom_right">
				
					<div id="bottom_nav_title">
						<ul id="menu">
							<li onMouseOver="secBoard(0)" class="sec2"><strong>公司新闻</strong></li>
							<li onMouseOver="secBoard(1)" class="sec1"><strong>行业新闻</strong></li>
							<li onMouseOver="secBoard(1)" class="sec1"></li>
							<div style="clear:both;"></div>
						</ul>
						<div style="clear:both;"></div>
						
						
						<ul id="bottom_nav_text">
							<li class="block">
                            	<%Sub News(InfoType,TopNum)%>
							<%
                            sqlstr="select top " & TopNum & " * from [Informations] where status=1 and Info_Type = " & InfoType & " order by show_order desc,id desc"
                                                
                            Dim rs_News : Set rs_News = db.getRecordBySQL(sqlstr)
                        
                            if not (rs_News.eof or rs_News.bof) then
                                do until rs_News.eof
                            %>
    
                                <div class="index_left_text">
								<div class="index_left_textdata">[<%response.Write(formatdatetime(rs_News("add_Time"),2))%>]</div>
								<div class="index_left_texttext"><a href="news/newsdetail.asp?id=<%=rs_News("ID")%>"><%=rs_News("Title")%></a></div>
								<div style="clear:both;"></div>
							</div>
                            
                            								
                            <%
                                rs_News.movenext
                                loop
                            else
                                
                            %>
                                没有数据
                            <%
                            end if											
                            %>
                            
                            <%
                            db.C(rs_News)
                            %>
                            <%end sub%>
                            <%Call News(3,3)%>

								
								<div class="index_left_text">
									<div class="index_left_texttext"><!--<a href="#">更多...</a>--></div>
								</div>
								
							</li>
							
							<li class="unblock">
							
								<%Call News(6,3)%>
								
								<div class="index_left_text">
									<div class="index_left_texttext"><!--<a href="#">更多...</a>--></div>
								</div>
								
							</li>
						</ul>
						</div>
				
				</div>
				<div style="clear:both;"></div>
			</div>