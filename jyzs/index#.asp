﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="Class/DBCtrl.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"")
db.OpenConn()

Dim path : path =""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="author" content="#" /> 
	<meta name="robots" content="all" />
	<meta name="keywords" content="" />
	<meta name="description" content="" />
	<title>久友木门</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="style/basic.css" rel="stylesheet" type="text/css" />
	<link href="style/main.css" rel="stylesheet" type="text/css" />
	<link href="style/default.css" rel="stylesheet" type="text/css" />
	<script language=javascript>
    function secBoard(n)
    {
    for(i=0;i<menu.childNodes.length;i++)
    menu.childNodes[i].className="sec1";
    menu.childNodes[n].className="sec2";
    for(i=0;i<bottom_nav_text.childNodes.length;i++)
    bottom_nav_text.childNodes[i].style.display="none";
    bottom_nav_text.childNodes[n].style.display="block";
    }
    </script>
</head>

<body>

<!--顶部开始-->
	<!--#include file="inc/header.asp"-->
<!--顶部结束-->

<!--容器开始-->
	<div id="container">
<!--主体开始-->
		<div id="index_main">
		
<!--主体上部flash开始-->
			<!--#include file="inc/main_boxtop.html"-->
<!--主体上部flash结束-->

<!--主体上部页面位置开始-->
			<div id="main_position">当前位置 > 首页</div>
<!--主体上部页面位置结束-->

<!--首页主体开始-->
			<div id="index_box">
			
<!--首页主体左侧开始-->
				<div id="index_left">
				
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
								<div class="index_left_texttext"><a href="news/index.asp?type=3">更多...</a></div>
							</div>
							
						</li>
						
						<li class="unblock">
						
							<%Call News(6,3)%>
							<div class="index_left_text">
								<div class="index_left_texttext"><a href="news/index.asp?type=6">更多...</a></div>
							</div>
							
						</li>
					</ul>
					</div>
				
					<div id="index_left_bottom"><img src="images/05.jpg" width="270" height="90" /></div>
				</div>
<!--首页主体左侧结束-->

<!--首页主体中部开始-->
				<div id="index_center">
					<div id="bottom_nav_title" style="height:25px;">
						<ul id="menu1">
							<li class="sec2"><strong>产品图示</strong></li>
							<li class="sec1" style="width:200px;"></li>
							<div style="clear:both;"></div>
						</ul>
					</div>
					<div style="clear:both;"></div>
					<div style="height:10px;"></div>
					
					<div id="index_center_box">
                    	<%Sub ProductsList(InfoType,PageNum)%>
						<%
						sqlstr="select top "&PageNum&" * from [Products] where status=1 order by show_order desc,id desc"
											
						Dim rs_Products : Set rs_Products = db.getRecordBySQL(sqlstr)
					
						if not (rs_Products.eof or rs_Products.bof) then
							do until rs_Products.eof
						%>
                        <div class="index_center_pic"><a href="product/productdetail.asp?id=<%=rs_Products("ID")%>"><img src="<%=Config.ImgUrl()%>Pictures/<%=rs_Products("Pic_View")%>" width="125" height="70" /></a></div>
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
                        

                        <%Call ProductsList(0,6)%>
						
						<!--<div class="index_center_pic"><a href="#"><img src="images/05.jpg" width="125" height="70" /></a></div>
						<div class="index_center_pic"><a href="#"><img src="images/05.jpg" width="125" height="70" /></a></div>
						<div class="index_center_pic"><a href="#"><img src="images/05.jpg" width="125" height="70" /></a></div>
						<div class="index_center_pic"><a href="#"><img src="images/05.jpg" width="125" height="70" /></a></div>
						<div class="index_center_pic"><a href="#"><img src="images/05.jpg" width="125" height="70" /></a></div>-->
						<div style="clear:both;"></div>
					</div>
					<div id="index_center_more"><a href="product/index.asp">更多...</a></div>
				</div>
<!--首页主体中部结束-->

<!--首页主体右部开始-->
				<div id="index_right">
					<div id="index_right_box">
						<div class="index_right_pic"><img src="images/15.jpg" width="180" height="100" /><span><a href="#">宏耐地板</a></span></div>
						<div class="index_right_pic"><img src="images/06.jpg" width="180" height="100" /><span><a href="#">世友地板</a></span></div>
					</div>
				</div>
<!--首页主体右部结束-->
				
				<div style="clear:both;"></div>
			</div>
<!--首页主体结束-->

		</div>	
<!--主体结束-->

	</div>
<!--容器结束-->			

<!--尾部容器开始-->
	<!--#include file="inc/footer.html"-->
<!--尾部容器结束-->
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>