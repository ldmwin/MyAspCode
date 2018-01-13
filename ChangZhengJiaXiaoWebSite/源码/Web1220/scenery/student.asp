﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim path : path ="../"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="author" content="#" /> 
	<meta name="robots" content="all" />
	<meta name="keywords" content="长征驾校，邢台，邢台市长征机动车驾驶员培训学校，驾照" />
	<meta name="description" content="邢台市长征驾驶员培训学校,是一所以培训驾驶员为主,办理各种驾驶、行驶证件、年审、换证、申批等业务。本校师资力量雄厚，办学历史悠久，学校规模宏大，管理正规，曾连续几年被评为“优秀驾校”、“办学先进单位”、“信得过单位”。 本校现有从业多年的优秀教练数十名，“耐心教育，细心指导，把每位学员都培养成合格人才。”是本校的办学宗旨。为让学员更为深刻的学习和了解专业知识，最近学校又斥资数万元，首家引进最新模拟教学设备，相信，来本校学习会让你学习到最新、最专业的知识，长征驾校会让你踏遍理想的沃土，世界的角落。" />
	<title>邢台市长征驾驶员培训学校-长征驾校欢迎您!</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="../style/basic.css" rel="stylesheet" type="text/css" />
	<link href="../style/main.css" rel="stylesheet" type="text/css" />
	<link href="../style/about.css" rel="stylesheet" type="text/css" />

</head>

<body>
	<!--顶部开始-->
		<!--#include file="../scontrol/top.asp"-->
	<!--顶部结束-->

<!--容器开始-->
	<div id="container">
		
		<div><img src="images/about_top_pic.jpg" width="1000" height="200" /></div>
		
<!--主体开始-->
		<div id="main_top">
			<div id="main_top_left">
				<div id="main_top_lefttitle">长征风采</div>
				<div id="main_top_lefttitle2">inter</div>
			</div>
			<div id="main_top_right">
				<div id="main_topright_logo"><img src="../images/main_righttop_logo.jpg" /></div>
				<div class="main_righttop_title">
					<div class="main_righttop_titleleft"><strong>学员风采</strong> |</div>
					<div class="main_righttop_titleright"><img src="../images/main_leftnav_textbg.jpg" /></div>
					<div style="clear:both;"></div>
				</div>
			</div>
		</div>
			
		<div id="main">
			
			
			<div id="main_left">
				<div id="main_left_top">&nbsp;</div>
				<div id="main_leftnav_box">
					<div class="main_leftnav_text"><a href="index.asp">教练风采</a></div>
					<div class="main_leftnav_text"><a href="student.asp">学员风采</a></div>
				</div>
				<div style="height:50px;"></div>
				<div class="main_left_pic"><img src="images/scenery_left_pic1.jpg" width="200" height="80" /></div>
				<div class="main_left_pic"><img src="images/scenery_left_pic2.jpg" width="200" height="80" /></div>
				<div class="main_left_pic"><img src="images/scenery_left_pic3.jpg" width="200" height="80" /></div>
				<!--<div><img src="../images/main_leftnav_bgbottom.jpg" /></div>-->
			</div>
			<div id="main_right">
				<div id="main_right_box">
					<!--<div id="about_title">驾校荣誉</div>-->
						
					<!--<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>
					<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>
					<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>
					<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>
					<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>
					<div class="honor_pic"><img src="#" width="150" height="150" /><span>荣誉证书名称</span></div>-->
					<%			
					Dim rs_Picture
					Set rs_Picture = db.getRecordBySQL("select * from [Album_Pictures] where Album_ID=2 and status<>4 order by Show_Order desc,id desc") 
			
					if not (rs_Picture.eof or rs_Picture.bof) then
					do while not rs_Picture.eof%>
					<div class="honor_pic"><a href="<%=Config.ImgUrl()%>Pictures/<%=rs_Picture("Picture")%>" target="_blank"><img src="<%=Config.ImgUrl()%>Pictures/<%=rs_Picture("Picture")%>" width="150" height="150" /></a><span><%=rs_Picture("Title")%></span></div>
					<%
					  rs_Picture.movenext()
						
					Loop
					
					'response.Write("</ul>")
					
					end if
					
					db.C(rs_Picture)
					  
				  %>
					<div style="clear:both;"></div>	
					<!--<div id="page">共X页，当前第y页&nbsp; <a href="#">[ 首页 ]</a> | <a href="#">[ 上一页 ]</a> | <a href="#">[ 下一页 ]</a> | <a href="#">[ 尾页 ]</a></div>-->
				</div>
			</div>
			<div style="clear:both;"></div>
			
		<!--</div>-->
<!--主体结束-->
<!--尾部开始-->
		<!--#include file="../scontrol/footer.html"-->
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>