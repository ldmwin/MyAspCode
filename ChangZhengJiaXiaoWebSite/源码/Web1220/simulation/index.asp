<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
	<link href="../style/scenery.css" rel="stylesheet" type="text/css" />

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
				<div id="main_top_lefttitle">模拟考试</div>
				<div id="main_top_lefttitle2">inter</div>
			</div>
			<div id="main_top_right">
				<div id="main_topright_logo"><img src="../images/main_righttop_logo.jpg" /></div>
				<div class="main_righttop_title">
					<div class="main_righttop_titleleft"><strong>模拟考试</strong> |</div>
					<div class="main_righttop_titleright"><img src="../images/main_leftnav_textbg.jpg" /></div>
					<div style="clear:both;"></div>
				</div>
			</div>
		</div>
			
	  <div id="main">
			
			
			<script>var location="";var navigate="";</script>
					<!--<iframe height=1042 marginheight=0 border=0 src="http://jxks.jxedt.com/exam/exam.asp?type=c" frameborder=0 width=1000 marginwidth=0></iframe>-->
					
					<IFRAME id=play height=1042 marginHeight=0 border=0 src="http://jxks.jxedt.com/exam/exam.asp?type=c" frameBorder=0 width=1000 name=play marginWidth=0 scrolling=yes></IFRAME>
			
		</div>
		<div style="height:50px;"><!--</div>-->
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