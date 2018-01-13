<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
	<meta name="keywords" content="长征驾校，邢台，邢台市长征机动车驾驶员培训学校，驾照" />
	<meta name="description" content="邢台市长征驾驶员培训学校,是一所以培训驾驶员为主,办理各种驾驶、行驶证件、年审、换证、申批等业务。本校师资力量雄厚，办学历史悠久，学校规模宏大，管理正规，曾连续几年被评为“优秀驾校”、“办学先进单位”、“信得过单位”。 本校现有从业多年的优秀教练数十名，“耐心教育，细心指导，把每位学员都培养成合格人才。”是本校的办学宗旨。为让学员更为深刻的学习和了解专业知识，最近学校又斥资数万元，首家引进最新模拟教学设备，相信，来本校学习会让你学习到最新、最专业的知识，长征驾校会让你踏遍理想的沃土，世界的角落。" />
	<title>邢台市长征驾驶员培训学校-长征驾校欢迎您!</title>
	<link rel="shortcut icon" href="images/favicon.ico" type="image/x-icon" />
	<link href="style/basic.css" rel="stylesheet" type="text/css" />
	<link href="style/main.css" rel="stylesheet" type="text/css" />
	<link href="style/default.css" rel="stylesheet" type="text/css" />

<style type="text/css"> 
.container, .container *{margin:0px;}

.container{width:1000px; height:370px; overflow:hidden; position:relative; margin:0px; padding:0px;}

.slider{
	position:absolute;
	left: 0px;
	top: 0px;
}
.slider li{ list-style:none;display:inline; margin:0; padding:0;}
.slider img{ width:1000px; height:370px;}

.slider2{width:10000px;}
.slider2 li{float:left;}

.num{ position:absolute; right:5px; bottom:5px;}
.num li{
	float: left;
	color: #FF7300;
	text-align: center;
	line-height: 16px;
	width: 16px;
	height: 16px;
	font-family: Arial;
	font-size: 12px;
	cursor: pointer;
	overflow: hidden;
	margin: 3px 1px;
	border: 1px solid #FF7300;
	background-color: #fff;
}
.num li.on{
	color: #fff;
	line-height: 21px;
	width: 21px;
	height: 21px;
	font-size: 16px;
	margin: 0 1px;
	border: 0;
	background-color: #FF7300;
	font-weight: bold;
}
</style>
 <script type="text/javascript">
var $ = function (id) {
	return "string" == typeof id ? document.getElementById(id) : id;
};

var Class = {
  create: function() {
	return function() {
	  this.initialize.apply(this, arguments);
	}
  }
}

Object.extend = function(destination, source) {
	for (var property in source) {
		destination[property] = source[property];
	}
	return destination;
}

var TransformView = Class.create();
TransformView.prototype = {
  //容器对象,滑动对象,切换参数,切换数量
  initialize: function(container, slider, parameter, count, options) {
	if(parameter <= 0 || count <= 0) return;
	var oContainer = $(container), oSlider = $(slider), oThis = this;

	this.Index = 0;//当前索引
	
	this._timer = null;//定时器
	this._slider = oSlider;//滑动对象
	this._parameter = parameter;//切换参数
	this._count = count || 0;//切换数量
	this._target = 0;//目标参数
	
	this.SetOptions(options);
	
	this.Up = !!this.options.Up;
	this.Step = Math.abs(this.options.Step);
	this.Time = Math.abs(this.options.Time);
	this.Auto = !!this.options.Auto;
	this.Pause = Math.abs(this.options.Pause);
	this.onStart = this.options.onStart;
	this.onFinish = this.options.onFinish;
	
	oContainer.style.overflow = "hidden";
	oContainer.style.position = "relative";
	
	oSlider.style.position = "absolute";
	oSlider.style.top = oSlider.style.left = 0;
  },
  //设置默认属性
  SetOptions: function(options) {
	this.options = {//默认值
		Up:			true,//是否向上(否则向左)
		Step:		5,//滑动变化率
		Time:		10,//滑动延时
		Auto:		true,//是否自动转换
		Pause:		2000,//停顿时间(Auto为true时有效)
		onStart:	function(){},//开始转换时执行
		onFinish:	function(){}//完成转换时执行
	};
	Object.extend(this.options, options || {});
  },
  //开始切换设置
  Start: function() {
	if(this.Index < 0){
		this.Index = this._count - 1;
	} else if (this.Index >= this._count){ this.Index = 0; }
	
	this._target = -1 * this._parameter * this.Index;
	this.onStart();
	this.Move();
  },
  //移动
  Move: function() {
	clearTimeout(this._timer);
	var oThis = this, style = this.Up ? "top" : "left", iNow = parseInt(this._slider.style[style]) || 0, iStep = this.GetStep(this._target, iNow);
	
	if (iStep != 0) {
		this._slider.style[style] = (iNow + iStep) + "px";
		this._timer = setTimeout(function(){ oThis.Move(); }, this.Time);
	} else {
		this._slider.style[style] = this._target + "px";
		this.onFinish();
		if (this.Auto) { this._timer = setTimeout(function(){ oThis.Index++; oThis.Start(); }, this.Pause); }
	}
  },
  //获取步长
  GetStep: function(iTarget, iNow) {
	var iStep = (iTarget - iNow) / this.Step;
	if (iStep == 0) return 0;
	if (Math.abs(iStep) < 1) return (iStep > 0 ? 1 : -1);
	return iStep;
  },
  //停止
  Stop: function(iTarget, iNow) {
	clearTimeout(this._timer);
	this._slider.style[this.Up ? "top" : "left"] = this._target + "px";
  }
};

window.onload=function(){
	function Each(list, fun){
		for (var i = 0, len = list.length; i < len; i++) { fun(list[i], i); }
	};
	
	//var objs = $("idNum").getElementsByTagName("li");
//	
//	var tv = new TransformView("idTransformView", "idSlider", 370, 3, {
//		onStart : function(){ Each(objs, function(o, i){ o.className = tv.Index == i ? "on" : ""; }) }//按钮样式
//	});
//	
//	tv.Start();
//	
//	Each(objs, function(o, i){
//		o.onmouseover = function(){
//			o.className = "on";
//			tv.Auto = false;
//			tv.Index = i;
//			tv.Start();
//		}
//		o.onmouseout = function(){
//			o.className = "";
//			tv.Auto = true;
//			tv.Start();
//		}
//	})
	
	////////////////////////test2
	
	var objs2 = $("idNum2").getElementsByTagName("li");
	
	var tv2 = new TransformView("idTransformView2", "idSlider2", 1000, 4, {
		onStart: function(){ Each(objs2, function(o, i){ o.className = tv2.Index == i ? "on" : ""; }) },//按钮样式
		Up: false
	});
	
	tv2.Start();
	
	Each(objs2, function(o, i){
		o.onmouseover = function(){
			o.className = "on";
			tv2.Auto = false;
			tv2.Index = i;
			tv2.Start();
		}
		o.onmouseout = function(){
			o.className = "";
			tv2.Auto = true;
			tv2.Start();
		}
	})
	
	$("idStop").onclick = function(){ tv2.Auto = false; tv2.Stop(); }
	$("idStart").onclick = function(){ tv2.Auto = true; tv2.Start(); }
	$("idNext").onclick = function(){ tv2.Index++; tv2.Start(); }
	$("idPre").onclick = function(){ tv2.Index--;tv2.Start(); }
	$("idFast").onclick = function(){ if(--tv2.Step <= 0){tv2.Step = 1;} }
	$("idSlow").onclick = function(){ if(++tv2.Step >= 10){tv2.Step = 10;} }
	$("idReduce").onclick = function(){ tv2.Pause-=1000; if(tv2.Pause <= 0){tv2.Pause = 0;} }
	$("idAdd").onclick = function(){ tv2.Pause+=1000; if(tv2.Pause >= 5000){tv2.Pause = 5000;} }
	
	$("idReset").onclick = function(){
		tv2.Step = Math.abs(tv2.options.Step);
		tv2.Time = Math.abs(tv2.options.Time);
		tv2.Auto = !!tv2.options.Auto;
		tv2.Pause = Math.abs(tv2.options.Pause);
	}
	
}
</script>
</head>

<body>
	<!--顶部开始-->
		<!--#include file="scontrol/top.asp"-->
	<!--顶部结束-->

<!--容器开始-->
	<div id="container">
		
		<div>
		<div class="container" id="idTransformView2">
  <ul class="slider slider2" id="idSlider2">
    <li><img src="images/DSC_0116.JPG"/></li>
    <li><img src="images/DSC_0287.JPG"/></li>
    <li><img src="images/DSC_0385.JPG"/></li>
    <li><img src="images/DSC_0402.JPG"/></li>
  </ul>
  <ul class="num" id="idNum2">
    <li>1</li>
    <li>2</li>
    <li>3</li>
    <li>4</li>
  </ul>
</div>
		</div>
		
<!--主体开始-->
		<div id="main1" style="background:#ededed;">
			<div><!--中间左右开始-->
			<div id="index_main_left">
				<div id="index_mainleft_left">
					<div class="title">
						<div class="title_left">驾校简介 |</div>
						<div class="title_right">more >></div>
					</div>
					<div id="index_intro">
						<div id="index_intro_left"><img src="images/index_intro.jpg" width="170" height="120" /></div>
						<div id="index_intro_right">
							<div id="index_intro_righttitle">邢台市长征驾校简介</div>
							<div id="index_intro_righttext">长征驾校成立于1986年9月，是全市创办最早的驾校之一。二十五年来，该校坚持“学员至上”的办学宗旨，坚持“诚信办学、正规....[<a href="about/about.asp?id=5">详细</a>]</div>
						</div>
						<div style="clear:both;"></div>
					</div>
				</div>
				<div id="index_mainleft_right">
					<div class="title" style="width:270px;">
						<div class="title_left">校内新闻 |</div>
						<div class="title_right">more >></div>
						<div style="clear:both;"></div>
					</div>
					<div id="index_news">
						<%Sub News()%>
						<%
						sqlstr="select top 5 * from [Informations] where status=1 and Info_Type = 3 order by show_order desc,id desc"
											
						Dim rs_News : Set rs_News = db.getRecordBySQL(sqlstr)
					
						if not (rs_News.eof or rs_News.bof) then
							do until rs_News.eof
						%>

						<div class="index_news_text"><a href="news/newsdetail.asp?id=<%=rs_News("ID")%>" target="_blank"><%=rs_News("Title")%></a></div>
						<!--综合新闻列表循环结束-->
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
						<%Call News()%>
						
					</div>
				</div>
				<div style="clear:both;"></div>
				<div id="index_mainleft_bottom">
					<div class="index_mainleft_bottombox" style="margin-left:21px;"><a href="canon/index.asp"><img src="images/index_use.jpg" width="218" height="98" /><span>用车须知</span></a></div>
					<div class="index_mainleft_bottombox" style="margin-left:21px;"><a href="registration/index.asp?id=11"><img src="images/index_baoming.jpg" width="218" height="98" /><span>报名指南</span></a></div>
					<div class="index_mainleft_bottombox" style="margin-left:21px;"><a href="contact/index.asp?id=16"><img src="images/index_contact.jpg" width="218" height="98" /><span>联系方式</span></a></div>
					<div style="clear:both;"></div>
				</div>
			</div>
			<div id="index_main_right">
				<div id="index_main_righttitle"><img src="images/right_top_pic.jpg" width="250" height="55" /></div>
				<div id="index_main_rightpic">
					<div style="height:200px;"></div>
					<div style="width:215px; line-height:23px; margin-left:12px;">邢台市长征驾校邢台市长征驾校邢台市长征驾校邢台市长征驾校邢台市长征驾校邢台市长征驾校...[<a href="simulation/index.asp" target="_blank">详细</a>]</div>
				</div>
			</div>
			
			</div><!--中间左右结束-->

			<div style="clear:both;"></div>	
			
			<div id="index_bottom">
				<div class="title" style="width:960px;">
					<div class="title_left">教学风采 |</div>
					<div class="title_right">more >></div>
					<div style="clear:both;"></div>
				</div>
				
				<div id="index_bottom_box">
					<%			
					Dim rs_Picture
					Set rs_Picture = db.getRecordBySQL("select top 8 * from [Album_Pictures] where Album_ID=1 and status<>4 order by Show_Order desc,id desc") 
			
						if not (rs_Picture.eof or rs_Picture.bof) then
						do while not rs_Picture.eof%>
						<div class="index_bottom_boxpic"><a href="<%=Config.ImgUrl()%>Pictures/<%=rs_Picture("Picture")%>" target="_blank"><img src="<%=Config.ImgUrl()%>Pictures/<%=rs_Picture("Picture")%>" width="100" height="100" /></a><span><%=rs_Picture("Title")%></span></div>
						<%
					  rs_Picture.movenext()
						
					Loop
					
					'response.Write("</ul>")
					
					end if
					
					db.C(rs_Picture)
				  
					%>
					<!--<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>
					<div class="index_bottom_boxpic"><img src="images/index_flash.jpg" width="100" height="100" /><span>风采</span></div>-->
					<div style="clear:both;"></div>
				</div>
				
			</div>
			
			
			<div id="column_box">
				<div class="column_box_pic"><a href="simulation/index.asp" target="_blank"><img src="images/index_kaoshi.jpg" width="218" height="98" /><span>模拟考试</span></a></div>
				<div class="column_box_pic"><a href="scenery/index.asp"><img src="images/index_flash.jpg" width="218" height="98" /><span>长征风采</span></a></div>
				<div class="column_box_pic"><a href="registration/index.asp?id=11"><img src="images/index_baoming.jpg" width="218" height="98" /><span>报名指南</span></a></div>
				<div class="column_box_pic"><a href="contact/index.asp?id=16"><img src="images/index_contact.jpg" width="218" height="98" /><span>联系方式</span></a></div>
				<div style="clear:both;"></div>
			</div>
		
		</div>
<!--主体结束-->
<!--尾部开始-->
		
		
		
		
	</div>
<!--尾部结束-->
		<!--<div style="clear:both;"></div>-->
<!--主体结束-->
	</div>
<!--容器结束-->
		<div id="footer">
			<!--<div class="main_line"></div>-->
			<div id="footer_nav"><div style="width:1000px; margin:0 auto;">友情网站：&nbsp; <a href="http://www.xtjdcjsr.com/Index.html" target="_blank">驾驶人信息网</a> | <a href="http://mnks.jxedt.com/" target="_blank">驾校一点通</a> | <a href="http://www.xingtai.net" target="_blank">邢台信息港</a></div></div>
			<div id="footer_copyright">联系电话： 0319 - 2298758/2298668   联系地址：邢台市桥西区兴达路     Copyright &nbsp;&nbsp;</div>
		</div>
		
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>