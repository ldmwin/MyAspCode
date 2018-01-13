<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>

<!--#include file="Class/DBCtrl.asp" -->
<!--#include file="Class/UserCtrl.asp" -->
<!--#include file="Class/LogCtrl.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"")
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	'response.Write("用户校验成功")
	'response.End()
	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>网站管理系统</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/mainframe.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" language="javascript">
var barstatus = 1;
</script>
<script type="text/javascript" src="js/function.js" language="javascript"></script>
</head>

<body>
<table border=0 cellpadding=0 cellspacing=0 height="100%" width="100%" style="background:#C3DAF9;">
	<tr>
		<td height="58" colspan="3">
			<iframe frameborder="0" id="top" name="top" scrolling="no" src="top.asp" style="height: 100%; visibility: inherit;width: 100%;"></iframe>
		</td>
	</tr>
	<tr>
		<td height="30" colspan="3">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr height="32">
					<td background="images/bg2.gif"width="28" style="padding-left:30px;"><img src="images/arrow.gif" alt="" align="absmiddle" /></td>
					<td background="images/bg2.gif"><span style="color:#c00;font-weight:bold;float:left;margin-top:2px;">公告：</span><span style="color:#135294;font-weight:bold;float:left;width:300px;" id="dcannounce"></span></td>
					<td background="images/bg2.gif" style="text-align:right;color:#135294;padding-right:20px;">
					<%=User.UserName%>(<%=User.RealName%>) | <a href="index.asp" target='_top'>后台首页</a> | <a href="" target="_blank">网站首页</a> | <a href="logout.asp" target="_top">退出</a></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="middle" id="frmTitle" valign="top" name="fmtitle" style="background:#c9defa">
			<iframe frameborder="0" id="frmleft" name="frmleft" src="left.asp" style="height: 100%; visibility: inherit;width: 185px;background:url(images/leftop.gif) no-repeat" allowtransparency="true"></iframe>
		</td>
		<td style="width:0px;" valign="middle">
			<div onClick="switchSysBar()">
				<span class="navpoint" id="switchPoint" title="关闭/打开左栏"><img src="images/right.gif" alt="" /></span>
			</div>
		</td>
		<td style="width: 100%" valign="top">
			<iframe frameborder="0" id="frmright" name="frmright" scrolling="yes" src="System/SysIndex.asp" style="height: 100%; visibility: inherit; width:100%; z-index: 1"></iframe>
		</td>
	</tr>
		<td height="30" colspan="3">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background:url(images/botbg.gif)">
				<tr height="32">
					<td style="padding-left:30px; font-family:arial; font-size:11px;"><b><%=Config.SysName%></b> Version <b><%=Config.SysVersion%></b> &nbsp;&nbsp;&nbsp;Powered By <b><%=Config.SysAuthor%></b> &nbsp;&nbsp;&nbsp;Copyright 2011 <b><%=Config.SysCopyRight%></b> All Rights Reserved</td>
					<td style="text-align:right;color:#135294;padding-right:20px;"><a href="http://xunzong.net" target="_blank"></a></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<div id="dvbbsannounce_true" style="display:none;">

</div>


<SCRIPT LANGUAGE="JavaScript">
<!--
document.getElementById("dcannounce").innerHTML = "<marquee width='300px' scrollamount=2>欢迎登录管理系统</marquee>";
//-->
</SCRIPT>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()	
	set db=nothing
%>