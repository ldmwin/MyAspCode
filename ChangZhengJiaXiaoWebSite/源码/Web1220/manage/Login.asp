<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<%
'-----------------------------------
'文 件 名 : /login.asp
'功    能 : 登陆及校验
'作    者 : Mr.Lion
'建立时间 : 2011/05/06
'-----------------------------------
%>
<!--#include file="Class/DBCtrl.asp" -->
<!--#include file="Class/UserCtrl.asp" -->
<!--#include file="Class/LogCtrl.asp" -->
<!--#include file="Class/InputCheck.asp" -->
<!--#include file="Inc/MD5.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
	Dim Config: Set Config = New ClsConfig
	
	Dim db : Set db = New DbCtrl
	db.dbConnStr = Config.ConnStr(1,"")
	db.OpenConn()
	
	'response.Write("链接数据库成功")
	
	Dim user : Set user = New UserCtrl
	Dim IC: Set IC = New InputCheck
	Dim EventLog: Set EventLog = New LogCtrl


	
	If Request.form("reaction")="chklogin" Then
		Call ChkLogin()
	Else
		Call AdminLoginMain()
	End If

%>

<%Sub AdminLoginMain()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>网站管理系统</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/login.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/function.js" language="javascript"></script>
</head>

<body>
<div id="nifty">
<b class="rtop"><b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b>
<div style="width:403px; height:26px; line-height:26px; background:none; font-size:12px; text-align:left;"><%=Config.SysCNName%> -- 管理登录</div>
<div style="width:403px; height:46px; background:#166CA3;"><img src="images/login.gif" alt="" /></div>
<div style="width:401px !important; width:403px; height:auto; background:#fff; border-left:1px solid #649EB2; border-right:1px solid #649EB2; ">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
<form action="login.asp" method="post">
<input name="reaction" type="hidden" value="chklogin" />
	<tr>
		<td align="right"><b>用户名：</b></td>
		<td align="left"><input name="username" type="text" tabindex="4"/></td>
	</tr>
	<tr>
		<td align="right"><b>密　码：</b></td>
		<td align="left"><input name="password" type="password" tabindex="5"/></td>
	</tr>
	<tr>	
		<td align="right"><b>附加码：</b></td>
		<td align="left"><input type="text" name="codestr" id="codestr" size="8" maxlength="4" tabindex="6" onFocus="get_Code();this.onfocus=null;" /><span id="imgid" style="color:red">点击获取验证码</span><span id="isok_codestr"></span></td>
	</tr>
	<tr>
	<td align="right"></td>
	<td align="left"><input  class="button" type="submit" name="submit" value="登 录"/></td>
	</tr>	
  </form>
</table>
</div>
<div style="width:401px !important; width:403px; height:20px; background:#F7F7E7; border:1px solid #649EB2; border-top:1px solid #ddd; margin-bottom:5px; font-size:12px; line-height:20px; "> <%=Config.SysName%> Version <%=Config.SysVersion%> </div>
<b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
</div>
</body>
</html>
<%end sub%>

<%
sub ChkLogin()

	Dim UserName
	Dim PassWord
	Dim Code
	UserName = Replace(Request.Form("username"),"'","")
	PassWord = Request.Form("password")
	Code = Request.Form("codestr")
'	
'	'缺一个数据校验类校验和过滤，以后补上。
'	先检查校验码


	if not IC.CodeCheck(Code) then
		Call EventLog.LogAdd(2,0,"system:login fail 校验码错误")
		response.Redirect("inc/error.asp?msg=校验码输入错误。&errurl=1")
		Exit Sub
	end if

'	检查用户登录	
	if not User.CheckUser(UserName,PassWord) then
		
		Call EventLog.LogAdd(2,0,"system:login fail " & User.UserErr)
		response.Redirect("inc/error.asp?msg=" & User.UserErr & "。&errurl=1")
		Exit Sub
	else
		
		if User.UserLogin(UserName) then
			
			Call EventLog.LogAdd(1,User.UserID,"system:login sucess")			
			response.Redirect("main.asp")
		
		else
			Call EventLog.LogAdd(2,0,"system:login fail" & User.UserErr)
			response.Redirect("inc/error.asp?msg=" & User.UserErr & "。&errurl=1")
			Exit Sub
		
		end if
		
		
	end if

'	

end sub
	
	set User = nothing
	set IC = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>