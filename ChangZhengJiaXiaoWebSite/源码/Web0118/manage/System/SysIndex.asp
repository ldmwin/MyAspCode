<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<%
'-----------------------------------
'文 件 名 : /SysIndex.asp
'功    能 : 系统登录首页
'作    者 : Mr.Lion
'建立时间 : 2011/08/19
'-----------------------------------
%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Class/UserCtrl.asp" -->
<!--#include file="../Class/LogCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="padding:2px 0px 2px 20px;" colspan="2">信息统计</th>
	</tr>
	<tr class="tr2">
		<td width="50%">服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
		<td>脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="tr2">
		<td width="50%">IIS 版本：(<%=Request.ServerVariables("SERVER_SOFTWARE")%>)</td>
		<td><a href="system.asp"><!--查看更详细服务器信息检测--></a>&nbsp;</td>
	</tr>
	<tr class="tr2">
		<td colspan="2">数据定期备份：请注意做好定期数据备份，数据的定期备份可最大限度的保障网站数据的安全 </td>
	</tr>
</table>

<br/>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="padding:2px 0px 2px 20px;" colspan="2">软件信息</th>
	</tr>
	<tr class="tr2">
		<td width="30%">产品名称：</td>
		<td width="70%"><%=Config.SysCNName%>/<%=Config.SysName%></td>
	</tr>
	<tr class="tr2">
		<td>产品开发：</td>
		<td><%=Config.SysAuthor%></td>
	</tr>
</table>

<br/>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="padding:2px 0px 2px 20px;" colspan="2">服务支持</th>
	</tr>
	<tr class="tr2">
		<td width="30%">版本检测：</td>
		<td width="70%">
			当前版本：V <%=Config.SysVersion%> &nbsp;
		</td>
	</tr>
	<tr class="tr2">
		<td>在线咨询：</td>
		<td><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=3715198&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:3715198:45" alt="点击这里给我发消息" title="点击这里给我发QQ消息"></a>&nbsp;</td>
	</tr>
</table>

</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>