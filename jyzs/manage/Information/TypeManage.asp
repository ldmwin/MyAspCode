<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>

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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>无标题文档</title>
<link href="../css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="99%" height="100%" border=0 align="center" cellpadding=0 cellspacing=0 style="background:#C3DAF9; border-color:#C9DEFA;">
	
<!--	<tr>
		<td height="10" colspan="3">&nbsp;</td>
	</tr>-->
	<tr>
		<td align="middle" id="frmTitle" valign="top" name="fmtitle" style="background:#c9defa; border-color:#C9DEFA;">
			<iframe frameborder="0" id="Typemagleft" name="Typemagleft" src="TypeTree.asp" style="height: 100%; visibility: inherit;width: 230px;background:url(images/leftop.gif) no-repeat" allowtransparency="true"></iframe></td>
		<!--<td style="width:5px; border-color:#C9DEFA;" valign="middle">&nbsp;</td>-->
		<td style="width: 100%; border-color:#C9DEFA;" valign="top">
			<iframe frameborder="0" id="Typemagright" name="Typemagright" scrolling="yes" src="TypeList.asp" style="height: 100%; visibility: inherit; width:100%; z-index: 1"></iframe></td>
<!--	</tr>
		<td height="10" colspan="3">&nbsp;</td>
	</tr>-->
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