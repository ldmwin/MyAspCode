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
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
</head>

<body>
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="4"><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="4" style="text-align:center;">用户基本信息</th>
		</tr>
		<tr class="tr1">
			<td width="20%">用户名：</td>
			<td width="30%"><%=User.UserName%>&nbsp;</td>
		    <td width="20%">真实姓名：</td>
		    <td width="30%"><%=User.RealName%>&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td>所属部门：</td>
			<td><%=User.UserBaseProfile(User.UserID,"Org")%>&nbsp;</td>
	        <td>状态：</td>
	        <td><%=User.StatusShow%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td>创建时间：</td>
			<td><%=User.UserBaseProfile(User.UserID,"CreateTime")%>&nbsp;</td>
	        <td>登陆次数：</td>
	        <td><%=User.UserBaseProfile(User.UserID,"LoginNum")%>&nbsp;</td>
	  </tr>
		<tr class="tr2">
			<td>最后登录时间：</td>
			<td><%=User.UserBaseProfile(User.UserID,"LastLoginTime")%>&nbsp;</td>
	        <td>最后登陆IP：</td>
	        <td><%=User.UserBaseProfile(User.UserID,"LastLoginIP")%>&nbsp;</td>
	  </tr>
	  		<tr>
	  <th colspan="4" style="text-align:center;">用户角色</th>
		</tr>

		<tr class="tr2">
			<td>角色列表：</td>
			<td colspan="3">&nbsp;&nbsp;</td>
        </tr>
</table>
	<br />
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>