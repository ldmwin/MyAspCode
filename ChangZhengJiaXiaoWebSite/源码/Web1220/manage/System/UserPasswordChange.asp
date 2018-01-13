<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>

<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Class/UserCtrl.asp" -->
<!--#include file="../Class/LogCtrl.asp" -->
<!--#include file="../Inc/MD5.asp" -->
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

Dim OldPW,PassWord,PassWordSure

Sub PageShow()
	Dim Action:Action = request.QueryString("action")
	
	if Action = "save" then
		Call PasswordSave()
	else
		Call PasswordChange()
	end if
End Sub

'数据接收及校验
Sub ReceiveDate()

OldPW = ParaEncode(trim(request.Form("OldPW")))
PassWord = ParaEncode(trim(request.Form("PassWord")))
PassWordSure = ParaEncode(trim(request.Form("PassWordSure")))

if OldPW ="" then
	response.write "<script>alert('原密码不能为空');history.go(-1);</Script>"
	response.End()
end if


if PassWordSure <> PassWord then
	response.write "<script>alert('两次输入的密码不符，请重新输入');history.go(-1);</Script>"
	response.End()
end if


if PassWord = "" or (not InfoRegularCheck(PassWord,"^[a-zA-Z0-9]{6,20}$")) then
	response.write "<script>alert('新密码不符合要求，请重新输入');history.go(-1);</Script>"
	response.End()
end if

End Sub

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
<%Call PageShow()%>
<%Sub PasswordChange()%>
<form name="PW" method="post" action="UserPasswordChange.asp?action=save">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add">			</td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">修改密码</th>
		</tr>
		<tr class="tr1">
			<td width="30%">用户名：</td>
			<td width="70%"><%=User.UserName%>&nbsp;</td>
		</tr>
		<tr class="tr2">
		  <td>原密码：</td>
		  <td><input name="OldPW" type="password" id="OldPW" size="50" /> *</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">新密码：</td>
		  <td width="70%"><input name="PassWord" type="password" id="PassWord" size="50" />
		    * 长度6-20位，可由大、小写英文字母和数字进行组合</td>
	  </tr>
		<tr class="tr2">
			<td width="30%">再次输入新密码：</td>
		  <td width="70%"><input name="PassWordSure" type="password" id="PassWordSure" size="50" />
		    *</td>
		</tr>
  </table>
</form>
<%End Sub%>
<%Sub PasswordSave()%>
<%

	Call ReceiveDate()
	
	if User.CheckUserByID(User.UserID,OldPW) then
	
		Call User.PasswordChange(User.UserID,PassWord)
		response.write("<script>alert('密码修改成功，请重新登录');window.top.location.href = '../logout.asp'; </Script>")
		
	else
		'response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=3")
		response.write("<script>alert('密码修改失败，" & User.UserErr & "');history.go(-1);</script>")
	end if

%>
<%End Sub%>

</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>