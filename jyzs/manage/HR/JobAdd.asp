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
<form name="JobAdd" method="post" action="JobSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('JobManage.asp');"/>
                <input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add">			</td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建职位</th>
		</tr>
		<tr class="tr1">
			<td width="30%">职位：</td>
			<td width="70%"><input name="Job_Title" type="text" id="Job_Title" size="50" />
			*</td>
		</tr>
		
		<tr class="tr2">
		  <td>招聘人数：</td>
		  <td><input name="Need_Num" type="text" id="Need_Num" value="若干" size="10" /></td>
	  </tr>
		<tr class="tr1">
			<td width="30%">薪水：</td>
			<td width="70%"><input name="Salary" type="text" id="Salary" value="面议" size="20" /></td>
		</tr>
		
		<tr class="tr2">
		  <td> 岗位职责 ：</td>
		  <td><input type="hidden" name="Responsibility"><iframe id="Responsibility" src="../editor/eWebEditor.asp?id=Responsibility&style=s_blue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
	  </tr>
		<tr class="tr1">
			<td width="30%"> 职位要求 ：</td>
			<td width="70%"><input type="hidden" name="Require"><iframe id="Require" src="../editor/eWebEditor.asp?id=Require&style=s_blue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
		</tr>
		<tr class="tr2">
		  <td>其他要求：</td>
		  <td><textarea name="Remark" cols="60" rows="6" id="Remark"></textarea></td>
	  </tr>
		<tr class="tr1">
		  <td>显示顺序：</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" value="0" size="10" />
*</td>
	  </tr>
  </table>
</form>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>