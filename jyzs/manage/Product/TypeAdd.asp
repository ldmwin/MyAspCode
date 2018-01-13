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

dim p_id,p_name

P_ID=request("P_ID")

if P_ID="" or not isnumeric(P_ID) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if
P_ID=cint(P_ID)

if P_ID=0 then
	p_name="无"
else 
Dim rs_Type : Set rs_Type = db.getRecordBySQL("select * from Product_Type where id = " & P_ID)
	p_name=rs_Type("Type_Name")
db.C(rs_Type)

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
<form name="TypeAdd" method="post" action="TypeSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="history.go(-1);"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add"><input type="hidden" name="p_id" value="<%=p_id%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建分类</th>
		</tr>
		<tr class="tr1">
			<td width="30%">上级分类：</td>
			<td width="70%"><%=p_name%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类名称：</td>
			<td width="70%"><input name="Type_Name" type="text" id="Type_Name" size="50" /> *</td>
	  </tr>
		
		
		

		<tr class="tr1">
		  <td>显示顺序</td>
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