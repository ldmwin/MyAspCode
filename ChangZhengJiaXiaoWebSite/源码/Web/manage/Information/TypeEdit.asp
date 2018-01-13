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

dim info_id,p_name

info_id=request("ID")

if info_id="" or not isnumeric(info_id) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if
info_id=cint(info_id)


Dim rs_Type : Set rs_Type = db.getRecordBySQL("select *,(select Type_Name from Information_Type a where a.id=c.parent_id) as p_name from Information_Type c where id = " & info_id)




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
				<input type="hidden" name="action" value="edit"><input type="hidden" name="id" value="<%=info_id%>"><input type="hidden" name="p_id" value="<%=rs_Type("parent_id")%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建分类</th>
		</tr>
		<tr class="tr1">
			<td width="30%">上级分类：</td>
			<td width="70%"><%=rs_Type("p_name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类名称：</td>
			<td width="70%"><input name="Type_Name" type="text" id="Type_Name" size="50" value="<%=rs_Type("Type_Name")%>"/></td>
	  </tr>
		
		
		

		<tr class="tr1">
		  <td>显示顺序</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" size="10" value="<%=rs_Type("Show_Order")%>"/>
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
	
	db.C(rs_Type)

	db.CloseConn()
	
	set db=nothing
%>