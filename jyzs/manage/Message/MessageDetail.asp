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

dim Info_id

Info_id=request.QueryString("id")

if Info_id="" or Info_id=0 or not isnumeric(Info_id) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_id=cint(Info_id)


sqlstr = "select * from Messages where status<>4 and id = " & Info_id

'response.Write(sqlstr)
'response.End()


Dim rs_Message : Set rs_Message = db.getRecordBySQL(sqlstr)

if rs_Message.eof or rs_Message.bof then
	response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
	response.End()
end if		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript" src="../JS/My97DatePicker/WdatePicker.js"></script>
</head>

<body>
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('MessageList.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">留言详情</th>
		</tr>
		<tr class="tr1">
			<td width="30%">留言人：</td>
			<td width="70%"><%=rs_Message("RealName")%></td>
		</tr>
		
		<tr class="tr2">
		  <td>联系电话：</td>
		  <td><%=rs_Message("Tel")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">联系邮箱：</td>
			<td width="70%">
			<%=rs_Message("Email")%>&nbsp;
            </td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">留言内容：</td>
			<td width="70%">
			<%=rs_Message("Content")%>
            </td>
		</tr>
		<tr class="tr1">
		  <td>留言时间：</td>
		  <td><%=rs_Message("Add_Time")%></td>
	  </tr>
	  <tr class="tr2">
		  <td>留言IP：</td>
		  <td><%=rs_Message("Adder_IP")%></td>
	  </tr>
		
</table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.C(rs_Message)

	db.CloseConn()
	
	set db=nothing
%>