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


sqlstr = "select * from Jobs where status<>4 and id = " & Info_id


Dim rs_Job : Set rs_Job = db.getRecordBySQL(sqlstr)

if rs_Job.eof or rs_Job.bof then
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
				<%if rs_Job("status")=0 or rs_Job("status")=2 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('JobSave.asp?action=singleup&id=<%=rs_Job("id")%>');"/><input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('JobEdit.asp?id=<%=rs_Job("id")%>');"/>
                <%elseif rs_Job("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('JobSave.asp?action=singledown&id=<%=rs_Job("id")%>');"/>
			
			<%end if%>
            <input type="button" name="del" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('JobSave.asp?action=singlecancel&id=<%=rs_Job("id")%>','_self','确定作废？');"/>
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('JobManage.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">职位预览</th>
		</tr>
		<tr class="tr1">
			<td width="30%">职位：</td>
			<td width="70%"><%=rs_Job("Job_Title")%></td>
		</tr>
		
		<tr class="tr2">
		  <td>招聘人数：</td>
		  <td><%=rs_Job("Need_Num")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">薪水：</td>
			<td width="70%"><%=rs_Job("Salary")%></td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">岗位职责：</td>
			<td width="70%"><%=rs_Job("Responsibility")%>&nbsp;</td>
		</tr>
		<tr class="tr1">
		  <td>职位要求：</td>
		  <td><%=rs_Job("Require")%>&nbsp;</td>
	  </tr>
		<tr class="tr2">
		  <td>其他要求：</td>
		  <td><%=rs_Job("Remark")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>状态：</td>
		  <td><%=StatusResult(1,rs_Job("status"))%></td>
	  </tr>
	  <tr class="tr2">
		  <td>显示顺序：</td>
		  <td><%=rs_Job("Show_Order")%></td>
	  </tr>
		
</table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.C(rs_Job)

	db.CloseConn()
	
	set db=nothing
%>