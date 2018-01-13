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


'接收查询数据
Dim SearchStartTime,SearchEndTime,sqlstr
'Dim Proname:Proname="Information_List"
Dim Current_page
Dim DataRs
Dim SubmitUrl:SubmitUrl = "MessageList.asp"

'接收查询参数
Sub SearchItemReceive()

if Request.ServerVariables("HTTP_METHOD")="POST" then 

	SearchStartTime = Request.Form("SearchStartTime")
	SearchEndTime = Request.Form("SearchEndTime")
	Current_page = Request.Form("me_page")	

	if Current_Page = "" then
		Current_Page = 1
	elseif not isnumeric(Current_Page) then
		Current_Page = 1
	else
		Current_page = cint(Current_Page)
	end if
	
	sqlstr = "select * from Messages where  " & TranslateDateTime(SearchStartTime,0) & " <=Add_Time and Add_Time<=" & TranslateDateTime(SearchEndTime,1) & " and status<>4"	
	
	sqlstr = sqlstr & " order by id desc"
	
else

	SearchStartTime = formatdatetime(Now()-7,2)
	SearchEndTime = formatdatetime(Now(),2)
	Current_Page = 1
	
	sqlstr = "select * from Messages where  " & TranslateDateTime(formatdatetime(Now()-30,2),0) & " <=Add_Time and Add_Time<=" & TranslateDateTime(formatdatetime(Now(),2),1) & " and status<>4 order by id desc"
	
end if

End Sub


Sub DateSearch()
	
	Set DataRs = db.getRecordBySQL(sqlstr)

	
	DataRs.PageSize=15'每页显示的记录数

	if (Current_page > DataRs.PageCount) and (DataRs.PageCount>0) then 
	  Current_page = DataRs.PageCount
	elseif Current_page < 1 then
	  Current_page = 1
	end if
     
	if not (DataRs.eof or DataRs.bof) then
		DataRs.AbsolutePage = Current_page
	end if
	 
	'db.DestroyProCmd()
	
End Sub
%>

<%
'分页跳转提交过程，查询条件全部提交
Sub PageSearchItem()
%>
<form action="<%=SubmitUrl%>" method="post" name="frm_page">
<input type="hidden" name="me_page" />
<input type="hidden" name="SearchStartTime" value="<%=SearchStartTime%>" />
<input type="hidden" name="SearchEndTime" value="<%=SearchEndTime%>" />
</form>
<%End Sub%>

<%
'其他操作时，页面返回查询时参数记录于session中
Sub OperateSearchItem()
	session("back_page") = Current_page
	session("back_SearchStartTime") = SearchStartTime
	session("back_SearchEndTime") = SearchEndTime
End Sub
%>

<%
Call SearchItemReceive()
Call DateSearch()
Call OperateSearchItem()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript" language="javascript"></script>
<script src="../js/function.js" type="text/javascript" language="javascript"></script>
<script language="javascript" type="text/javascript" src="../JS/My97DatePicker/WdatePicker.js"></script>
</head>

<body>
<%
'载入分页查询提交form
Call PageSearchItem()
%>

<!--载入查询提交form Start-->
 <form name="frm_search" action="<%=SubmitUrl%>" method="post">
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td colspan="6">
            <input type="submit" name="put" class="button" value="查询" style="margin-left:20px;" onClick="document.frm_search.me_page.value=1;"/><!--切记，submit型input,name属性不能用submit！-->
			<input type="button" name="refresh" class="button" value="刷新" style="margin-left:20px;" onClick="Refresh('_submit');"/>
			</td>
		</tr>
		<tr class="tr2">
			<td width="10%">提交开始时间：</td>
			<td width="23%"><input type="text" name="SearchStartTime" id="SearchStartTime" onFocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchStartTime%>" /></td>
		    <td width="10%">提交结束时间：</td>
		    <td width="23%"><input type="text" name="SearchEndTime" id="SearchEndTime" onFocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchEndTime%>" /><input type="hidden" name="me_page" value="<%=trim(Current_page)%>"/></td>
		    <td width="10%">&nbsp;</td>
		    <td width="24%">&nbsp;</td>
		</tr>
</table>
</form>
<!--载入查询提交form End-->

<!--主数据载入form Start-->
<form name="frm_list" method="post" action="<%=SubmitUrl%>">
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="7" style="text-align:center;">留言列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();" /></td>
		<td width="5%"><b>ID</b></td>
		<td width="20%"><strong>留言人</strong></td>
		<td width="20%"><strong>联系电话</strong></td>
		<td width="20%"><strong>提交时间</strong></td>
		<td width="20%"><strong>提交IP</strong></td>
		<td width="10%"><B>操作</B></td>
	</tr>
<%

for i = 1 to DataRs.pagesize
'	On Error Resume Next
	if DataRs.bof or DataRs.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" >
		<td align="center"><input class="checkbox" type="checkbox" name="info_id" value="<%=DataRs("ID")%>"></td>
		<td align="center"><span><%=DataRs("ID")%></span></td>
		<td><a target="_self" href="MessageDetail.asp?id=<%=DataRs("id")%>"><%=DataRs("RealName")%></a></td>
		<td align="center"><%=DataRs("Tel")%></td>
		<td align="center"><%=DataRs("Add_Time")%></td>
		<td align="center"><%=DataRs("Adder_IP")%></td>
		<td align="center">&nbsp;	    </td>
	</tr>
<%
	DataRs.movenext()
next

%>
	<tr class="tr2">
		<td colspan="7" align="right">
			<div style="margin-right:50px; width:500px;">
				<%=ShowPage(DataRs.RecordCount,DataRs.PageCount,Current_page)%>
             </div>
         </td>
	</tr>
</table>
</form>
<!--主数据载入form End-->

</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.C(DataRs)
	
	db.CloseConn()
	
	set db=nothing
%>