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
Dim Keyword,RsStatus,SearchStartTime,SearchEndTime,AlbumType,Sqlstr
'Dim Proname:Proname="Album_List"
Dim Current_page
Dim DataRs:Set DataRs = Server.CreateObject("adodb.recordset")
Dim SubmitUrl:SubmitUrl = "AlbumManage.asp"

'接收查询参数
Sub SearchItemReceive()

if Request.ServerVariables("HTTP_METHOD")="POST" then 
	Keyword = Request.Form("Keyword")
	RsStatus = Request.Form("RsStatus")
	AlbumType = Request.Form("AlbumType")
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
	'response.Write(RsStatus & "yes")
	'response.Write(AlbumType & "yes")
	'Sqlstr = "select *,(select type_name from Album_Type where Album_Type.id=Albums.Album_Type) as Albumtypeshow,(select UserName from Sys_User where Sys_User.id= Albums.Adder_ID) as Adder from Albums where  " & TranslateDateTime(SearchStartTime,0) & " <=Add_Time and Add_Time<=" & TranslateDateTime(SearchEndTime,1) & " and status<>4"

	Sqlstr = "select *,(select type_name from Album_Type where Album_Type.id=Albums.Album_Type) as Albumtypeshow,(select UserName from Sys_User where Sys_User.id= Albums.Adder_ID) as Adder from Albums where status<>4"
	
	if Keyword <>"" then
		sqlstr = sqlstr & " and Title like '%" & Keyword & "%'"
	end if
	
	if RsStatus <>"10" then
		sqlstr = sqlstr & " and status = " & RsStatus
	end if
	
	if AlbumType <>"0" then
		sqlstr = sqlstr & " and Album_Type = " & AlbumType
	end if		
	
	sqlstr = sqlstr & " order by show_order desc,id desc"
	
	
else

	Keyword = null
	RsStatus = 10
	AlbumType = 0
	SearchStartTime = formatdatetime(Now()-90,2)
	SearchEndTime = formatdatetime(Now(),2)
	Current_Page = 1
	
	'sqlstr = "select *,(select type_name from Album_Type where Album_Type.id=Albums.Album_Type) as Albumtypeshow,(select UserName from Sys_User where Sys_User.id= Albums.Adder_ID) as Adder from Albums where  " & TranslateDateTime(formatdatetime(Now()-30,2),0) & " <=Add_Time and Add_Time<=" & TranslateDateTime(formatdatetime(Now(),2),1) & " and status<>4 order by show_order desc,id desc"
	sqlstr = "select *,(select type_name from Album_Type where Album_Type.id=Albums.Album_Type) as Albumtypeshow,(select UserName from Sys_User where Sys_User.id= Albums.Adder_ID) as Adder from Albums where status<>4 order by show_order desc,id desc"
	
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
<input type="hidden" name="Keyword" value="<%=Keyword%>" />
<input type="hidden" name="RsStatus" value="<%=RsStatus%>" />
<input type="hidden" name="AlbumType" value="<%=AlbumType%>" />
<input type="hidden" name="SearchStartTime" value="<%=SearchStartTime%>" />
<input type="hidden" name="SearchEndTime" value="<%=SearchEndTime%>" />
</form>
<%End Sub%>

<%
'其他操作时，页面返回查询时参数记录于session中
Sub OperateSearchItem()
	session("back_page") = Current_page
	session("back_Keyword") = Keyword
	session("back_RSstatus") = RSstatus
	session("back_AlbumType") = AlbumType
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
			<!--<input type="button" name="add" class="button" value="新建" style="margin-left:20px;" onClick="GotoUrl('AlbumAdd.asp');"/>-->
			<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('AlbumSave.asp?action=up','确认上线？');"/>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('AlbumSave.asp?action=down','确认下线？');"/>
			<input name="cancel" type="button" class="button" id="cancel" style="margin-left:20px;" value="作废" onClick="GoToUrl_MorePrm('AlbumSave.asp?action=cancel','确认作废？');"/>
			<input type="submit" name="put" class="button" value="查询" style="margin-left:20px;" onClick="document.frm_search.me_page.value=1;"/><!--切记，submit型input,name属性不能用submit！-->
			<input type="button" name="refresh" class="button" value="刷新" style="margin-left:20px;" onClick="Refresh('_submit');"/>
			</td>
		</tr>
		<tr class="tr1">
			<td width="10%">关键字：</td>
			<td width="23%">
				<input type="text" name="Keyword" value="<%=trim(Keyword)%>" />
			</td>
		    <td width="10%">状态：</td>
		    <td width="23%">
			<select size="1" name="RsStatus">
	  				<option value="10" <%=object_selected(cint(RsStatus),10)%>>所有</option>
					<option value="5" <%=object_selected(cint(RsStatus),5)%>>推荐</option>
					<option value="3" <%=object_selected(cint(RsStatus),3)%>>测试</option>
                    <option value="2" <%=object_selected(cint(RsStatus),2)%>>下线</option>
					<option value="1" <%=object_selected(cint(RsStatus),1)%>>上线</option>  
					<option value="0" <%=object_selected(cint(RsStatus),0)%>>初始</option>
      		</select>
			</td>
		    <td width="10%">信息类型：</td>
		    <td width="24%">
			<select name="AlbumType">
			<option value="0" <%=object_selected(cint(AlbumType),0)%>>所有</option>
			<%
			Set rs = db.getRecordBySQL("select * from Album_Type where status=1 and Type_Level=2 order by show_order desc,id desc")
			if not(rs.eof and rs.bof) then
			do until rs.eof
			response.Write("<option value=""" & rs("id") & """ " & object_selected(cint(AlbumType),cint(rs("id"))) & ">" & rs("Type_Name") & "</option>")
			rs.movenext
			loop
			end if
			db.C(rs)
		  %>
		  </select></td>
		</tr>
		<tr class="tr2">
			<td>添加开始时间：</td>
			<td><input type="text" name="SearchStartTime" id="SearchStartTime" onfocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchStartTime%>" /></td>
		    <td>添加结束时间：</td>
		    <td><input type="text" name="SearchEndTime" id="SearchEndTime" onfocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchEndTime%>" /><input type="hidden" name="me_page" value="<%=trim(Current_page)%>"/></td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		</tr>
</table>
</form>
<!--载入查询提交form End-->

<!--主数据载入form Start-->
<form name="frm_list" method="post" action="<%=SubmitUrl%>">
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="10" style="text-align:center;">图集列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();" /></td>
		<td width="5%"><b>ID</b></td>
		<td width="23%"><strong>标题</strong></td>
		<td width="9%"><strong>信息类型</strong></td>
		<td width="9%"><strong>状态</strong></td>
		<td width="15%"><strong>添加时间</strong></td>
		<td width="10%"><strong>添加人</strong></td>
		<td width="7%"><B>点击次数</B></td>
		<td width="7%"><B>排序级别</B></td>
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
		<td><a target="_self" href="AlbumView.asp?id=<%=DataRs("id")%>"><%=DataRs("Album_Name")%></a></td>
		<td align="center"><%=DataRs("AlbumTypeshow")%></td>
		<td align="center"><%=StatusResult(1,DataRs("status"))%></td>
		<td align="center"><%=DataRs("Add_Time")%></td>
		<td align="center"><%=DataRs("Adder")%></td>
		<td align="center"><%=DataRs("Album_Hits")%></td>
		<td align="center"><%=DataRs("Show_Order")%></td>
		<td align="center"><%if DataRs("status")=0 or DataRs("status")=2 then%>
						<input type="button" name="edit" class="button" value="编辑" style="margin-left:5px;" onClick="GotoUrl('AlbumEdit.asp?id=<%=DataRs("id")%>');"/>
						<%end if%>&nbsp;
	  </td>
	</tr>
<%
	DataRs.movenext()
next

%>
	<tr class="tr2">
		<td colspan="10" align="right">
			<div style="margin-right:50px; width:500px;">
				<%=ShowPage(DataRs.RecordCount,DataRs.PageCount,Current_page)%>			</div>		</td>
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