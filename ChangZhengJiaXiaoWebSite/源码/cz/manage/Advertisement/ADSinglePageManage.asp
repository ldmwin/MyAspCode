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
db.dbConnStr = Config.ConnStr(0)
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

'接收查询数据
Dim Keyword,AD_ID,RsStatus,SearchStartTime,SearchEndTime
Dim Proname:Proname="Advertisement_SinglePage_List"
Dim Current_page
Dim DataRs:Set DataRs = Server.CreateObject("adodb.recordset")
Dim SubmitUrl:SubmitUrl = "ADSinglePageManage.asp"

AD_ID = Request("AD_ID")

if AD_ID = "" or AD_ID = 0 or not isnumeric(AD_ID) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

AD_ID = cint(AD_ID)

'response.Write(sqlstr)
'response.End()

'接收查询参数
Sub SearchItemReceive()

if Request.ServerVariables("HTTP_METHOD")="POST" then 
	Keyword = Request.Form("Keyword")
	RsStatus = Request.Form("RsStatus")
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
	
else

	Keyword = null
	RsStatus = 10
	SearchStartTime = formatdatetime(Now()-7,2)
	SearchEndTime = formatdatetime(Now(),2)
	
	Current_Page = 1
	
end if

End Sub


Sub DateSearch()

	db.CreateProCmd(Proname)
	
'	response.Write(Keyword)
'	response.Write(RsStatus)
'	response.Write(SearchStartTime)
'	response.Write(SearchEndTime)
'	response.Write(AD_ID)

	Call db.ProCmdAddPra("Keyword",200,1,50,Keyword)
	Call db.ProCmdAddPra("RsStatus",200,1,50,RsStatus)
	Call db.ProCmdAddPra("Starttime",200,1,50,SearchStartTime)
	Call db.ProCmdAddPra("Endtime",200,1,50,SearchEndTime)
	Call db.ProCmdAddPra("AD_ID",200,1,50,AD_ID)
	
	if not db.ProCmdExcute() then
	
		'response.Redirect("/inc/error.asp?msg=数据查询失败。&errurl=1")
	
	else
		
		Set DataRs = db.ProCmdGetOutRS()
		
	end if
	
	DataRs.PageSize=15'每页显示的记录数

	if (Current_page > DataRs.PageCount) and (DataRs.PageCount>0) then 
	  Current_page = DataRs.PageCount
	elseif Current_page < 1 then
	  Current_page = 1
	end if
     
	if not (DataRs.eof or DataRs.bof) then
		DataRs.AbsolutePage = Current_page
	end if
	 
	db.DestroyProCmd()
	
End Sub
%>

<%
'分页跳转提交过程，查询条件全部提交
Sub PageSearchItem()
%>
<form action="<%=SubmitUrl%>" method=post name="frm_page">
<input type="hidden" name="me_page" />
<input type="hidden" name="Keyword" value="<%=Keyword%>" />
<input type="hidden" name="RsStatus" value="<%=RsStatus%>" />
<input type="hidden" name="SearchStartTime" value="<%=SearchStartTime%>" />
<input type="hidden" name="SearchEndTime" value="<%=SearchEndTime%>" />
<input type="hidden" name="AD_ID" value="<%=AD_ID%>" />
</form>
<%End Sub%>

<%
'其他操作时，页面返回查询时参数记录于session中
Sub OperateSearchItem()
	session("back_page") = Current_page
	session("back_Keyword") = Keyword
	session("back_RSstatus") = RSstatus
	session("back_SearchStartTime") = SearchStartTime
	session("back_SearchEndTime") = SearchEndTime
	session("back_AD_ID") = AD_ID
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

sqlstr = "select *,dbo.StatusShow(8,status) as statusshow,(select sort_name from Advertisement_Sort where Advertisement_Sort.id=Advertisement.Sort_ID) as adtypeshow from Advertisement where status<>4 and id = " & AD_ID

Dim rs_ad : Set rs_ad = db.getRecordBySQL(sqlstr)

if rs_ad.eof or rs_ad.bof then
	response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
	response.End()
end if
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
		  <th colspan="4" style="text-align:center;">对应页面广告</th>
		</tr>
		<tr class="tr1">
			<td width="20%">广告名称：</td>
			<td width="30%"><%=rs_ad("Advertisement_Name")%></td>
	      <td width="20%">广告分类：</td>
		    <td width="30%"><%=rs_ad("adtypeshow")%></td>
	    </tr>
</table>
<%
db.C(rs_ad)
%>

<!--载入查询提交form Start-->
 <form name="frm_search" action="<%=SubmitUrl%>" method="post">
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td colspan="6"><input name="return" type="button" class="button" id="return" onClick="GotoUrl('ADSinglePageAdd.asp?ad_id=<%=AD_ID%>');" value="新建项目" style="margin-left:20px;"/>
			<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('ADSinglePageSave.asp?action=up','确认上线？');"/>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('ADSinglePageSave.asp?action=down','确认下线？');"/>
			<input name="cancel" type="button" class="button" id="cancel" style="margin-left:20px;" value="作废" onClick="GoToUrl_MorePrm('ADSinglePageSave.asp?action=cancel','确认作废？');"/>
			<input type="submit" name="set" class="button" value="查询" style="margin-left:20px;" onClick="document.frm_search.me_page.value=1;"/><input name="return" type="button" class="button" id="return" onClick="GotoUrl('ADView.asp?id=<%=AD_ID%>');" value="返回广告" style="margin-left:20px;"/><input type="button" name="returnlist" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('ADManage.asp');"/>
			<input type="button" name="refresh" class="button" value="刷新" style="margin-left:20px;" onClick="Refresh('_submit');"/>
		  </td>
		</tr>
		<tr class="tr1">
			<td width="10%">关键字：</td>
			<td width="23%">
				<input type="text" name="Keyword" value="<%=trim(Keyword)%>" />
			</td>
		    <td width="10%"><input type="hidden" name="AD_ID" value="<%=trim(AD_ID)%>" /><input type="hidden" name="me_page" value="<%=trim(Current_page)%>"/>状态：</td>
		    <td width="23%"><select size="1" name="RsStatus">
              <option value="10" <%=object_selected(cint(RsStatus),10)%>>所有</option>
              <option value="3" <%=object_selected(cint(RsStatus),3)%>>测试</option>
			  <option value="2" <%=object_selected(cint(RsStatus),2)%>>下线</option>
              <option value="1" <%=object_selected(cint(RsStatus),1)%>>上线</option>
			  <option value="0" <%=object_selected(cint(RsStatus),0)%>>初始</option>
            </select></td>
		    <td width="10%">&nbsp;</td>
		    <td width="24%">&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td>开始时间：</td>
			<td><input type="text" name="SearchStartTime" id="SearchStartTime" onfocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchStartTime%>" /></td>
		    <td>结束时间：</td>
		    <td><input type="text" name="SearchEndTime" id="SearchEndTime" onfocus="WdatePicker({dateFmt:'yyyy-M-d'})" class="Wdate" value="<%=SearchEndTime%>" /></td>
		    <td>&nbsp;</td>
	      <td>&nbsp;</td>
		</tr>
</table>
</form>
<!--载入查询提交form End-->

<!--主数据载入form Start-->
<form name="frm_list" method="post" action="<%=SubmitUrl%>">
<input type="hidden" name="AD_ID" value="<%=trim(AD_ID)%>" />
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="10" style="text-align:center;">广告项目列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();"></td>
		<td width="5%"><b>ID</b></td>
		<td width="15%"><strong>项目名称</strong></td>
		<td width="7%"><strong>状态</strong></td>
		<td width="15%"><strong>对应链接</strong></td>
		<td width="7%"><strong>打开方式</strong></td>
		<td width="7%"><strong>图片</strong></td>
		<td width="21%"><strong>展示开始时间/展示结束时间</strong></td>
		<td width="7%"><b>排序级别</b></td>
		<td width="11%"><strong>操作</strong></td>
	</tr>
<%

for i = 1 to DataRs.pagesize
'	On Error Resume Next
	if DataRs.bof or DataRs.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" >
		<td align="center"><input class="checkbox" type="checkbox" name="info_id" id="info_id" value="<%=DataRs("ID")%>"></td>
		<td align="center"><span><%=DataRs("ID")%></span></td>
		<td align="center"><%=DataRs("Item_Name")%></td>
		<td align="center"><%=DataRs("statusshow")%></td>
		<td align="center"><%=DataRs("Link")%></td>
		<td align="center"><%=DataRs("Target")%></td>
		<td align="center"><%if DataRs("Pic_Url")<>"" then%>
               <a href="<%=Config.ImgUrl()%>AD/Picture/<%=DataRs("Pic_Url")%>" target="_blank">点击预览</a>
                <%else%>
无
<%end if%>&nbsp;</td>
		<td align="center"><%
		if DataRs("Show_Start_Time") <> "" then%>
		<%=DataRs("Show_Start_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;>><%if DataRs("Show_End_Time")<>"" then%>
		<%=DataRs("Show_End_Time")%>
		<%else%>
		结束时间无限制
		<%end if%></td>
		<td align="center"><%=DataRs("Show_Order")%></td>
		<td align="center"><%if DataRs("status")=0 or DataRs("status")=2 then%>
						<input type="button" name="edit" class="button" value="编辑" style="margin-left:5px;" onClick="GotoUrl('ADSinglePageEdit.asp?id=<%=DataRs("id")%>');"/>
						<%end if%>&nbsp;&nbsp;</td>
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