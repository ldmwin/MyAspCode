<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>  
<%
'-----------------------------------
'文 件 名 : 
'功    能 : 左侧导航，调用树状结构除最高一级以外的级别，由top.asp传入参数，默认载入系统管理 0
'作    者 : Mr.Lion
'建立时间 : 2011/05/12
'页面权限： system:login
'-----------------------------------
%>
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

dim Info_ID

Info_ID=request.QueryString("Site_ID")

if Info_ID="" or Info_ID=0 or not isnumeric(Info_ID) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_ID=cint(Info_ID)


sqlstr = "select * from Sys_Sites where status<>4 and id = " & Info_ID

Dim rs_Site : Set rs_Site = db.getRecordBySQL(sqlstr)

if rs_Site.eof or rs_Site.bof then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Site_Name = rs_Site("Site_Name")
Site_Version = rs_Site("Version")
Site_Status = StatusResult(3,rs_Site("status"))

db.C(rs_Site)


Function BuildXMLStr(pid,str) '递归类别及其子类别存入字符串
	Dim rs_menutree,tempStr,i
	'Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum from Site_Columns m where m.status<>4 and m.parent_id=" & pid & " order by m.show_order desc,m.id desc") 
	Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum,(select Type_Name from site_column_type where m.column_type = site_column_type.id) as Type_Name from Site_Columns m where m.status<>4 and m.parent_id=" & pid & " and m.site_id=" & Info_ID & " order by m.show_order desc,m.id desc")
	i  = 0
	do while not rs_menutree.eof
		if i = 0 then
			str = str & "<ul>" & vbcrlf
		end if

		if rs_menutree("childnodenum") > 0 then
			'有子目录

			str = str & "<li>"
			str = str & "<div style=""float:left; display:inline;""><input type=""checkbox"" name=""info_id"" id=""info_id"" value=""" & rs_menutree("id") & """ style=""border-width:0;""></div>"
			str = str & "<div style=""display:inline; width:100%;""><div style=""float:left;""><a href=""#""><b>" & rs_menutree("Column_Name") & "</b></a>&nbsp;(" & rs_menutree("Type_Name") & ")(" & StatusResult(3,rs_menutree("status")) & ")</div><div style=""float:right; margin-right: 50px; width:150px;""><a href=""#"" id=""addcolumn" & rs_menutree("id") & """>添加</a>"
			if rs_menutree("status")<>1 then 
				str = str & "&nbsp;&nbsp;&nbsp;<a href=""#"" id=""editcolumn" & rs_menutree("id") & """>编辑</a>"
			end if
			
			if rs_menutree("Column_Type") = 2 then
				str = str & "&nbsp;&nbsp;&nbsp;<a href=""#"" id=""columncontent" & rs_menutree("id") & """>内容管理</a><script type=""text/javascript"">$(""#columncontent" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:750,height:500},target:""ColumnContentEdit.asp?ID=" & rs_menutree("id") & """});</script>"
			end if
			
			str = str & "<script type=""text/javascript"">$(""#addcolumn" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:500,height:350},target:""ColumnAdd.asp?Parent_id=" & rs_menutree("id") & "&Site_ID=" & rs_menutree("Site_ID") &"""});</script><script type=""text/javascript"">$(""#editcolumn" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:500,height:350},target:""ColumnEdit.asp?ID=" & rs_menutree("id") & """});</script>"
		
			str = str & "</div></div>"
		else
			'无子目录
			str = str & "<li class=""Child"">"
			str = str & "<div style=""float:left; display:inline;""><input type=""checkbox"" name=""info_id"" id=""info_id"" value=""" & rs_menutree("id") & """  style=""border-width:0;""></div>"
			str = str & "<div style=""display:inline; width:100%;""><div style=""float:left;""><a href=""#""><b>" & rs_menutree("Column_Name") & "</b></a>&nbsp;(" & rs_menutree("Type_Name") & ")(" & StatusResult(3,rs_menutree("status")) & ")</div><div style=""float:right; margin-right: 50px; width:150px;""><a href=""#"" id=""addcolumn" & rs_menutree("id") & """>添加</a>"
			if rs_menutree("status")<>1 then 
				str = str & "&nbsp;&nbsp;&nbsp;<a href=""#"" id=""editcolumn" & rs_menutree("id") & """>编辑</a><script type=""text/javascript"">$(""#editcolumn" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:500,height:350},target:""ColumnEdit.asp?ID=" & rs_menutree("id") & """});</script>"
			end if
			
			if rs_menutree("Column_Type") = 2 then
				str = str & "&nbsp;&nbsp;&nbsp;<a href=""#"" id=""columncontent" & rs_menutree("id") & """>内容管理</a><script type=""text/javascript"">$(""#columncontent" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:750,height:500},target:""ColumnContentEdit.asp?ID=" & rs_menutree("id") & """});</script>"
			end if
			
			str = str & "&nbsp;&nbsp;&nbsp;<a href=""ColumnSave.asp?action=singlecancel&ID=" & rs_menutree("id") & """>作废</a>"
			str = str & "<script type=""text/javascript"">$(""#addcolumn" & rs_menutree("id") & """).wBox({requestType:""iframe_refresh"",iframeWH:{width:500,height:350},target:""ColumnAdd.asp?Parent_id=" & rs_menutree("id") & "&Site_ID=" & rs_menutree("Site_ID") &"""});</script>"
			
			str = str & "</div></div>" & vbcrlf
		end if
		Call BuildXMLStr(rs_menutree("ID"),str) '递归调用
		rs_menutree.movenext()
		i = i + 1
		str = str & "</li>" & vbcrlf
		if rs_menutree.eof then str = str & "</ul>" & vbcrlf
	Loop
	BuildXMLStr = str
	db.C(rs_menutree)

End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>无标题文档</title>
<link rel="stylesheet" href="../css/jquery.treeview.css" />
 
<!--<script src="../js/jquery.js" type="text/javascript"></script>-->
<script type="text/javascript" src="../js/wbox/jquery1.4.2.js"></script>
<script src="../js/jquery.cookie.js" type="text/javascript"></script>
<script src="../js/jquery.treeview.js" type="text/javascript"></script>
 
<script type="text/javascript">
		$(function() {
			$("#tree").treeview({
				collapsed: false,
				animated: "medium",
				control:"#sidetreecontrol",
				persist: "location"
			});
		})
		
	</script>
	
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
<link href="../css/main.css" rel="stylesheet" type="text/css" />
 
<script type="text/javascript" src="../js/wbox/wbox.js"></script>
<link rel="stylesheet" type="text/css" href="../js/wbox/wbox/wbox.css" />
</head>

<body>
<table border="0" align="center" cellpadding="0" cellspacing="1" width="100%">
  <tr>
    <th  colspan="4" style="text-align:center;">站点信息</th>
  </tr>
  <tr  class="tr2">
    <td colspan="4">
      <input name="return" type="button" class="button" id="return" onClick="GotoUrl('ShopView.asp?id=<%=Info_id%>');" value="返回" style="margin-left:20px;"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
  </tr>
    <tr class="tr1">
      <td width="20%"  align="right">网站名称：</td>
      <td width="30%" ><%response.Write(Site_Name)%>(V <%response.Write(Site_Version)%>)</td>
      <td width="20%" align="right" >状态：</td>
      <td width="30%" ><%response.Write(Site_Status)%></td>
  </tr>
</table>
<form name="frm_list" method="post" action="ColumnSave.asp">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <tr> 
          <th colspan="4" style="text-align:center;">栏目管理</th>
        </tr>    
	  
        <tr class="tr2">
          <td align="left" >
           <input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('ColumnSave.asp?action=up','确认上线？');"/>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GoToUrl_MorePrm('ColumnSave.asp?action=down','确认下线？');"/>
			<input type="button" name="refreshPage" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();" style="margin-left:20px;">全选<input type="hidden" value="<%=Info_ID%>" name="Site_ID"></td>
        </tr>
        
          
          <tr align="left" class="tr1" > 
            <td height="340"  style="padding-left:40px;">
              <br />
    <ul id="tree" style="width:650px;">
		<li>
			<span>
				<div style="float:left;">
				<strong><a href="#" target="_self"><%=Site_Name%></a></strong></div>
				<div style="float:right; margin-right:110px;"><a href="#" id="AddFirColumn">添加子栏目</a>
					<script type="text/javascript"> 
						$("#AddFirColumn").wBox({requestType:"iframe_refresh",iframeWH:{width:500,height:350},target:"ColumnAdd.asp?Parent_id=0&Site_ID=<%=Info_ID%>"});
					</script>	
				</div>
			</span>
		<%response.Write(BuildXMLStr(0,str))%>
		</li>
	</ul>
</td>
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
